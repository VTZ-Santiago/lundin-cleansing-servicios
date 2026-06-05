from dataclasses import dataclass, field
from pathlib import Path
import re
import unicodedata

import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter


OUTPUT_HEADERS: list[str] = [
    "Planta",
    "PR/SOLPED",
    "Contrato Marco",
    "Documento compras",
    "Posición",
    "Material",
    "Proveedor",
    "Texto Breve",
    "Tipo de Posición",
    "Cl Docto compras",
    "Grupo de liberación",
    "Por entregar (cantidad)",
    "Por entregar (valor)",
    "Fecha documento",
    "Fecha de Entrega",
    "Fecha de Termino",
    "Indicador de borrado",
]

DATE_HEADERS = ["Fecha documento", "Fecha de Entrega", "Fecha de Termino"]
ID_HEADERS = ["PR/SOLPED", "Contrato Marco", "Documento compras", "Posición", "Material"]
NUMERIC_HEADERS = ["Por entregar (cantidad)", "Por entregar (valor)"]
REQUIRED_HEADERS = ["Documento compras", "Posición"]
RAW_TO_OUTPUT: dict[str, str] = {
    "centro": "Planta",
    "plant": "Planta",
    "solicitud de pedido": "PR/SOLPED",
    "purchase requisition": "PR/SOLPED",
    "purchase requisition.1": "PR/SOLPED",
    "contrato marco": "Contrato Marco",
    "outline agreement": "Contrato Marco",
    "documento compras": "Documento compras",
    "purchasing document": "Documento compras",
    "posicion": "Posición",
    "item": "Posición",
    "material": "Material",
    "proveedor/centro suministrador": "Proveedor",
    "vendor/supplying plant": "Proveedor",
    "texto breve": "Texto Breve",
    "short text": "Texto Breve",
    "tipo de posicion": "Tipo de Posición",
    "item category": "Tipo de Posición",
    "item category.1": "Tipo de Posición",
    "cl.documento compras": "Cl Docto compras",
    "purchasing doc. type": "Cl Docto compras",
    "purch. doc. category": "Cl Docto compras",
    "grupo de liberación": "Grupo de liberación",
    "grupo de liberacion": "Grupo de liberación",
    "release group": "Grupo de liberación",
    "por entregar (cantidad)": "Por entregar (cantidad)",
    "still to be delivered (qty)": "Por entregar (cantidad)",
    "por entregar (valor)": "Por entregar (valor)",
    "still to be delivered (value)": "Por entregar (valor)",
    "fecha documento": "Fecha documento",
    "document date": "Fecha documento",
    "fecha de entrega": "Fecha de Entrega",
    "fecha entrega": "Fecha de Entrega",
    "delivery date": "Fecha de Entrega",
    "fin periodo validez": "Fecha de Termino",
    "validity period end": "Fecha de Termino",
    "indicador de borrado": "Indicador de borrado",
    "deletion flag": "Indicador de borrado",
    "deletion indicator": "Indicador de borrado",
}

RAW_HEADER_PRIORITY: dict[str, int] = {
    "purchasing doc. type": 20,
    "purch. doc. category": 10,
    "purchase requisition.1": 20,
    "purchase requisition": 10,
    "tipo de posicion": 20,
    "item category.1": 20,
    "item category": 10,
}

RAW_HEADER_OCCURRENCE_PRIORITY: dict[tuple[str, int], int] = {
    ("purchase requisition", 2): 20,
    ("purchase requisition", 1): 10,
    ("item category", 2): 20,
    ("item category", 1): 10,
}


@dataclass
class FileConsolidationStats:
    operation: str
    source_file: str
    rows_read: int
    rows_kept: int
    missing_headers: list[str] = field(default_factory=list)
    skipped: bool = False
    skip_reason: str = ""


@dataclass
class ConsolidationResult:
    output_path: Path
    dataframe: pd.DataFrame
    file_stats: list[FileConsolidationStats]

    @property
    def rows_read(self) -> int:
        return sum(stat.rows_read for stat in self.file_stats)

    @property
    def rows_kept(self) -> int:
        return sum(stat.rows_kept for stat in self.file_stats)

    @property
    def rows_discarded(self) -> int:
        return self.rows_read - self.rows_kept


def _normalize_header(header: object) -> str:
    normalized = unicodedata.normalize("NFKD", str(header).strip())
    ascii_header = "".join(ch for ch in normalized if not unicodedata.combining(ch))
    ascii_header = re.sub(r"\s+", " ", ascii_header)
    return ascii_header.lower()


def _map_header(header: object) -> str | None:
    return RAW_TO_OUTPUT.get(_normalize_header(header))


def _header_priority(header: object, occurrence: int = 1) -> int:
    normalized = _normalize_header(header)
    return RAW_HEADER_OCCURRENCE_PRIORITY.get(
        (normalized, occurrence),
        RAW_HEADER_PRIORITY.get(normalized, 10),
    )


def _excel_files(inputs_root: Path, operation: str) -> list[Path]:
    directory = inputs_root / operation / "ordenes-compra"
    if not directory.exists():
        raise FileNotFoundError(f"Input directory not found: {directory}")

    files = [
        path
        for path in list(directory.glob("*.XLSX")) + list(directory.glob("*.xlsx"))
        if not path.name.startswith("~$")
    ]
    files = sorted({path.resolve(): path for path in files}.values(), key=lambda path: path.name.lower())
    if not files:
        raise FileNotFoundError(f"No Excel files found in {directory}")
    return files


def _selected_columns(raw_headers: tuple[object, ...]) -> dict[str, tuple[int, object]]:
    selected: dict[str, tuple[object, int]] = {}
    occurrences: dict[str, int] = {}
    for idx, raw_col in enumerate(raw_headers):
        normalized = _normalize_header(raw_col)
        occurrences[normalized] = occurrences.get(normalized, 0) + 1
        occurrence = occurrences[normalized]
        output_col = RAW_TO_OUTPUT.get(normalized)
        if output_col is None:
            continue
        priority = _header_priority(raw_col, occurrence)
        current = selected.get(output_col)
        if current is None or priority > current[1]:
            selected[output_col] = (idx, priority)
    return {output_col: (idx, raw_headers[idx]) for output_col, (idx, _) in selected.items()}


def _row_has_selected_values(row_values: tuple[object, ...], selected_columns: dict[str, tuple[int, object]]) -> bool:
    for idx, _ in selected_columns.values():
        if idx < len(row_values) and not pd.isna(row_values[idx]):
            return True
    return False


def _clean_identifier_value(value: object) -> object:
    if pd.isna(value):
        return None
    if isinstance(value, str):
        text = value.strip()
    else:
        text = str(value).strip()
    if not text:
        return None
    numeric = pd.to_numeric(text, errors="coerce")
    if pd.notna(numeric) and float(numeric).is_integer():
        return str(int(numeric))
    if text.endswith(".0"):
        without_decimal = text[:-2]
        if without_decimal.isdigit():
            return without_decimal
    return text


def _clean_text_value(value: object) -> object:
    if pd.isna(value):
        return None
    text = str(value).strip()
    return text or None


def _clean_date_value(value: object) -> object:
    parsed = pd.to_datetime(value, errors="coerce")
    if pd.isna(parsed):
        return pd.NaT
    return parsed


def _clean_numeric_value(value: object) -> object:
    parsed = pd.to_numeric(value, errors="coerce")
    if pd.isna(parsed):
        return None
    return float(parsed)


def _clean_output_row(raw_values: dict[str, object], last_purchase_document: object) -> tuple[dict[str, object], object]:
    row: dict[str, object] = {}
    for header in OUTPUT_HEADERS:
        value = raw_values.get(header)
        if header in DATE_HEADERS:
            row[header] = _clean_date_value(value)
        elif header in NUMERIC_HEADERS:
            row[header] = _clean_numeric_value(value)
        elif header in ID_HEADERS:
            row[header] = _clean_identifier_value(value)
        else:
            row[header] = _clean_text_value(value)

    if row["Documento compras"] is None:
        row["Documento compras"] = last_purchase_document
    else:
        last_purchase_document = row["Documento compras"]
    return row, last_purchase_document


def _qualifies_for_output(row: dict[str, object]) -> bool:
    return not any(row[header] is None for header in REQUIRED_HEADERS)


def _missing_headers(selected_columns: dict[str, tuple[int, object]]) -> list[str]:
    return [header for header in OUTPUT_HEADERS if header not in selected_columns]


def _stream_source_file(path: Path, operation: str) -> tuple[list[dict[str, object]], FileConsolidationStats]:
    wb = load_workbook(path, read_only=True, data_only=True)
    try:
        ws = wb[wb.sheetnames[0]]
        header_row = next(ws.iter_rows(min_row=1, max_row=1, values_only=True), None)
        if header_row is None:
            stat = FileConsolidationStats(
                operation=operation,
                source_file=path.name,
                rows_read=0,
                rows_kept=0,
                skipped=True,
                skip_reason="Empty workbook",
            )
            return [], stat

        selected_columns = _selected_columns(header_row)
        missing = _missing_headers(selected_columns)
        missing_required = [header for header in REQUIRED_HEADERS if header in missing]
        if missing_required:
            stat = FileConsolidationStats(
                operation=operation,
                source_file=path.name,
                rows_read=0,
                rows_kept=0,
                missing_headers=missing,
                skipped=True,
                skip_reason=f"Missing required headers: {', '.join(missing_required)}",
            )
            return [], stat

        kept_rows: list[dict[str, object]] = []
        rows_read = 0
        last_purchase_document = None
        for row_values in ws.iter_rows(min_row=2, values_only=True):
            if not _row_has_selected_values(row_values, selected_columns):
                continue
            rows_read += 1
            raw_values = {
                output_col: row_values[idx] if idx < len(row_values) else None
                for output_col, (idx, _) in selected_columns.items()
            }
            row, last_purchase_document = _clean_output_row(raw_values, last_purchase_document)
            if _qualifies_for_output(row):
                kept_rows.append(row)

        stat = FileConsolidationStats(
            operation=operation,
            source_file=path.name,
            rows_read=rows_read,
            rows_kept=len(kept_rows),
            missing_headers=missing,
        )
        return kept_rows, stat
    finally:
        wb.close()


def consolidate_purchase_orders(inputs_root: Path, operations: list[str]) -> tuple[pd.DataFrame, list[FileConsolidationStats]]:
    rows: list[dict[str, object]] = []
    stats: list[FileConsolidationStats] = []

    for operation in operations:
        op = operation.upper()
        for path in _excel_files(inputs_root, op):
            kept_rows, stat = _stream_source_file(path, op)
            rows.extend(kept_rows)
            stats.append(stat)

    if not rows:
        return pd.DataFrame(columns=OUTPUT_HEADERS), stats

    combined = pd.DataFrame(rows, columns=OUTPUT_HEADERS)
    return combined[OUTPUT_HEADERS], stats


def _write_headers(ws) -> None:
    header_fill = PatternFill(start_color="1E3A5F", end_color="1E3A5F", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF")
    for col_idx, header in enumerate(OUTPUT_HEADERS, 1):
        cell = ws.cell(row=1, column=col_idx, value=header)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def _write_excel(df: pd.DataFrame, output_path: Path) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb = Workbook()
    ws = wb.active
    ws.title = "PO_2025plus"

    _write_headers(ws)
    for row_idx, row in enumerate(df.itertuples(index=False, name=None), 2):
        for col_idx, value in enumerate(row, 1):
            cell = ws.cell(row=row_idx, column=col_idx)
            if pd.isna(value):
                cell.value = None
            elif OUTPUT_HEADERS[col_idx - 1] in DATE_HEADERS:
                cell.value = value.to_pydatetime() if hasattr(value, "to_pydatetime") else value
                cell.number_format = "DD-MM-YYYY"
            elif OUTPUT_HEADERS[col_idx - 1] in NUMERIC_HEADERS:
                cell.value = value
                cell.number_format = "#,##0.00"
            else:
                cell.value = value
            cell.alignment = Alignment(vertical="center")

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(OUTPUT_HEADERS))}{max(len(df) + 1, 1)}"
    ws.row_dimensions[1].height = 32
    for column in ws.columns:
        width = 10
        for cell in column:
            if cell.value is not None:
                width = max(width, min(len(str(cell.value)) + 2, 42))
        ws.column_dimensions[get_column_letter(column[0].column)].width = width

    wb.save(output_path)


def build_po_consolidation(inputs_root: Path, output_path: Path, operations: list[str]) -> ConsolidationResult:
    df, stats = consolidate_purchase_orders(inputs_root, operations)
    _write_excel(df, output_path)
    return ConsolidationResult(output_path=output_path, dataframe=df, file_stats=stats)