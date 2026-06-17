from dataclasses import dataclass, field
from pathlib import Path
import re
import unicodedata

import pandas as pd
from openpyxl.cell import WriteOnlyCell
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from src.ordenes_compra.exchange_rates import load_usd_rates


FX_RATE_DATE = "2026-06-08"
MIN_DELIVERY_DATE = pd.Timestamp("2025-01-01")
# Maestro de contratos marco (ME3N). Solo se usa para ENRIQUECER el consolidado
# (rellenar Fecha de Termino por contrato marco); no afecta entregables ni control points.
MARCO_MASTER_SUBDIR = "contratos-marco-me3n"
MARCO_MASTER_KEY_HEADERS = ["documento compras", "purchasing document", "contrato marco", "outline agreement"]
MARCO_MASTER_DATE_HEADERS = ["fin periodo validez", "validity period end"]
NET_ORDER_VALUE_HEADER = "Valor neto de pedido"
CURRENCY_HEADER = "Moneda"
USD_RATE_HEADER = f"Tasa cambio USD ({FX_RATE_DATE})"
NET_ORDER_VALUE_USD_HEADER = "Valor neto de pedido USD"
VIGENCIA_DATE_HEADER = "Fecha vigencia"
DERIVED_HEADERS = {USD_RATE_HEADER, NET_ORDER_VALUE_USD_HEADER, VIGENCIA_DATE_HEADER}
DEFAULT_FX_RATES_PATH = Path(__file__).resolve().parents[2] / "resources" / f"fx_rates_{FX_RATE_DATE}.json"


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
    NET_ORDER_VALUE_HEADER,
    CURRENCY_HEADER,
    USD_RATE_HEADER,
    NET_ORDER_VALUE_USD_HEADER,
    "Fecha documento",
    "Fecha de Entrega",
    "Fecha de Termino",
    VIGENCIA_DATE_HEADER,
    "Indicador de borrado",
]

DATE_HEADERS = ["Fecha documento", "Fecha de Entrega", "Fecha de Termino", VIGENCIA_DATE_HEADER]
ID_HEADERS = ["PR/SOLPED", "Contrato Marco", "Documento compras", "Posición", "Material"]
NUMERIC_HEADERS = [
    "Por entregar (cantidad)",
    "Por entregar (valor)",
    NET_ORDER_VALUE_HEADER,
    USD_RATE_HEADER,
    NET_ORDER_VALUE_USD_HEADER,
]
NUMERIC_FORMATS = {
    USD_RATE_HEADER: "0.000000",
}
REQUIRED_HEADERS = ["Documento compras", "Posición"]
SOURCE_HEADERS = [header for header in OUTPUT_HEADERS if header not in DERIVED_HEADERS]
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
    "valor neto de pedido": NET_ORDER_VALUE_HEADER,
    "net order value": NET_ORDER_VALUE_HEADER,
    "moneda": CURRENCY_HEADER,
    "currency": CURRENCY_HEADER,
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


def _period_end_from_filename(path: Path) -> pd.Timestamp | None:
    matches = re.findall(r"(\d{2})\.(\d{2})\.(\d{4})", path.stem)
    if not matches:
        return None
    dates = [pd.Timestamp(year=int(year), month=int(month), day=int(day)) for day, month, year in matches]
    return max(dates)


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


def _clean_currency_value(value: object) -> object:
    text = _clean_text_value(value)
    if text is None:
        return None
    return str(text).upper()


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
        elif header == CURRENCY_HEADER:
            row[header] = _clean_currency_value(value)
        else:
            row[header] = _clean_text_value(value)

    if row["Documento compras"] is None:
        row["Documento compras"] = last_purchase_document
    else:
        last_purchase_document = row["Documento compras"]
    return row, last_purchase_document


def _qualifies_for_output(row: dict[str, object]) -> bool:
    # Sin filtro de vigencia ni de fecha: el consolidado lleva TODOS los documentos.
    # Solo se exige la clave estructural (documento de compra + posicion).
    return not any(row[header] is None for header in REQUIRED_HEADERS)


def _missing_headers(selected_columns: dict[str, tuple[int, object]]) -> list[str]:
    return [header for header in SOURCE_HEADERS if header not in selected_columns]


def _has_qualifying_delivery_date(ws, delivery_col_idx: int) -> bool:
    for (value,) in ws.iter_rows(
        min_row=2,
        min_col=delivery_col_idx,
        max_col=delivery_col_idx,
        values_only=True,
    ):
        delivery_date = _clean_date_value(value)
        if pd.notna(delivery_date) and pd.Timestamp(delivery_date) >= MIN_DELIVERY_DATE:
            return True
    return False


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
        max_selected_col = max(idx for idx, _ in selected_columns.values()) + 1
        for row_values in ws.iter_rows(min_row=2, max_col=max_selected_col, values_only=True):
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
            print(f"  Leyendo [{op}] {path.name}")
            kept_rows, stat = _stream_source_file(path, op)
            rows.extend(kept_rows)
            stats.append(stat)

    if not rows:
        return pd.DataFrame(columns=OUTPUT_HEADERS), stats

    combined = pd.DataFrame(rows, columns=OUTPUT_HEADERS)
    return combined[OUTPUT_HEADERS], stats


def _enrich_with_usd_values(df: pd.DataFrame) -> pd.DataFrame:
    result = df.copy()
    if CURRENCY_HEADER not in result.columns:
        result[CURRENCY_HEADER] = None
    if NET_ORDER_VALUE_HEADER not in result.columns:
        result[NET_ORDER_VALUE_HEADER] = None

    currency = result[CURRENCY_HEADER].apply(_clean_currency_value)
    result[CURRENCY_HEADER] = currency
    currencies = sorted({value for value in currency.dropna().unique() if value})
    rates = load_usd_rates(currencies, DEFAULT_FX_RATES_PATH, FX_RATE_DATE)

    result[USD_RATE_HEADER] = currency.map(rates)
    net_value = pd.to_numeric(result[NET_ORDER_VALUE_HEADER], errors="coerce")
    usd_rate = pd.to_numeric(result[USD_RATE_HEADER], errors="coerce")
    result[NET_ORDER_VALUE_USD_HEADER] = net_value * usd_rate
    return result[OUTPUT_HEADERS]


def _write_headers(ws) -> list[int]:
    header_fill = PatternFill(start_color="1E3A5F", end_color="1E3A5F", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF")
    widths: list[int] = []
    cells = []
    for col_idx, header in enumerate(OUTPUT_HEADERS, 1):
        cell = WriteOnlyCell(ws, value=header)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cells.append(cell)
        widths.append(min(max(len(header) + 2, 10), 42))
    ws.append(cells)
    return widths


def _write_excel(df: pd.DataFrame, output_path: Path) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb = Workbook(write_only=True)
    ws = wb.create_sheet("Consolidado")
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(OUTPUT_HEADERS))}{max(len(df) + 1, 1)}"
    ws.row_dimensions[1].height = 32

    widths = _write_headers(ws)
    for row in df.itertuples(index=False, name=None):
        cells = []
        for col_idx, value in enumerate(row, 1):
            cell = WriteOnlyCell(ws)
            if pd.isna(value):
                cell.value = None
            elif OUTPUT_HEADERS[col_idx - 1] in DATE_HEADERS:
                cell.value = value.to_pydatetime() if hasattr(value, "to_pydatetime") else value
                cell.number_format = "DD-MM-YYYY"
            elif OUTPUT_HEADERS[col_idx - 1] in NUMERIC_HEADERS:
                cell.value = value
                cell.number_format = NUMERIC_FORMATS.get(OUTPUT_HEADERS[col_idx - 1], "#,##0.00")
            else:
                cell.value = value
            cell.alignment = Alignment(vertical="center")
            if cell.value is not None:
                widths[col_idx - 1] = max(widths[col_idx - 1], min(len(str(cell.value)) + 2, 42))
            cells.append(cell)
        ws.append(cells)

    for col_idx, width in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(col_idx)].width = width

    wb.save(output_path)


def _load_marco_end_dates(inputs_root: Path, operation: str) -> dict[str, pd.Timestamp]:
    """Mapa contrato marco -> fecha de fin de validez, desde el maestro ME3N.

    Lee `inputs/<OP>/contratos-marco-me3n/*.xlsx`. La clave es el número de
    contrato marco (en ME3N viene en 'Documento compras', 46*) y el valor es la
    máxima 'Fin período validez' encontrada para ese marco. Devuelve {} si la
    carpeta no existe (p. ej. MLCC), dejando el consolidado intacto.
    """
    directory = inputs_root / operation.upper() / MARCO_MASTER_SUBDIR
    if not directory.exists():
        return {}
    files = [
        path
        for path in list(directory.glob("*.XLSX")) + list(directory.glob("*.xlsx"))
        if not path.name.startswith("~$")
    ]
    files = sorted({path.resolve(): path for path in files}.values(), key=lambda p: p.name.lower())

    end_dates: dict[str, pd.Timestamp] = {}
    for path in files:
        wb = load_workbook(path, read_only=True, data_only=True)
        try:
            ws = wb[wb.sheetnames[0]]
            rows = ws.iter_rows(values_only=True)
            header_row = next(rows, None)
            if header_row is None:
                continue
            headers = {_normalize_header(v): idx for idx, v in enumerate(header_row) if v is not None}
            key_idx = next((headers[h] for h in MARCO_MASTER_KEY_HEADERS if h in headers), None)
            date_idx = next((headers[h] for h in MARCO_MASTER_DATE_HEADERS if h in headers), None)
            if key_idx is None or date_idx is None:
                continue
            last_key = None
            for row in rows:
                key = _clean_identifier_value(row[key_idx]) if key_idx < len(row) else None
                # ME3N repite el marco solo en la primera fila del bloque; se arrastra.
                if key is None:
                    key = last_key
                else:
                    last_key = key
                if key is None:
                    continue
                end = _clean_date_value(row[date_idx]) if date_idx < len(row) else pd.NaT
                if pd.isna(end):
                    continue
                previous = end_dates.get(key)
                if previous is None or end > previous:
                    end_dates[key] = end
        finally:
            wb.close()
    return end_dates


def _enrich_marco_validity(df: pd.DataFrame, marco_end_dates: dict[str, pd.Timestamp]) -> tuple[pd.DataFrame, int]:
    """Rellena 'Fecha de Termino' (donde falte) con el fin de validez del marco.

    Solo toca filas con 'Contrato Marco' presente y 'Fecha de Termino' vacía; no
    sobrescribe fechas ya pobladas. Devuelve (df, filas_rellenadas).
    """
    if not marco_end_dates or "Contrato Marco" not in df.columns:
        return df, 0
    result = df.copy()
    termino = pd.to_datetime(result.get("Fecha de Termino"), errors="coerce")
    marco = result["Contrato Marco"].map(_clean_identifier_value)
    mapped = marco.map(marco_end_dates)
    fill_mask = termino.isna() & mapped.notna()
    result.loc[fill_mask, "Fecha de Termino"] = mapped[fill_mask]
    return result, int(fill_mask.sum())


def _derive_vigencia_date(df: pd.DataFrame, operations: list[str]) -> pd.DataFrame:
    """Fecha de referencia del consolidado.

    MLCC: fecha de entrega con fallback a fin de periodo de validez.
    CCMC (y otros): el export no trae entrega ni periodo de validez, por lo que
    se usa la fecha de documento (unica fecha poblada).
    """
    result = df.copy()
    op = operations[0].upper() if len(operations) == 1 else ""
    entrega = pd.to_datetime(result.get("Fecha de Entrega"), errors="coerce")
    termino = pd.to_datetime(result.get("Fecha de Termino"), errors="coerce")
    documento = pd.to_datetime(result.get("Fecha documento"), errors="coerce")
    if op == "MLCC":
        vigencia = entrega.fillna(termino).fillna(documento)
    else:
        vigencia = termino.fillna(documento)
    result[VIGENCIA_DATE_HEADER] = vigencia
    return result


def build_po_consolidation(inputs_root: Path, output_path: Path, operations: list[str]) -> ConsolidationResult:
    df, stats = consolidate_purchase_orders(inputs_root, operations)
    df = _enrich_with_usd_values(df)
    if len(operations) == 1:
        marco_end_dates = _load_marco_end_dates(inputs_root, operations[0])
        if marco_end_dates:
            df, filled = _enrich_marco_validity(df, marco_end_dates)
            print(
                f"  Cruce contratos-marco (ME3N): {len(marco_end_dates):,} marcos con fin de validez; "
                f"Fecha de Termino rellenada en {filled:,} filas."
            )
    df = _derive_vigencia_date(df, operations)
    print(f"  Escribiendo workbook: {output_path} ({len(df):,} filas)")
    _write_excel(df, output_path)
    return ConsolidationResult(output_path=output_path, dataframe=df, file_stats=stats)