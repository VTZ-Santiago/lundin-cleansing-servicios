from __future__ import annotations

import json
import re
from dataclasses import dataclass, field
from datetime import date
from pathlib import Path

from openpyxl import load_workbook
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_AUTO_SHAPE_TYPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.util import Inches, Pt

from src.segmentacion.characterization import load_segment_rules


ORANGE = RGBColor.from_string("F97316")
ORANGE_LIGHT = RGBColor.from_string("FFF7ED")
CHARCOAL = RGBColor.from_string("222222")
SLATE = RGBColor.from_string("475569")
GRAY = RGBColor.from_string("E5E7EB")
GRAY_LIGHT = RGBColor.from_string("F8FAFC")
BLUE = RGBColor.from_string("1D4ED8")
BLUE_LIGHT = RGBColor.from_string("DBEAFE")
GREEN = RGBColor.from_string("0F766E")
GREEN_LIGHT = RGBColor.from_string("CCFBF1")
RED = RGBColor.from_string("BE123C")
RED_LIGHT = RGBColor.from_string("FFE4E6")
PURPLE = RGBColor.from_string("4C1D95")
PURPLE_LIGHT = RGBColor.from_string("EDE9FE")
TEAL = RGBColor.from_string("0E7490")
TEAL_LIGHT = RGBColor.from_string("CFFAFE")

# Secondary accent used consistently for "tipo de posición" en ambas operaciones
# (la operación marca el color primario: MLCC azul, CCMC verde).
POS_ACCENT = ORANGE
POS_ACCENT_LIGHT = ORANGE_LIGHT

# Nombre legible de cada tipo de posición SAP (letra canónica del loader).
POS_TYPE_LABELS: dict[str, str] = {
    "D": "Servicio (D)",
    "L": "Subcontratación (L)",
    "C": "Consignación (C)",
    "K": "Consignación (K)",
    "V": "Traslado (V)",
    "U": "Traslado (U)",
    "P": "Tope / límite (P)",
    "Sin tipo": "Estándar / Stock",
}


def _pos_type_label(code: str) -> str:
    code = (code or "").strip()
    if not code:
        return POS_TYPE_LABELS["Sin tipo"]
    return POS_TYPE_LABELS.get(code, code)


def _fmt_usd(value: float) -> str:
    """USD compacto para tarjetas/tablas: USD 1400,0MM / USD 98,5MM / USD 720K."""
    if value is None or value <= 0.0:
        return "—"
    if value >= 1_000_000:
        return f"USD {value / 1_000_000:.1f}MM".replace(".", ",")
    if value >= 1_000:
        return f"USD {value / 1_000:.0f}K"
    return f"USD {value:.0f}"


@dataclass
class SegmentStats:
    label: str
    c1_rows: int = 0
    c1_documents: int = 0
    c2_documents: int = 0
    c2_no_migra_documents: int = 0
    c1_outline_contracts: int = 0
    c2_outline_contracts: int = 0
    c2_no_migra_outline_contracts: int = 0
    c2_rows: int = 0
    c2_no_migra_rows: int = 0
    reconciliation_ok: bool = False
    excluded_pct: float = 0.0
    doc_class_counts: dict[str, int] = field(default_factory=dict)


@dataclass
class OperationSummary:
    operation: str
    rows_read: int = 0
    total_rows: int = 0
    classified_rows: int = 0
    unclassified_rows: int = 0
    overlap_rows: int = 0
    criteria_count: int = 0
    source_files: list[str] = field(default_factory=list)
    segments: dict[str, SegmentStats] = field(default_factory=dict)


@dataclass
class POStats:
    operation: str
    periods: list[str] = field(default_factory=list)
    total_positions: int = 0
    unique_documents: int = 0
    framework_contracts: int = 0
    po_with_framework: int = 0
    po_without_framework: int = 0
    pct_po_with_framework: float = 0.0
    doc_classes: dict[str, int] = field(default_factory=dict)
    position_types: dict[str, int] = field(default_factory=dict)
    expired_count: int = 0
    active_count: int = 0
    no_date_count: int = 0
    deletion_flags: dict[str, int] = field(default_factory=dict)
    top_frameworks: list[tuple[str, int, int]] = field(default_factory=list)

    def top_doc_classes(self, limit: int = 5) -> list[tuple[str, int]]:
        return sorted(self.doc_classes.items(), key=lambda item: (-item[1], item[0]))[:limit]

    def top_position_types(self, limit: int = 4) -> list[tuple[str, int]]:
        return sorted(self.position_types.items(), key=lambda item: (-item[1], item[0]))[:limit]


@dataclass
class TempOCSnapshot:
    operation: str
    total_base: int = 0
    plant_counts: dict[str, int] = field(default_factory=dict)
    type_counts: dict[str, int] = field(default_factory=dict)
    with_material: int = 0
    without_material: int = 0
    with_material_expired: int = 0
    with_material_active: int = 0
    with_material_ls: int = 0
    without_material_expired: int = 0
    without_material_active: int = 0
    without_material_ls: int = 0


@dataclass
class SuministrosData:
    operation: str
    vigente_rows: int = 0
    vigente_usd: float = 0.0
    vcs_rows: int = 0
    vcs_usd: float = 0.0
    vss_rows: int = 0
    prefix_counts: dict[str, int] = field(default_factory=dict)
    con_material: int = 0
    sin_material: int = 0


@dataclass
class SupplySegmentRow:
    segment: str
    vigente_rows: int = 0
    vcs_rows: int = 0
    vss_rows: int = 0
    vigente_usd: float = 0.0
    vcs_usd: float = 0.0

    @property
    def total_rows(self) -> int:
        return self.vigente_rows + self.vcs_rows + self.vss_rows


@dataclass
class SegmentValor:
    label: str = ""
    c2_rows: int = 0
    c2_usd: float = 0.0
    vencidos_rows: int = 0
    vencidos_usd: float = 0.0


@dataclass
class MigracionValores:
    """Valores en USD del universo D (Contratos / Órdenes de Servicio).

    Se leen del sidecar outputs/control_points/valores_migracion_<OP>.json escrito
    por run_segmentacion_mlcc.py.
    """
    operation: str = ""
    contratos: SegmentValor = field(default_factory=SegmentValor)
    ordenes_servicio: SegmentValor = field(default_factory=SegmentValor)
    total_c2_rows: int = 0
    total_c2_usd: float = 0.0
    total_vencidos_rows: int = 0
    total_vencidos_usd: float = 0.0
    ost_rows: int = 0
    ost_usd: float = 0.0
    usd_available: bool = False


@dataclass
class PlantRow:
    planta: str
    vigente_filas: int = 0
    vigente_usd: float = 0.0
    vcs_filas: int = 0
    vcs_usd: float = 0.0
    vss_filas: int = 0


@dataclass
class SegmentCriteria:
    position_type: str = ""
    marco: str = ""
    doc_classes: list[str] = field(default_factory=list)


def _as_text(value: object) -> str:
    if value is None:
        return ""
    return " ".join(str(value).replace("\xa0", " ").split())


def _cell_primary_line(value: object) -> str:
    if value is None:
        return ""
    lines = [line.strip() for line in str(value).replace("\xa0", " ").splitlines() if line.strip()]
    return lines[0] if lines else ""


def _as_int(value: object) -> int:
    if value in (None, ""):
        return 0
    if isinstance(value, bool):
        return int(value)
    if isinstance(value, (int, float)):
        return int(round(float(value)))
    text = str(value).strip()
    if not text:
        return 0
    digits = re.sub(r"[^0-9-]", "", text)
    return int(digits) if digits not in {"", "-"} else 0


def _as_float(value: object) -> float:
    if value in (None, ""):
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip().replace("%", "")
    if not text:
        return 0.0
    text = text.replace(",", ".")
    try:
        return float(text)
    except ValueError:
        return 0.0


def _format_int(value: int) -> str:
    return f"{value:,}".replace(",", ".")


def _format_pct(value: float) -> str:
    return f"{value:.1f}%".replace(".", ",")


def _month_name_es(value: date) -> str:
    months = {
        1: "ENERO",
        2: "FEBRERO",
        3: "MARZO",
        4: "ABRIL",
        5: "MAYO",
        6: "JUNIO",
        7: "JULIO",
        8: "AGOSTO",
        9: "SEPTIEMBRE",
        10: "OCTUBRE",
        11: "NOVIEMBRE",
        12: "DICIEMBRE",
    }
    return months[value.month]


def _compact_period(period: str) -> str:
    years = re.findall(r"\d{4}", period)
    if len(years) >= 2:
        return f"{years[0]}-{years[-1]}"
    return period.strip()


def _compact_periods(periods: list[str]) -> str:
    labels = [_compact_period(period) for period in periods if period.strip()]
    return " | ".join(labels)


def parse_summary_markdown(path: Path) -> OperationSummary:
    lines = path.read_text(encoding="utf-8").splitlines()
    summary = OperationSummary(operation=lines[0].replace("# Resumen segmentación", "").strip())
    current_section: str | None = None
    current_segment: SegmentStats | None = None
    nested_target: str | None = None

    for raw_line in lines[1:]:
        line = raw_line.rstrip()
        stripped = line.strip()
        if not stripped:
            continue
        if line.startswith("  - "):
            nested_content = line[4:].strip()
            if current_segment is None:
                if nested_target == "global_source_files":
                    summary.source_files.append(nested_content)
                continue
            if nested_target == "doc_class_counts" and ":" in nested_content:
                key, value = [part.strip() for part in nested_content.split(":", 1)]
                current_segment.doc_class_counts[key] = _as_int(value)
            elif nested_target == "source_file_counts":
                continue
            continue
        if stripped.startswith("## "):
            current_section = stripped[3:].strip()
            current_segment = None if current_section == "Auditoría global" else SegmentStats(label=current_section)
            if current_segment is not None:
                summary.segments[current_section] = current_segment
            nested_target = None
            continue
        if not stripped.startswith("- "):
            continue

        content = stripped[2:]
        if current_section == "Auditoría global":
            if content == "Archivos fuente:":
                nested_target = "global_source_files"
                continue
            if ":" not in content:
                continue
            key, value = [part.strip() for part in content.split(":", 1)]
            if key == "Filas leídas desde OC":
                summary.rows_read = _as_int(value)
            elif key == "Filas limpias cargadas":
                summary.total_rows = _as_int(value)
            elif key == "Filas clasificadas en segmentos":
                summary.classified_rows = _as_int(value)
            elif key == "Filas no clasificadas":
                summary.unclassified_rows = _as_int(value)
            elif key == "Filas con solapamiento entre segmentos":
                summary.overlap_rows = _as_int(value)
            elif key == "Criterios de segmentación cargados":
                summary.criteria_count = _as_int(value)
            continue

        if current_segment is None:
            continue
        if content == "Cl. documento compras:":
            nested_target = "doc_class_counts"
            continue
        if content == "Archivos fuente:":
            nested_target = "source_file_counts"
            continue
        if ":" not in content:
            continue
        key, value = [part.strip() for part in content.split(":", 1)]
        if key == "C1 filas":
            current_segment.c1_rows = _as_int(value)
        elif key == "Documentos únicos C1":
            current_segment.c1_documents = _as_int(value)
        elif key == "Documentos únicos C2":
            current_segment.c2_documents = _as_int(value)
        elif key == "Documentos únicos C2_NO_MIGRA":
            current_segment.c2_no_migra_documents = _as_int(value)
        elif key == "Outline contracts únicos C1":
            current_segment.c1_outline_contracts = _as_int(value)
        elif key == "Outline contracts únicos C2":
            current_segment.c2_outline_contracts = _as_int(value)
        elif key == "Outline contracts únicos C2_NO_MIGRA":
            current_segment.c2_no_migra_outline_contracts = _as_int(value)
        elif key == "C2 migra":
            current_segment.c2_rows = _as_int(value)
        elif key == "C2 no migra":
            current_segment.c2_no_migra_rows = _as_int(value)
        elif key == "Porcentaje no migra":
            current_segment.excluded_pct = _as_float(value)
        elif key == "Reconciliación C1 = C2 + C2_NO_MIGRA":
            current_segment.reconciliation_ok = value == "OK"

    return summary


def _fill_down_headers(ws, row_number: int) -> dict[int, str]:
    current = ""
    headers: dict[int, str] = {}
    for column in range(1, ws.max_column + 1):
        value = _as_text(ws.cell(row_number, column).value)
        if value:
            current = value
        headers[column] = current
    return headers


def _load_po_stats(path: Path, operation: str) -> POStats:
    data = POStats(operation=operation)
    if not path.exists():
        return data

    workbook = load_workbook(path, read_only=True, data_only=True)
    worksheet = workbook[workbook.sheetnames[0]]
    section_headers = _fill_down_headers(worksheet, 2)
    total_row = next((row for row in range(4, worksheet.max_row + 1) if _as_text(worksheet.cell(row, 1).value) == "TOTAL"), None)
    if total_row is None:
        return data

    for row in range(4, total_row):
        label = _as_text(worksheet.cell(row, 1).value)
        if label and not label.startswith("CRUCE"):
            data.periods.append(label)

    for column in range(1, worksheet.max_column + 1):
        section = section_headers[column]
        primary = _cell_primary_line(worksheet.cell(3, column).value)
        value = worksheet.cell(total_row, column).value
        if not section or not primary or value in (None, ""):
            continue

        if section == "Volumen":
            if primary == "Total de posiciones":
                data.total_positions = _as_int(value)
            elif primary in {"PO únicas", "Contratos únicos"}:
                data.unique_documents = _as_int(value)
        elif section == "Cl. documento de compras" and primary.startswith("Clase "):
            data.doc_classes[primary.replace("Clase ", "", 1).strip()] = _as_int(value)
        elif section == "Contrato Marco":
            if primary == "Contratos Marco":
                data.framework_contracts = _as_int(value)
            elif primary == "PO con marco":
                data.po_with_framework = _as_int(value)
            elif primary == "PO sin marco":
                data.po_without_framework = _as_int(value)
            elif primary == "% PO con marco":
                data.pct_po_with_framework = _as_float(value)
        elif section.startswith("Tipo de Posición"):
            if primary.startswith("Tipo "):
                data.position_types[primary.replace("Tipo ", "", 1).strip()] = _as_int(value)
            elif primary.startswith("Sin tipo"):
                data.position_types["Sin tipo"] = _as_int(value)
        elif section.startswith("Vigencia del Contrato"):
            if primary.startswith("Vencidos"):
                data.expired_count = _as_int(value)
            elif primary.startswith("Vigentes en"):
                data.active_count += _as_int(value)
            elif primary.startswith("Sin fecha"):
                data.no_date_count = _as_int(value)
        elif section.startswith("Indicador de Borrado"):
            if primary.startswith("Libre"):
                data.deletion_flags["Libre"] = _as_int(value)
            elif primary.startswith("Marcado"):
                data.deletion_flags["Marcado"] = _as_int(value)
            elif primary.startswith("Sin indicador"):
                data.deletion_flags["Sin indicador"] = _as_int(value)

    if "Contrato Marco" in workbook.sheetnames:
        framework_sheet = workbook["Contrato Marco"]
        for row in range(2, min(framework_sheet.max_row, 6) + 1):
            contract = _as_text(framework_sheet.cell(row, 1).value)
            if not contract:
                continue
            data.top_frameworks.append((
                contract,
                _as_int(framework_sheet.cell(row, 2).value),
                _as_int(framework_sheet.cell(row, 3).value),
            ))

    return data


def _load_universo_oc(output_root: Path, operation: str) -> POStats | None:
    """Construye POStats desde outputs/control_points/universo_oc_<OP>.json.

    Ese sidecar cuenta tipos de posición y clases documentales sobre el TOTAL de
    posiciones (OC totales), con el loader vigente de la segmentación — a
    diferencia del reporte estadístico, que agrega por documento y está fechado.
    Devuelve None si el sidecar no existe (el deck cae al reporte estadístico).
    """
    path = output_root / "control_points" / f"universo_oc_{operation}.json"
    if not path.exists():
        return None
    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return None
    data = POStats(operation=operation)
    data.periods = [str(p) for p in payload.get("periods", []) if str(p).strip()]
    data.total_positions = _as_int(payload.get("total_positions"))
    data.unique_documents = _as_int(payload.get("unique_documents"))
    data.framework_contracts = _as_int(payload.get("framework_contracts"))
    data.po_with_framework = _as_int(payload.get("po_with_framework"))
    data.po_without_framework = _as_int(payload.get("po_without_framework"))
    data.pct_po_with_framework = _as_float(payload.get("pct_po_with_framework"))
    data.doc_classes = {str(k): _as_int(v) for k, v in (payload.get("by_doc_class") or {}).items()}
    data.position_types = {str(k): _as_int(v) for k, v in (payload.get("by_position_type") or {}).items()}
    return data


def _load_temp_snapshot(path: Path, operation: str) -> TempOCSnapshot:
    snapshot = TempOCSnapshot(operation=operation)
    if not path.exists():
        return snapshot

    workbook = load_workbook(path, read_only=True, data_only=True)
    if operation not in workbook.sheetnames:
        return snapshot
    worksheet = workbook[operation]

    if operation == "MLCC":
        snapshot.total_base = _as_int(worksheet.cell(2, 4).value)
        plant = _as_text(worksheet.cell(5, 1).value)
        if plant:
            snapshot.plant_counts[plant] = _as_int(worksheet.cell(5, 2).value)
        for column in range(3, 8):
            label = _as_text(worksheet.cell(4, column).value)
            if label:
                snapshot.type_counts[label] = _as_int(worksheet.cell(5, column).value)
        snapshot.with_material = _as_int(worksheet.cell(5, 12).value)
        snapshot.without_material = _as_int(worksheet.cell(5, 13).value)
        snapshot.with_material_expired = _as_int(worksheet.cell(5, 14).value)
        snapshot.with_material_active = _as_int(worksheet.cell(5, 15).value)
        snapshot.with_material_ls = _as_int(worksheet.cell(5, 16).value)
        snapshot.without_material_expired = _as_int(worksheet.cell(5, 17).value)
        snapshot.without_material_active = _as_int(worksheet.cell(5, 18).value)
        snapshot.without_material_ls = _as_int(worksheet.cell(5, 19).value)
        return snapshot

    snapshot.total_base = _as_int(worksheet.cell(10, 2).value)
    for row in range(6, 10):
        plant = _as_text(worksheet.cell(row, 1).value)
        quantity = _as_int(worksheet.cell(row, 2).value)
        if plant and quantity:
            snapshot.plant_counts[plant] = quantity
        snapshot.with_material += _as_int(worksheet.cell(row, 10).value)
        snapshot.without_material += _as_int(worksheet.cell(row, 11).value)
        snapshot.with_material_expired += _as_int(worksheet.cell(row, 12).value)
        snapshot.with_material_active += _as_int(worksheet.cell(row, 13).value)
        snapshot.without_material_expired += _as_int(worksheet.cell(row, 14).value)
        snapshot.without_material_active += _as_int(worksheet.cell(row, 15).value)

    current_type = ""
    for row in range(1, worksheet.max_row + 1):
        marker = _as_text(worksheet.cell(row, 1).value)
        if marker.startswith("TIPO POSICION TODAS"):
            current_type = "TODAS"
            continue
        if marker.startswith("TIPO POSICION"):
            current_type = _as_text(worksheet.cell(row, 2).value) or "TODAS"
            continue
        if current_type and not marker:
            total_value = _as_int(worksheet.cell(row, 2).value)
            if total_value:
                snapshot.type_counts[current_type] = total_value
                current_type = ""

    return snapshot


def _load_suministros_data(output_root: Path, operation: str) -> SuministrosData:
    data = SuministrosData(operation=operation)
    pattern = f"reportes/suministros_{operation.lower()}_*.xlsx"
    candidates = sorted(
        output_root.glob(pattern),
        key=lambda p: p.stat().st_mtime,
        reverse=True,
    )
    if not candidates:
        return data

    workbook = load_workbook(candidates[0], read_only=True, data_only=True)

    if "Resumen" in workbook.sheetnames:
        ws = workbook["Resumen"]
        headers = [_as_text(ws.cell(1, c).value) for c in range(1, ws.max_column + 1)]
        # Row 2 is always TOTAL (first entry in plant list)
        total_vals = [ws.cell(2, c).value for c in range(1, ws.max_column + 1)]
        for idx, hdr in enumerate(headers):
            val = total_vals[idx]
            if hdr == "VIGENTE — filas":
                data.vigente_rows = _as_int(val)
            elif hdr == "VIGENTE — valor USD":
                data.vigente_usd = _as_float(val)
            elif hdr == "NO_VIGENTE_CON_SALDO — filas":
                data.vcs_rows = _as_int(val)
            elif hdr == "NO_VIGENTE_CON_SALDO — valor USD":
                data.vcs_usd = _as_float(val)
            elif hdr == "NO_VIGENTE_SIN_SALDO — filas":
                data.vss_rows = _as_int(val)

    if "Apertura_Codigo" in workbook.sheetnames:
        ws2 = workbook["Apertura_Codigo"]
        headers2 = [_as_text(ws2.cell(1, c).value) for c in range(1, ws2.max_column + 1)]
        _skip = {"Categoría", "Total filas", "Valor pendiente", "Valor USD", "Con material", "Sin material", ""}
        for row in range(2, ws2.max_row + 1):
            if _as_text(ws2.cell(row, 1).value) == "TOTAL":
                row_vals = [ws2.cell(row, c).value for c in range(1, ws2.max_column + 1)]
                for idx, hdr in enumerate(headers2):
                    val = row_vals[idx]
                    if hdr == "Con material":
                        data.con_material = _as_int(val)
                    elif hdr == "Sin material":
                        data.sin_material = _as_int(val)
                    elif hdr not in _skip:
                        data.prefix_counts[hdr] = _as_int(val)
                break

    workbook.close()
    return data


def _load_suministros_segments(output_root: Path, operation: str) -> list[SupplySegmentRow]:
    """Lee la hoja Por_Subsegmento del reporte de Suministros más reciente.

    Devuelve una fila por sub-bloque (excluye TOTAL y SUBTOTAL), con el conteo
    por categoría de vigencia.
    """
    pattern = f"reportes/suministros_{operation.lower()}_*.xlsx"
    candidates = sorted(
        output_root.glob(pattern),
        key=lambda p: p.stat().st_mtime,
        reverse=True,
    )
    if not candidates:
        return []

    workbook = load_workbook(candidates[0], read_only=True, data_only=True)
    if "Por_Subsegmento" not in workbook.sheetnames:
        workbook.close()
        return []

    ws = workbook["Por_Subsegmento"]
    headers = [_as_text(ws.cell(1, c).value) for c in range(1, ws.max_column + 1)]
    idx = {h: i for i, h in enumerate(headers)}
    col_seg = idx.get("Sub-bloque")
    col_cat = idx.get("Categoría migración")
    col_rows = idx.get("Filas")
    col_usd = idx.get("Valor USD")
    if col_seg is None or col_cat is None or col_rows is None:
        workbook.close()
        return []

    by_seg: dict[str, SupplySegmentRow] = {}
    for row in range(2, ws.max_row + 1):
        seg = _as_text(ws.cell(row, col_seg + 1).value)
        cat = _as_text(ws.cell(row, col_cat + 1).value)
        if not seg or seg == "TOTAL" or cat.startswith("SUBTOTAL") or not cat:
            continue
        filas = _as_int(ws.cell(row, col_rows + 1).value)
        usd = _as_float(ws.cell(row, col_usd + 1).value) if col_usd is not None else 0.0
        entry = by_seg.setdefault(seg, SupplySegmentRow(segment=seg))
        if cat == "VIGENTE":
            entry.vigente_rows = filas
            entry.vigente_usd = usd
        elif cat == "NO_VIGENTE_CON_SALDO":
            entry.vcs_rows = filas
            entry.vcs_usd = usd
        elif cat == "NO_VIGENTE_SIN_SALDO":
            entry.vss_rows = filas

    workbook.close()
    return list(by_seg.values())


def _load_valores_migracion(output_root: Path, operation: str) -> MigracionValores:
    """Lee outputs/control_points/valores_migracion_<OP>.json (USD del universo D)."""
    data = MigracionValores(operation=operation)
    path = output_root / "control_points" / f"valores_migracion_{operation}.json"
    if not path.exists():
        return data
    try:
        payload = json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return data

    def _seg(node: dict) -> SegmentValor:
        return SegmentValor(
            label=_as_text(node.get("label")),
            c2_rows=_as_int(node.get("c2_rows")),
            c2_usd=_as_float(node.get("c2_usd")),
            vencidos_rows=_as_int(node.get("vencidos_rows")),
            vencidos_usd=_as_float(node.get("vencidos_usd")),
        )

    segments = payload.get("segments", {})
    data.contratos = _seg(segments.get("contratos", {}))
    data.ordenes_servicio = _seg(segments.get("ordenes_servicio", {}))
    totals = payload.get("totals", {})
    data.total_c2_rows = _as_int(totals.get("c2_rows"))
    data.total_c2_usd = _as_float(totals.get("c2_usd"))
    data.total_vencidos_rows = _as_int(totals.get("vencidos_rows"))
    data.total_vencidos_usd = _as_float(totals.get("vencidos_usd"))
    ost = payload.get("ost_sin_marco")
    if isinstance(ost, dict):
        data.ost_rows = _as_int(ost.get("rows"))
        data.ost_usd = _as_float(ost.get("usd"))
    data.usd_available = bool(payload.get("usd_rates_available"))
    return data


def _load_suministros_plants(output_root: Path, operation: str) -> list[PlantRow]:
    pattern = f"reportes/suministros_{operation.lower()}_*.xlsx"
    candidates = sorted(
        output_root.glob(pattern),
        key=lambda p: p.stat().st_mtime,
        reverse=True,
    )
    if not candidates:
        return []

    workbook = load_workbook(candidates[0], read_only=True, data_only=True)
    if "Resumen" not in workbook.sheetnames:
        workbook.close()
        return []

    ws = workbook["Resumen"]
    headers = [_as_text(ws.cell(1, c).value) for c in range(1, ws.max_column + 1)]

    col_planta = col_vig_f = col_vig_usd = col_vcs_f = col_vcs_usd = col_vss_f = None
    for idx, hdr in enumerate(headers):
        if hdr == "Planta":
            col_planta = idx
        elif hdr == "VIGENTE — filas":
            col_vig_f = idx
        elif hdr == "VIGENTE — valor USD":
            col_vig_usd = idx
        elif hdr == "NO_VIGENTE_CON_SALDO — filas":
            col_vcs_f = idx
        elif hdr == "NO_VIGENTE_CON_SALDO — valor USD":
            col_vcs_usd = idx
        elif hdr == "NO_VIGENTE_SIN_SALDO — filas":
            col_vss_f = idx

    if col_planta is None:
        workbook.close()
        return []

    rows = []
    for row in range(2, ws.max_row + 1):
        vals = [ws.cell(row, c).value for c in range(1, ws.max_column + 1)]
        planta = _as_text(vals[col_planta])
        if not planta:
            continue
        rows.append(PlantRow(
            planta=planta,
            vigente_filas=_as_int(vals[col_vig_f]) if col_vig_f is not None else 0,
            vigente_usd=_as_float(vals[col_vig_usd]) if col_vig_usd is not None else 0.0,
            vcs_filas=_as_int(vals[col_vcs_f]) if col_vcs_f is not None else 0,
            vcs_usd=_as_float(vals[col_vcs_usd]) if col_vcs_usd is not None else 0.0,
            vss_filas=_as_int(vals[col_vss_f]) if col_vss_f is not None else 0,
        ))

    workbook.close()
    return rows


def _load_characterization_criteria(project_root: Path) -> dict[str, SegmentCriteria]:
    candidates = sorted((project_root / "tmp").glob("Car*.xlsx"))
    if not candidates:
        return {}

    try:
        rules = load_segment_rules(candidates[0])
    except Exception:
        return {}

    criteria: dict[str, SegmentCriteria] = {}
    for rule in rules:
        entry = criteria.setdefault(rule.segment_id, SegmentCriteria())
        for condition in rule.conditions:
            if condition.column == "position_type" and condition.operator == "equals":
                entry.position_type = condition.value
            elif condition.column == "outline_contract":
                entry.marco = "con contrato marco" if condition.operator == "present" else "sin contrato marco"
            elif (
                condition.column == "purchase_doc_class"
                and condition.operator == "equals"
                and condition.value not in entry.doc_classes
            ):
                entry.doc_classes.append(condition.value)

    for entry in criteria.values():
        entry.doc_classes.sort()
    return criteria


def _characterization_lines(criteria: dict[str, SegmentCriteria]) -> list[str]:
    lines: list[str] = []
    for segment_id, label in (("contratos", "Contratos"), ("ordenes_servicio", "Órdenes de Servicio")):
        entry = criteria.get(segment_id)
        if entry is None:
            continue
        parts = ["posiciones de servicio"]
        if entry.marco:
            parts.append(entry.marco)
        if entry.doc_classes:
            parts.append("clases: " + ", ".join(entry.doc_classes))
        lines.append(f"{label} — " + " · ".join(parts))

    if not lines:
        lines = [
            "Contratos — posiciones de servicio con contrato marco",
            "Órdenes de Servicio — posiciones de servicio sin contrato marco",
        ]
    lines.append(
        "En CCMC la separación entre Contratos y Órdenes de Servicio usa la presencia de contrato marco. "
        "Las posiciones de reparación con contrato marco quedan fuera de este universo."
    )
    return lines


def _find_blank_layout(prs: Presentation):
    for layout in prs.slide_layouts:
        if layout.name.strip().lower() in {"blank", "en blanco"}:
            return layout
    return min(prs.slide_layouts, key=lambda candidate: len(candidate.placeholders))


def _delete_slide(prs: Presentation, index: int) -> None:
    slide_id = prs.slides._sldIdLst[index]
    prs.part.drop_rel(slide_id.rId)
    del prs.slides._sldIdLst[index]


def _set_shape_text(shape, text: str) -> None:
    text_frame = shape.text_frame
    text_frame.clear()
    paragraph = text_frame.paragraphs[0]
    run = paragraph.add_run()
    run.text = text


def _set_run_text_preserve(shape, text: str) -> None:
    """Cambia el texto conservando el formato del primer run (color/tamaño)."""
    paragraph = shape.text_frame.paragraphs[0]
    runs = paragraph.runs
    if runs:
        runs[0].text = text
        for extra in runs[1:]:
            extra.text = ""
    else:
        run = paragraph.add_run()
        run.text = text


def _update_cover_slide(slide, generated_on: date) -> None:
    date_pattern = re.compile(r"^\d{1,2}\s+[A-Za-zÁÉÍÓÚÑáéíóúñ]+\s+\d{4}$")
    subtitle_done = False
    for shape in slide.shapes:
        if not getattr(shape, "has_text_frame", False):
            continue
        text = shape.text.strip()
        normalized = text.replace("\xa0", " ")
        if normalized.startswith("Data Cleansing"):
            _set_run_text_preserve(shape, "Data Cleansing · Servicios")
        elif date_pattern.match(normalized):
            _set_run_text_preserve(shape, f"{generated_on.day:02d} {_month_name_es(generated_on)} {generated_on.year}")
        elif "Hallazgos" in normalized or "órdenes de compra" in normalized.lower():
            _set_run_text_preserve(shape, "Bases de datos · Método · Suministros · Contratos y Órdenes de Servicio · Registros de compras")
            subtitle_done = True

    if not subtitle_done:
        subtitle = slide.shapes.add_textbox(Inches(0.95), Inches(4.65), Inches(8.2), Inches(0.72))
        subtitle_frame = subtitle.text_frame
        subtitle_frame.clear()
        subtitle_frame.word_wrap = True
        paragraph = subtitle_frame.paragraphs[0]
        run = paragraph.add_run()
        run.text = "Bases de datos · Método · Suministros · Contratos y Órdenes de Servicio · Registros de compras"
        run.font.size = Pt(16)
        run.font.color.rgb = CHARCOAL


# Etiqueta de sección que se muestra en el kicker del encabezado.
SECTION_INTRO = "Servicios"
SECTION_BASES = "Bases de datos"
SECTION_METODO = "Método y criterios para migración"
SECTION_CONTRATOS = "Contratos y Órdenes de Servicio"
SECTION_SUMINISTROS = "Suministros"
SECTION_PIR = "Registros de compras (PIR)"


def _add_header(slide, slide_width: int, title: str, subtitle: str | None = None, section: str = SECTION_INTRO) -> None:
    banner = slide.shapes.add_shape(
        MSO_AUTO_SHAPE_TYPE.RECTANGLE,
        Inches(0),
        Inches(0),
        slide_width,
        Inches(0.42),
    )
    banner.fill.solid()
    banner.fill.fore_color.rgb = ORANGE
    banner.line.fill.background()

    kicker = slide.shapes.add_textbox(Inches(0.55), Inches(0.07), Inches(9.0), Inches(0.24))
    kicker_frame = kicker.text_frame
    kicker_frame.clear()
    paragraph = kicker_frame.paragraphs[0]
    run = paragraph.add_run()
    run.text = f"Data Cleansing  ·  {section}".upper()
    run.font.size = Pt(12)
    run.font.bold = True
    run.font.color.rgb = RGBColor(255, 255, 255)

    title_box = slide.shapes.add_textbox(Inches(0.7), Inches(0.72), Inches(11.8), Inches(0.6))
    title_frame = title_box.text_frame
    title_frame.clear()
    title_frame.word_wrap = True
    paragraph = title_frame.paragraphs[0]
    run = paragraph.add_run()
    run.text = title
    run.font.size = Pt(24)
    run.font.bold = True
    run.font.color.rgb = CHARCOAL

    if subtitle:
        subtitle_box = slide.shapes.add_textbox(Inches(0.7), Inches(1.25), Inches(11.5), Inches(0.48))
        subtitle_frame = subtitle_box.text_frame
        subtitle_frame.clear()
        subtitle_frame.word_wrap = True
        paragraph = subtitle_frame.paragraphs[0]
        run = paragraph.add_run()
        run.text = subtitle
        run.font.size = Pt(11)
        run.font.color.rgb = SLATE

    rule = slide.shapes.add_shape(
        MSO_AUTO_SHAPE_TYPE.RECTANGLE,
        Inches(0.7),
        Inches(1.74),
        Inches(2.1),
        Inches(0.04),
    )
    rule.fill.solid()
    rule.fill.fore_color.rgb = ORANGE
    rule.line.fill.background()


def _add_footnote(slide, text: str) -> None:
    box = slide.shapes.add_textbox(Inches(0.7), Inches(7.0), Inches(11.7), Inches(0.24))
    frame = box.text_frame
    frame.clear()
    paragraph = frame.paragraphs[0]
    run = paragraph.add_run()
    run.text = text
    run.font.size = Pt(9)
    run.font.color.rgb = SLATE


def _add_card(
    slide,
    left: float,
    top: float,
    width: float,
    height: float,
    title: str,
    value: str,
    note: str,
    fill: RGBColor,
    accent: RGBColor,
) -> None:
    shape = slide.shapes.add_shape(
        MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE,
        Inches(left),
        Inches(top),
        Inches(width),
        Inches(height),
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill
    shape.line.color.rgb = accent
    shape.line.width = Pt(1.35)

    title_box = slide.shapes.add_textbox(Inches(left + 0.18), Inches(top + 0.12), Inches(width - 0.36), Inches(0.28))
    title_frame = title_box.text_frame
    title_frame.clear()
    paragraph = title_frame.paragraphs[0]
    run = paragraph.add_run()
    run.text = title
    run.font.size = Pt(13)
    run.font.bold = True
    run.font.color.rgb = accent

    value_box = slide.shapes.add_textbox(Inches(left + 0.18), Inches(top + 0.42), Inches(width - 0.36), Inches(0.42))
    value_frame = value_box.text_frame
    value_frame.clear()
    paragraph = value_frame.paragraphs[0]
    run = paragraph.add_run()
    run.text = value
    run.font.size = Pt(24)
    run.font.bold = True
    run.font.color.rgb = CHARCOAL

    note_box = slide.shapes.add_textbox(Inches(left + 0.18), Inches(top + 0.84), Inches(width - 0.36), Inches(height - 0.98))
    note_frame = note_box.text_frame
    note_frame.clear()
    note_frame.word_wrap = True
    paragraph = note_frame.paragraphs[0]
    run = paragraph.add_run()
    run.text = note
    run.font.size = Pt(11)
    run.font.color.rgb = SLATE


def _add_callout_box(
    slide,
    left: float,
    top: float,
    width: float,
    height: float,
    title: str,
    lines: list[str],
    fill: RGBColor,
    accent: RGBColor,
    font_size: int = 10,
) -> None:
    shape = slide.shapes.add_shape(
        MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE,
        Inches(left),
        Inches(top),
        Inches(width),
        Inches(height),
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill
    shape.line.color.rgb = accent
    shape.line.width = Pt(1.2)

    title_box = slide.shapes.add_textbox(Inches(left + 0.18), Inches(top + 0.12), Inches(width - 0.36), Inches(0.22))
    title_frame = title_box.text_frame
    title_frame.clear()
    paragraph = title_frame.paragraphs[0]
    run = paragraph.add_run()
    run.text = title
    run.font.size = Pt(12)
    run.font.bold = True
    run.font.color.rgb = accent

    text_box = slide.shapes.add_textbox(Inches(left + 0.18), Inches(top + 0.34), Inches(width - 0.36), Inches(height - 0.44))
    text_frame = text_box.text_frame
    text_frame.clear()
    text_frame.word_wrap = True
    for index, line in enumerate(lines):
        paragraph = text_frame.paragraphs[0] if index == 0 else text_frame.add_paragraph()
        paragraph.space_after = Pt(2)
        run = paragraph.add_run()
        run.text = line
        run.font.size = Pt(font_size)
        run.font.color.rgb = SLATE


def _add_bar_box(
    slide,
    left: float,
    top: float,
    width: float,
    height: float,
    title: str,
    items: list[tuple[str, int]],
    fill: RGBColor,
    accent: RGBColor,
) -> None:
    shape = slide.shapes.add_shape(
        MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE,
        Inches(left),
        Inches(top),
        Inches(width),
        Inches(height),
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill
    shape.line.color.rgb = accent
    shape.line.width = Pt(1.2)

    title_box = slide.shapes.add_textbox(Inches(left + 0.18), Inches(top + 0.14), Inches(width - 0.36), Inches(0.28))
    title_frame = title_box.text_frame
    title_frame.clear()
    paragraph = title_frame.paragraphs[0]
    run = paragraph.add_run()
    run.text = title
    run.font.size = Pt(13)
    run.font.bold = True
    run.font.color.rgb = accent

    clean_items = [(label, value) for label, value in items if value > 0]
    if not clean_items:
        return
    max_value = max(value for _, value in clean_items) or 1
    row_height = (height - 0.7) / len(clean_items)

    for index, (label, value) in enumerate(clean_items):
        current_top = top + 0.52 + index * row_height
        label_box = slide.shapes.add_textbox(Inches(left + 0.18), Inches(current_top), Inches(2.0), Inches(0.22))
        label_frame = label_box.text_frame
        label_frame.clear()
        paragraph = label_frame.paragraphs[0]
        run = paragraph.add_run()
        run.text = label
        run.font.size = Pt(10)
        run.font.bold = True
        run.font.color.rgb = CHARCOAL

        bar_left = left + 2.1
        bar_width = width - 3.2
        bg_bar = slide.shapes.add_shape(
            MSO_AUTO_SHAPE_TYPE.RECTANGLE,
            Inches(bar_left),
            Inches(current_top + 0.04),
            Inches(bar_width),
            Inches(0.14),
        )
        bg_bar.fill.solid()
        bg_bar.fill.fore_color.rgb = RGBColor(255, 255, 255)
        bg_bar.line.color.rgb = GRAY

        fg_bar = slide.shapes.add_shape(
            MSO_AUTO_SHAPE_TYPE.RECTANGLE,
            Inches(bar_left),
            Inches(current_top + 0.04),
            Inches(max(bar_width * value / max_value, 0.08)),
            Inches(0.14),
        )
        fg_bar.fill.solid()
        fg_bar.fill.fore_color.rgb = accent
        fg_bar.line.fill.background()

        value_box = slide.shapes.add_textbox(Inches(left + width - 1.0), Inches(current_top - 0.02), Inches(0.82), Inches(0.24))
        value_frame = value_box.text_frame
        value_frame.clear()
        paragraph = value_frame.paragraphs[0]
        paragraph.alignment = PP_ALIGN.RIGHT
        run = paragraph.add_run()
        run.text = _format_int(value)
        run.font.size = Pt(10)
        run.font.color.rgb = SLATE


def _add_table(
    slide,
    rows: list[list[str]],
    left: float,
    top: float,
    width: float,
    height: float,
    col_widths: list[float],
) -> None:
    table = slide.shapes.add_table(len(rows), len(rows[0]), Inches(left), Inches(top), Inches(width), Inches(height)).table
    for index, current_width in enumerate(col_widths):
        table.columns[index].width = Inches(current_width)

    for row_index, row_values in enumerate(rows):
        for col_index, value in enumerate(row_values):
            cell = table.cell(row_index, col_index)
            cell.text = value
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            cell.text_frame.word_wrap = True
            paragraph = cell.text_frame.paragraphs[0]
            paragraph.alignment = PP_ALIGN.LEFT if col_index < 2 else PP_ALIGN.CENTER
            for run in paragraph.runs:
                run.font.size = Pt(10 if row_index > 0 else 9)
                run.font.bold = row_index == 0
                run.font.color.rgb = RGBColor(255, 255, 255) if row_index == 0 else CHARCOAL
            cell.fill.solid()
            cell.fill.fore_color.rgb = ORANGE if row_index == 0 else (GRAY_LIGHT if row_index % 2 else RGBColor(255, 255, 255))


def _get_segment(summary: OperationSummary, token: str) -> SegmentStats:
    for label, segment in summary.segments.items():
        if token.lower() in label.lower():
            return segment
    raise KeyError(f"No se encontro el segmento '{token}' en {summary.operation}")


def _add_index_slide(prs: Presentation) -> None:
    slide = prs.slides.add_slide(_find_blank_layout(prs))
    _add_header(
        slide,
        prs.slide_width,
        "Contenido",
        "Cómo está organizada esta presentación.",
        section=SECTION_INTRO,
    )
    sections = [
        ("1", "Entregables", "Paquete de salidas por bloque de negocio y propósito de uso."),
        ("2", "Bases de datos", "Universo de órdenes de compra y estructura por operación."),
        ("3", "Método y criterios para migración", "Cómo se decide qué migra: vigencia, saldo y rotulación."),
        ("4", "Suministros", "Stock, Consignación, Reparación y Traslado: vigencia, valor y apertura."),
        ("5", "Contratos y Órdenes de Servicio", "Resultados de migración y valor (USD) por operación."),
        ("6", "Registros de compras (PIR)", "Registros de información de compras (consignación y pipeline)."),
    ]
    top = 2.05
    for number, title, desc in sections:
        _add_step_box(slide, 0.9, top, 11.5, 0.82, number, f"{title}  —  {desc}", GRAY_LIGHT, ORANGE)
        top += 0.95
    _add_footnote(slide, "Data Cleansing · Servicios — MLCC (Caserones) y CCMC (Candelaria).")


def _add_intro_alcance_slide(prs: Presentation) -> None:
    slide = prs.slides.add_slide(_find_blank_layout(prs))
    _add_header(
        slide,
        prs.slide_width,
        "Aumento de alcance del servicio",
        "Identificación y limpieza de paquetes de datos en SAP, más allá del alcance original licitado.",
        section=SECTION_INTRO,
    )
    _add_callout_box(
        slide, 0.78, 2.0, 11.44, 1.2,
        "El encargo",
        [
            "Se incorporan tareas de identificación y limpieza en SAP de los datos de servicios, "
            "compras, bodega e ingeniería de materiales.",
            "Esta presentación se concentra en Suministros, Contratos y Órdenes de Servicio, y Registros de Compras (PIR).",
        ],
        BLUE_LIGHT, BLUE, font_size=12,
    )
    areas = [
        ("Convenios", "Materiales ZZ, duplicados o sin movimiento, convenios vigentes sin uso y contratos vencidos con saldo."),
        ("Órdenes de Compra", "Material catalogado y cargo directo: entregas vencidas, sin Incoterm y con atrasos a concluir."),
        ("Órdenes de Servicio", "OS vencidas con saldo, HES pendientes y OS vigentes sin movimiento."),
        ("PR / SOLPED · Reservas", "Solicitudes de servicios y materiales ya innecesarios, y reservas sin utilización."),
    ]
    top = 3.4
    for title, desc in areas:
        _add_callout_box(slide, 0.78, top, 11.44, 0.78, title, [desc], GRAY_LIGHT, CHARCOAL, font_size=11)
        top += 0.88
    _add_footnote(slide, "Fuente: Aumento de alcance — Servicios de Data Cleansing (Vantaz Analytics).")


def _add_entregables_slide(prs: Presentation) -> None:
    slide = prs.slides.add_slide(_find_blank_layout(prs))
    _add_header(
        slide,
        prs.slide_width,
        "Entregables",
        "Paquete de salidas operativas para migración y seguimiento.",
        section=SECTION_INTRO,
    )

    rows = [
        ["Bloque", "Entregable", "Contenido", "Ubicación"],
        ["Contratos / OS", "OC_migra_contratos_{OP}.xlsx", "Posiciones vigentes que migran a SAP.", "outputs/entregables/"],
        ["Contratos / OS", "OC_migra_ordenes_servicio_{OP}.xlsx", "Órdenes de servicio vigentes a migrar.", "outputs/entregables/"],
        ["Contratos / OS", "OC_vencidos_saldo_contratos_{OP}.xlsx", "Contratos vencidos con saldo para seguimiento.", "outputs/entregables/"],
        ["Contratos / OS", "OC_vencidos_saldo_ordenes_servicio_{OP}.xlsx", "OS vencidas con saldo para seguimiento.", "outputs/entregables/"],
        ["Suministros", "OC_migra_suministros_{OP}.xlsx", "Posiciones vigentes de Suministros que migran a SAP.", "outputs/entregables/"],
        ["Suministros", "OC_vencidos_saldo_suministros_{OP}.xlsx", "Suministros vencidos con saldo para seguimiento.", "outputs/entregables/"],
        ["Suministros", "suministros_{op}_*.xlsx", "Reporte de control: resumen, aperturas y detalle por categoría.", "outputs/reportes/"],
        ["Suministros", "resumen_segmentacion_{OP}.md", "Resumen ejecutivo por operación para trazabilidad.", "outputs/control_points/"],
        ["Presentación", "presentacion_servicios_YYYYMMDD_v1.pptx", "Síntesis ejecutiva consolidada del servicio.", "outputs/entregables/"],
    ]
    _add_table(slide, rows, 0.78, 2.05, 11.44, 3.2, [1.8, 3.0, 4.05, 2.59])

    _add_callout_box(
        slide, 0.78, 5.45, 11.44, 1.0,
        "Notas de entrega",
        [
            "{OP} representa MLCC y CCMC. Las salidas se publican por operación para uso directo de migración.",
            "Los archivos con prefijo LAST corresponden a personalizaciones manuales y no se sobrescriben en la generación estándar.",
        ],
        GRAY_LIGHT, CHARCOAL, font_size=11,
    )
    _add_footnote(slide, "Cada entregable conserva trazabilidad hacia las bases de control_points y reportes operativos.")


def _add_resumen_ejecutivo_slide(
    prs: Presentation,
    mlcc_val: MigracionValores,
    ccmc_val: MigracionValores,
    mlcc_comp: SuministrosData,
    ccmc_comp: SuministrosData,
) -> None:
    """Resumen ejecutivo: cifras finales (USD) consolidadas solo por operación.

    Migra a SAP = C2 de Contratos + C2 de Órdenes de Servicio + Suministros vigente.
    Vencidos con saldo = vencidos de Contratos/OS + Suministros con saldo (seguimiento).
    Sin más apertura que MLCC y CCMC.
    """
    slide = prs.slides.add_slide(_find_blank_layout(prs))
    _add_header(
        slide,
        prs.slide_width,
        "Resumen ejecutivo",
        "Resultado consolidado del data cleansing de Servicios — cifras finales por operación.",
        section=SECTION_INTRO,
    )

    mlcc_con_migra_usd = mlcc_val.total_c2_usd
    ccmc_con_migra_usd = ccmc_val.total_c2_usd
    mlcc_con_migra_rows = mlcc_val.total_c2_rows
    ccmc_con_migra_rows = ccmc_val.total_c2_rows
    mlcc_con_venc_usd = mlcc_val.total_vencidos_usd
    ccmc_con_venc_usd = ccmc_val.total_vencidos_usd
    mlcc_con_venc_rows = mlcc_val.total_vencidos_rows
    ccmc_con_venc_rows = ccmc_val.total_vencidos_rows

    mlcc_sum_migra_usd = mlcc_comp.vigente_usd
    ccmc_sum_migra_usd = ccmc_comp.vigente_usd
    mlcc_sum_migra_rows = mlcc_comp.vigente_rows
    ccmc_sum_migra_rows = ccmc_comp.vigente_rows
    mlcc_sum_venc_usd = mlcc_comp.vcs_usd
    ccmc_sum_venc_usd = ccmc_comp.vcs_usd
    mlcc_sum_venc_rows = mlcc_comp.vcs_rows
    ccmc_sum_venc_rows = ccmc_comp.vcs_rows

    mlcc_migra_usd = mlcc_con_migra_usd + mlcc_sum_migra_usd
    ccmc_migra_usd = ccmc_con_migra_usd + ccmc_sum_migra_usd
    mlcc_migra_rows = mlcc_con_migra_rows + mlcc_sum_migra_rows
    ccmc_migra_rows = ccmc_con_migra_rows + ccmc_sum_migra_rows
    mlcc_venc_usd = mlcc_con_venc_usd + mlcc_sum_venc_usd
    ccmc_venc_usd = ccmc_con_venc_usd + ccmc_sum_venc_usd
    mlcc_venc_rows = mlcc_con_venc_rows + mlcc_sum_venc_rows
    ccmc_venc_rows = ccmc_con_venc_rows + ccmc_sum_venc_rows

    total_migra_usd = mlcc_migra_usd + ccmc_migra_usd
    total_venc_usd = mlcc_venc_usd + ccmc_venc_usd
    total_migra_rows = mlcc_migra_rows + ccmc_migra_rows
    total_venc_rows = mlcc_venc_rows + ccmc_venc_rows
    total_con_migra_usd = mlcc_con_migra_usd + ccmc_con_migra_usd
    total_con_migra_rows = mlcc_con_migra_rows + ccmc_con_migra_rows
    total_con_venc_usd = mlcc_con_venc_usd + ccmc_con_venc_usd
    total_con_venc_rows = mlcc_con_venc_rows + ccmc_con_venc_rows
    total_sum_migra_usd = mlcc_sum_migra_usd + ccmc_sum_migra_usd
    total_sum_migra_rows = mlcc_sum_migra_rows + ccmc_sum_migra_rows
    total_sum_venc_usd = mlcc_sum_venc_usd + ccmc_sum_venc_usd
    total_sum_venc_rows = mlcc_sum_venc_rows + ccmc_sum_venc_rows

    _add_metric_panel(
        slide, 0.78, 2.0, 5.72, 1.22, "Migra a SAP", _fmt_usd(total_migra_usd),
        f"{_format_int(total_migra_rows)} posiciones vigentes · Contratos, OS y Suministros",
        GREEN_LIGHT, GREEN, value_pt=30,
    )
    _add_metric_panel(
        slide, 6.62, 2.0, 5.6, 1.22, "Vencidos con saldo", _fmt_usd(total_venc_usd),
        f"{_format_int(total_venc_rows)} posiciones · se trasladan como seguimiento",
        POS_ACCENT_LIGHT, POS_ACCENT, value_pt=30,
    )

    rows = [
        ["Operación", "Bloque", "Migra a SAP (USD)", "Posiciones", "Vencidos c/ saldo (USD)", "Posiciones"],
        ["MLCC", "Contratos / OS", _fmt_usd(mlcc_con_migra_usd), _format_int(mlcc_con_migra_rows), _fmt_usd(mlcc_con_venc_usd), _format_int(mlcc_con_venc_rows)],
        ["MLCC", "Suministros", _fmt_usd(mlcc_sum_migra_usd), _format_int(mlcc_sum_migra_rows), _fmt_usd(mlcc_sum_venc_usd), _format_int(mlcc_sum_venc_rows)],
        ["CCMC", "Contratos / OS", _fmt_usd(ccmc_con_migra_usd), _format_int(ccmc_con_migra_rows), _fmt_usd(ccmc_con_venc_usd), _format_int(ccmc_con_venc_rows)],
        ["CCMC", "Suministros", _fmt_usd(ccmc_sum_migra_usd), _format_int(ccmc_sum_migra_rows), _fmt_usd(ccmc_sum_venc_usd), _format_int(ccmc_sum_venc_rows)],
        ["Total", "Contratos / OS", _fmt_usd(total_con_migra_usd), _format_int(total_con_migra_rows), _fmt_usd(total_con_venc_usd), _format_int(total_con_venc_rows)],
        ["Total", "Suministros", _fmt_usd(total_sum_migra_usd), _format_int(total_sum_migra_rows), _fmt_usd(total_sum_venc_usd), _format_int(total_sum_venc_rows)],
    ]
    _add_table(slide, rows, 0.78, 3.45, 11.45, 2.45, [1.45, 2.1, 2.15, 1.3, 3.05, 1.4])

    lines = [
        "Se clasificó el universo de OC de ambas operaciones (Suministros, Contratos y Órdenes de Servicio) y se "
        "determinó qué migra a SAP según vigencia (fecha de término) y saldo pendiente, valorizado en USD.",
        "MLCC ahora considera todas sus plantas en Contratos y Órdenes de Servicio (se retiró el filtro por planta 8000).",
        f"Migran {_fmt_usd(total_migra_usd)} a SAP; adicionalmente {_fmt_usd(total_venc_usd)} en posiciones vencidas con "
        "saldo se trasladan como seguimiento.",
    ]
    _add_callout_box(slide, 0.78, 6.0, 11.44, 1.1, "Método", lines, GRAY_LIGHT, CHARCOAL, font_size=10)
    _add_footnote(slide, "USD = valor pendiente convertido a dólares. El detalle por segmento y planta se presenta en las secciones siguientes.")


def _add_universe_slide(
    prs: Presentation,
    mlcc_stats: POStats,
    ccmc_stats: POStats,
    mlcc_val: MigracionValores,
    ccmc_val: MigracionValores,
    mlcc_comp: SuministrosData,
    ccmc_comp: SuministrosData,
) -> None:
    slide = prs.slides.add_slide(_find_blank_layout(prs))
    total_positions = mlcc_stats.total_positions + ccmc_stats.total_positions
    total_documents = mlcc_stats.unique_documents + ccmc_stats.unique_documents
    total_framework = mlcc_stats.po_with_framework + ccmc_stats.po_with_framework
    pct_marco = (total_framework / total_documents * 100.0) if total_documents else 0.0
    total_migra_usd = (
        mlcc_val.total_c2_usd + ccmc_val.total_c2_usd
        + mlcc_val.total_vencidos_usd + ccmc_val.total_vencidos_usd
        + mlcc_comp.vigente_usd + ccmc_comp.vigente_usd
        + mlcc_comp.vcs_usd + ccmc_comp.vcs_usd
    )
    _add_header(
        slide,
        prs.slide_width,
        "Universo de órdenes de compra",
        "Cobertura combinada MLCC + CCMC · período 2007–2026.",
        section=SECTION_BASES,
    )
    _add_card(slide, 0.78, 2.0, 2.65, 1.4, "Posiciones de OC", _format_int(total_positions), "Universo consolidado MLCC + CCMC", BLUE_LIGHT, BLUE)
    _add_card(slide, 3.58, 2.0, 2.65, 1.4, "Documentos únicos", _format_int(total_documents), "Documentos de compra distintos", GREEN_LIGHT, GREEN)
    _add_card(slide, 6.38, 2.0, 2.65, 1.4, "OC con contrato marco", _format_int(total_framework), f"{_format_pct(pct_marco)} de los documentos", POS_ACCENT_LIGHT, POS_ACCENT)
    _add_card(slide, 9.18, 2.0, 3.05, 1.4, "Valor a migrar (USD)", _fmt_usd(total_migra_usd), "Contratos, OS y Suministros que migran", PURPLE_LIGHT, PURPLE)

    rows = [
        ["Operación", "Períodos", "Posiciones OC", "Documentos", "Con marco", "Sin marco"],
        [
            "MLCC",
            _compact_periods(mlcc_stats.periods),
            _format_int(mlcc_stats.total_positions),
            _format_int(mlcc_stats.unique_documents),
            _format_int(mlcc_stats.po_with_framework),
            _format_int(mlcc_stats.po_without_framework),
        ],
        [
            "CCMC",
            _compact_periods(ccmc_stats.periods),
            _format_int(ccmc_stats.total_positions),
            _format_int(ccmc_stats.unique_documents),
            _format_int(ccmc_stats.po_with_framework),
            _format_int(ccmc_stats.po_without_framework),
        ],
    ]
    _add_table(slide, rows, 0.78, 3.62, 11.45, 1.75, [1.25, 3.45, 1.55, 1.55, 1.35, 1.35])

    findings = [
        f"El valor que migra a SAP asciende a {_fmt_usd(total_migra_usd)} (Contratos, Órdenes de Servicio y Suministros vigentes o con saldo pendiente).",
        f"Solo el {_format_pct(pct_marco)} de los documentos de compra cuelga de un contrato marco; el resto se evalúa por documento.",
        f"{_format_int(total_documents)} documentos distintos a lo largo de ~20 años: el grueso del volumen son Suministros; Contratos y Órdenes de Servicio son una fracción acotada.",
    ]
    _add_callout_box(slide, 0.78, 5.5, 11.44, 1.35, "Lo que conviene destacar", findings, GRAY_LIGHT, CHARCOAL, font_size=10)
    _add_footnote(slide, "Fuente: universo de OC (carga vigente MLCC + CCMC). USD = valor pendiente convertido a dólares.")


def _operation_callout_lines(stats: POStats, snapshot: TempOCSnapshot) -> list[str]:
    pct_marco = stats.pct_po_with_framework or (
        (stats.po_with_framework / stats.unique_documents * 100.0) if stats.unique_documents else 0.0
    )
    lines = [
        f"{_format_int(stats.framework_contracts)} contratos marco distintos · {_format_int(stats.po_with_framework)} "
        f"documentos con marco ({_format_pct(pct_marco)}) · {_format_int(stats.po_without_framework)} sin marco.",
    ]
    top_class = stats.top_doc_classes(1)
    if top_class:
        cls, n = top_class[0]
        pct = (n / stats.total_positions * 100.0) if stats.total_positions else 0.0
        lines.append(f"Clase documental más frecuente: {cls} con {_format_int(n)} posiciones ({_format_pct(pct)}).")
    return lines


def _add_operation_structure_slide(prs: Presentation, stats: POStats, snapshot: TempOCSnapshot, accent: RGBColor, fill: RGBColor) -> None:
    slide = prs.slides.add_slide(_find_blank_layout(prs))
    nombre = "MLCC (Caserones)" if stats.operation == "MLCC" else "CCMC (Candelaria)"
    _add_header(
        slide,
        prs.slide_width,
        f"{nombre} — Estructura de las órdenes de compra",
        "Contratos marco, clases documentales y tipos de posición — conteos sobre el total de OC (posiciones).",
        section=SECTION_BASES,
    )
    _add_card(slide, 0.78, 2.0, 3.2, 1.35, "Posiciones de OC", _format_int(stats.total_positions), "Universo de la operación", fill, accent)
    _add_card(slide, 4.15, 2.0, 3.2, 1.35, "Documentos únicos", _format_int(stats.unique_documents), "Documentos de compra distintos", GREEN_LIGHT, GREEN)
    _add_card(
        slide,
        7.52,
        2.0,
        4.7,
        1.35,
        "Contratos marco",
        _format_int(stats.framework_contracts),
        f"OC con marco {_format_int(stats.po_with_framework)} · sin marco {_format_int(stats.po_without_framework)}",
        PURPLE_LIGHT,
        PURPLE,
    )
    pos_items = [
        (_pos_type_label(code), n)
        for code, n in sorted(stats.position_types.items(), key=lambda kv: (-kv[1], kv[0]))
    ]
    _add_bar_box(slide, 0.78, 3.6, 5.65, 2.35, "Clases documentales · sobre OC totales", stats.top_doc_classes(6), fill, accent)
    _add_bar_box(slide, 6.6, 3.6, 5.62, 2.35, "Tipos de posición · sobre OC totales", pos_items, POS_ACCENT_LIGHT, POS_ACCENT)
    _add_callout_box(slide, 0.78, 6.12, 11.44, 0.78, "Lectura de la estructura", _operation_callout_lines(stats, snapshot), GRAY_LIGHT, CHARCOAL, font_size=10)
    _add_footnote(slide, f"Fuente: universo de OC vigente de {stats.operation}. Clases y tipos contados sobre el total de posiciones.")


def _segment_title(slide, left: float, top: float, width: float, text: str, accent: RGBColor) -> None:
    box = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(0.4))
    frame = box.text_frame
    frame.clear()
    paragraph = frame.paragraphs[0]
    run = paragraph.add_run()
    run.text = text
    run.font.size = Pt(20)
    run.font.bold = True
    run.font.color.rgb = accent


def _add_metric_panel(
    slide,
    left: float,
    top: float,
    width: float,
    height: float,
    title: str,
    value: str,
    note: str,
    fill: RGBColor,
    accent: RGBColor,
    value_pt: int = 34,
) -> None:
    shape = slide.shapes.add_shape(
        MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE,
        Inches(left), Inches(top), Inches(width), Inches(height),
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill
    shape.line.color.rgb = accent
    shape.line.width = Pt(1.5)

    title_box = slide.shapes.add_textbox(Inches(left + 0.22), Inches(top + 0.13), Inches(width - 0.44), Inches(0.3))
    title_frame = title_box.text_frame
    title_frame.clear()
    paragraph = title_frame.paragraphs[0]
    run = paragraph.add_run()
    run.text = title
    run.font.size = Pt(15)
    run.font.bold = True
    run.font.color.rgb = accent

    value_box = slide.shapes.add_textbox(Inches(left + 0.22), Inches(top + 0.45), Inches(width - 0.44), Inches(0.62))
    value_frame = value_box.text_frame
    value_frame.clear()
    paragraph = value_frame.paragraphs[0]
    run = paragraph.add_run()
    run.text = value
    run.font.size = Pt(value_pt)
    run.font.bold = True
    run.font.color.rgb = CHARCOAL

    if note:
        note_box = slide.shapes.add_textbox(Inches(left + 0.22), Inches(top + height - 0.4), Inches(width - 0.44), Inches(0.34))
        note_frame = note_box.text_frame
        note_frame.clear()
        note_frame.word_wrap = True
        paragraph = note_frame.paragraphs[0]
        run = paragraph.add_run()
        run.text = note
        run.font.size = Pt(12)
        run.font.color.rgb = SLATE


def _add_step_box(slide, left: float, top: float, width: float, height: float, number: str, text: str, fill: RGBColor, accent: RGBColor) -> None:
    shape = slide.shapes.add_shape(
        MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE,
        Inches(left), Inches(top), Inches(width), Inches(height),
    )
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill
    shape.line.color.rgb = accent
    shape.line.width = Pt(1.2)

    num_box = slide.shapes.add_textbox(Inches(left + 0.18), Inches(top), Inches(0.7), Inches(height))
    num_frame = num_box.text_frame
    num_frame.clear()
    num_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
    paragraph = num_frame.paragraphs[0]
    run = paragraph.add_run()
    run.text = number
    run.font.size = Pt(28)
    run.font.bold = True
    run.font.color.rgb = accent

    text_box = slide.shapes.add_textbox(Inches(left + 0.95), Inches(top), Inches(width - 1.15), Inches(height))
    text_frame = text_box.text_frame
    text_frame.clear()
    text_frame.word_wrap = True
    text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
    paragraph = text_frame.paragraphs[0]
    run = paragraph.add_run()
    run.text = text
    run.font.size = Pt(15)
    run.font.color.rgb = CHARCOAL


def _add_criterios_slide(prs: Presentation, criteria_lines: list[str]) -> None:
    slide = prs.slides.add_slide(_find_blank_layout(prs))
    _add_header(
        slide,
        prs.slide_width,
        "Criterios de migración",
        "Qué migra y cómo se decide — mismo método para MLCC y CCMC.",
        section=SECTION_METODO,
    )

    _add_callout_box(
        slide, 0.78, 2.0, 11.44, 1.7,
        "Vigencia y saldo — la regla",
        [
            "Una posición MIGRA si su contrato está vigente al corte del 30 de junio de 2026.",
            "Contratos y Órdenes de Servicio: la vigencia surge del cruce con el reporte de Contratos Vigentes "
            "(por contrato marco; si no existe, por documento de compra).",
            "Suministros: vigencia por fecha de entrega/validez. Si está vencida pero conserva saldo pendiente (> 0), "
            "igual migra como caso con saldo; sin saldo pendiente, no migra.",
        ],
        BLUE_LIGHT, BLUE, font_size=12,
    )

    crit = [ln for ln in criteria_lines if ln]
    _add_callout_box(
        slide, 0.78, 3.85, 5.6, 2.95,
        "Contratos y Órdenes de Servicio",
        crit or ["Contratos — tipo servicio, con contrato marco", "Órdenes de Servicio — tipo servicio, sin contrato marco"],
        GREEN_LIGHT, GREEN, font_size=10,
    )
    _add_callout_box(
        slide, 6.62, 3.85, 5.6, 2.95,
        "Suministros — rótulos por tipo de posición",
        [
            "Reparación — subcontratación (tipo L)",
            "Consignación — tipo C (MLCC) / K (CCMC)",
            "Traslado — tipo V (MLCC) / U (CCMC)",
            "Stock — sin tipo, con contrato marco",
            "Cargo Directo — sin tipo, sin contrato marco",
            "La rotulación es ortogonal a la vigencia: dentro de cada bloque hay posiciones que migran y que no.",
        ],
        POS_ACCENT_LIGHT, POS_ACCENT, font_size=10,
    )
    _add_footnote(slide, "Reporte de Contratos Vigentes: CCMC mayo 2026 · MLCC mayo 2026. Corte de vigencia: 30 jun 2026.")


def _no_migra_rows(summary: OperationSummary) -> int:
    return _get_segment(summary, "Contrato").c2_no_migra_rows + _get_segment(summary, "Servicio").c2_no_migra_rows


def _add_contratos_overview_slide(
    prs: Presentation,
    mlcc_summary: OperationSummary,
    ccmc_summary: OperationSummary,
    mlcc_val: MigracionValores,
    ccmc_val: MigracionValores,
) -> None:
    slide = prs.slides.add_slide(_find_blank_layout(prs))
    _add_header(
        slide,
        prs.slide_width,
        "Contratos y Órdenes de Servicio — Visión general",
        "Cuánto migra y cuánto vale (USD), combinando MLCC y CCMC. Corte de vigencia: 30 jun 2026.",
        section=SECTION_CONTRATOS,
    )

    migra_rows = mlcc_val.total_c2_rows + ccmc_val.total_c2_rows
    migra_usd = mlcc_val.total_c2_usd + ccmc_val.total_c2_usd
    venc_rows = mlcc_val.total_vencidos_rows + ccmc_val.total_vencidos_rows
    venc_usd = mlcc_val.total_vencidos_usd + ccmc_val.total_vencidos_usd
    no_migra = _no_migra_rows(mlcc_summary) + _no_migra_rows(ccmc_summary)

    _add_metric_panel(
        slide, 0.78, 1.95, 5.85, 1.5, "USD que migra a SAP", _fmt_usd(migra_usd),
        f"{_format_int(migra_rows)} posiciones con contrato vigente", GREEN_LIGHT, GREEN, value_pt=40,
    )
    _add_metric_panel(
        slide, 6.85, 1.95, 5.4, 1.5, "USD vencidos con saldo", _fmt_usd(venc_usd),
        f"{_format_int(venc_rows)} posiciones · migran como seguimiento", POS_ACCENT_LIGHT, POS_ACCENT, value_pt=36,
    )

    rows = [
        ["Operación", "Migran", "Migran USD", "Venc. c/saldo", "Venc. USD", "No migran"],
        [
            "MLCC",
            _format_int(mlcc_val.total_c2_rows), _fmt_usd(mlcc_val.total_c2_usd),
            _format_int(mlcc_val.total_vencidos_rows), _fmt_usd(mlcc_val.total_vencidos_usd),
            _format_int(_no_migra_rows(mlcc_summary)),
        ],
        [
            "CCMC",
            _format_int(ccmc_val.total_c2_rows), _fmt_usd(ccmc_val.total_c2_usd),
            _format_int(ccmc_val.total_vencidos_rows), _fmt_usd(ccmc_val.total_vencidos_usd),
            _format_int(_no_migra_rows(ccmc_summary)),
        ],
        [
            "TOTAL",
            _format_int(migra_rows), _fmt_usd(migra_usd),
            _format_int(venc_rows), _fmt_usd(venc_usd),
            _format_int(no_migra),
        ],
    ]
    _add_table(slide, rows, 0.78, 3.75, 11.44, 1.85, [1.6, 1.6, 2.1, 1.9, 2.1, 2.14])

    lines = [
        "Migran = posiciones con contrato vigente. Vencidos con saldo = sin vigencia pero con saldo pendiente (> 0); igual migran.",
        "Valores en USD = valor pendiente convertido a dólares al tipo de cambio del día.",
    ]
    if ccmc_val.ost_rows:
        lines.append(
            f"Adicional CCMC — OST vigentes sin contrato marco: {_format_int(ccmc_val.ost_rows)} posiciones · {_fmt_usd(ccmc_val.ost_usd)} (se anexan al entregable)."
        )
    _add_callout_box(slide, 0.78, 5.75, 11.44, 1.1, "Cómo leer estas cifras", lines, GRAY_LIGHT, CHARCOAL, font_size=10)
    _add_footnote(slide, "Fuente: resúmenes de segmentación y valores_migracion_<op>.json (USD).")


def _add_operation_results_slide(
    prs: Presentation,
    summary: OperationSummary,
    operation: str,
    accent: RGBColor,
    fill: RGBColor,
    valores: MigracionValores,
) -> None:
    slide = prs.slides.add_slide(_find_blank_layout(prs))
    contracts = _get_segment(summary, "Contrato")
    services = _get_segment(summary, "Servicio")
    nombre = "CCMC (Candelaria)" if operation.upper() == "CCMC" else "MLCC (Caserones)"

    _add_header(
        slide,
        prs.slide_width,
        f"{nombre} — Contratos y Órdenes de Servicio",
        "Posiciones, vigencia y valor (USD) de lo que migra. Corte: 30 jun 2026.",
        section=SECTION_CONTRATOS,
    )

    columns = [
        (0.78, "Contratos", contracts, valores.contratos),
        (6.85, "Órdenes de Servicio", services, valores.ordenes_servicio),
    ]
    for left, title, segment, seg_val in columns:
        _segment_title(slide, left, 1.92, 5.62, title, accent)
        _add_metric_panel(
            slide, left, 2.34, 5.62, 1.35,
            "Posiciones (C1)", _format_int(segment.c1_rows),
            f"{_format_int(segment.c1_documents)} documentos únicos", fill, accent, value_pt=34,
        )
        _add_metric_panel(
            slide, left, 3.79, 5.62, 1.35,
            "Migran a SAP", _format_int(segment.c2_rows),
            f"Contrato vigente  ·  {_fmt_usd(seg_val.c2_usd)}", GREEN_LIGHT, GREEN, value_pt=30,
        )
        _add_metric_panel(
            slide, left, 5.24, 5.62, 1.35,
            "Vencidos con saldo", _format_int(seg_val.vencidos_rows),
            f"Migran como seguimiento  ·  {_fmt_usd(seg_val.vencidos_usd)}", POS_ACCENT_LIGHT, POS_ACCENT, value_pt=30,
        )

    nota = (
        "Valores en USD = valor pendiente convertido a dólares. "
        "Las posiciones que caen en ambos segmentos se cuentan en cada uno."
    )
    if valores.ost_rows and operation.upper() == "CCMC":
        nota += f"  ·  OST vigentes sin contrato marco: {_format_int(valores.ost_rows)} ({_fmt_usd(valores.ost_usd)})."
    _add_footnote(slide, nota)


def _add_fuentes_contratos_slide(prs: Presentation) -> None:
    """Tabla de bases de datos usadas en Contratos y Órdenes de Servicio."""
    slide = prs.slides.add_slide(_find_blank_layout(prs))
    _add_header(
        slide,
        prs.slide_width,
        "Bases de datos — Contratos y Órdenes de Servicio",
        "Qué archivo alimenta cada categoría, contra qué se cruza la vigencia y la fecha de la base.",
        section=SECTION_CONTRATOS,
    )

    rows = [
        ["Categoría", "Oper.", "Base de datos (carpeta · archivos)", "Cruce de vigencia", "Fecha base"],
        ["Contratos", "MLCC", "ordenes-compra · ME2N_SERVICIOS (7 arch.)",
         "Reporte Contratos Vigentes Mayo-26 · hoja VIGENTES (Sitio CAS)", "jun-26"],
        ["Contratos", "CCMC", "ordenes-compra · ME2N (4 arch.)",
         "CONTRATOS VIGENTES CCMC Mayo-26 · Report Cttos Vigentes + Report Pedidos", "jun-26"],
        ["Órdenes de Servicio", "MLCC", "suministros-ordenes-compra (4 arch.)",
         "Sin cruce · vigencia por fecha de entrega", "may-26"],
        ["Órdenes de Servicio", "CCMC", "ordenes-compra · ME2N (4 arch.)",
         "CONTRATOS VIGENTES CCMC Mayo-26 · Report Pedidos", "jun-26"],
    ]
    _add_table(slide, rows, 0.78, 2.05, 11.44, 2.55, [1.7, 0.85, 3.2, 4.55, 1.14])

    _add_callout_box(
        slide, 0.78, 4.95, 11.44, 1.25,
        "Hoja extra del cliente (Órdenes de Servicio MLCC)",
        [
            "Se anexa la hoja «Sitio Caserones_BD mar26» (catálogo del cliente, mar-26) al entregable de OS MLCC.",
            "Aporta las OS vigentes por «Fin período validez» que el flujo no marca vigentes al medir por fecha de entrega.",
        ],
        GRAY_LIGHT, CHARCOAL, font_size=11,
    )
    _add_footnote(
        slide,
        "Vigencia D = cruce con CONTRATOS VIGENTES (clave: contrato marco; si no, documento de compra). "
        "OS MLCC = fecha de entrega de la base de suministros. Corte de migración: 30 jun 2026.",
    )


def _add_fuentes_suministros_slide(prs: Presentation) -> None:
    """Tabla de bases de datos usadas en Suministros."""
    slide = prs.slides.add_slide(_find_blank_layout(prs))
    _add_header(
        slide,
        prs.slide_width,
        "Bases de datos — Suministros",
        "Fuente única por operación; la vigencia se mide en la propia base, sin cruce con contratos.",
        section=SECTION_SUMINISTROS,
    )

    rows = [
        ["Operación", "Base de datos (carpeta · archivos)", "Período de datos", "Criterio de vigencia", "Cruce"],
        ["MLCC", "suministros-ordenes-compra (4 arch.)", "2010 – 13 may 2026", "Fecha de entrega → Fin validez", "Sin cruce"],
        ["CCMC", "suministros-ordenes-compra (3 arch.)", "2007 – 20 may 2026", "Fecha de entrega → Fin validez", "Sin cruce"],
    ]
    _add_table(slide, rows, 0.78, 2.15, 11.44, 1.55, [1.3, 3.55, 2.35, 3.0, 1.24])

    _add_callout_box(
        slide, 0.78, 4.05, 11.44, 1.7,
        "Cómo se usa esta base",
        [
            "Es la misma fuente para las cinco aperturas: Stock, Consignación, Reparación, Traslado y Cargo Directo.",
            "Vigencia por fecha de entrega (fallback: fin de validez); con saldo pendiente migra, sin saldo no migra.",
            "No se cruza contra los reportes de contratos vigentes: el estado de vigencia vive en la propia base.",
        ],
        GRAY_LIGHT, CHARCOAL, font_size=11,
    )
    _add_footnote(
        slide,
        "Las posiciones tipo D (Contratos y Órdenes de Servicio) se procesan aparte y no entran en Suministros. "
        "Corte de migración: 30 jun 2026.",
    )


def _add_suministros_slide(
    prs: Presentation,
    mlcc_data: SuministrosData,
    ccmc_data: SuministrosData,
) -> None:
    slide = prs.slides.add_slide(_find_blank_layout(prs))

    total_vigente = mlcc_data.vigente_rows + ccmc_data.vigente_rows
    total_vcs = mlcc_data.vcs_rows + ccmc_data.vcs_rows
    total_vss = mlcc_data.vss_rows + ccmc_data.vss_rows
    total_vigente_usd = mlcc_data.vigente_usd + ccmc_data.vigente_usd
    total_vcs_usd = mlcc_data.vcs_usd + ccmc_data.vcs_usd

    _add_header(
        slide,
        prs.slide_width,
        "Suministros — Visión general",
        "Stock, Consignación, Reparación y Traslado, segmentados por vigencia y saldo. Corte: 30 jun 2026.",
        section=SECTION_SUMINISTROS,
    )

    # KPI cards
    _add_card(slide, 0.78, 2.0, 3.72, 1.3, "Vigentes", _format_int(total_vigente), f"Migran  ·  {_fmt_usd(total_vigente_usd)}", GREEN_LIGHT, GREEN)
    _add_card(slide, 4.72, 2.0, 3.72, 1.3, "Vencidos con saldo", _format_int(total_vcs), f"Migran por saldo  ·  {_fmt_usd(total_vcs_usd)}", POS_ACCENT_LIGHT, POS_ACCENT)
    _add_card(slide, 8.66, 2.0, 3.89, 1.3, "Vencidos sin saldo", _format_int(total_vss), "No migran", RED_LIGHT, RED)

    # Table: per operation
    rows = [
        ["Operación", "Vigentes", "Vig. USD", "Venc. c/saldo", "USD c/saldo", "Venc. s/saldo"],
        [
            "MLCC",
            _format_int(mlcc_data.vigente_rows),
            _fmt_usd(mlcc_data.vigente_usd),
            _format_int(mlcc_data.vcs_rows),
            _fmt_usd(mlcc_data.vcs_usd),
            _format_int(mlcc_data.vss_rows),
        ],
        [
            "CCMC",
            _format_int(ccmc_data.vigente_rows),
            _fmt_usd(ccmc_data.vigente_usd),
            _format_int(ccmc_data.vcs_rows),
            _fmt_usd(ccmc_data.vcs_usd),
            _format_int(ccmc_data.vss_rows),
        ],
        [
            "TOTAL",
            _format_int(total_vigente),
            _fmt_usd(total_vigente_usd),
            _format_int(total_vcs),
            _fmt_usd(total_vcs_usd),
            _format_int(total_vss),
        ],
    ]
    _add_table(slide, rows, 0.78, 3.55, 11.44, 1.85, [1.45, 1.55, 1.95, 1.9, 1.95, 1.84])

    migra_rows = total_vigente + total_vcs
    migra_usd = total_vigente_usd + total_vcs_usd
    lines = [
        f"Migran {_format_int(migra_rows)} posiciones (vigentes + vencidas con saldo) por un total de {_fmt_usd(migra_usd)}.",
        f"No migran {_format_int(total_vss)} posiciones vencidas sin saldo pendiente — el grueso del volumen, pero sin valor a migrar.",
    ]
    _add_callout_box(slide, 0.78, 5.55, 11.44, 1.15, "Cuánto migra y cuánto vale", lines, GRAY_LIGHT, CHARCOAL, font_size=11)

    _add_footnote(slide, "Fuente: reportes de Suministros en outputs/reportes/. Valores USD = valor pendiente convertido a dólares. Corte: 30 jun 2026.")


def _add_suministros_segment_slide(
    prs: Presentation,
    mlcc_segs: list[SupplySegmentRow],
    ccmc_segs: list[SupplySegmentRow],
) -> None:
    """Apertura de Suministros por sub-bloque (Stock, Consignación, Reparación, ...)."""
    slide = prs.slides.add_slide(_find_blank_layout(prs))
    _add_header(
        slide,
        prs.slide_width,
        "Suministros — Apertura por sub-bloque",
        "Stock, Consignación, Reparación y Traslado (MLCC + CCMC), con valor a migrar en USD. Corte: 30 jun 2026.",
        section=SECTION_SUMINISTROS,
    )

    order = ["Stock", "Cargo Directo", "Consignación", "Reparación", "Traslado/Transporte", "Otros"]
    combined: dict[str, SupplySegmentRow] = {}
    for seg_row in [*mlcc_segs, *ccmc_segs]:
        entry = combined.setdefault(seg_row.segment, SupplySegmentRow(segment=seg_row.segment))
        entry.vigente_rows += seg_row.vigente_rows
        entry.vcs_rows += seg_row.vcs_rows
        entry.vss_rows += seg_row.vss_rows
        entry.vigente_usd += seg_row.vigente_usd
        entry.vcs_usd += seg_row.vcs_usd

    present = [seg for seg in order if seg in combined]
    present += [seg for seg in combined if seg not in order]

    rows = [["Sub-bloque", "Vigentes", "Vig. USD", "Venc. c/saldo", "USD c/saldo", "Venc. s/saldo"]]
    tot = SupplySegmentRow(segment="TOTAL")
    for seg in present:
        e = combined[seg]
        tot.vigente_rows += e.vigente_rows
        tot.vcs_rows += e.vcs_rows
        tot.vss_rows += e.vss_rows
        tot.vigente_usd += e.vigente_usd
        tot.vcs_usd += e.vcs_usd
        rows.append([
            seg,
            _format_int(e.vigente_rows),
            _fmt_usd(e.vigente_usd),
            _format_int(e.vcs_rows),
            _fmt_usd(e.vcs_usd),
            _format_int(e.vss_rows),
        ])
    rows.append([
        "TOTAL",
        _format_int(tot.vigente_rows),
        _fmt_usd(tot.vigente_usd),
        _format_int(tot.vcs_rows),
        _fmt_usd(tot.vcs_usd),
        _format_int(tot.vss_rows),
    ])

    n_rows = len(rows)
    tbl_height = min(4.2, max(1.4, n_rows * 0.46))
    _add_table(slide, rows, 0.78, 2.1, 11.44, tbl_height, [2.7, 1.5, 1.95, 1.84, 1.95, 1.5])

    _add_callout_box(
        slide, 0.78, 6.35, 11.44, 0.6,
        "Migran las posiciones vigentes y las vencidas con saldo (columnas en USD); las vencidas sin saldo no migran.",
        ["Reparación = subcontratación (L) · Consignación = C/K · Traslado = V/U · Stock = sin tipo con marco · Cargo Directo = sin tipo sin marco."],
        GRAY_LIGHT, CHARCOAL, font_size=9,
    )
    _add_footnote(slide, "Fuente: hoja Por_Subsegmento de los reportes de Suministros en outputs/reportes/. Corte: 30 jun 2026.")


def _add_suministros_plant_slide(prs: Presentation, operation: str, plant_rows: list[PlantRow]) -> None:
    slide = prs.slides.add_slide(_find_blank_layout(prs))
    nombre = "MLCC (Caserones)" if operation.upper() == "MLCC" else "CCMC (Candelaria)"
    _add_header(
        slide,
        prs.slide_width,
        f"{nombre} — Suministros por planta",
        "Posiciones de suministros, por vigencia, saldo y valor (USD). Corte: 30 jun 2026.",
        section=SECTION_SUMINISTROS,
    )

    total_row = next((r for r in plant_rows if r.planta == "TOTAL"), None)

    def _sort_key(r: PlantRow) -> tuple:
        if r.vigente_filas > 0:
            return (0, -(r.vigente_filas + r.vcs_filas))
        if r.vcs_filas > 0:
            return (1, -r.vcs_filas)
        return (2, -r.vss_filas)

    plant_data = sorted(
        [r for r in plant_rows if r.planta != "TOTAL"],
        key=_sort_key,
    )

    ordered: list[tuple[PlantRow, bool]] = []  # (row, is_total)
    if total_row:
        ordered.append((total_row, True))
    ordered.extend((r, False) for r in plant_data)

    n_rows = 1 + len(ordered)
    tbl_height = min(5.1, max(1.2, n_rows * 0.26))
    font_size = 8 if len(plant_data) > 10 else 10

    col_widths = [1.3, 1.65, 2.2, 1.65, 2.2, 2.44]  # 11.44 total

    table = slide.shapes.add_table(
        n_rows, 6, Inches(0.78), Inches(1.9), Inches(11.44), Inches(tbl_height)
    ).table

    for idx, cw in enumerate(col_widths):
        table.columns[idx].width = Inches(cw)

    for c_idx, hdr_text in enumerate(["Planta", "Vigentes", "Vig. USD", "Venc. c/saldo", "USD c/saldo", "Venc. s/saldo"]):
        cell = table.cell(0, c_idx)
        cell.text = hdr_text
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE
        para = cell.text_frame.paragraphs[0]
        para.alignment = PP_ALIGN.LEFT if c_idx == 0 else PP_ALIGN.CENTER
        for run in para.runs:
            run.font.size = Pt(9)
            run.font.bold = True
            run.font.color.rgb = RGBColor(255, 255, 255)
        cell.fill.solid()
        cell.fill.fore_color.rgb = ORANGE

    for r_idx, (row_data, is_total) in enumerate(ordered):
        row_index = r_idx + 1
        no_migra = not is_total and row_data.vigente_filas == 0 and row_data.vcs_filas == 0
        base_bg = GRAY_LIGHT if row_index % 2 else RGBColor(255, 255, 255)

        row_cells = [
            row_data.planta,
            _format_int(row_data.vigente_filas),
            _fmt_usd(row_data.vigente_usd),
            _format_int(row_data.vcs_filas),
            _fmt_usd(row_data.vcs_usd),
            _format_int(row_data.vss_filas),
        ]

        for c_idx, val in enumerate(row_cells):
            cell = table.cell(row_index, c_idx)
            cell.text = val
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            para = cell.text_frame.paragraphs[0]
            para.alignment = PP_ALIGN.LEFT if c_idx == 0 else PP_ALIGN.CENTER
            for run in para.runs:
                run.font.size = Pt(font_size)
                run.font.bold = is_total or c_idx == 5
                run.font.color.rgb = CHARCOAL
            cell.fill.solid()
            if no_migra and c_idx in (4, 5):
                cell.fill.fore_color.rgb = RED_LIGHT
            else:
                cell.fill.fore_color.rgb = base_bg

    _add_footnote(slide, "Fuente: reporte de Suministros más reciente en outputs/reportes/. Corte vigencia: 30 Jun 2026.")


def _slide_text_blob(slide) -> str:
    parts = []
    for shape in slide.shapes:
        if getattr(shape, "has_text_frame", False) and shape.text_frame.text.strip():
            parts.append(shape.text_frame.text)
    return " \n ".join(parts).lower()


def _find_slide_index(prs: Presentation, *needles: str) -> int | None:
    for index, slide in enumerate(prs.slides):
        blob = _slide_text_blob(slide)
        if any(needle.lower() in blob for needle in needles):
            return index
    return None


def _reorder_slides_by_index(prs: Presentation, desired_indices: list[int]) -> None:
    """Reordena las láminas según índices del orden documental actual."""
    sld_lst = prs.slides._sldIdLst
    elements = list(sld_lst)
    for element in elements:
        sld_lst.remove(element)
    for idx in desired_indices:
        sld_lst.append(elements[idx])


def generate_project_presentation(
    template_path: Path,
    output_path: Path,
    mlcc_summary_path: Path,
    ccmc_summary_path: Path,
    output_root: Path,
    generated_on: date | None = None,
) -> Path:
    generated_on = generated_on or date.today()
    project_root = output_root.parent
    mlcc_stats_path = output_root / "ordenes_compra" / "reporte_estadistico_MLCC-PO.xlsx"
    ccmc_stats_path = output_root / "ordenes_compra" / "reporte_estadistico_CCMC-PO.xlsx"
    tmp_stats_path = project_root / "inputs" / "Estadistica de OC.xlsx"
    if not tmp_stats_path.exists():
        tmp_stats_path = project_root / "tmp" / "Estadistica de OC.xlsx"

    prs = Presentation(str(template_path))
    mlcc_summary = parse_summary_markdown(mlcc_summary_path)
    ccmc_summary = parse_summary_markdown(ccmc_summary_path)
    mlcc_stats = _load_universo_oc(output_root, "MLCC") or _load_po_stats(mlcc_stats_path, "MLCC")
    ccmc_stats = _load_universo_oc(output_root, "CCMC") or _load_po_stats(ccmc_stats_path, "CCMC")
    mlcc_tmp = _load_temp_snapshot(tmp_stats_path, "MLCC")
    ccmc_tmp = _load_temp_snapshot(tmp_stats_path, "CCMC")
    mlcc_comp = _load_suministros_data(output_root, "MLCC")
    ccmc_comp = _load_suministros_data(output_root, "CCMC")
    mlcc_plants = _load_suministros_plants(output_root, "MLCC")
    ccmc_plants = _load_suministros_plants(output_root, "CCMC")
    mlcc_segs = _load_suministros_segments(output_root, "MLCC")
    ccmc_segs = _load_suministros_segments(output_root, "CCMC")
    mlcc_val = _load_valores_migracion(output_root, "MLCC")
    ccmc_val = _load_valores_migracion(output_root, "CCMC")
    criteria_lines = _characterization_lines(_load_characterization_criteria(project_root))

    # Láminas hechas a mano que se conservan de la plantilla base (no se re-codean):
    # portada (índice 0), diagrama de flujo (Método) y PIR (última sección, con gráfico).
    flow_idx = _find_slide_index(prs, "flujo de clasificación", "migra / no migra")
    pir_idx = _find_slide_index(prs, "registro de información de compras", "(pir)")
    keep = {0}
    if flow_idx is not None:
        keep.add(flow_idx)
    if pir_idx is not None:
        keep.add(pir_idx)
    for index in range(len(prs.slides) - 1, -1, -1):
        if index not in keep:
            _delete_slide(prs, index)

    # Tras la poda, los supervivientes mantienen el orden: portada, [flujo], [pir].
    cover_pos = 0
    pos = 1
    flow_pos = None
    if flow_idx is not None:
        flow_pos = pos
        pos += 1
    pir_pos = None
    if pir_idx is not None:
        pir_pos = pos
        pos += 1
    built_start = pos

    _update_cover_slide(prs.slides[cover_pos], generated_on)

    # Láminas nuevas (se anexan al final en este orden).
    _add_index_slide(prs)                                                                       # b0
    _add_intro_alcance_slide(prs)                                                               # b1
    _add_entregables_slide(prs)                                                                 # b2
    _add_universe_slide(prs, mlcc_stats, ccmc_stats, mlcc_val, ccmc_val, mlcc_comp, ccmc_comp)  # b3
    _add_operation_structure_slide(prs, mlcc_stats, mlcc_tmp, BLUE, BLUE_LIGHT)                  # b4
    _add_operation_structure_slide(prs, ccmc_stats, ccmc_tmp, GREEN, GREEN_LIGHT)               # b5
    _add_criterios_slide(prs, criteria_lines)                                                   # b6
    _add_suministros_slide(prs, mlcc_comp, ccmc_comp)                                           # b7
    _add_suministros_segment_slide(prs, mlcc_segs, ccmc_segs)                                   # b8
    _add_suministros_plant_slide(prs, "MLCC", mlcc_plants)                                      # b9
    _add_suministros_plant_slide(prs, "CCMC", ccmc_plants)                                     # b10
    _add_contratos_overview_slide(prs, mlcc_summary, ccmc_summary, mlcc_val, ccmc_val)          # b11
    _add_operation_results_slide(prs, ccmc_summary, "CCMC", GREEN, GREEN_LIGHT, ccmc_val)      # b12
    _add_operation_results_slide(prs, mlcc_summary, "MLCC", BLUE, BLUE_LIGHT, mlcc_val)        # b13
    _add_resumen_ejecutivo_slide(prs, mlcc_val, ccmc_val, mlcc_comp, ccmc_comp)                 # b14
    _add_fuentes_contratos_slide(prs)                                                           # b15
    _add_fuentes_suministros_slide(prs)                                                         # b16
    b = list(range(built_start, built_start + 17))  # índices de las 17 nuevas

    # Orden final por secciones (Contratos y Órdenes de Servicio al final).
    desired = [cover_pos, b[0], b[1], b[2], b[14], b[3], b[4], b[5]]  # portada · índice · alcance · entregables · resumen ejecutivo · universo · MLCC · CCMC
    if flow_pos is not None:
        desired.append(flow_pos)                         # Método: flujo (conservado)
    desired.append(b[6])                                 # Método: criterios
    desired += [b[16], b[7], b[8], b[9], b[10]]          # Suministros: bases de datos · general · sub-bloque · MLCC · CCMC
    desired += [b[15], b[11], b[12], b[13]]              # Contratos/OS: bases de datos · general · CCMC · MLCC
    if pir_pos is not None:
        desired.append(pir_pos)                          # PIR (conservada)
    _reorder_slides_by_index(prs, desired)

    # Renumera los partnames de las láminas (slide1..slideN) en el orden final.
    # Necesario porque borrar y luego agregar láminas puede colisionar nombres
    # (p.ej. slide6/slide13 de las conservadas); esto los deja únicos y limpios.
    prs.part.rename_slide_parts([sldId.rId for sldId in prs.slides._sldIdLst])

    output_path.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(output_path))
    return output_path
