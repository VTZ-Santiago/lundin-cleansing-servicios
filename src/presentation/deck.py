from __future__ import annotations

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
class ComplementariaData:
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


def _load_complementaria_data(output_root: Path, operation: str) -> ComplementariaData:
    data = ComplementariaData(operation=operation)
    pattern = f"reportes/complementaria_{operation.lower()}_*.xlsx"
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


def _load_complementaria_plants(output_root: Path, operation: str) -> list[PlantRow]:
    pattern = f"reportes/complementaria_{operation.lower()}_*.xlsx"
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
        parts = [f"tipo posición {entry.position_type or 'D'}"]
        if entry.marco:
            parts.append(entry.marco)
        if entry.doc_classes:
            parts.append("clases: " + ", ".join(entry.doc_classes))
        lines.append(f"{label} — " + " · ".join(parts))

    if not lines:
        lines = [
            "Contratos — tipo posición D · con contrato marco",
            "Órdenes de Servicio — tipo posición D · sin contrato marco",
        ]
    lines.append(
        "Criterios de la hoja CASERONES (planta 8000); en CCMC la apertura usa la presencia de contrato marco. "
        "Bloque de órdenes de servicio de reparación (tipo L) excluido."
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


def _update_cover_slide(slide, generated_on: date) -> None:
    for shape in slide.shapes:
        if not getattr(shape, "has_text_frame", False):
            continue
        text = shape.text.strip()
        if text == "Data Cleansing":
            _set_shape_text(shape, "Data Cleansing Servicios")
        elif text == "30 ABRIL\xa02026":
            _set_shape_text(shape, f"{generated_on.day:02d} {_month_name_es(generated_on)} {generated_on.year}")

    subtitle = slide.shapes.add_textbox(Inches(0.95), Inches(4.65), Inches(7.2), Inches(0.72))
    subtitle_frame = subtitle.text_frame
    subtitle_frame.clear()
    subtitle_frame.word_wrap = True
    paragraph = subtitle_frame.paragraphs[0]
    run = paragraph.add_run()
    run.text = "Hallazgos vigentes sobre órdenes de compra, contratos y órdenes de servicio"
    run.font.size = Pt(18)
    run.font.color.rgb = CHARCOAL


def _add_header(slide, slide_width: int, title: str, subtitle: str | None = None) -> None:
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

    kicker = slide.shapes.add_textbox(Inches(0.55), Inches(0.07), Inches(3.25), Inches(0.22))
    kicker_frame = kicker.text_frame
    kicker_frame.clear()
    paragraph = kicker_frame.paragraphs[0]
    run = paragraph.add_run()
    run.text = "Data Cleansing | Avance General"
    run.font.size = Pt(13)
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
        label_box = slide.shapes.add_textbox(Inches(left + 0.18), Inches(current_top), Inches(1.5), Inches(0.22))
        label_frame = label_box.text_frame
        label_frame.clear()
        paragraph = label_frame.paragraphs[0]
        run = paragraph.add_run()
        run.text = label
        run.font.size = Pt(10)
        run.font.bold = True
        run.font.color.rgb = CHARCOAL

        bar_left = left + 1.6
        bar_width = width - 2.7
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


def _add_findings_slide(
    prs: Presentation,
    mlcc_summary: OperationSummary,
    ccmc_summary: OperationSummary,
    mlcc_stats: POStats,
    ccmc_stats: POStats,
    mlcc_tmp: TempOCSnapshot,
) -> None:
    slide = prs.slides.add_slide(_find_blank_layout(prs))
    total_positions = mlcc_stats.total_positions + ccmc_stats.total_positions
    total_framework = mlcc_stats.po_with_framework + ccmc_stats.po_with_framework
    total_segmented = mlcc_summary.classified_rows + ccmc_summary.classified_rows
    contracts_total = _get_segment(mlcc_summary, "Contrato").c1_rows + _get_segment(ccmc_summary, "Contrato").c1_rows
    services_total = _get_segment(mlcc_summary, "Servicio").c1_rows + _get_segment(ccmc_summary, "Servicio").c1_rows

    _add_header(
        slide,
        prs.slide_width,
        "Hallazgos principales",
        "La lectura se centra en OC y conteos absolutos; se dejan fuera las tasas de no migración de la versión anterior.",
    )
    _add_card(slide, 0.78, 2.05, 5.45, 1.75, "Posiciones de OC", _format_int(total_positions), "Cobertura combinada 2007-2026 en MLCC y CCMC", BLUE_LIGHT, BLUE)
    _add_card(slide, 6.55, 2.05, 5.95, 1.75, "OC con contrato marco", _format_int(total_framework), f"{_format_int(ccmc_stats.po_with_framework)} CCMC | {_format_int(mlcc_stats.po_with_framework)} MLCC", GREEN_LIGHT, GREEN)
    _add_card(slide, 0.78, 4.2, 5.45, 1.75, "Posiciones segmentadas", _format_int(total_segmented), f"{_format_int(contracts_total)} contratos | {_format_int(services_total)} órdenes de servicio", ORANGE_LIGHT, ORANGE)
    _add_card(slide, 6.55, 4.2, 5.95, 1.75, "Base MLCC sin posiciones D", _format_int(mlcc_tmp.total_base), f"{_format_int(mlcc_tmp.type_counts.get('Vacio', 0))} registros sin tipo de posición", PURPLE_LIGHT, PURPLE)
    _add_footnote(slide, "Fuentes: reporte estadístico de OC, resúmenes de segmentación y estadística ad hoc de OC.")


def _add_universe_slide(prs: Presentation, mlcc_stats: POStats, ccmc_stats: POStats) -> None:
    slide = prs.slides.add_slide(_find_blank_layout(prs))
    total_positions = mlcc_stats.total_positions + ccmc_stats.total_positions
    total_documents = mlcc_stats.unique_documents + ccmc_stats.unique_documents
    total_framework = mlcc_stats.po_with_framework + ccmc_stats.po_with_framework
    _add_header(
        slide,
        prs.slide_width,
        "Universo de OC analizado",
        "Cobertura por operación y períodos considerados, sin detallar archivos fuente.",
    )
    _add_card(slide, 0.78, 2.0, 2.65, 1.4, "Posiciones de OC", _format_int(total_positions), "Posiciones del universo consolidado", BLUE_LIGHT, BLUE)
    _add_card(slide, 3.58, 2.0, 2.65, 1.4, "Documentos únicos", _format_int(total_documents), "Documentos de compra distintos", GREEN_LIGHT, GREEN)
    _add_card(slide, 6.38, 2.0, 2.65, 1.4, "OC con marco", _format_int(total_framework), "Documentos con contrato marco asociado", ORANGE_LIGHT, ORANGE)
    _add_card(slide, 9.18, 2.0, 3.05, 1.4, "Período cubierto", "2007-2026", "MLCC 2010-2026 | CCMC 2007-2026", PURPLE_LIGHT, PURPLE)

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
    _add_table(slide, rows, 0.78, 3.78, 11.45, 2.55, [1.25, 3.45, 1.55, 1.55, 1.35, 1.35])
    _add_footnote(slide, "Períodos consolidados desde los reportes estadísticos oficiales de OC para MLCC y CCMC.")


def _operation_callout_lines(stats: POStats, snapshot: TempOCSnapshot) -> list[str]:
    if stats.operation == "MLCC":
        top_framework = stats.top_frameworks[0] if stats.top_frameworks else ("Sin dato", 0, 0)
        return [
            f"Estadística ad hoc sin posiciones D: {_format_int(snapshot.total_base)} registros",
            f"Con código material: {_format_int(snapshot.with_material)} | Sin código material: {_format_int(snapshot.without_material)}",
            f"Marco con mayor volumen: {top_framework[0]} con {_format_int(top_framework[1])} OC asociadas",
        ]

    top_framework = stats.top_frameworks[0] if stats.top_frameworks else ("Sin dato", 0, 0)
    plant_summary = " | ".join(
        f"{plant} {_format_int(value)}" for plant, value in sorted(snapshot.plant_counts.items(), key=lambda item: (-item[1], item[0]))[:3]
    )
    return [
        f"Distribución por planta en estadística ad hoc: {plant_summary}",
        f"Con código material: {_format_int(snapshot.with_material)} | Sin código material: {_format_int(snapshot.without_material)}",
        f"Marco con mayor volumen: {top_framework[0]} con {_format_int(top_framework[1])} OC asociadas",
    ]


def _add_operation_structure_slide(prs: Presentation, stats: POStats, snapshot: TempOCSnapshot, accent: RGBColor, fill: RGBColor) -> None:
    slide = prs.slides.add_slide(_find_blank_layout(prs))
    _add_header(
        slide,
        prs.slide_width,
        f"{stats.operation} | Estructura de OC",
        "Conteos absolutos del reporte estadístico y de la estadística ad hoc de OC.",
    )
    _add_card(slide, 0.78, 2.0, 3.2, 1.35, "Posiciones de OC", _format_int(stats.total_positions), "Universo inicial antes de segmentación", fill, accent)
    _add_card(slide, 4.15, 2.0, 3.2, 1.35, "Documentos únicos", _format_int(stats.unique_documents), "Documentos de compra distintos", GREEN_LIGHT, GREEN)
    _add_card(
        slide,
        7.52,
        2.0,
        4.7,
        1.35,
        "Contratos marco",
        _format_int(stats.framework_contracts),
        f"OC con marco {_format_int(stats.po_with_framework)} | sin marco {_format_int(stats.po_without_framework)}",
        ORANGE_LIGHT,
        ORANGE,
    )
    _add_bar_box(slide, 0.78, 3.6, 5.65, 2.35, "Clases documentales con mayor volumen", stats.top_doc_classes(5), fill, accent)
    _add_bar_box(slide, 6.6, 3.6, 5.62, 2.35, "Tipo de posición dominante", stats.top_position_types(4), GREEN_LIGHT, GREEN)
    _add_callout_box(slide, 0.78, 6.12, 11.44, 0.72, "Lo que aporta Estadística de OC.xlsx", _operation_callout_lines(stats, snapshot), GRAY_LIGHT, CHARCOAL, font_size=10)
    _add_footnote(slide, f"Fuentes: reporte_estadistico_{stats.operation}-PO.xlsx y tmp/Estadistica de OC.xlsx.")


def _segment_table_row(operation: str, label: str, segment: SegmentStats) -> list[str]:
    return [
        operation,
        label,
        _format_int(segment.c1_rows),
        _format_int(segment.c2_rows),
        _format_int(segment.c2_no_migra_rows),
        _format_pct(segment.excluded_pct),
        _format_int(segment.c1_documents),
    ]


def _add_servicios_slide(
    prs: Presentation,
    mlcc_summary: OperationSummary,
    ccmc_summary: OperationSummary,
    criteria: dict[str, SegmentCriteria],
) -> None:
    slide = prs.slides.add_slide(_find_blank_layout(prs))
    mlcc_contracts = _get_segment(mlcc_summary, "Contrato")
    mlcc_services = _get_segment(mlcc_summary, "Servicio")
    ccmc_contracts = _get_segment(ccmc_summary, "Contrato")
    ccmc_services = _get_segment(ccmc_summary, "Servicio")
    segments = [mlcc_contracts, mlcc_services, ccmc_contracts, ccmc_services]

    total_c1 = sum(segment.c1_rows for segment in segments)
    total_c2 = sum(segment.c2_rows for segment in segments)
    total_no_migra = sum(segment.c2_no_migra_rows for segment in segments)
    total_contracts = mlcc_contracts.c1_rows + ccmc_contracts.c1_rows
    total_services = mlcc_services.c1_rows + ccmc_services.c1_rows
    pct_no_migra = (total_no_migra / total_c1 * 100) if total_c1 else 0.0

    _add_header(
        slide,
        prs.slide_width,
        "CONTRATOS Y ÓRDENES DE SERVICIO — Universo tipificado como D",
        "Vigencia por cruce con los reportes de contratos vigentes (contrato marco o documento de compra) — corte 30 Jun 2026 (E01)",
    )
    _add_card(
        slide, 0.78, 2.0, 3.72, 1.3,
        "Posiciones D segmentadas",
        _format_int(total_c1),
        f"Contratos {_format_int(total_contracts)} | Órdenes de servicio {_format_int(total_services)}",
        BLUE_LIGHT, BLUE,
    )
    _add_card(
        slide, 4.72, 2.0, 3.72, 1.3,
        "C2 — migra",
        _format_int(total_c2),
        f"MLCC {_format_int(mlcc_contracts.c2_rows + mlcc_services.c2_rows)} | CCMC {_format_int(ccmc_contracts.c2_rows + ccmc_services.c2_rows)}",
        GREEN_LIGHT, GREEN,
    )
    _add_card(
        slide, 8.66, 2.0, 3.89, 1.3,
        "C2 — no migra",
        _format_int(total_no_migra),
        f"{_format_pct(pct_no_migra)} del universo D segmentado",
        RED_LIGHT, RED,
    )

    rows = [
        ["Operación", "Segmento", "Posiciones C1", "C2 migra", "C2 no migra", "% no migra", "Docs únicos"],
        _segment_table_row("MLCC", "Contratos", mlcc_contracts),
        _segment_table_row("MLCC", "Órdenes de Servicio", mlcc_services),
        _segment_table_row("CCMC", "Contratos", ccmc_contracts),
        _segment_table_row("CCMC", "Órdenes de Servicio", ccmc_services),
    ]
    _add_table(slide, rows, 0.78, 3.55, 11.44, 1.85, [1.15, 2.35, 1.65, 1.45, 1.65, 1.5, 1.69])

    _add_callout_box(
        slide, 0.78, 5.55, 11.44, 1.2,
        "Caracterización aplicada al universo D",
        _characterization_lines(criteria),
        GRAY_LIGHT, CHARCOAL, font_size=9,
    )
    _add_footnote(
        slide,
        "Fuentes: resúmenes de segmentación, Caracterización de Contratos.xlsx (bloque de reparación excluido) "
        "y reportes CONTRATOS VIGENTES (CCMC mayo 26 | MLCC abril 26).",
    )


def _add_servicios_doc_class_slide(prs: Presentation, mlcc_summary: OperationSummary, ccmc_summary: OperationSummary) -> None:
    slide = prs.slides.add_slide(_find_blank_layout(prs))
    _add_header(
        slide,
        prs.slide_width,
        "CONTRATOS Y ÓRDENES DE SERVICIO — Apertura por clase documental",
        "Posiciones C1 del universo D por clase de documento de compras y segmento.",
    )

    boxes = [
        (0.78, 2.0, 5.65, "MLCC — Contratos", _get_segment(mlcc_summary, "Contrato"), BLUE_LIGHT, BLUE),
        (6.6, 2.0, 5.62, "MLCC — Órdenes de Servicio", _get_segment(mlcc_summary, "Servicio"), BLUE_LIGHT, BLUE),
        (0.78, 4.5, 5.65, "CCMC — Contratos", _get_segment(ccmc_summary, "Contrato"), GREEN_LIGHT, GREEN),
        (6.6, 4.5, 5.62, "CCMC — Órdenes de Servicio", _get_segment(ccmc_summary, "Servicio"), GREEN_LIGHT, GREEN),
    ]
    for left, top, width, title, segment, fill, accent in boxes:
        items = sorted(segment.doc_class_counts.items(), key=lambda item: (-item[1], item[0]))[:5]
        _add_bar_box(slide, left, top, width, 2.3, title, items, fill, accent)

    _add_footnote(slide, "Fuente: resúmenes de segmentación vigentes por operación (conteo de posiciones C1).")


def _add_complementaria_slide(
    prs: Presentation,
    mlcc_data: ComplementariaData,
    ccmc_data: ComplementariaData,
) -> None:
    slide = prs.slides.add_slide(_find_blank_layout(prs))

    total_vigente = mlcc_data.vigente_rows + ccmc_data.vigente_rows
    total_vcs = mlcc_data.vcs_rows + ccmc_data.vcs_rows
    total_vss = mlcc_data.vss_rows + ccmc_data.vss_rows
    total_vigente_usd = mlcc_data.vigente_usd + ccmc_data.vigente_usd
    total_vcs_usd = mlcc_data.vcs_usd + ccmc_data.vcs_usd

    def _fmt_usd(value: float) -> str:
        if value <= 0.0:
            return "—"
        if value >= 1_000_000:
            return f"USD {value / 1_000_000:.1f}M"
        if value >= 1_000:
            return f"USD {value / 1_000:.0f}K"
        return f"USD {value:.0f}"

    _add_header(
        slide,
        prs.slide_width,
        "COMPLEMENTARIA — Universo no tipificado como D",
        "Tipos de posición C, V, K, L, P y en blanco — segmentados por vigencia y saldo pendiente (cutoff: 30 Jun 2026)",
    )

    # KPI cards
    _add_card(slide, 0.78, 2.0, 3.72, 1.3, "VIGENTE", _format_int(total_vigente), f"Deben migrar  ·  {_fmt_usd(total_vigente_usd)}", GREEN_LIGHT, GREEN)
    _add_card(slide, 4.72, 2.0, 3.72, 1.3, "NO VIGENTE c/ saldo", _format_int(total_vcs), f"No vigentes con balance  ·  {_fmt_usd(total_vcs_usd)}", ORANGE_LIGHT, ORANGE)
    _add_card(slide, 8.66, 2.0, 3.89, 1.3, "NO VIGENTE s/ saldo", _format_int(total_vss), "No vigentes sin balance — NO migran", RED_LIGHT, RED)

    # Table: per operation
    rows = [
        ["Operación", "VIGENTE", "VIGENTE USD", "NVCS filas", "NVCS USD", "NVSS filas"],
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
    _add_table(slide, rows, 0.78, 3.55, 11.44, 1.85, [1.25, 1.65, 2.0, 1.55, 2.0, 1.55])

    # Callout: document prefix breakdown
    all_prefix_labels = ["46XXXX (Marco)", "45XXXX", "44XXXX", "49XXXX", "Otros"]

    def _pfx(data: ComplementariaData, lbl: str) -> str:
        return _format_int(data.prefix_counts.get(lbl, 0))

    lines = [
        "46XXXX (Marco) — MLCC: {m46}  |  CCMC: {c46}".format(
            m46=_pfx(mlcc_data, "46XXXX (Marco)"), c46=_pfx(ccmc_data, "46XXXX (Marco)")
        ),
        (
            "45XXXX — MLCC: {m45} | CCMC: {c45}   ·   "
            "44XXXX — MLCC: {m44} | CCMC: {c44}   ·   "
            "49XXXX — MLCC: {m49} | CCMC: {c49}   ·   "
            "Otros — MLCC: {mot} | CCMC: {cot}"
        ).format(
            m45=_pfx(mlcc_data, "45XXXX"), c45=_pfx(ccmc_data, "45XXXX"),
            m44=_pfx(mlcc_data, "44XXXX"), c44=_pfx(ccmc_data, "44XXXX"),
            m49=_pfx(mlcc_data, "49XXXX"), c49=_pfx(ccmc_data, "49XXXX"),
            mot=_pfx(mlcc_data, "Otros"),  cot=_pfx(ccmc_data, "Otros"),
        ),
        (
            "Con material — MLCC: {mc} | CCMC: {cc}   ·   "
            "Sin material — MLCC: {ms} | CCMC: {cs}"
        ).format(
            mc=_format_int(mlcc_data.con_material), cc=_format_int(ccmc_data.con_material),
            ms=_format_int(mlcc_data.sin_material), cs=_format_int(ccmc_data.sin_material),
        ),
    ]
    _add_callout_box(slide, 0.78, 5.55, 11.44, 1.2, "Apertura por código de documento de compra", lines, GRAY_LIGHT, CHARCOAL, font_size=9)

    _add_footnote(slide, "Fuente: reportes complementaria más recientes en outputs/reportes/.  Cutoff vigencia: 30 Jun 2026.")


def _add_complementaria_plant_slide(prs: Presentation, operation: str, plant_rows: list[PlantRow]) -> None:
    slide = prs.slides.add_slide(_find_blank_layout(prs))
    _add_header(
        slide,
        prs.slide_width,
        f"{operation} | Complementaria — Apertura por Planta",
        "Posiciones no tipificadas como D, segmentadas por vigencia y saldo pendiente (cutoff: 30 Jun 2026)",
    )

    def _fmt_usd(value: float) -> str:
        if value <= 0.0:
            return "—"
        if value >= 1_000_000:
            return f"USD {value / 1_000_000:.1f}M"
        if value >= 1_000:
            return f"USD {value / 1_000:.0f}K"
        return f"USD {value:.0f}"

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

    for c_idx, hdr_text in enumerate(["Planta", "VIGENTE", "VIGENTE USD", "NVCS filas", "NVCS USD", "NVSS filas"]):
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

    _add_footnote(slide, "Fuente: reporte complementaria más reciente en outputs/reportes/.  Cutoff vigencia: 30 Jun 2026.")


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
    tmp_stats_path = project_root / "tmp" / "Estadistica de OC.xlsx"

    prs = Presentation(str(template_path))
    mlcc_summary = parse_summary_markdown(mlcc_summary_path)
    ccmc_summary = parse_summary_markdown(ccmc_summary_path)
    mlcc_stats = _load_po_stats(mlcc_stats_path, "MLCC")
    ccmc_stats = _load_po_stats(ccmc_stats_path, "CCMC")
    mlcc_tmp = _load_temp_snapshot(tmp_stats_path, "MLCC")
    ccmc_tmp = _load_temp_snapshot(tmp_stats_path, "CCMC")
    mlcc_comp = _load_complementaria_data(output_root, "MLCC")
    ccmc_comp = _load_complementaria_data(output_root, "CCMC")
    mlcc_plants = _load_complementaria_plants(output_root, "MLCC")
    ccmc_plants = _load_complementaria_plants(output_root, "CCMC")
    characterization = _load_characterization_criteria(project_root)

    for index in range(len(prs.slides) - 1, 0, -1):
        _delete_slide(prs, index)

    _update_cover_slide(prs.slides[0], generated_on)
    _add_findings_slide(prs, mlcc_summary, ccmc_summary, mlcc_stats, ccmc_stats, mlcc_tmp)
    _add_universe_slide(prs, mlcc_stats, ccmc_stats)
    _add_operation_structure_slide(prs, mlcc_stats, mlcc_tmp, BLUE, BLUE_LIGHT)
    _add_operation_structure_slide(prs, ccmc_stats, ccmc_tmp, GREEN, GREEN_LIGHT)
    _add_servicios_slide(prs, mlcc_summary, ccmc_summary, characterization)
    _add_servicios_doc_class_slide(prs, mlcc_summary, ccmc_summary)
    _add_complementaria_slide(prs, mlcc_comp, ccmc_comp)
    _add_complementaria_plant_slide(prs, "MLCC", mlcc_plants)
    _add_complementaria_plant_slide(prs, "CCMC", ccmc_plants)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(output_path))
    return output_path
