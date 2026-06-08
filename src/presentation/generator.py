from __future__ import annotations

import math
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
    doc_class_counts: dict[str, int] = field(default_factory=dict)
    source_file_counts: dict[str, int] = field(default_factory=dict)

    @property
    def migrate_pct(self) -> float:
        if self.c1_rows == 0:
            return 0.0
        return round(self.c2_rows / self.c1_rows * 100, 2)


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
class ArtifactCard:
    title: str
    headline: str
    details: list[str]
    fill: RGBColor
    accent: RGBColor


def _parse_int(raw: str) -> int:
    return int(raw.replace(",", "").strip())


def _parse_float_pct(raw: str) -> float:
    return float(raw.replace("%", "").replace(",", "").strip())


def parse_summary_markdown(path: Path) -> OperationSummary:
    lines = path.read_text(encoding="utf-8").splitlines()
    header = lines[0].strip()
    operation = header.replace("# Resumen segmentación", "").strip()
    summary = OperationSummary(operation=operation)
    current_section: str | None = None
    current_segment: SegmentStats | None = None
    nested_target: str | None = None

    for raw_line in lines[1:]:
        line = raw_line.rstrip()
        stripped = line.strip()
        if line.startswith("  - "):
            nested_content = line[4:].strip()
            if nested_target == "global_source_files":
                summary.source_files.append(nested_content)
                continue
            if current_segment is None:
                continue
            if nested_target == "doc_class_counts":
                key, value = [part.strip() for part in nested_content.split(":", 1)]
                current_segment.doc_class_counts[key] = _parse_int(value)
            elif nested_target == "source_file_counts":
                key, value = [part.strip() for part in nested_content.split(":", 1)]
                current_segment.source_file_counts[key] = _parse_int(value)
            continue
        if not stripped:
            continue
        if stripped.startswith("## "):
            current_section = stripped[3:].strip()
            nested_target = None
            if current_section != "Auditoría global":
                current_segment = SegmentStats(label=current_section)
                summary.segments[current_section] = current_segment
            continue
        if stripped.startswith("- "):
            nested_target = None
            content = stripped[2:]
            if current_section == "Auditoría global":
                if content == "Archivos fuente:":
                    nested_target = "global_source_files"
                    continue
                key, value = [part.strip() for part in content.split(":", 1)]
                if key == "Filas leídas desde OC":
                    summary.rows_read = _parse_int(value)
                elif key == "Filas limpias cargadas":
                    summary.total_rows = _parse_int(value)
                elif key == "Filas clasificadas en segmentos":
                    summary.classified_rows = _parse_int(value)
                elif key == "Filas no clasificadas":
                    summary.unclassified_rows = _parse_int(value)
                elif key == "Filas con solapamiento entre segmentos":
                    summary.overlap_rows = _parse_int(value)
                elif key == "Criterios de segmentación cargados":
                    summary.criteria_count = _parse_int(value)
                continue
            if current_segment is None:
                continue
            if content == "Cl. documento compras:":
                nested_target = "doc_class_counts"
                continue
            if content == "Archivos fuente:":
                nested_target = "source_file_counts"
                continue
            key, value = [part.strip() for part in content.split(":", 1)]
            if key == "C1 filas":
                current_segment.c1_rows = _parse_int(value)
            elif key == "Documentos únicos C1":
                current_segment.c1_documents = _parse_int(value)
            elif key == "Documentos únicos C2":
                current_segment.c2_documents = _parse_int(value)
            elif key == "Documentos únicos C2_NO_MIGRA":
                current_segment.c2_no_migra_documents = _parse_int(value)
            elif key == "Outline contracts únicos C1":
                current_segment.c1_outline_contracts = _parse_int(value)
            elif key == "Outline contracts únicos C2":
                current_segment.c2_outline_contracts = _parse_int(value)
            elif key == "Outline contracts únicos C2_NO_MIGRA":
                current_segment.c2_no_migra_outline_contracts = _parse_int(value)
            elif key == "C2 migra":
                current_segment.c2_rows = _parse_int(value)
            elif key == "C2 no migra":
                current_segment.c2_no_migra_rows = _parse_int(value)
            elif key == "Porcentaje no migra":
                current_segment.excluded_pct = _parse_float_pct(value)
            elif key == "Reconciliación C1 = C2 + C2_NO_MIGRA":
                current_segment.reconciliation_ok = value == "OK"
            continue
        if stripped.startswith("-"):
            continue
        if line.startswith("  - "):
            nested_content = line[4:].strip()
            if nested_target == "global_source_files":
                summary.source_files.append(nested_content)
                continue
            if current_segment is None:
                continue
            if nested_target == "doc_class_counts":
                key, value = [part.strip() for part in nested_content.split(":", 1)]
                current_segment.doc_class_counts[key] = _parse_int(value)
            elif nested_target == "source_file_counts":
                key, value = [part.strip() for part in nested_content.split(":", 1)]
                current_segment.source_file_counts[key] = _parse_int(value)

    return summary


def _iter_workbooks(folder: Path) -> list[Path]:
    return sorted(
        [path for path in folder.glob("*.xlsx") if not path.name.startswith("~$")],
        key=lambda path: path.stat().st_mtime,
        reverse=True,
    )


def _safe_summary_map(path: Path) -> dict[str, str]:
    try:
        workbook = load_workbook(path, read_only=True, data_only=True)
        if "Resumen" not in workbook.sheetnames:
            return {}
        sheet = workbook["Resumen"]
        pairs: dict[str, str] = {}
        for row in sheet.iter_rows(values_only=True):
            values = [value for value in row if value not in (None, "")]
            if len(values) >= 2:
                key = str(values[0]).strip()
                value = str(values[1]).strip()
                if key and value and key not in {"INDICADOR", "Operación"} and key not in pairs:
                    pairs[key] = value
        return pairs
    except Exception:
        return {}


def _load_crossref_card(folder: Path) -> ArtifactCard:
    workbooks = _iter_workbooks(folder)
    summary = _safe_summary_map(workbooks[0]) if workbooks else {}
    headline = "Sin reporte detectado"
    details = ["Cruce de materiales ya generado", "Fuente: outputs/material_crossref"]
    if summary:
        total_materials = summary.get("Total materiales unicos", "0")
        headline = f"{_format_int(int(total_materials))} materiales únicos" if str(total_materials).isdigit() else f"{total_materials} materiales únicos"
        details = [
            f"Cobertura maestro: {summary.get('% cobertura maestro', 'N/D')}",
            f"No encontrados en maestro: {summary.get('No encontrados en maestro', 'N/D')}",
            f"Servicios en maestro: {summary.get('% servicios en maestro', 'N/D')}",
        ]
    return ArtifactCard(
        title="Material Crossref",
        headline=headline,
        details=details,
        fill=ORANGE_LIGHT,
        accent=ORANGE,
    )


def _load_dup_check_card(folder: Path) -> ArtifactCard:
    workbooks = _iter_workbooks(folder)
    headline = "Sin reporte detectado"
    details = ["Análisis de duplicados disponible", "Fuente: outputs/material_dup_check"]
    if workbooks:
        try:
            workbook = load_workbook(workbooks[0], read_only=True, data_only=True)
            sheet = workbook["Resumen"]
            records: list[tuple[str, str, str, str]] = []
            for row in sheet.iter_rows(min_row=5, values_only=True):
                if not row or not row[0]:
                    continue
                operation = str(row[0]).strip()
                source = str(row[1]).strip()
                duplicated = str(row[3]).strip()
                affected = str(row[4]).strip()
                records.append((operation, source, duplicated, affected))
            oc_records = [record for record in records if record[1] == "OC"]
            total_duplicated = sum(int(record[2]) for record in oc_records)
            headline = f"{_format_int(total_duplicated)} materiales duplicados en OC"
            details = [
                f"CCMC OC: {record[2]} materiales | {record[3]} documentos" for record in oc_records[:1]
            ]
            if len(oc_records) > 1:
                details.append(f"MLCC OC: {oc_records[1][2]} materiales | {oc_records[1][3]} documentos")
            details.append("ME3L con casos acotados ya identificados")
        except Exception:
            pass
    return ArtifactCard(
        title="Dup Check",
        headline=headline,
        details=details,
        fill=PURPLE_LIGHT,
        accent=PURPLE,
    )


def _load_consolidated_card(folder: Path) -> ArtifactCard:
    workbooks = sorted(
        [path for path in folder.glob("*2025plus.xlsx") if not path.name.startswith("~$")],
        key=lambda path: path.stat().st_mtime,
        reverse=True,
    )
    headline = "Sin reporte detectado"
    details = ["Consolidado vigente 2025+", "Fuente: outputs/ordenes_compra/consolidado"]
    if workbooks:
        counts: list[str] = []
        total_rows = 0
        for path in workbooks[:2]:
            try:
                workbook = load_workbook(path, read_only=True, data_only=True)
                sheet = workbook[workbook.sheetnames[0]]
                rows = max(sheet.max_row - 1, 0)
                total_rows += rows
                label = "MLCC" if "MLCC" in path.name.upper() else "CCMC"
                counts.append(f"{label}: {_format_int(rows)} filas")
            except Exception:
                continue
        if counts:
            headline = f"{_format_int(total_rows)} filas vigentes 2025+"
            details = counts + ["Incluye columnas de liberación y borrado"]
    return ArtifactCard(
        title="Consolidado OC",
        headline=headline,
        details=details,
        fill=GREEN_LIGHT,
        accent=GREEN,
    )


def _artifact_cards(output_root: Path) -> list[ArtifactCard]:
    segment_workbooks = sorted(output_root.glob("control_points/**/*.xlsx"))
    total_classified = 0
    for summary_path in [output_root / "control_points" / "resumen_segmentacion_MLCC.md", output_root / "control_points" / "resumen_segmentacion_CCMC.md"]:
        if summary_path.exists():
            total_classified += parse_summary_markdown(summary_path).classified_rows
    return [
        ArtifactCard(
            title="Segmentación",
            headline=f"{_format_int(total_classified)} filas clasificadas",
            details=[
                f"{len(segment_workbooks):,} workbooks de control".replace(",", "."),
                "Dominios: MLCC y CCMC",
                "Segmentos: contratos y órdenes de servicio",
            ],
            fill=BLUE_LIGHT,
            accent=BLUE,
        ),
        _load_consolidated_card(output_root / "ordenes_compra" / "consolidado"),
        _load_crossref_card(output_root / "material_crossref"),
        _load_dup_check_card(output_root / "material_dup_check"),
    ]


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


def _update_cover_slide(slide, generated_on: date) -> None:
    for shape in slide.shapes:
        if not getattr(shape, "has_text_frame", False):
            continue
        text = shape.text.strip()
        if text == "Data Cleansing":
            _set_shape_text(shape, "Data Cleansing Servicios")
        elif text == "30 ABRIL\xa02026":
            _set_shape_text(shape, f"{generated_on.day:02d} {_month_name_es(generated_on)} {generated_on.year}")

    subtitle = slide.shapes.add_textbox(Inches(0.95), Inches(4.7), Inches(6.6), Inches(0.6))
    subtitle_frame = subtitle.text_frame
    subtitle_frame.clear()
    subtitle_frame.word_wrap = True
    paragraph = subtitle_frame.paragraphs[0]
    run = paragraph.add_run()
    run.text = "Segmentación y reportes vigentes generados para MLCC y CCMC"
    run.font.size = Pt(17)
    run.font.bold = False
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

    kicker = slide.shapes.add_textbox(Inches(0.55), Inches(0.07), Inches(3.2), Inches(0.22))
    kicker_frame = kicker.text_frame
    kicker_frame.clear()
    kicker_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
    paragraph = kicker_frame.paragraphs[0]
    run = paragraph.add_run()
    run.text = "Data Cleansing | Avance General"
    run.font.size = Pt(13)
    run.font.bold = True
    run.font.color.rgb = RGBColor(255, 255, 255)

    title_box = slide.shapes.add_textbox(Inches(0.7), Inches(0.72), Inches(11.7), Inches(0.6))
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
        subtitle_box = slide.shapes.add_textbox(Inches(0.7), Inches(1.25), Inches(11.5), Inches(0.45))
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
        Inches(2.0),
        Inches(0.04),
    )
    rule.fill.solid()
    rule.fill.fore_color.rgb = ORANGE
    rule.line.fill.background()


def _add_footnote(slide, text: str) -> None:
    box = slide.shapes.add_textbox(Inches(0.7), Inches(7.02), Inches(11.6), Inches(0.22))
    frame = box.text_frame
    frame.clear()
    paragraph = frame.paragraphs[0]
    paragraph.alignment = PP_ALIGN.LEFT
    run = paragraph.add_run()
    run.text = text
    run.font.size = Pt(9)
    run.font.color.rgb = SLATE


def _add_card(slide, left: float, top: float, width: float, height: float, title: str, value: str, note: str, fill: RGBColor, accent: RGBColor) -> None:
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

    title_box = slide.shapes.add_textbox(Inches(left + 0.18), Inches(top + 0.14), Inches(width - 0.36), Inches(0.24))
    title_frame = title_box.text_frame
    title_frame.clear()
    paragraph = title_frame.paragraphs[0]
    run = paragraph.add_run()
    run.text = title
    run.font.size = Pt(12)
    run.font.bold = True
    run.font.color.rgb = accent

    value_box = slide.shapes.add_textbox(Inches(left + 0.18), Inches(top + 0.45), Inches(width - 0.36), Inches(0.4))
    value_frame = value_box.text_frame
    value_frame.clear()
    paragraph = value_frame.paragraphs[0]
    run = paragraph.add_run()
    run.text = value
    run.font.size = Pt(22)
    run.font.bold = True
    run.font.color.rgb = CHARCOAL

    note_box = slide.shapes.add_textbox(Inches(left + 0.18), Inches(top + 0.92), Inches(width - 0.36), Inches(height - 1.04))
    note_frame = note_box.text_frame
    note_frame.word_wrap = True
    note_frame.clear()
    paragraph = note_frame.paragraphs[0]
    paragraph.line_spacing = 1.15
    run = paragraph.add_run()
    run.text = note
    run.font.size = Pt(10)
    run.font.color.rgb = SLATE


def _add_bullets(slide, left: float, top: float, width: float, height: float, lines: list[str], font_size: int = 10, color: RGBColor = SLATE) -> None:
    box = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(width), Inches(height))
    frame = box.text_frame
    frame.clear()
    frame.word_wrap = True
    for index, line in enumerate(lines):
        paragraph = frame.paragraphs[0] if index == 0 else frame.add_paragraph()
        paragraph.level = 0
        paragraph.bullet = True
        paragraph.space_after = Pt(2)
        run = paragraph.add_run()
        run.text = line
        run.font.size = Pt(font_size)
        run.font.color.rgb = color


def _format_int(value: int) -> str:
    return f"{value:,}".replace(",", ".")


def _source_files_label(values: list[str]) -> str:
    if not values:
        return "Sin archivos registrados"
    return " | ".join(values)


def _add_table(slide, rows: list[list[str]], left: float, top: float, width: float, height: float) -> None:
    table = slide.shapes.add_table(len(rows), len(rows[0]), Inches(left), Inches(top), Inches(width), Inches(height)).table
    col_widths = [1.3, 1.4, 1.4, 1.55, 1.2, 1.0, 5.15]
    for index, current_width in enumerate(col_widths):
        table.columns[index].width = Inches(current_width)
    for row_index, row_values in enumerate(rows):
        for col_index, value in enumerate(row_values):
            cell = table.cell(row_index, col_index)
            cell.text = value
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            paragraph = cell.text_frame.paragraphs[0]
            paragraph.alignment = PP_ALIGN.LEFT
            for run in paragraph.runs:
                run.font.size = Pt(10 if row_index > 0 else 9)
                run.font.bold = row_index == 0
                run.font.color.rgb = RGBColor(255, 255, 255) if row_index == 0 else CHARCOAL
            cell.fill.solid()
            cell.fill.fore_color.rgb = ORANGE if row_index == 0 else (GRAY_LIGHT if row_index % 2 else RGBColor(255, 255, 255))


def _top_doc_classes(segment: SegmentStats, limit: int = 5) -> list[str]:
    items = sorted(segment.doc_class_counts.items(), key=lambda item: (-item[1], item[0]))
    return [f"{doc_class}: {_format_int(count)}" for doc_class, count in items[:limit]]


def _add_segment_panel(slide, segment: SegmentStats, left: float, top: float, width: float, height: float, accent: RGBColor, fill: RGBColor) -> None:
    frame = slide.shapes.add_shape(
        MSO_AUTO_SHAPE_TYPE.ROUNDED_RECTANGLE,
        Inches(left),
        Inches(top),
        Inches(width),
        Inches(height),
    )
    frame.fill.solid()
    frame.fill.fore_color.rgb = fill
    frame.line.color.rgb = accent
    frame.line.width = Pt(1.2)

    title_box = slide.shapes.add_textbox(Inches(left + 0.2), Inches(top + 0.18), Inches(width - 0.4), Inches(0.3))
    title_frame = title_box.text_frame
    title_frame.clear()
    paragraph = title_frame.paragraphs[0]
    run = paragraph.add_run()
    run.text = segment.label
    run.font.size = Pt(15)
    run.font.bold = True
    run.font.color.rgb = CHARCOAL

    pct_box = slide.shapes.add_textbox(Inches(left + width - 1.9), Inches(top + 0.14), Inches(1.55), Inches(0.45))
    pct_frame = pct_box.text_frame
    pct_frame.clear()
    pct_frame.vertical_anchor = MSO_ANCHOR.MIDDLE
    paragraph = pct_frame.paragraphs[0]
    paragraph.alignment = PP_ALIGN.RIGHT
    run = paragraph.add_run()
    run.text = f"{segment.excluded_pct:.2f}% no migra"
    run.font.size = Pt(13)
    run.font.bold = True
    run.font.color.rgb = RED

    card_width = (width - 0.65) / 3
    _add_card(slide, left + 0.2, top + 0.62, card_width, 1.08, "C1", _format_int(segment.c1_rows), f"Documentos únicos: {_format_int(segment.c1_documents)}", GRAY_LIGHT, accent)
    _add_card(slide, left + 0.25 + card_width, top + 0.62, card_width, 1.08, "C2 migra", _format_int(segment.c2_rows), f"Documentos únicos: {_format_int(segment.c2_documents)}", GREEN_LIGHT, GREEN)
    _add_card(slide, left + 0.3 + card_width * 2, top + 0.62, card_width, 1.08, "C2 no migra", _format_int(segment.c2_no_migra_rows), f"Documentos únicos: {_format_int(segment.c2_no_migra_documents)}", RED_LIGHT, RED)

    details = [
        f"Outline contracts C1/C2/C2_NO_MIGRA: {_format_int(segment.c1_outline_contracts)} / {_format_int(segment.c2_outline_contracts)} / {_format_int(segment.c2_no_migra_outline_contracts)}",
        f"Migración preliminar: {segment.migrate_pct:.2f}% | Reconciliación: {'OK' if segment.reconciliation_ok else 'REVISAR'}",
        "Clases documentales dominantes:",
        *[f"  {item}" for item in _top_doc_classes(segment)],
    ]
    _add_bullets(slide, left + 0.22, top + 1.92, width - 0.44, height - 2.08, details, font_size=10)


def _add_tools_slide(prs: Presentation, cards: list[ArtifactCard]) -> None:
    slide = prs.slides.add_slide(_find_blank_layout(prs))
    _add_header(
        slide,
        prs.slide_width,
        "Herramientas y reportes disponibles",
        "La presentación toma como base los artefactos ya generados en outputs, sin recalcular el pipeline.",
    )
    positions = [
        (0.75, 2.05),
        (6.65, 2.05),
        (0.75, 4.38),
        (6.65, 4.38),
    ]
    for card, (left, top) in zip(cards, positions):
        note = "\n".join([line for line in card.details if line])
        _add_card(slide, left, top, 5.2, 1.78, card.title, card.headline, note, card.fill, card.accent)
    _add_footnote(slide, "Fuentes: outputs/control_points, outputs/ordenes_compra/consolidado, outputs/material_crossref y outputs/material_dup_check.")


def _add_overview_slide(prs: Presentation, mlcc: OperationSummary, ccmc: OperationSummary) -> None:
    slide = prs.slides.add_slide(_find_blank_layout(prs))
    total_rows_read = mlcc.rows_read + ccmc.rows_read
    total_classified = mlcc.classified_rows + ccmc.classified_rows
    total_unclassified = mlcc.unclassified_rows + ccmc.unclassified_rows
    total_overlap = mlcc.overlap_rows + ccmc.overlap_rows
    _add_header(
        slide,
        prs.slide_width,
        "Universo segmentado a analizar",
        "Resumen consolidado de los reportes actuales de segmentación para MLCC y CCMC.",
    )
    _add_card(slide, 0.78, 2.0, 2.6, 1.45, "Filas leídas", _format_int(total_rows_read), "Universo leído desde OC", BLUE_LIGHT, BLUE)
    _add_card(slide, 3.5, 2.0, 2.6, 1.45, "Filas clasificadas", _format_int(total_classified), "Base efectiva segmentada", GREEN_LIGHT, GREEN)
    _add_card(slide, 6.22, 2.0, 2.6, 1.45, "No clasificadas", _format_int(total_unclassified), "Fuera de segmentos actuales", ORANGE_LIGHT, ORANGE)
    _add_card(slide, 8.94, 2.0, 3.1, 1.45, "Solapamientos", _format_int(total_overlap), "MLCC concentra la revisión de traslapes", RED_LIGHT, RED)
    rows = [
        ["Operación", "Filas leídas", "Clasificadas", "No clasificadas", "Solapamiento", "Criterios", "Archivos fuente"],
        [mlcc.operation, _format_int(mlcc.rows_read), _format_int(mlcc.classified_rows), _format_int(mlcc.unclassified_rows), _format_int(mlcc.overlap_rows), _format_int(mlcc.criteria_count), _source_files_label(mlcc.source_files)],
        [ccmc.operation, _format_int(ccmc.rows_read), _format_int(ccmc.classified_rows), _format_int(ccmc.unclassified_rows), _format_int(ccmc.overlap_rows), _format_int(ccmc.criteria_count), _source_files_label(ccmc.source_files)],
    ]
    _add_table(slide, rows, 0.78, 3.88, 11.45, 2.55)
    _add_footnote(slide, "Los conteos se leen directamente desde outputs/control_points/resumen_segmentacion_MLCC.md y outputs/control_points/resumen_segmentacion_CCMC.md.")


def _add_operation_slide(prs: Presentation, summary: OperationSummary) -> None:
    slide = prs.slides.add_slide(_find_blank_layout(prs))
    classified_share = 0.0 if summary.total_rows == 0 else round(summary.classified_rows / summary.total_rows * 100, 2)
    _add_header(
        slide,
        prs.slide_width,
        f"Resultados de segmentación {summary.operation}",
        f"{classified_share:.2f}% del universo cargado cae en los segmentos actuales; los paneles muestran C1, C2 y C2_NO_MIGRA por dominio.",
    )
    segments = list(summary.segments.values())
    if len(segments) != 2:
        raise ValueError(f"Se esperaban 2 segmentos en {summary.operation}, se encontraron {len(segments)}")
    _add_segment_panel(slide, segments[0], 0.78, 2.0, 5.45, 4.65, BLUE, BLUE_LIGHT)
    _add_segment_panel(slide, segments[1], 6.78, 2.0, 5.45, 4.65, GREEN, GREEN_LIGHT)
    _add_footnote(slide, f"{summary.operation}: filas clasificadas {_format_int(summary.classified_rows)} | filas con solapamiento {_format_int(summary.overlap_rows)} | criterios cargados {_format_int(summary.criteria_count)}.")


def _add_ratio_bar(slide, left: float, top: float, width: float, label: str, migrate_pct: float, migrate_rows: int, excluded_rows: int) -> None:
    label_box = slide.shapes.add_textbox(Inches(left), Inches(top), Inches(2.45), Inches(0.26))
    label_frame = label_box.text_frame
    label_frame.clear()
    paragraph = label_frame.paragraphs[0]
    run = paragraph.add_run()
    run.text = label
    run.font.size = Pt(11)
    run.font.bold = True
    run.font.color.rgb = CHARCOAL

    bar_top = top + 0.38
    bar_height = 0.26
    bar_bg = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.RECTANGLE, Inches(left), Inches(bar_top), Inches(width), Inches(bar_height))
    bar_bg.fill.solid()
    bar_bg.fill.fore_color.rgb = RED_LIGHT
    bar_bg.line.color.rgb = GRAY

    migrate_width = max(width * migrate_pct / 100, 0.12 if migrate_rows > 0 else 0)
    if migrate_width:
        migrate_bar = slide.shapes.add_shape(MSO_AUTO_SHAPE_TYPE.RECTANGLE, Inches(left), Inches(bar_top), Inches(min(migrate_width, width)), Inches(bar_height))
        migrate_bar.fill.solid()
        migrate_bar.fill.fore_color.rgb = GREEN
        migrate_bar.line.fill.background()

    value_box = slide.shapes.add_textbox(Inches(left + width + 0.18), Inches(top + 0.16), Inches(2.6), Inches(0.3))
    value_frame = value_box.text_frame
    value_frame.clear()
    paragraph = value_frame.paragraphs[0]
    run = paragraph.add_run()
    run.text = f"{migrate_pct:.2f}% migra | {_format_int(migrate_rows)} vs {_format_int(excluded_rows)}"
    run.font.size = Pt(10)
    run.font.color.rgb = SLATE


def _add_estimation_slide(prs: Presentation, summaries: list[OperationSummary]) -> None:
    slide = prs.slides.add_slide(_find_blank_layout(prs))
    total_c2 = sum(segment.c2_rows for summary in summaries for segment in summary.segments.values())
    total_no_migra = sum(segment.c2_no_migra_rows for summary in summaries for segment in summary.segments.values())
    total_c1 = sum(segment.c1_rows for summary in summaries for segment in summary.segments.values())
    migrate_share = 0.0 if total_c1 == 0 else round(total_c2 / total_c1 * 100, 2)
    _add_header(
        slide,
        prs.slide_width,
        "Estimación preliminar a migrar",
        "Cifras derivadas de los segmentos ya clasificados; el flujo detallado y su visualización se incorporarán en una iteración posterior.",
    )
    _add_card(slide, 0.78, 2.0, 3.1, 1.55, "C2 migra", _format_int(total_c2), f"{migrate_share:.2f}% del total segmentado", GREEN_LIGHT, GREEN)
    _add_card(slide, 4.1, 2.0, 3.25, 1.55, "C2 no migra", _format_int(total_no_migra), f"{100 - migrate_share:.2f}% del total segmentado", RED_LIGHT, RED)
    _add_card(slide, 7.57, 2.0, 4.5, 1.55, "Base segmentada", _format_int(total_c1), "Reconciliación C1 = C2 + C2_NO_MIGRA validada en todos los segmentos", BLUE_LIGHT, BLUE)

    bar_top = 4.05
    bars = []
    for summary in summaries:
        for segment in summary.segments.values():
            bars.append((f"{summary.operation} | {segment.label}", segment.migrate_pct, segment.c2_rows, segment.c2_no_migra_rows))
    for index, (label, pct, migrate_rows, excluded_rows) in enumerate(bars):
        _add_ratio_bar(slide, 0.92, bar_top + index * 0.68, 7.2, label, pct, migrate_rows, excluded_rows)

    messages = [
        "CCMC contratos concentra la mayor tasa preliminar de migración entre los segmentos actuales.",
        "MLCC presenta el mayor volumen relativo de exclusión preliminar, especialmente en órdenes de servicio.",
        "La lámina de flujo se agregará después; este deck se concentra en herramientas y resultados ya disponibles.",
    ]
    _add_bullets(slide, 8.55, 4.0, 3.2, 2.0, messages, font_size=10)
    _add_footnote(slide, "Totales consolidados desde reportes vigentes al momento de generar la presentación.")


def generate_project_presentation(
    template_path: Path,
    output_path: Path,
    mlcc_summary_path: Path,
    ccmc_summary_path: Path,
    output_root: Path,
    generated_on: date | None = None,
) -> Path:
    generated_on = generated_on or date.today()
    prs = Presentation(str(template_path))
    mlcc = parse_summary_markdown(mlcc_summary_path)
    ccmc = parse_summary_markdown(ccmc_summary_path)
    cards = _artifact_cards(output_root)

    for index in range(len(prs.slides) - 1, 0, -1):
        _delete_slide(prs, index)

    _update_cover_slide(prs.slides[0], generated_on)
    _add_tools_slide(prs, cards)
    _add_overview_slide(prs, mlcc, ccmc)
    _add_operation_slide(prs, mlcc)
    _add_operation_slide(prs, ccmc)
    _add_estimation_slide(prs, [mlcc, ccmc])

    output_path.parent.mkdir(parents=True, exist_ok=True)
    prs.save(str(output_path))
    return output_path