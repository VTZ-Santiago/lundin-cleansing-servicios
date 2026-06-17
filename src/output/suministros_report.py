"""Generate SUMINISTROS summary report: posiciones no-D por categoría de
migración, sub-bloque y planta.

Sheet layout per segment Excel:
  Resumen              — pivot planta × categoría: filas, docs, valor USD
  Por_Subsegmento      — pivot sub-bloque × categoría: filas, docs, valor USD
  Por_Tipo_Posicion    — pivot (planta × categoría × tipo posición): filas, docs, USD
  Apertura_Codigo      — pivot categoría × prefijo documento + presencia material
  VIGENTE              — detalle, ordenado por planta
  NO_VIG_SIN_SALDO     — detalle; estos NO migran
  SALDO_PEND_{YYYY}    — una hoja por año de entrega para NO_VIGENTE_CON_SALDO (migran)
"""
from __future__ import annotations

from datetime import datetime
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from src.segmentacion.supply_segments import (
    SUPPLY_SEGMENT_COLUMN,
    SUPPLY_SEGMENT_ORDER,
    classify_supply_segments,
)


_DARK = "1E3A5F"
_WHITE = "FFFFFF"
_GREEN = "C6EFCE"
_RED = "FFC7CE"
_ORANGE = "FFEB9C"
_GREY = "D9D9D9"

_KNOWN_PREFIXES: dict[str, str] = {
    "46": "46XXXX (Marco)",
    "45": "45XXXX",
    "44": "44XXXX",
    "49": "49XXXX",
}

CATEGORY_VIGENTE = "VIGENTE"
CATEGORY_NO_VIGENTE_CON_SALDO = "NO_VIGENTE_CON_SALDO"
CATEGORY_NO_VIGENTE_SIN_SALDO = "NO_VIGENTE_SIN_SALDO"

_CATEGORY_FILL = {
    CATEGORY_VIGENTE: _GREEN,
    CATEGORY_NO_VIGENTE_CON_SALDO: _ORANGE,
    CATEGORY_NO_VIGENTE_SIN_SALDO: _RED,
}

_DETAIL_COLUMNS = [
    "plant_code",
    "purchase_document",
    "position",
    "short_text",
    "position_type",
    SUPPLY_SEGMENT_COLUMN,
    "purchase_doc_class",
    "purchase_group",
    "vendor",
    "pending_planned_qty",
    "pending_delivery_value",
    "delivery_date",
    "validity_end",
    "delivery_year",
    "migration_category",
    "deletion_flag",
    "outline_contract",
]


def _hdr_font(color: str = _WHITE) -> Font:
    return Font(bold=True, color=color)


def _fill(color: str) -> PatternFill:
    return PatternFill(start_color=color, end_color=color, fill_type="solid")


def _write_sheet(ws, df: pd.DataFrame) -> None:
    """Write dataframe to worksheet with dark frozen header + autofilter."""
    for col_idx, header in enumerate(df.columns, 1):
        cell = ws.cell(row=1, column=col_idx, value=header)
        cell.font = _hdr_font()
        cell.fill = _fill(_DARK)
        cell.alignment = Alignment(wrap_text=False)

    ws.freeze_panes = "A2"

    display = df.copy()
    for col in display.columns:
        if pd.api.types.is_datetime64_any_dtype(display[col]):
            display[col] = display[col].apply(lambda x: x.date() if pd.notna(x) else None)

    # Convert Int64 nullable int to plain int/None for openpyxl
    for col in display.columns:
        if hasattr(display[col], "dtype") and str(display[col].dtype) == "Int64":
            display[col] = display[col].astype(object).where(display[col].notna(), other=None)

    display = display.astype(object).where(pd.notnull(display), other=None)

    for row_data in display.values.tolist():
        ws.append(row_data)

    if not df.empty:
        ws.auto_filter.ref = ws.dimensions

    for col_cells in ws.columns:
        width = 10
        for cell in col_cells:
            if cell.value is not None:
                width = max(width, min(len(str(cell.value)) + 2, 45))
        ws.column_dimensions[get_column_letter(col_cells[0].column)].width = width


_LABEL_SIN_TIPO = "(sin tipo)"


def _display_pos_type(series: pd.Series) -> pd.Series:
    """Normalize blank/NaN position_type to '(sin tipo)' for display."""
    return series.fillna("").astype(str).str.strip().replace("", _LABEL_SIN_TIPO)


def _doc_prefix_label(series: pd.Series) -> pd.Series:
    """Map first 2 chars of purchase_document to a group label."""
    prefix = series.fillna("").astype(str).str[:2]
    return prefix.map(lambda p: _KNOWN_PREFIXES.get(p, "Otros"))


def _apply_usd_column(
    df: pd.DataFrame,
    value_col: str,
    usd_rates: dict[str, float],
) -> str | None:
    """Add {value_col}_usd to df in-place. Returns column name or None if rates unavailable."""
    if not usd_rates or value_col not in df.columns or "currency" not in df.columns:
        return None
    usd_col = f"{value_col}_usd"
    rates = df["currency"].fillna("").astype(str).str.strip().str.upper().map(
        lambda c: usd_rates.get(c, 0.0)
    )
    df[usd_col] = pd.to_numeric(df[value_col], errors="coerce").fillna(0.0) * rates
    return usd_col


def _build_subsegment_breakdown(
    df: pd.DataFrame,
    cat_col: str,
    value_col: str,
    usd_col: str | None = None,
) -> pd.DataFrame:
    """Pivot: (sub-bloque × categoría) con filas, docs únicos y valor USD.

    Filas por cada (sub-bloque, categoría), un SUBTOTAL por sub-bloque y un
    TOTAL final. Solo aparecen los sub-bloques presentes en el universo no-D.
    """
    categories = [CATEGORY_VIGENTE, CATEGORY_NO_VIGENTE_CON_SALDO, CATEGORY_NO_VIGENTE_SIN_SALDO]
    seg_series = df.get(SUPPLY_SEGMENT_COLUMN, pd.Series("Otros", index=df.index))
    present = [seg for seg in SUPPLY_SEGMENT_ORDER if (seg_series == seg).any()]

    def _row(label: str, cat: str, sub: pd.DataFrame) -> dict:
        row: dict = {
            "Sub-bloque": label,
            "Categoría migración": cat,
            "Filas": len(sub),
            "Docs únicos": int(sub["purchase_document"].nunique(dropna=True)) if "purchase_document" in sub.columns else 0,
        }
        if usd_col and usd_col in sub.columns:
            row["Valor USD"] = round(float(pd.to_numeric(sub[usd_col], errors="coerce").sum()), 2)
        return row

    rows: list[dict] = []
    # TOTAL primero (consumo directo en la presentación).
    for cat in categories:
        rows.append(_row("TOTAL", cat, df[df[cat_col] == cat]))
    rows.append(_row("TOTAL", "SUBTOTAL", df))

    for seg in present:
        seg_df = df[seg_series == seg]
        for cat in categories:
            rows.append(_row(seg, cat, seg_df[seg_df[cat_col] == cat]))
        rows.append(_row(seg, "SUBTOTAL", seg_df))

    return pd.DataFrame(rows)


def _build_type_breakdown(
    df: pd.DataFrame,
    cat_col: str,
    value_col: str,
    usd_col: str | None = None,
) -> pd.DataFrame:
    """Pivot: (plant × category × position_type) with row counts, docs and USD value."""
    categories = [CATEGORY_VIGENTE, CATEGORY_NO_VIGENTE_CON_SALDO, CATEGORY_NO_VIGENTE_SIN_SALDO]
    plant_col = "plant_code"
    pos_col = "position_type"
    pos_display = _display_pos_type(df.get(pos_col, pd.Series("", index=df.index)))
    types = sorted(pos_display.unique().tolist())
    plants = ["TOTAL"] + sorted(df[plant_col].fillna("SIN_PLANTA").astype(str).unique().tolist())

    rows = []
    for plant in plants:
        plant_df = df if plant == "TOTAL" else df[df[plant_col].fillna("SIN_PLANTA").astype(str) == plant]
        plant_pos = pos_display.loc[plant_df.index]

        for cat in categories:
            cat_df = plant_df[plant_df[cat_col] == cat]
            cat_pos = pos_display.loc[cat_df.index]

            for pos in types:
                sub = cat_df[cat_pos == pos]
                row: dict = {
                    "Planta": plant,
                    "Categoría migración": cat,
                    "Tipo posición": pos,
                    "Filas": len(sub),
                    "Docs únicos": int(sub["purchase_document"].nunique(dropna=True)) if "purchase_document" in sub.columns else 0,
                }
                if usd_col and usd_col in sub.columns:
                    row["Valor USD"] = round(float(pd.to_numeric(sub[usd_col], errors="coerce").sum()), 2)
                rows.append(row)

            # subtotal per category within plant
            row = {
                "Planta": plant,
                "Categoría migración": f"SUBTOTAL {cat}",
                "Tipo posición": "",
                "Filas": len(cat_df),
                "Docs únicos": int(cat_df["purchase_document"].nunique(dropna=True)) if "purchase_document" in cat_df.columns else 0,
            }
            if usd_col and usd_col in cat_df.columns:
                row["Valor USD"] = round(float(pd.to_numeric(cat_df[usd_col], errors="coerce").sum()), 2)
            rows.append(row)

    return pd.DataFrame(rows)


def _build_summary(
    df: pd.DataFrame,
    cat_col: str,
    year_col: str,
    value_col: str,
    usd_col: str | None = None,
) -> pd.DataFrame:
    """Pivot: plant × category with rows, unique purchase_documents and USD value."""
    plant_col = "plant_code"
    categories = [CATEGORY_VIGENTE, CATEGORY_NO_VIGENTE_CON_SALDO, CATEGORY_NO_VIGENTE_SIN_SALDO]

    plants = ["TOTAL"] + sorted(
        df[plant_col].fillna("SIN_PLANTA").astype(str).unique().tolist()
    )

    rows = []
    for plant in plants:
        if plant == "TOTAL":
            sub = df
        else:
            sub = df[df[plant_col].fillna("SIN_PLANTA").astype(str) == plant]

        row: dict = {"Planta": plant}
        for cat in categories:
            cat_df = sub[sub[cat_col] == cat]
            n_rows = len(cat_df)
            n_docs = (
                int(cat_df["purchase_document"].nunique(dropna=True))
                if "purchase_document" in cat_df.columns
                else 0
            )
            row[f"{cat} — filas"] = n_rows
            row[f"{cat} — docs únicos"] = n_docs
            if usd_col and usd_col in cat_df.columns:
                total_usd = pd.to_numeric(cat_df[usd_col], errors="coerce").sum()
                row[f"{cat} — valor USD"] = round(float(total_usd), 2)

        rows.append(row)

    return pd.DataFrame(rows)


def _build_doc_prefix_breakdown(
    df: pd.DataFrame,
    cat_col: str,
    value_col: str,
    usd_col: str | None = None,
    material_col: str = "material",
) -> pd.DataFrame:
    """Apertura: rows = migration category + TOTAL, cols = doc prefix groups + material presence."""
    categories = [CATEGORY_VIGENTE, CATEGORY_NO_VIGENTE_CON_SALDO, CATEGORY_NO_VIGENTE_SIN_SALDO, "TOTAL"]
    prefix_labels = list(_KNOWN_PREFIXES.values()) + ["Otros"]

    doc_col = "purchase_document"
    prefix_series = _doc_prefix_label(
        df.get(doc_col, pd.Series("", index=df.index))
    )
    has_material = (
        df.get(material_col, pd.Series("", index=df.index))
        .fillna("").astype(str).str.strip().ne("")
    )

    rows = []
    for cat in categories:
        sub = df if cat == "TOTAL" else df[df[cat_col] == cat]
        sub_prefix = prefix_series.loc[sub.index]
        sub_mat = has_material.loc[sub.index]

        row: dict = {"Categoría": cat}
        for lbl in prefix_labels:
            row[lbl] = int((sub_prefix == lbl).sum())
        row["Con material"] = int(sub_mat.sum())
        row["Sin material"] = int((~sub_mat).sum())
        row["Total filas"] = len(sub)
        if value_col in df.columns:
            row["Valor pendiente"] = round(
                float(pd.to_numeric(sub[value_col], errors="coerce").sum()), 2
            )
        if usd_col and usd_col in df.columns:
            row["Valor USD"] = round(
                float(pd.to_numeric(sub[usd_col], errors="coerce").sum()), 2
            )
        rows.append(row)

    return pd.DataFrame(rows)


def _write_summary_sheet(ws, summary_df: pd.DataFrame) -> None:
    """Write summary with colour-coded category column groups."""
    n_cols = len(summary_df.columns)

    # Header row
    for col_idx, header in enumerate(summary_df.columns, 1):
        cell = ws.cell(row=1, column=col_idx, value=header)
        cell.font = _hdr_font()
        fill_color = _DARK
        for cat, color in _CATEGORY_FILL.items():
            if header.startswith(cat):
                fill_color = color
                break
        cell.fill = _fill(fill_color)
        if fill_color != _DARK:
            cell.font = Font(bold=True, color="000000")
        cell.alignment = Alignment(wrap_text=False)

    ws.freeze_panes = "B2"

    # Data rows — highlight TOTAL row
    display = summary_df.astype(object).where(pd.notnull(summary_df), other=None)
    for row_idx, row_data in enumerate(display.values.tolist(), start=2):
        for col_idx, value in enumerate(row_data, start=1):
            cell = ws.cell(row=row_idx, column=col_idx, value=value)
            if row_data[0] == "TOTAL":
                cell.font = Font(bold=True)
                cell.fill = _fill(_GREY)

    if not summary_df.empty:
        from openpyxl.utils import get_column_letter as gcl
        last_col = gcl(n_cols)
        ws.auto_filter.ref = f"A1:{last_col}{len(summary_df) + 1}"

    for col_cells in ws.columns:
        width = 12
        for cell in col_cells:
            if cell.value is not None:
                width = max(width, min(len(str(cell.value)) + 2, 45))
        ws.column_dimensions[get_column_letter(col_cells[0].column)].width = width


def generate_suministros_report(
    segment_id: str,
    label: str,
    df: pd.DataFrame,
    m01_config: dict,
    output_dir: Path,
    usd_rates: dict[str, float] | None = None,
) -> Path | None:
    """Generate SUMINISTROS Excel report for one operation.

    Args:
        segment_id: e.g. "mlcc"
        label: human-readable name
        df: full ruled dataframe, must contain migration_category (y opcionalmente
            supply_segment; si falta se calcula aquí)
        m01_config: the config block from the M01 rule definition
        output_dir: directory where the Excel is written
        usd_rates: optional mapping of currency code → USD rate (1 unit = X USD)

    Returns:
        Path to the created file, or None if migration_category column is missing.
    """
    cat_col = m01_config.get("output_column", "migration_category")
    year_col = m01_config.get("year_column", "delivery_year")
    value_col = m01_config.get("value_column", "pending_delivery_value")
    excluded_cat = m01_config.get("category_excluded_type", "TIPO_D")

    if cat_col not in df.columns:
        return None

    output_dir.mkdir(parents=True, exist_ok=True)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_path = output_dir / f"suministros_{segment_id}_{ts}.xlsx"

    # Exclude TIPO_D rows — solo se reporta el universo SUMINISTROS (no-D)
    df_comp = df[df[cat_col] != excluded_cat].copy()

    # Sub-bloque (rótulo); si no viene precalculado, se deriva aquí.
    if SUPPLY_SEGMENT_COLUMN not in df_comp.columns:
        df_comp[SUPPLY_SEGMENT_COLUMN] = classify_supply_segments(df_comp)

    # Apply USD conversion if rates were provided
    usd_col = _apply_usd_column(df_comp, value_col, usd_rates or {})

    detail_cols = [c for c in _DETAIL_COLUMNS if c in df_comp.columns]

    wb = Workbook()
    wb.remove(wb.active)

    # --- Sheet: Resumen ---
    ws_resumen = wb.create_sheet("Resumen")
    summary_df = _build_summary(df_comp, cat_col, year_col, value_col, usd_col=usd_col)
    _write_summary_sheet(ws_resumen, summary_df)

    # --- Sheet: Por_Subsegmento (sub-bloque × categoría) ---
    ws_sub = wb.create_sheet("Por_Subsegmento")
    sub_df = _build_subsegment_breakdown(df_comp, cat_col, value_col, usd_col=usd_col)
    _write_sheet(ws_sub, sub_df)

    # --- Sheet: Por_Tipo_Posicion (breakdown plant × category × position_type) ---
    ws_tipos = wb.create_sheet("Por_Tipo_Posicion")
    tipos_df = _build_type_breakdown(df_comp, cat_col, value_col, usd_col=usd_col)
    _write_sheet(ws_tipos, tipos_df)

    # --- Sheet: Apertura_Codigo (breakdown by purchase_document prefix + material) ---
    ws_codigo = wb.create_sheet("Apertura_Codigo")
    codigo_df = _build_doc_prefix_breakdown(df_comp, cat_col, value_col, usd_col=usd_col)
    _write_sheet(ws_codigo, codigo_df)

    # --- Sheet: VIGENTE ---
    ws_vigente = wb.create_sheet("VIGENTE")
    vigente_df = (
        df_comp[df_comp[cat_col] == CATEGORY_VIGENTE][detail_cols]
        .sort_values(["plant_code"] if "plant_code" in detail_cols else [])
        .reset_index(drop=True)
    )
    _write_sheet(ws_vigente, vigente_df)

    # --- Sheet: NO_VIG_SIN_SALDO ---
    ws_nvss = wb.create_sheet("NO_VIG_SIN_SALDO")
    sort_by_nvss = [c for c in ["plant_code", year_col] if c in detail_cols]
    nvss_df = (
        df_comp[df_comp[cat_col] == CATEGORY_NO_VIGENTE_SIN_SALDO][detail_cols]
        .sort_values(sort_by_nvss)
        .reset_index(drop=True)
    )
    _write_sheet(ws_nvss, nvss_df)

    # --- Sheets: SALDO_PEND_{YYYY} per year ---
    nvcs_df = df_comp[df_comp[cat_col] == CATEGORY_NO_VIGENTE_CON_SALDO].copy()
    if not nvcs_df.empty and year_col in nvcs_df.columns:
        years = sorted(
            y for y in nvcs_df[year_col].dropna().unique()
            if str(y) not in {"", "None", "nan", "<NA>"}
        )
        for year in years:
            year_int = int(year)
            year_df = (
                nvcs_df[nvcs_df[year_col] == year][detail_cols]
                .sort_values(["plant_code"] if "plant_code" in detail_cols else [])
                .reset_index(drop=True)
            )
            ws_year = wb.create_sheet(f"SALDO_PEND_{year_int}")
            _write_sheet(ws_year, year_df)
        # Catch rows with no year
        no_year = nvcs_df[nvcs_df[year_col].isna()][detail_cols]
        if not no_year.empty:
            ws_ny = wb.create_sheet("SALDO_PEND_SIN_FECHA")
            _write_sheet(ws_ny, no_year.reset_index(drop=True))
    elif not nvcs_df.empty:
        ws_nvcs = wb.create_sheet("SALDO_PENDIENTE")
        _write_sheet(ws_nvcs, nvcs_df[detail_cols].reset_index(drop=True))

    wb.save(output_path)
    return output_path
