import re
from datetime import datetime
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter


# ── Paleta ──────────────────────────────────────────────────────────────────
_SEC_HDR_BG = {
    "periodo":  "374151",
    "volumen":  "1D4ED8",
    "compras":  "0F766E",
    "tipo":     "065F46",
    "vigencia": "92400E",
    "cruce":    "7C2D12",
    "borrado":  "4C1D95",
}
_SEC_DATA_BG = {
    "periodo":  "F3F4F6",
    "volumen":  "DBEAFE",
    "compras":  "CCFBF1",
    "tipo":     "D1FAE5",
    "vigencia": "FEF3C7",
    "cruce":    "FFEDD5",
    "borrado":  "EDE9FE",
}
_SEC_LABELS = {
    "periodo":  "",
    "volumen":  "Volumen",
    "compras":  "Grupo de Compras y Monto",
    "tipo":     "Tipo de Posición  (tipo dominante por contrato)",
    "vigencia": "Vigencia del Contrato  (fecha de vencimiento más reciente por contrato)",
    "cruce":    "CRUCE  —  Tipo de Posición  ×  Año de Vencimiento",
    "borrado":  "Indicador de Borrado  (valor dominante por contrato)",
}
_SEP_BG    = "D1D5DB"
_TOTALS_BG = "1E3A5F"
_GRAY_CELL = "E9ECEF"   # celdas vacías en bloques adicionales

# Bloques adicionales (C, V, VACIO) que se replican bajo el bloque D
_EXTRA_BLOCKS = [
    ("C",  "c",    "TIPO C  —  Consignación"),
    ("V",  "v",    "TIPO V  —  Valor límite"),
    ("",   "vacio","SIN TIPO  —  Posición sin tipo asignado"),
]


def _fill(hex_bg: str) -> PatternFill:
    return PatternFill(start_color=hex_bg, end_color=hex_bg, fill_type="solid")


def _extract_period(source_file: str) -> str:
    stem = Path(source_file).stem
    m = re.search(r"\d{4}[-_]\d{4}", stem)
    return m.group().replace("-", "_") if m else stem


def _is_sep(sec: str) -> bool:
    return sec.startswith("_sep")


def _contracts_table(subset: pd.DataFrame) -> pd.DataFrame:
    tmp = pd.DataFrame({
        "purchase_document": subset["purchase_document"].values,
        "validity_end": pd.to_datetime(
            subset["validity_end"].values if "validity_end" in subset.columns
            else [pd.NaT] * len(subset),
            errors="coerce",
        ),
        "position_type": (
            subset["position_type"].fillna("").astype(str).str.strip().values
            if "position_type" in subset.columns else [""] * len(subset)
        ),
        "deletion_flag": (
            subset["deletion_flag"].fillna("").astype(str).str.strip().values
            if "deletion_flag" in subset.columns else [""] * len(subset)
        ),
        "purchase_group": (
            subset["purchase_group"].fillna("SIN_GRUPO").astype(str).str.strip().values
            if "purchase_group" in subset.columns else ["SIN_GRUPO"] * len(subset)
        ),
        "estimated_value": pd.to_numeric(
            subset["estimated_value"].values if "estimated_value" in subset.columns
            else [0.0] * len(subset),
            errors="coerce",
        ),
    })
    tmp["purchase_group"] = tmp["purchase_group"].where(tmp["purchase_group"] != "", other="SIN_GRUPO")
    contracts = tmp.groupby("purchase_document").agg(
        max_ve=("validity_end", "max"),
        dom_type=("position_type", lambda s: s.mode().iloc[0] if not s.empty else ""),
        dom_flag=("deletion_flag", lambda s: s.mode().iloc[0] if not s.empty else ""),
        dom_purchase_group=("purchase_group", lambda s: s.mode().iloc[0] if not s.empty else "SIN_GRUPO"),
        estimated_value_sum=("estimated_value", "sum"),
    )
    contracts["yr"] = contracts["max_ve"].dt.year
    return contracts


def _period_stats(subset: pd.DataFrame, future_years: list[int]) -> dict:
    c = _contracts_table(subset)
    stats: dict = {
        "nro_lineas":    len(subset),
        "nro_contratos": len(c),
        "grupos_compra": int(c["dom_purchase_group"].nunique()),
        "monto_total_usd": round(float(c["estimated_value_sum"].sum()), 2),
        "pos_D":         int((c["dom_type"] == "D").sum()),
        "pos_C":         int((c["dom_type"] == "C").sum()),
        "pos_V":         int((c["dom_type"] == "V").sum()),
        "pos_VACIO":     int((c["dom_type"] == "").sum()),
        "yr_leq2025":    int((c["yr"] <= 2025).sum()),
    }
    for yr in future_years:
        stats[f"yr_{yr}"] = int((c["yr"] == yr).sum())
    stats["yr_sin_fecha"] = int(c["yr"].isna().sum())

    # Bloque D
    stats["d_leq2025"] = int(((c["dom_type"] == "D") & (c["yr"] <= 2025)).sum())
    for yr in future_years:
        stats[f"d_{yr}"] = int(((c["dom_type"] == "D") & (c["yr"] == yr)).sum())

    # Bloques C, V, VACIO
    for typ, pfx, _ in _EXTRA_BLOCKS:
        stats[f"{pfx}_leq2025"] = int(((c["dom_type"] == typ) & (c["yr"] <= 2025)).sum())
        for yr in future_years:
            stats[f"{pfx}_{yr}"] = int(((c["dom_type"] == typ) & (c["yr"] == yr)).sum())

    stats["flag_L"]   = int((c["dom_flag"] == "L").sum())
    stats["flag_S"]   = int((c["dom_flag"] == "S").sum())
    stats["flag_sin"] = int((c["dom_flag"] == "").sum())
    return stats


def _remap_cruce(row_data: dict, pfx: str, future_years: list[int]) -> dict:
    """Return row_data with d_* cruce keys replaced by pfx_* values."""
    remapped = dict(row_data)
    remapped["d_leq2025"] = row_data.get(f"{pfx}_leq2025", 0)
    for yr in future_years:
        remapped[f"d_{yr}"] = row_data.get(f"{pfx}_{yr}", 0)
    return remapped


def _col_defs(future_years: list[int]) -> list[tuple[str, str, str, str]]:
    d: list[tuple[str, str, str, str]] = []

    d += [("periodo", "periodo", "Período", "Archivo fuente")]
    d += [("_sep_0", "_sep_0", "", "")]

    d += [
        ("volumen", "nro_lineas",    "Total de posiciones",  "(filas en el dataset)"),
        ("volumen", "nro_contratos", "Contratos únicos",     "(Documentos de compra distintos)"),
    ]
    d += [("_sep_1", "_sep_1", "", "")]

    d += [
        ("compras", "grupos_compra", "Grupos de compras", "(grupos distintos)"),
        ("compras", "monto_total_usd", "Monto total USD", "(estimated_value)"),
    ]
    d += [("_sep_1b", "_sep_1b", "", "")]

    d += [
        ("tipo", "pos_D",     "Tipo D",             "Servicio / límite de valor"),
        ("tipo", "pos_C",     "Tipo C",             "Consignación"),
        ("tipo", "pos_V",     "Tipo V",             "Valor límite"),
        ("tipo", "pos_VACIO", "Sin tipo asignado",  "(position_type vacío)"),
    ]
    d += [("_sep_2", "_sep_2", "", "")]

    d += [("vigencia", "yr_leq2025", "Vencidos hasta 2025",
           "(cualquier año ≤ 2025 — agrupados)")]
    for yr in future_years:
        d.append(("vigencia", f"yr_{yr}", f"Vigentes en {yr}", f"(vencimiento en {yr})"))
    d.append(("vigencia", "yr_sin_fecha", "Sin fecha de vencimiento", "(validity_end vacío)"))
    d += [("_sep_3", "_sep_3", "", "")]

    # Cruce — columnas genéricas (el bloque activo lo indica la fila de etiqueta)
    d.append(("cruce", "d_leq2025", "× Vencidos ≤ 2025",
              "(tipo según bloque — vencimiento agrupado)"))
    for yr in future_years:
        d.append(("cruce", f"d_{yr}", f"× Vigentes {yr}", f"(tipo según bloque — año {yr})"))
    d += [("_sep_4", "_sep_4", "", "")]

    d += [
        ("borrado", "flag_L",   "Libre — sin borrado (L)",   "(indicador L — no marcado)"),
        ("borrado", "flag_S",   "Marcado para borrar (S)",   "(indicador S — en proceso de baja)"),
        ("borrado", "flag_sin", "Sin indicador de borrado",  "(campo deletion_flag vacío)"),
    ]
    return d


def _write_section_headers(ws, col_defs: list) -> None:
    n = len(col_defs)
    groups: list[tuple[str, int, int]] = []
    cur_sec, start = col_defs[0][0], 1
    for i, (sec, *_) in enumerate(col_defs[1:], 2):
        if sec != cur_sec:
            groups.append((cur_sec, start, i - 1))
            cur_sec, start = sec, i
    groups.append((cur_sec, start, n))

    for sec, c_start, c_end in groups:
        if c_end > c_start:
            ws.merge_cells(start_row=2, start_column=c_start,
                           end_row=2, end_column=c_end)
        cell = ws.cell(row=2, column=c_start)
        if _is_sep(sec):
            cell.fill = _fill(_SEP_BG)
        else:
            cell.value = _SEC_LABELS.get(sec, sec)
            cell.font = Font(bold=True, color="FFFFFF", size=10)
            cell.fill = _fill(_SEC_HDR_BG.get(sec, "374151"))
            cell.alignment = Alignment(horizontal="center", vertical="center")


def _write_data_rows(
    ws,
    all_rows: list[dict],
    col_defs: list,
    start_row: int,
    only_cruce: bool = False,
) -> int:
    """Write data rows and return the next available row number."""
    row_n = start_row
    for row_data in all_rows:
        is_total = row_data["periodo"] == "TOTAL"
        for col_idx, (sec, key, *_) in enumerate(col_defs, 1):
            cell = ws.cell(row=row_n, column=col_idx)
            if _is_sep(sec):
                cell.fill = _fill(_SEP_BG)
                continue
            if key == "periodo":
                cell.value = row_data.get("periodo", "")
                if is_total:
                    cell.font = Font(bold=True, size=10, color="FFFFFF")
                    cell.fill = _fill(_TOTALS_BG)
                else:
                    cell.fill = _fill(_SEC_DATA_BG["periodo"])
                    cell.font = Font(bold=True, size=10)
                cell.alignment = Alignment(horizontal="left", vertical="center")
                continue
            if only_cruce and sec != "cruce":
                # En bloques adicionales, vaciar secciones que no son cruce
                cell.fill = _fill(_GRAY_CELL)
                continue
            val = row_data.get(key, 0)
            cell.value = val
            if is_total:
                cell.font = Font(bold=True, size=10, color="FFFFFF")
                cell.fill = _fill(_TOTALS_BG)
            else:
                cell.fill = _fill(_SEC_DATA_BG.get(sec, "FFFFFF"))
                cell.font = Font(bold=(key == "periodo"), size=10)
            cell.alignment = Alignment(
                horizontal="left" if key == "periodo" else "right",
                vertical="center",
            )
            if isinstance(val, int) and key != "periodo":
                cell.number_format = "#,##0"
            elif isinstance(val, float):
                cell.number_format = "#,##0.00"
        ws.row_dimensions[row_n].height = 18
        row_n += 1
    return row_n


def _write_block_label(ws, row_n: int, col_defs: list, label: str) -> int:
    """Label row: merges only the cruce-section columns; other cells get a neutral fill."""
    cruce_start = cruce_end = None
    for col_idx, (sec, *_) in enumerate(col_defs, 1):
        if sec == "cruce":
            if cruce_start is None:
                cruce_start = col_idx
            cruce_end = col_idx

    for col_idx, (sec, *_) in enumerate(col_defs, 1):
        cell = ws.cell(row=row_n, column=col_idx)
        if _is_sep(sec):
            cell.fill = _fill(_SEP_BG)
        elif sec == "cruce":
            cell.fill = _fill(_SEC_HDR_BG["cruce"])  # covered by merge below
        else:
            cell.fill = _fill(_GRAY_CELL)

    if cruce_start is not None and cruce_end is not None and cruce_end >= cruce_start:
        ws.merge_cells(start_row=row_n, start_column=cruce_start,
                       end_row=row_n, end_column=cruce_end)
    lc = ws.cell(row=row_n, column=cruce_start or 1,
                 value=f"  CRUCE  —  {label}")
    lc.font = Font(bold=True, size=10, color="FFFFFF")
    lc.fill = _fill(_SEC_HDR_BG["cruce"])
    lc.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[row_n].height = 18
    return row_n + 1


def generate_stats_report(
    df: pd.DataFrame,
    operation: str,
    output_dir: Path,
    dataset_label: str = "Universo Completo — C1 (pre-reglas)",
) -> Path:
    if "validity_end" in df.columns:
        ve = pd.to_datetime(df["validity_end"], errors="coerce")
        future_years = sorted({int(y) for y in ve.dropna().dt.year.unique() if y >= 2026})
    else:
        future_years = []

    # Per-period stats — ordenados por período extraído (no por nombre de archivo)
    period_rows: list[dict] = []
    if "_source_file" in df.columns:
        src_files = sorted(df["_source_file"].unique(),
                           key=lambda s: _extract_period(s))
        for src_file in src_files:
            subset = df[df["_source_file"] == src_file]
            stats = _period_stats(subset, future_years)
            stats["periodo"] = _extract_period(str(src_file))
            period_rows.append(stats)
    else:
        stats = _period_stats(df, future_years)
        stats["periodo"] = operation
        period_rows.append(stats)

    numeric_keys = [k for k in period_rows[0] if k != "periodo"]
    total = {"periodo": "TOTAL"} | {k: sum(r[k] for r in period_rows) for k in numeric_keys}
    if "purchase_group" in df.columns:
        purchase_group = df["purchase_group"].fillna("SIN_GRUPO").astype(str).str.strip()
        purchase_group = purchase_group.where(purchase_group != "", other="SIN_GRUPO")
        total["grupos_compra"] = int(purchase_group.nunique())
    all_rows = period_rows + [total]

    col_defs = _col_defs(future_years)
    ncols = len(col_defs)

    # ── Workbook ─────────────────────────────────────────────────────────────
    wb = Workbook()
    ws = wb.active
    ws.title = f"Estadísticas {operation}"

    # Fila 1 — Banner
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ncols)
    t = ws.cell(row=1, column=1,
                value=(f"Estadísticas de Contratos — Operación {operation}   |   "
                       f"{dataset_label}   |   "
                       f"Generado: {datetime.now().strftime('%Y-%m-%d %H:%M')}"))
    t.font = Font(bold=True, size=12, color="FFFFFF")
    t.fill = _fill("1E3A5F")
    t.alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[1].height = 22

    # Fila 2 — Secciones fusionadas
    _write_section_headers(ws, col_defs)
    ws.row_dimensions[2].height = 20

    # Fila 3 — Sub-encabezados
    for col_idx, (sec, key, hdr1, hdr2) in enumerate(col_defs, 1):
        cell = ws.cell(row=3, column=col_idx)
        if _is_sep(sec):
            cell.fill = _fill(_SEP_BG)
        else:
            cell.value = f"{hdr1}\n{hdr2}"
            cell.font = Font(bold=True, size=9, color="FFFFFF")
            cell.fill = _fill(_SEC_HDR_BG.get(sec, "374151"))
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[3].height = 42

    # ── Bloque D (completo, todas las columnas) ──────────────────────────────
    row_n = 4
    row_n = _write_block_label(ws, row_n, col_defs, "TIPO D  —  Servicio / límite de valor")
    row_n = _write_data_rows(ws, all_rows, col_defs, row_n, only_cruce=False)

    # ── Bloques C, V, VACIO (sólo cruce — otras columnas en gris) ────────────
    for typ, pfx, label in _EXTRA_BLOCKS:
        # Fila en blanco
        for col_idx in range(1, ncols + 1):
            ws.cell(row=row_n, column=col_idx).fill = _fill("FFFFFF")
        ws.row_dimensions[row_n].height = 8
        row_n += 1

        # Etiqueta del bloque
        row_n = _write_block_label(ws, row_n, col_defs, label)

        # Datos con cruce remapeado
        remapped_rows = [_remap_cruce(r, pfx, future_years) for r in all_rows]
        row_n = _write_data_rows(ws, remapped_rows, col_defs, row_n, only_cruce=True)

    # ── Anchos de columna ────────────────────────────────────────────────────
    for col_idx, (sec, key, *_) in enumerate(col_defs, 1):
        if _is_sep(sec):
            ws.column_dimensions[get_column_letter(col_idx)].width = 2
        elif key == "periodo":
            ws.column_dimensions[get_column_letter(col_idx)].width = 14
        elif key in ("nro_lineas", "nro_contratos", "grupos_compra", "monto_total_usd", "yr_leq2025",
                     "yr_sin_fecha", "d_leq2025", "flag_sin"):
            ws.column_dimensions[get_column_letter(col_idx)].width = 14
        else:
            ws.column_dimensions[get_column_letter(col_idx)].width = 11

    ws.freeze_panes = "B4"

    path = output_dir / f"reporte_estadistico_{operation}.xlsx"
    wb.save(str(path))
    return path
