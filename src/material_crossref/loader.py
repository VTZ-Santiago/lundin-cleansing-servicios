"""Load and normalize the materials master and consumption data.

Supports MLCC and CCMC operations, which have different sheet names,
key column names, and column structures.
"""
from pathlib import Path

import pandas as pd


# Per-operation configuration for the materials master
# col_map: raw column name → canonical name used throughout the module
_OPERATION_MASTER_CONFIGS: dict[str, dict] = {
    "MLCC": {
        "sheet": "Catálogo Materiales MLCC",
        "key_col": "Código",
        "col_map": {
            "Código":                                  "material_key",
            "Descripción":                             "descripcion",
            "Texto de Compra":                         "texto_compra",
            "Tipo de material":                        "tipo_material",
            "Centro":                                  "centro_master",
            "Status":                                  "status_master",
            "Obs 1":                                   "obs_1",
            "Obs 2":                                   "obs_2",
            "Obs 3":                                   "obs_3",
            "Grupo de artículos":                      "grupo_articulos",
            "Segmento":                                "segmento",
            "Característica Planificación Necesaria":  "planning_type",
            "Indicador: Parte crítica":                "critical_part",
            "Criticidad Operativa (Rotación)":         "operational_criticality",
            "Fecha Último Movimiento":                 "last_movement_date",
            "Indicador rotación Sugerido":             "ind_rotacion_sugerido",
            "Cantidad Consumida Últ 24 Meses":         "ctd_consumida_24m",
            "Frecuencia Últ 24 Meses":                "frecuencia_24m_master",
            "Cantidad x vez":                          "ctd_x_vez_master",
            "Precio medio variable USD":               "precio_medio_usd",
            "Unidad de medida base":                   "uom_base",
        },
    },
    "CCMC": {
        "sheet": "CCMC",
        "key_col": "Material",
        "col_map": {
            "Material":                    "material_key",
            "Texto breve de material":     "descripcion",
            "Texto de Compra":             "texto_compra",
            "Tipo material":               "tipo_material",
            "Centro":                      "centro_master",
            "Status":                      "status_master",
            "Obs1":                        "obs_1",
            "Obs2":                        "obs_2",
            "Obs3":                        "obs_3",
            "Grupo de artículos":          "grupo_articulos",
            "Grupo articulos":             "grupo_articulos",
            "Segmento":                    "segmento",
            "Caract.planif.nec.":          "planning_type",
            "Parte crítica":               "critical_part",
            "Rotación (Labor/oficina)":    "operational_criticality",
            "Fecha Último Movimiento":     "last_movement_date",
            "Fecha último Movimiento":     "last_movement_date",
            "Indicador rotación Sugerido": "ind_rotacion_sugerido",
            "Cantidad Consumida Últ 24 Meses": "ctd_consumida_24m",
            "Frecuencia Últ 24 Meses":    "frecuencia_24m_master",
            "Cantidad x vez":              "ctd_x_vez_master",
            "Precio Unitario USD":         "precio_medio_usd",
            "Precio variable":             "precio_medio_usd",
            "Unidad medida base":          "uom_base",
            "Unidad de medida base":       "uom_base",
        },
    },
}

# Per-operation configuration for the consumption file
_OPERATION_CONSUMOS_CONFIGS: dict[str, dict] = {
    "MLCC": {
        "sheet": "Consumos Mensuales",
        # key col has "digo" AND "SAP"
        "key_matcher": lambda c: "digo" in str(c) and "SAP" in str(c),
        # monthly cols are int (202401)
        "monthly_matcher": lambda c: isinstance(c, int) and 200000 < c < 300000,
        "month_to_str": str,
    },
    "CCMC": {
        "sheet": "Consumos Mensuales",
        # key col has "digo" but NOT "SAP"
        "key_matcher": lambda c: "digo" in str(c) and "SAP" not in str(c),
        # monthly cols are str ('202401')
        "monthly_matcher": lambda c: isinstance(c, str) and c.isdigit() and 200000 < int(c) < 300000,
        "month_to_str": str,
    },
}


def _to_material_key(series: pd.Series) -> pd.Series:
    """Normalize SAP material codes: float/int/str → clean integer string (e.g. '32001085')."""
    numeric = pd.to_numeric(series, errors="coerce")
    return numeric.apply(lambda v: str(int(v)) if pd.notna(v) else None)


def _match_columns(df_cols: list, col_map: dict) -> dict:
    """Match DataFrame columns to col_map keys, returning {raw_col: canonical} for found ones."""
    result = {}
    for raw_col in df_cols:
        if raw_col in col_map:
            result[raw_col] = col_map[raw_col]
    return result


def load_material_master(path: Path | str, operation: str = "MLCC") -> pd.DataFrame:
    """Load the materials master for the given operation.

    Returns one deduplicated row per material code with canonical column names.
    """
    path = Path(path)
    cfg = _OPERATION_MASTER_CONFIGS[operation.upper()]

    df = pd.read_excel(path, sheet_name=cfg["sheet"], engine="openpyxl", dtype=object)

    col_map = cfg["col_map"]
    key_col = cfg["key_col"]

    # Keep only columns present in this file that have a mapping
    keep = [c for c in df.columns if c in col_map]
    df = df[keep].copy()

    # Normalize material key
    if key_col not in df.columns:
        raise ValueError(f"Key column '{key_col}' not found in {path} sheet '{cfg['sheet']}'.")
    df["material_key"] = _to_material_key(df[key_col])
    df = df.dropna(subset=["material_key"])
    df = df[df["material_key"] != "0"]

    # Deduplicate per material_key (keep first after sort)
    df = df.sort_values("material_key").drop_duplicates(subset=["material_key"], keep="first")

    # Rename to canonical names (drop original key col)
    rename = {k: v for k, v in col_map.items() if k in df.columns and k != key_col}
    df = df.rename(columns=rename).drop(columns=[key_col], errors="ignore")

    # Drop duplicate canonical columns that arise when multiple raw names map to the same target
    df = df.loc[:, ~df.columns.duplicated(keep="first")]

    return df.reset_index(drop=True)


def load_consumptions(path: Path | str, operation: str = "MLCC") -> pd.DataFrame:
    """Load monthly consumption data for the given operation.

    Aggregates monthly columns into yearly totals and normalises indicator columns.
    """
    path = Path(path)
    cfg = _OPERATION_CONSUMOS_CONFIGS[operation.upper()]

    df = pd.read_excel(path, sheet_name=cfg["sheet"], engine="openpyxl", dtype=object)

    key_matcher = cfg["key_matcher"]
    key_col = next((c for c in df.columns if key_matcher(c)), None)
    if key_col is None:
        raise ValueError(f"Cannot find consumption key column in {path}. Columns: {list(df.columns[:10])}")

    df["material_key"] = _to_material_key(df[key_col])
    df = df.dropna(subset=["material_key"])
    df = df.drop_duplicates(subset=["material_key"], keep="first")

    monthly_matcher = cfg["monthly_matcher"]
    monthly_cols = [c for c in df.columns if monthly_matcher(c)]

    for year in (2024, 2025, 2026):
        year_cols = [c for c in monthly_cols if cfg["month_to_str"](c).startswith(str(year))]
        if year_cols:
            df[f"total_consumo_{year}"] = (
                df[year_cols].apply(pd.to_numeric, errors="coerce").sum(axis=1)
            )

    # Rename summary indicator columns to canonical names
    consumo_rename: dict = {}
    extra_cols = [
        c for c in df.columns
        if c != key_col
        and not monthly_matcher(c)
        and c != "material_key"
        and not str(c).startswith("total_consumo_")
    ]
    for col in extra_cols:
        s = str(col).lower()
        if "frecuencia" in s:
            consumo_rename[col] = "frecuencia_consumo_24m"
        elif "cantidad" in s and "vez" in s:
            consumo_rename[col] = "ctd_x_vez_consumo"
        elif "rotaci" in s:
            consumo_rename[col] = "ind_rotacion_consumo"

    cols_out = (
        ["material_key"]
        + [f"total_consumo_{y}" for y in (2024, 2025, 2026) if f"total_consumo_{y}" in df.columns]
        + list(consumo_rename.keys())
    )
    df = df[[c for c in cols_out if c in df.columns]].copy()
    df = df.rename(columns=consumo_rename)

    return df.reset_index(drop=True)
