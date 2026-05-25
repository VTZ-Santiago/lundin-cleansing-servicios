"""Load and normalize the materials master and consumption data."""
from pathlib import Path

import pandas as pd


# Columns kept from Catálogo Materiales MLCC (subset most useful for enrichment)
_MASTER_COLS = [
    "Código",
    "Descripción",
    "Tipo de material",
    "Centro",
    "Status",
    "Obs 1",
    "Obs 2",
    "Obs 3",
    "Grupo de artículos",
    "Segmento",
    "Indicador rotación Sugerido",
    "Cantidad Consumida Últ 24 Meses",
    "Frecuencia Últ 24 Meses",
    "Cantidad x vez",
    "Precio medio variable USD",
    "Unidad de medida base",
]

# Canonical renames for master columns
_MASTER_RENAME = {
    "Código": "material_key",
    "Descripción": "descripcion",
    "Tipo de material": "tipo_material",
    "Centro": "centro_master",
    "Status": "status_master",
    "Obs 1": "obs_1",
    "Obs 2": "obs_2",
    "Obs 3": "obs_3",
    "Grupo de artículos": "grupo_articulos",
    "Segmento": "segmento",
    "Indicador rotación Sugerido": "ind_rotacion_sugerido",
    "Cantidad Consumida Últ 24 Meses": "ctd_consumida_24m",
    "Frecuencia Últ 24 Meses": "frecuencia_24m_master",
    "Cantidad x vez": "ctd_x_vez_master",
    "Precio medio variable USD": "precio_medio_usd",
    "Unidad de medida base": "uom_base",
}


def _to_material_key(series: pd.Series) -> pd.Series:
    """Normalize SAP material codes: float/int → clean integer string (e.g. '32001085')."""
    numeric = pd.to_numeric(series, errors="coerce")
    return numeric.where(numeric.notna()).apply(
        lambda v: str(int(v)) if pd.notna(v) else None
    )


def load_material_master(path: Path | str) -> pd.DataFrame:
    """Load the materials master from 'Catálogo Materiales MLCC' sheet.

    Returns one row per unique material code with key enrichment columns.
    """
    path = Path(path)
    df = pd.read_excel(
        path,
        sheet_name="Catálogo Materiales MLCC",
        engine="openpyxl",
        dtype=object,
    )

    # Keep only columns that exist in this file
    keep = [c for c in _MASTER_COLS if c in df.columns]
    df = df[keep].copy()

    # Normalize material key
    df["material_key"] = _to_material_key(df["Código"])
    df = df.dropna(subset=["material_key"])

    # Drop pure-header rows (Código == 0)
    df = df[df["material_key"] != "0"]

    # Deduplicate: keep first occurrence per material_key
    df = df.sort_values("material_key").drop_duplicates(subset=["material_key"], keep="first")

    # Rename to canonical names
    rename = {k: v for k, v in _MASTER_RENAME.items() if k in df.columns and k != "Código"}
    df = df.rename(columns=rename).drop(columns=["Código"], errors="ignore")

    return df.reset_index(drop=True)


def load_consumptions(path: Path | str) -> pd.DataFrame:
    """Load monthly consumption data from 'Consumos Mensuales' sheet.

    Aggregates monthly columns into yearly totals and keeps summary indicators.
    """
    path = Path(path)
    df = pd.read_excel(
        path,
        sheet_name="Consumos Mensuales",
        engine="openpyxl",
        dtype=object,
    )

    # Find the key column (Código SAP)
    key_col = next((c for c in df.columns if "digo" in str(c) and "SAP" in str(c)), None)
    if key_col is None:
        raise ValueError(f"Cannot find 'Código SAP' column in {path}. Columns: {list(df.columns)}")

    df["material_key"] = _to_material_key(df[key_col])
    df = df.dropna(subset=["material_key"])
    df = df.drop_duplicates(subset=["material_key"], keep="first")

    # Identify monthly columns (integer-like column names, e.g. 202401)
    monthly_cols = [c for c in df.columns if isinstance(c, int) and 200000 < c < 300000]

    # Compute totals by year
    for year in (2024, 2025, 2026):
        year_cols = [c for c in monthly_cols if str(c).startswith(str(year))]
        if year_cols:
            df[f"total_consumo_{year}"] = (
                df[year_cols].apply(pd.to_numeric, errors="coerce").sum(axis=1)
            )

    # Keep summary indicator columns (last non-monthly text columns after the monthly block)
    non_monthly_extra = [
        c for c in df.columns
        if c != key_col
        and not isinstance(c, int)
        and c != "material_key"
        and c not in (f"total_consumo_{y}" for y in (2024, 2025, 2026))
    ]
    # Map them by position to friendly names if they follow the expected order
    consumo_rename: dict = {}
    for col in non_monthly_extra:
        s = str(col).lower()
        if "frecuencia" in s:
            consumo_rename[col] = "frecuencia_consumo_24m"
        elif "cantidad" in s and "vez" in s:
            consumo_rename[col] = "ctd_x_vez_consumo"
        elif "rotaci" in s or "rotacion" in s:
            consumo_rename[col] = "ind_rotacion_consumo"

    keep_extra = list(consumo_rename.keys())
    cols_out = ["material_key"] + [f"total_consumo_{y}" for y in (2024, 2025, 2026) if f"total_consumo_{y}" in df.columns] + keep_extra
    df = df[[c for c in cols_out if c in df.columns]].copy()
    df = df.rename(columns=consumo_rename)

    return df.reset_index(drop=True)
