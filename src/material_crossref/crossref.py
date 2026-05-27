"""Extract source materials from contracts/OC and cross-reference with master + consumptions."""
from pathlib import Path

import pandas as pd

from src.config.domains import DOMAIN_CONFIGS
from src.schema.field_map_mlcc import MLCC_RAW_TO_CANONICAL, normalize_header

# Canonical columns we care about for the crossref
_WANTED_CANONICAL = {
    "material",
    "position_type",
    "account_assignment_type",
    "short_text",
    "validity_end",
    "purchase_document",
    "purchase_group",
    "plant_code",
}


def _usecols_filter(col_name: str) -> bool:
    """Return True if this raw column maps to one of our wanted canonical columns."""
    canonical = MLCC_RAW_TO_CANONICAL.get(normalize_header(str(col_name)))
    return canonical in _WANTED_CANONICAL


def _read_source_file(path: Path) -> pd.DataFrame:
    """Read a single Excel source file, keeping only relevant columns."""
    df = pd.read_excel(
        path,
        engine="openpyxl",
        dtype=object,
        usecols=_usecols_filter,
    )
    # Rename raw headers → canonical names
    rename = {}
    for raw_col in df.columns:
        canonical = MLCC_RAW_TO_CANONICAL.get(normalize_header(str(raw_col)))
        if canonical:
            rename[raw_col] = canonical
    df = df.rename(columns=rename)
    df["_source_file"] = path.name
    return df


def normalize_material_key(series: pd.Series) -> pd.Series:
    """Float or str material → clean integer string key (e.g. '32001085')."""
    numeric = pd.to_numeric(series, errors="coerce")
    return numeric.apply(lambda v: str(int(v)) if pd.notna(v) else None)


def _normalize_material_key(series: pd.Series) -> pd.Series:
    return normalize_material_key(series)


def extract_source_materials(domain: str, inputs_root: Path, operation: str = "MLCC") -> pd.DataFrame:
    """Load all source files for a domain/operation and return a flat DataFrame."""
    cfg = DOMAIN_CONFIGS.get(domain)
    if cfg is None:
        raise ValueError(f"Unknown domain: {domain}")

    source_dir = inputs_root / operation.upper() / cfg.input_subdir
    files = [
        f for f in sorted(source_dir.glob("*.XLSX")) + sorted(source_dir.glob("*.xlsx"))
        if not f.name.startswith("~$")
    ]
    if not files:
        raise FileNotFoundError(f"No Excel files found in {source_dir}")

    frames = []
    for f in files:
        df = _read_source_file(f)
        df["domain"] = domain
        df["operation"] = operation.upper()
        frames.append(df)

    master = pd.concat(frames, ignore_index=True, join="outer")

    # Normalize material key
    master["material_key"] = _normalize_material_key(master.get("material", pd.Series(dtype=object)))

    # Forward-fill purchase_document within each source file (SAP omits repeated values)
    if "purchase_document" in master.columns:
        master["purchase_document"] = master.groupby("_source_file")["purchase_document"].ffill()

    # Drop rows without a valid material code
    master = master.dropna(subset=["material_key"])
    master = master[master["material_key"] != "0"]

    return master.reset_index(drop=True)


def build_unique_materials(source_df: pd.DataFrame) -> pd.DataFrame:
    """Deduplicate source rows to one row per (material_key, operation, domain)."""
    group_cols = [c for c in ("material_key", "operation", "domain") if c in source_df.columns]
    rows = []
    for group_vals, grp in source_df.groupby(group_cols, sort=False):
        vals = group_vals if isinstance(group_vals, tuple) else (group_vals,)
        meta = dict(zip(group_cols, vals))

        pos_types    = grp["position_type"].dropna().unique().tolist() if "position_type" in grp.columns else []
        assign_types = grp["account_assignment_type"].dropna().unique().tolist() if "account_assignment_type" in grp.columns else []

        short_text_src = grp["short_text"].dropna().iloc[0]    if "short_text"    in grp.columns and grp["short_text"].notna().any()    else None
        plant          = grp["plant_code"].dropna().iloc[0]     if "plant_code"    in grp.columns and grp["plant_code"].notna().any()     else None
        pur_group_s    = grp["purchase_group"].dropna().mode()  if "purchase_group" in grp.columns else pd.Series(dtype=object)
        pur_group      = pur_group_s.iloc[0] if not pur_group_s.empty else None

        is_service_d   = "D" in pos_types
        is_direct_cost = "K" in assign_types

        rows.append({
            **meta,
            "n_registros":        len(grp),
            "n_documentos":       grp["purchase_document"].nunique() if "purchase_document" in grp.columns else 0,
            "tipos_posicion":     ", ".join(sorted(set(str(t) for t in pos_types)))    if pos_types    else "",
            "tipos_imputacion":   ", ".join(sorted(set(str(t) for t in assign_types))) if assign_types else "",
            "descripcion_fuente": short_text_src,
            "centro_fuente":      plant,
            "grupo_compras":      pur_group,
            "is_service":         is_service_d or is_direct_cost,
            "is_service_d":       is_service_d,
            "is_direct_cost":     is_direct_cost,
        })

    return pd.DataFrame(rows)


def build_crossref(
    unique_df: pd.DataFrame,
    master_df: pd.DataFrame,
    consumptions_df: pd.DataFrame,
) -> pd.DataFrame:
    """Join unique materials with master and consumption data, adding coverage flags."""
    enriched = unique_df.merge(master_df, on="material_key", how="left")
    enriched = enriched.merge(consumptions_df, on="material_key", how="left")

    enriched["in_master"] = enriched["descripcion"].notna()
    if "total_consumo_2024" in enriched.columns:
        enriched["has_consumption"] = enriched["total_consumo_2024"].notna()
    elif "total_consumo_2025" in enriched.columns:
        enriched["has_consumption"] = enriched["total_consumo_2025"].notna()
    else:
        enriched["has_consumption"] = False

    # Prefer master description; fall back to source short_text
    enriched["descripcion_efectiva"] = enriched["descripcion"].fillna(enriched["descripcion_fuente"])

    return enriched.reset_index(drop=True)
