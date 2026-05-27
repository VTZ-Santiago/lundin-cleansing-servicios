"""Row-level material enrichment for pipeline rules."""
from pathlib import Path

import pandas as pd

from src.material_crossref.crossref import normalize_material_key
from src.material_crossref.loader import load_consumptions, load_material_master


def _single_match(directory: Path, pattern: str) -> Path:
    matches = sorted(directory.glob(pattern))
    if not matches:
        raise FileNotFoundError(f"No file found in {directory} matching {pattern}")
    return matches[0]


def load_reference_tables(inputs_dir: Path) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Load material master and historical consumptions from the MLCC inputs folder."""
    mlcc_dir = inputs_dir / "MLCC"
    master_path = _single_match(mlcc_dir, "MLCC - An*lisis Completo.xlsx")
    consumptions_path = _single_match(mlcc_dir, "MLCC - Consumos Hist*ricos Marzo 2026.xlsx")
    return load_material_master(master_path), load_consumptions(consumptions_path)


def enrich_material_rows(
    df: pd.DataFrame,
    material_master: pd.DataFrame,
    consumptions: pd.DataFrame,
) -> pd.DataFrame:
    """Append material-reference columns without changing row order or row count."""
    result = df.copy()
    result["material_key"] = normalize_material_key(
        result.get("material", pd.Series(pd.NA, index=result.index))
    )
    result["_material_enrich_order"] = range(len(result))

    master_cols = [
        "material_key",
        "descripcion",
        "texto_compra",
        "tipo_material",
        "centro_master",
        "status_master",
        "planning_type",
        "critical_part",
        "operational_criticality",
        "last_movement_date",
        "ctd_consumida_24m",
        "frecuencia_24m_master",
        "ind_rotacion_sugerido",
        "precio_medio_usd",
        "uom_base",
    ]
    master_ref = material_master[[c for c in master_cols if c in material_master.columns]].copy()
    master_ref = master_ref.drop_duplicates(subset=["material_key"], keep="first")
    master_ref = master_ref.rename(columns={
        "descripcion": "material_master_description",
        "texto_compra": "material_purchase_text",
        "tipo_material": "material_type",
        "centro_master": "material_master_plant",
        "status_master": "material_master_status",
        "planning_type": "material_planning_type",
        "critical_part": "material_critical_part",
        "operational_criticality": "material_operational_criticality",
        "last_movement_date": "material_last_movement_date",
        "ctd_consumida_24m": "material_consumed_24m",
        "frecuencia_24m_master": "material_frequency_24m",
        "ind_rotacion_sugerido": "material_rotation_indicator",
        "precio_medio_usd": "material_avg_price_usd",
        "uom_base": "material_base_uom",
    })

    consumption_ref = consumptions.copy()
    consumption_ref = consumption_ref.drop_duplicates(subset=["material_key"], keep="first")
    consumption_ref = consumption_ref.rename(columns={
        col: f"material_{col}" for col in consumption_ref.columns if col != "material_key"
    })

    result = result.merge(master_ref, on="material_key", how="left")
    result = result.merge(consumption_ref, on="material_key", how="left")
    result = result.sort_values("_material_enrich_order").drop(columns=["_material_enrich_order"])

    result["material_in_master"] = result["material_master_description"].notna()
    total_cols = [c for c in result.columns if c.startswith("material_total_consumo_")]
    if total_cols:
        result["material_has_consumption_data"] = result[total_cols].notna().any(axis=1)
    else:
        result["material_has_consumption_data"] = False

    frequency = pd.to_numeric(
        result.get("material_frequency_24m", pd.Series(pd.NA, index=result.index)),
        errors="coerce",
    )
    result["material_no_movement_24m"] = result["material_in_master"] & frequency.fillna(0).le(0)

    critical_part = (
        result.get("material_critical_part", pd.Series("", index=result.index))
        .fillna("")
        .astype(str)
        .str.strip()
        .str.upper()
    )
    criticality = (
        result.get("material_operational_criticality", pd.Series("", index=result.index))
        .fillna("")
        .astype(str)
        .str.strip()
        .str.upper()
    )
    result["material_is_critical"] = critical_part.ne("") | criticality.isin({"A", "B"})
    result["material_effective_description"] = result["material_master_description"].fillna(
        result.get("short_text", pd.Series(pd.NA, index=result.index))
    )

    return result