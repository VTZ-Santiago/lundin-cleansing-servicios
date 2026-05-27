import pandas as pd

from src.rules.base import RuleResult
from src.utils.date_logic import effective_validity_dates


def apply(df: pd.DataFrame, operation: str, config: dict) -> RuleResult:
    cutoff = pd.Timestamp(config.get("active_after", "2025-12-31"))
    dates = effective_validity_dates(df)
    active = dates.isna() | dates.gt(cutoff)

    document = df.get("purchase_document", pd.Series("", index=df.index)).fillna("").astype(str).str.strip()
    material = df.get("material_key", pd.Series("", index=df.index)).fillna("").astype(str).str.strip()
    valid = active & document.ne("") & material.ne("")

    counts = (
        pd.DataFrame({"material": material[valid], "document": document[valid]})
        .groupby("material")["document"]
        .nunique()
    )
    duplicated_materials = set(counts[counts > 1].index)
    mask = valid & material.isin(duplicated_materials)

    mark = pd.Series("", index=df.index)
    mark.loc[mask] = "MATERIAL_EN_MULTIPLES_CONTRATOS_VIGENTES"
    return RuleResult(
        "C03",
        mask,
        mask.map({True: "C03", False: ""}),
        updates={"mark_material_multi_active_contract": mark},
    )
