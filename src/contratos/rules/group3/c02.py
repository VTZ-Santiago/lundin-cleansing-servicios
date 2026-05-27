import pandas as pd

from src.rules.base import RuleResult


def apply(df: pd.DataFrame, operation: str, config: dict) -> RuleResult:
    document = df.get("purchase_document", pd.Series("", index=df.index)).fillna("").astype(str).str.strip()
    material = df.get("material_key", pd.Series("", index=df.index)).fillna("").astype(str).str.strip()
    valid = document.ne("") & material.ne("")

    keys = pd.DataFrame({"document": document, "material": material}, index=df.index)
    counts = keys.groupby(["document", "material"], dropna=False)["material"].transform("size")
    mask = valid & counts.gt(1)

    mark = pd.Series("", index=df.index)
    mark.loc[mask] = "MATERIAL_DUPLICADO_MISMO_CONTRATO"
    return RuleResult(
        "C02",
        mask,
        mask.map({True: "C02", False: ""}),
        updates={"mark_contract_duplicate_material": mark},
    )
