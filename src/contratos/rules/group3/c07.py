import pandas as pd

from src.rules.base import RuleResult


def apply(df: pd.DataFrame, operation: str, config: dict) -> RuleResult:
    material = df.get("material_key", pd.Series("", index=df.index)).fillna("").astype(str).str.strip()
    vendor = df.get("vendor", pd.Series("", index=df.index)).fillna("").astype(str).str.strip()
    document = df.get("purchase_document", pd.Series("", index=df.index)).fillna("").astype(str).str.strip()
    valid = material.ne("") & vendor.ne("") & document.ne("")

    counts = (
        pd.DataFrame({"material": material[valid], "vendor": vendor[valid], "document": document[valid]})
        .groupby(["material", "vendor"])["document"]
        .nunique()
    )
    duplicate_pairs = set(counts[counts > 1].index)
    keys = list(zip(material, vendor))
    mask = valid & pd.Series([key in duplicate_pairs for key in keys], index=df.index)

    mark = pd.Series("", index=df.index)
    mark.loc[mask] = "PROVEEDOR_MATERIAL_EN_MULTIPLES_CONTRATOS"
    return RuleResult(
        "C07",
        mask,
        mask.map({True: "C07", False: ""}),
        updates={"mark_duplicate_vendor_material_contracts": mark},
    )
