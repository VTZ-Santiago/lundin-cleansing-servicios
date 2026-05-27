import pandas as pd

from src.rules.base import RuleResult


def apply(df: pd.DataFrame, operation: str, config: dict) -> RuleResult:
    planning_types = {
        str(value).strip().upper()
        for value in config.get("planning_types", ["ZZ"])
    }
    planning_type = (
        df.get("material_planning_type", pd.Series("", index=df.index))
        .fillna("")
        .astype(str)
        .str.strip()
        .str.upper()
    )
    mask = planning_type.isin(planning_types)
    return RuleResult("C01", mask, mask.map({True: "C01", False: ""}))
