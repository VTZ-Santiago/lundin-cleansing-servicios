import pandas as pd

from src.rules.base import RuleResult


def apply(df: pd.DataFrame, operation: str, config: dict) -> RuleResult:
    pt = df["position_type"].fillna("").astype(str).str.strip()
    mask = pt != "D"
    return RuleResult("R05", mask, pd.Series([""] * len(df), index=df.index))
