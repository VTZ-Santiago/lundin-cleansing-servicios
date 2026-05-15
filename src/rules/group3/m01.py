import pandas as pd

from src.rules.base import RuleResult


def apply(df: pd.DataFrame, operation: str, config: dict) -> RuleResult:
    col = df["position_type"].fillna("").astype(str).str.strip()
    mark = col.where(col != "", other="VACIO")
    return RuleResult(
        rule_id="M01",
        mask=pd.Series([True] * len(df), index=df.index),
        reason=pd.Series([""] * len(df), index=df.index),
        updates={"mark_position_type": mark},
    )