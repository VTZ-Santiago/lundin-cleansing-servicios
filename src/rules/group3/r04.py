import pandas as pd

from src.rules.base import RuleResult


def apply(df: pd.DataFrame, operation: str, config: dict) -> RuleResult:
    col = df["deletion_flag"].fillna("").astype(str).str.strip()
    mark = col.where(col != "", other="SIN_FLAG")
    return RuleResult(
        rule_id="R04",
        mask=pd.Series([True] * len(df), index=df.index),
        reason=pd.Series([""] * len(df), index=df.index),
        updates={"mark_deletion_flag": mark},
    )
