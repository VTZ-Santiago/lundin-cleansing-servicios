import pandas as pd

from src.rules.base import RuleResult


def apply(df: pd.DataFrame, operation: str, config: dict) -> RuleResult:
    qty = pd.to_numeric(df["pending_planned_qty"], errors="coerce")
    mask = qty.notna() & (qty > 0)
    return RuleResult(
        rule_id="R02",
        mask=mask,
        reason=pd.Series([""] * len(df), index=df.index),
    )
