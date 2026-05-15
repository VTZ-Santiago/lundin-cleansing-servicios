import pandas as pd

from src.rules.base import RuleResult


def apply(df: pd.DataFrame, operation: str, config: dict) -> RuleResult:
    column = config.get("column", "pending_planned_qty")
    cutoff = pd.Timestamp(config.get("cutoff_date", "2025-12-31"))

    dates = pd.to_datetime(df.get("validity_end", pd.Series(pd.NaT, index=df.index)), errors="coerce")
    qty = pd.to_numeric(df.get(column, pd.Series(pd.NA, index=df.index)), errors="coerce")

    active_or_unknown = dates.isna() | (dates > cutoff)
    mask = active_or_unknown & qty.notna() & (qty < 0)
    mark = pd.Series("", index=df.index)
    mark.loc[mask] = "PENDIENTE_NEGATIVO_NO_VENCIDO"

    return RuleResult(
        rule_id="R04",
        mask=mask,
        reason=mask.map({True: "R04", False: ""}),
        updates={"mark_pending_negative_active": mark},
    )
