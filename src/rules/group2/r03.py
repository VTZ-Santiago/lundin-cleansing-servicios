import pandas as pd

from src.rules.base import RuleResult
from src.utils.date_logic import effective_validity_dates


def apply(df: pd.DataFrame, operation: str, config: dict) -> RuleResult:
    column = config.get("column", "pending_planned_qty")
    cutoff = pd.Timestamp(config.get("cutoff_date", "2025-12-31"))
    allowed_position_type = str(config.get("allowed_position_type", "D")).strip().upper()
    rescue_reasons = set(config.get("rescue_exclusion_reasons", ["R01"]))

    dates = effective_validity_dates(df)
    qty = pd.to_numeric(df.get(column, pd.Series(pd.NA, index=df.index)), errors="coerce")
    position_type = (
        df.get("position_type", pd.Series("", index=df.index))
        .fillna("")
        .astype(str)
        .str.strip()
        .str.upper()
    )
    exclusion_reason = df.get("exclusion_reason", pd.Series("", index=df.index)).fillna("")

    mask = (
        dates.notna()
        & (dates <= cutoff)
        & qty.notna()
        & (qty > 0)
        & (position_type == allowed_position_type)
        & exclusion_reason.isin(rescue_reasons)
    )
    reason = mask.map({True: "R03", False: ""})
    return RuleResult("R03", mask, reason)