import pandas as pd

from src.rules.base import RuleResult


def apply(df: pd.DataFrame, operation: str, config: dict) -> RuleResult:
    """Exclude contracts with validity_end on or before cutoff_date.

    Rows with NaT validity_end are NOT excluded (unknown expiry treated as active).
    """
    cutoff = pd.Timestamp(config["cutoff_date"])
    dates = pd.to_datetime(df["validity_end"], errors="coerce")
    mask = dates.notna() & (dates <= cutoff)
    reason = mask.map({True: "R01", False: ""})
    return RuleResult("R01", mask, reason)
