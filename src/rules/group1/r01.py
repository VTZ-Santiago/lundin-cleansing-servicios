import pandas as pd

from src.rules.base import RuleResult
from src.utils.date_logic import effective_validity_dates


def apply(df: pd.DataFrame, operation: str, config: dict) -> RuleResult:
    """Exclude rows whose effective validity date is on or before cutoff_date.

    The flow uses ``validity_end`` first and falls back to ``delivery_date`` when
    the primary field is empty. Rows with no effective date are not excluded.
    """
    cutoff = pd.Timestamp(config["cutoff_date"])
    dates = effective_validity_dates(df)
    mask = dates.notna() & (dates <= cutoff)
    reason = mask.map({True: "R01", False: ""})
    return RuleResult("R01", mask, reason)
