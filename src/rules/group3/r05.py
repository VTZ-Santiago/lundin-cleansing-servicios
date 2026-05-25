import pandas as pd

from src.rules.base import RuleResult
from src.utils.date_logic import effective_validity_dates


def apply(df: pd.DataFrame, operation: str, config: dict) -> RuleResult:
    year = int(config.get("year", 2026))
    dates = effective_validity_dates(df)

    mask = dates.notna() & (dates.dt.year == year)
    mark = pd.Series("", index=df.index)
    mark.loc[mask] = f"VENCE_{year}"

    return RuleResult(
        rule_id="R05",
        mask=mask,
        reason=mask.map({True: "R05", False: ""}),
        updates={f"mark_validity_{year}": mark},
    )