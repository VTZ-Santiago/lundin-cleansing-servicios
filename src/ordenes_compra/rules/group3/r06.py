import pandas as pd

from src.rules.base import RuleResult


def apply(df: pd.DataFrame, operation: str, config: dict) -> RuleResult:
    assignment_types = {
        str(value).strip().upper()
        for value in config.get("account_assignment_types", ["K"])
    }
    account_assignment = (
        df.get("account_assignment_type", pd.Series("", index=df.index))
        .fillna("")
        .astype(str)
        .str.strip()
        .str.upper()
    )

    mask = account_assignment.isin(assignment_types)
    mark = pd.Series("", index=df.index)
    mark.loc[mask] = "CARGO_DIRECTO_CENTRO_COSTO"

    return RuleResult(
        rule_id="R06",
        mask=mask,
        reason=mask.map({True: "R06", False: ""}),
        updates={"mark_direct_cost_center_service": mark},
    )