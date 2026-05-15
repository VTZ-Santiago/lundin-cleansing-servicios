import pandas as pd

from src.rules.base import RuleResult


def apply(df: pd.DataFrame, operation: str, config: dict) -> RuleResult:
    allowed_position_type = str(config.get("allowed_position_type", "D")).strip().upper()

    if "position_type" in df.columns:
        position_type = df["position_type"].fillna("").astype(str).str.strip().str.upper()
    else:
        position_type = pd.Series([""] * len(df), index=df.index)

    mask = position_type != allowed_position_type
    reason = mask.map({True: "R02", False: ""})
    return RuleResult("R02", mask, reason)