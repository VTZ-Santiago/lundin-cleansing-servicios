import pandas as pd

from src.ordenes_compra.rules.group3.p01 import _open_balance_mask
from src.rules.base import RuleResult


def apply(df: pd.DataFrame, operation: str, config: dict) -> RuleResult:
    reference_date = pd.Timestamp(config.get("reference_date", "2026-05-25"))
    assignment_types = {
        str(value).strip().upper()
        for value in config.get("account_assignment_types", ["K"])
    }
    delivery = pd.to_datetime(df.get("delivery_date", pd.Series(pd.NaT, index=df.index)), errors="coerce")
    assignment = (
        df.get("account_assignment_type", pd.Series("", index=df.index))
        .fillna("")
        .astype(str)
        .str.strip()
        .str.upper()
    )
    mask = (
        delivery.notna()
        & delivery.lt(reference_date)
        & _open_balance_mask(df, config)
        & assignment.isin(assignment_types)
    )

    mark = pd.Series("", index=df.index)
    mark.loc[mask] = "CARGO_DIRECTO_ENTREGA_VENCIDA_NO_ENTREGADA"
    return RuleResult(
        "P05",
        mask,
        mask.map({True: "P05", False: ""}),
        updates={"mark_direct_cost_overdue_open_delivery": mark},
    )