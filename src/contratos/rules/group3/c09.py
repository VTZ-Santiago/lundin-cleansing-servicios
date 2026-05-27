import pandas as pd

from src.rules.base import RuleResult
from src.utils.date_logic import effective_validity_dates


def apply(df: pd.DataFrame, operation: str, config: dict) -> RuleResult:
    cutoff = pd.Timestamp(config.get("cutoff_date", "2025-12-31"))
    qty_column = config.get("pending_qty_column", "pending_planned_qty")

    dates = effective_validity_dates(df)
    qty = pd.to_numeric(df.get(qty_column, pd.Series(pd.NA, index=df.index)), errors="coerce")
    no_movement = df.get("material_no_movement_24m", pd.Series(False, index=df.index)).fillna(False).astype(bool)
    mask = dates.notna() & dates.le(cutoff) & qty.gt(0) & no_movement

    mark = pd.Series("", index=df.index)
    mark.loc[mask] = "CONTRATO_VENCIDO_CON_SALDO_SIN_CONSUMO"
    return RuleResult(
        "C09",
        mask,
        mask.map({True: "C09", False: ""}),
        updates={"mark_expired_balance_no_consumption": mark},
    )