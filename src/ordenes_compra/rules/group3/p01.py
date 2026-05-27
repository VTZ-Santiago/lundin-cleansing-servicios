import pandas as pd

from src.rules.base import RuleResult


def _open_balance_mask(df: pd.DataFrame, config: dict) -> pd.Series:
    qty_column = config.get("open_qty_column", "pending_delivery_qty")
    value_column = config.get("open_value_column", "pending_delivery_value")
    qty = pd.to_numeric(df.get(qty_column, pd.Series(pd.NA, index=df.index)), errors="coerce")
    value = pd.to_numeric(df.get(value_column, pd.Series(pd.NA, index=df.index)), errors="coerce")
    fallback_qty = pd.to_numeric(df.get("pending_planned_qty", pd.Series(pd.NA, index=df.index)), errors="coerce")
    return qty.gt(0) | value.gt(0) | fallback_qty.gt(0)


def _purchase_requisition_mask(df: pd.DataFrame, config: dict) -> pd.Series:
    if not config.get("require_purchase_requisition", True):
        return pd.Series(True, index=df.index)

    column = config.get("purchase_requisition_column", "purchase_requisition")
    requisition = df.get(column)
    if requisition is None:
        return pd.Series(False, index=df.index)

    normalized = requisition.fillna("").astype(str).str.strip()
    empty_values = {"", "0", "0.0", "nan", "none", "nat"}
    return ~normalized.str.lower().isin(empty_values)


def apply(df: pd.DataFrame, operation: str, config: dict) -> RuleResult:
    reference_date = pd.Timestamp(config.get("reference_date", "2026-05-25"))
    delivery = pd.to_datetime(df.get("delivery_date", pd.Series(pd.NaT, index=df.index)), errors="coerce")
    mask = (
        delivery.notna()
        & delivery.lt(reference_date)
        & _open_balance_mask(df, config)
        & _purchase_requisition_mask(df, config)
    )

    mark = pd.Series("", index=df.index)
    mark.loc[mask] = "OC_ENTREGA_VENCIDA_PENDIENTE"
    return RuleResult(
        "P01",
        mask,
        mask.map({True: "P01", False: ""}),
        updates={"mark_po_overdue_open_delivery": mark},
    )
