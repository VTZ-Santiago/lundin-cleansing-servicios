import pandas as pd

from src.rules.base import RuleResult


def _effective_date(df: pd.DataFrame, config: dict) -> pd.Series:
    primary_column = config.get("primary_date_column", "delivery_date")
    fallback_column = config.get("fallback_date_column", "validity_end")
    primary = pd.to_datetime(df.get(primary_column, pd.Series(pd.NaT, index=df.index)), errors="coerce")
    fallback = pd.to_datetime(df.get(fallback_column, pd.Series(pd.NaT, index=df.index)), errors="coerce")
    return primary.fillna(fallback)


def _numeric_column(df: pd.DataFrame, column: str) -> pd.Series:
    return pd.to_numeric(df.get(column, pd.Series(0, index=df.index)), errors="coerce").fillna(0)


def apply(df: pd.DataFrame, operation: str, config: dict) -> RuleResult:
    cutoff_date = pd.Timestamp(config.get("cutoff_date", "2026-06-30"))
    effective_date = _effective_date(df, config)
    pending_qty = _numeric_column(df, config.get("pending_qty_column", "pending_delivery_qty"))
    pending_value = _numeric_column(df, config.get("pending_value_column", "pending_delivery_value"))

    mask = (
        effective_date.notna()
        & effective_date.le(cutoff_date)
        & pending_qty.le(0)
        & pending_value.le(0)
    )

    return RuleResult("E01", mask, mask.map({True: "E01", False: ""}))
