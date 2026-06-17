import pandas as pd

from src.rules.base import RuleResult
from src.segmentacion.date_utils import to_datetime_ns


def _effective_date(df: pd.DataFrame, config: dict) -> pd.Series:
    primary_column = config.get("primary_date_column", "delivery_date")
    fallback_column = config.get("fallback_date_column", "validity_end")
    primary = to_datetime_ns(df.get(primary_column, pd.Series(pd.NaT, index=df.index)), index=df.index)
    if not fallback_column:
        return primary
    fallback = to_datetime_ns(df.get(fallback_column, pd.Series(pd.NaT, index=df.index)), index=df.index)
    return primary.fillna(fallback)


def _numeric_column(df: pd.DataFrame, column: str) -> pd.Series:
    return pd.to_numeric(df.get(column, pd.Series(0, index=df.index)), errors="coerce").fillna(0)


def _marked_for_deletion(df: pd.DataFrame, column: str) -> pd.Series:
    if column not in df.columns:
        return pd.Series(False, index=df.index, dtype="bool")
    values = df[column].fillna("").astype(str).str.strip().str.upper()
    return ~values.isin({"", "0", "NAN", "NONE", "NAT"})


def before_cutoff_mask(df: pd.DataFrame, config: dict) -> pd.Series:
    """Rows considered expired against the cutoff date.

    With `missing_date_means_expired`, rows without an effective date (e.g. sin
    cruce contra el reporte de contratos vigentes) also count as expired.
    """
    cutoff_raw = config.get("cutoff_date")
    if not cutoff_raw:
        return pd.Series(False, index=df.index, dtype="bool")
    cutoff = pd.to_datetime(cutoff_raw, errors="coerce")
    if pd.isna(cutoff):
        return pd.Series(False, index=df.index, dtype="bool")

    effective_date = _effective_date(df, config)
    row_mask = effective_date.notna() & effective_date.le(cutoff)
    if config.get("missing_date_means_expired", False):
        row_mask = row_mask | effective_date.isna()

    # Document-level override: if any position in a doc has a future date, keep ALL its positions
    if "purchase_document" in df.columns and row_mask.any():
        has_future = effective_date.notna() & effective_date.gt(cutoff)
        doc_has_future = has_future.groupby(df["purchase_document"]).transform("any")
        row_mask = row_mask & ~doc_has_future

    return row_mask


def apply(df: pd.DataFrame, operation: str, config: dict) -> RuleResult:
    pending_qty = _numeric_column(df, config.get("pending_qty_column", "pending_delivery_qty"))
    pending_value = _numeric_column(df, config.get("pending_value_column", "pending_delivery_value"))
    deletion_flag = _marked_for_deletion(df, config.get("deletion_flag_column", "deletion_flag"))
    before_cutoff = before_cutoff_mask(df, config)

    no_pending_balance = pending_qty.le(0) & pending_value.le(0)
    mask = before_cutoff | no_pending_balance | deletion_flag

    return RuleResult("E01", mask, mask.map({True: "E01", False: ""}))
