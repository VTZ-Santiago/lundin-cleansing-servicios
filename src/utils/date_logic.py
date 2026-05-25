import pandas as pd


def effective_validity_dates(
    df: pd.DataFrame,
    primary_column: str = "validity_end",
    fallback_column: str = "delivery_date",
) -> pd.Series:
    """Return the effective validity date used by the flow.

    The primary source is ``validity_end``. When it is empty or invalid,
    ``delivery_date`` is used as a fallback when present.
    """
    primary = pd.to_datetime(
        df.get(primary_column, pd.Series(pd.NaT, index=df.index)),
        errors="coerce",
    )
    fallback = pd.to_datetime(
        df.get(fallback_column, pd.Series(pd.NaT, index=df.index)),
        errors="coerce",
    )
    return primary.fillna(fallback)