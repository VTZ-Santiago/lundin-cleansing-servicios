"""Utilidades de fechas robustas para el flujo de segmentación."""
from __future__ import annotations

import pandas as pd


def to_datetime_ns(values, index=None) -> pd.Series:
    """Convierte a `datetime64[ns]`, volviendo NaT cualquier fecha fuera de rango.

    pandas 2.x puede devolver resoluciones no-ns (us/s) capaces de almacenar
    fechas fuera del rango de `datetime64[ns]` (p. ej. exports SAP con años
    basura como 3034). Esas fechas provocan `OutOfBoundsDatetime` al alinear o
    rellenar; aquí se descartan a NaT para no romper el flujo.
    """
    series = pd.to_datetime(values, errors="coerce")
    if not isinstance(series, pd.Series):
        series = pd.Series(series, index=index)
    if str(series.dtype) != "datetime64[ns]":
        try:
            in_range = (series >= pd.Timestamp.min) & (series <= pd.Timestamp.max)
            series = series.where(in_range)
        except TypeError:
            pass
        series = series.astype("datetime64[ns]")
    return series
