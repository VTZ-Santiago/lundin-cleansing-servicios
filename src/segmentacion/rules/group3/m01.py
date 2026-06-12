"""M01 — Categorise every row by migration status for COMPLEMENTARIA (non-D) positions.

Categories (written to `output_column`):
  VIGENTE              — effective_date > cutoff  (migra)
  NO_VIGENTE_CON_SALDO — effective_date <= cutoff AND pending_value > 0  (debe migrar)
  NO_VIGENTE_SIN_SALDO — effective_date <= cutoff AND pending_value = 0  (NO migra)
  TIPO_D (configurable) — position_type is in excluded_position_types

Also writes `year_column` with the year of the effective date (nullable int).
"""
from __future__ import annotations

import pandas as pd

from src.rules.base import RuleResult


def _effective_date(df: pd.DataFrame, config: dict) -> pd.Series:
    primary = pd.to_datetime(
        df.get(config.get("primary_date_column", "delivery_date"), pd.Series(pd.NaT, index=df.index)),
        errors="coerce",
    )
    fallback = pd.to_datetime(
        df.get(config.get("fallback_date_column", "validity_end"), pd.Series(pd.NaT, index=df.index)),
        errors="coerce",
    )
    return primary.fillna(fallback)


def apply(df: pd.DataFrame, operation: str, config: dict) -> RuleResult:
    cutoff = pd.to_datetime(config.get("cutoff_date", "2026-06-30"), errors="coerce")
    value_col = config.get("value_column", "pending_delivery_value")
    excluded_types = [str(t).upper() for t in config.get("excluded_position_types", ["D"])]

    cat_vigente = config.get("category_vigente", "VIGENTE")
    cat_vencido_saldo = config.get("category_vencido_con_saldo", "NO_VIGENTE_CON_SALDO")
    cat_vencido_sin_saldo = config.get("category_vencido_sin_saldo", "NO_VIGENTE_SIN_SALDO")
    cat_excluded = config.get("category_excluded_type", "TIPO_D")
    out_col = config.get("output_column", "migration_category")
    year_col = config.get("year_column", "delivery_year")

    included_types = [str(t).upper() for t in config.get("included_position_types", [])]

    pos_type = df.get("position_type", pd.Series("", index=df.index))
    normalized_pos = pos_type.fillna("").astype(str).str.strip().str.upper()

    if included_types:
        # Allowlist: only these types are COMPLEMENTARIA; everything else → cat_excluded
        is_excluded_type = ~normalized_pos.isin(included_types)
    else:
        # Denylist (default): only excluded_types → cat_excluded; vacíos incluidos
        is_excluded_type = normalized_pos.isin(excluded_types)

    effective = _effective_date(df, config)
    pending_value = pd.to_numeric(
        df.get(value_col, pd.Series(0.0, index=df.index)), errors="coerce"
    ).fillna(0.0)

    is_expired = effective.notna() & effective.le(cutoff)
    has_saldo = pending_value > 0
    non_d = ~is_excluded_type

    category = pd.Series("", index=df.index, dtype="object")
    category[is_excluded_type] = cat_excluded
    category[non_d & ~is_expired] = cat_vigente
    category[non_d & is_expired & has_saldo] = cat_vencido_saldo
    category[non_d & is_expired & ~has_saldo] = cat_vencido_sin_saldo

    year_series = effective.dt.year.where(effective.notna(), other=pd.NA).astype("Int64")

    return RuleResult(
        "M01",
        mask=pd.Series(True, index=df.index, dtype="bool"),
        reason=pd.Series("M01", index=df.index, dtype="object"),
        updates={out_col: category, year_col: year_series},
    )
