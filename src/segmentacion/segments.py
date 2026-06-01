from __future__ import annotations

from dataclasses import dataclass

import pandas as pd

from src.segmentacion.characterization import SegmentRule, normalize_text


SEGMENT_LABELS = {
    "contratos": "Contratos",
    "ordenes_servicio": "Ordenes de Servicio",
}


@dataclass(frozen=True)
class SegmentMatchResult:
    masks: dict[str, pd.Series]
    overlaps: pd.DataFrame
    unclassified: pd.DataFrame


def _empty_like(df: pd.DataFrame) -> pd.Series:
    return pd.Series(False, index=df.index, dtype="bool")


def _normalized_series(df: pd.DataFrame, column: str) -> pd.Series:
    if column not in df.columns:
        return pd.Series([""] * len(df), index=df.index, dtype="object")
    return df[column].map(normalize_text).fillna("")


def _present_mask(series: pd.Series) -> pd.Series:
    return ~series.isin({"", "0", "0 0", "NAN", "NONE", "NAT"})


def _rule_mask(df: pd.DataFrame, rule: SegmentRule) -> pd.Series:
    mask = pd.Series(True, index=df.index, dtype="bool")
    normalized_cache: dict[str, pd.Series] = {}
    for condition in rule.conditions:
        series = normalized_cache.setdefault(condition.column, _normalized_series(df, condition.column))
        if condition.operator == "present":
            mask &= _present_mask(series)
        elif condition.operator == "absent":
            mask &= ~_present_mask(series)
        elif condition.operator == "equals":
            mask &= series.eq(condition.value)
        else:
            raise ValueError(f"Unsupported condition operator: {condition.operator}")
    return mask


def apply_segment_rules(df: pd.DataFrame, rules: list[SegmentRule]) -> SegmentMatchResult:
    segment_ids = sorted({rule.segment_id for rule in rules})
    masks = {segment_id: _empty_like(df) for segment_id in segment_ids}
    for rule in rules:
        masks[rule.segment_id] |= _rule_mask(df, rule)

    match_count = sum(mask.astype(int) for mask in masks.values()) if masks else pd.Series(0, index=df.index)
    overlaps = df[match_count.gt(1)].copy()
    unclassified = df[match_count.eq(0)].copy()
    return SegmentMatchResult(masks=masks, overlaps=overlaps, unclassified=unclassified)


def segment_dataframe(df: pd.DataFrame, rules: list[SegmentRule]) -> tuple[dict[str, pd.DataFrame], SegmentMatchResult]:
    result = apply_segment_rules(df, rules)
    segments = {
        segment_id: df[mask].copy()
        for segment_id, mask in result.masks.items()
    }
    return segments, result
