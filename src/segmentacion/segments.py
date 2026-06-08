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
    resolved_overlap_rows: int = 0


def _empty_like(df: pd.DataFrame) -> pd.Series:
    return pd.Series(False, index=df.index, dtype="bool")


def _normalized_series(df: pd.DataFrame, column: str) -> pd.Series:
    if column not in df.columns:
        return pd.Series([""] * len(df), index=df.index, dtype="object")
    text = df[column].where(df[column].notna(), "").astype(str).str.strip()
    text = text.str.replace(r"\.0$", "", regex=True)
    text = text.str.normalize("NFKD").str.encode("ascii", "ignore").str.decode("ascii")
    text = text.str.replace(r"\s+", " ", regex=True)
    text = text.str.replace(".", " ", regex=False)
    text = text.str.replace(r"\s+", " ", regex=True)
    return text.str.strip().str.upper().fillna("")


def _present_mask(series: pd.Series) -> pd.Series:
    return ~series.isin({"", "0", "0 0", "NAN", "NONE", "NAT"})


def _rule_mask(df: pd.DataFrame, rule: SegmentRule, normalized_cache: dict[str, pd.Series]) -> pd.Series:
    mask = pd.Series(True, index=df.index, dtype="bool")
    for condition in rule.conditions:
        if condition.column not in normalized_cache:
            normalized_cache[condition.column] = _normalized_series(df, condition.column)
        series = normalized_cache[condition.column]
        if condition.operator == "present":
            mask &= _present_mask(series)
        elif condition.operator == "absent":
            mask &= ~_present_mask(series)
        elif condition.operator == "equals":
            mask &= series.eq(condition.value)
        else:
            raise ValueError(f"Unsupported condition operator: {condition.operator}")
    return mask


def _condition_mask(
    df: pd.DataFrame,
    condition: object,
    normalized_cache: dict[str, pd.Series],
) -> pd.Series:
    if condition.column not in normalized_cache:
        normalized_cache[condition.column] = _normalized_series(df, condition.column)
    series = normalized_cache[condition.column]
    if condition.operator == "present":
        return _present_mask(series)
    if condition.operator == "absent":
        return ~_present_mask(series)
    if condition.operator == "equals":
        return series.eq(condition.value)
    raise ValueError(f"Unsupported condition operator: {condition.operator}")


def _rule_mask_from_conditions(
    df: pd.DataFrame,
    rule: SegmentRule,
    condition_cache: dict[tuple[str, str, str], pd.Series],
    normalized_cache: dict[str, pd.Series],
) -> pd.Series:
    mask = pd.Series(True, index=df.index, dtype="bool")
    for condition in rule.conditions:
        key = (condition.column, condition.operator, condition.value)
        if key not in condition_cache:
            condition_cache[key] = _condition_mask(df, condition, normalized_cache)
        mask &= condition_cache[key]
    return mask


def _resolve_masks_by_priority(
    masks: dict[str, pd.Series],
    priority: list[str] | None,
) -> dict[str, pd.Series]:
    if not priority or not masks:
        return masks

    resolved = {segment_id: _empty_like(next(iter(masks.values()))) for segment_id in masks}
    unassigned = pd.Series(True, index=next(iter(masks.values())).index, dtype="bool")
    ordered = [segment_id for segment_id in priority if segment_id in masks]
    ordered.extend(segment_id for segment_id in sorted(masks) if segment_id not in ordered)
    for segment_id in ordered:
        selected = masks[segment_id] & unassigned
        resolved[segment_id] = selected
        unassigned &= ~selected
    return resolved


def apply_segment_rules(
    df: pd.DataFrame,
    rules: list[SegmentRule],
    priority: list[str] | None = None,
) -> SegmentMatchResult:
    segment_ids = sorted({rule.segment_id for rule in rules})
    raw_masks = {segment_id: _empty_like(df) for segment_id in segment_ids}
    normalized_cache: dict[str, pd.Series] = {}
    condition_cache: dict[tuple[str, str, str], pd.Series] = {}
    for rule in rules:
        raw_masks[rule.segment_id] |= _rule_mask_from_conditions(df, rule, condition_cache, normalized_cache)

    match_count = sum(mask.astype(int) for mask in raw_masks.values()) if raw_masks else pd.Series(0, index=df.index)
    overlap_mask = match_count.gt(1)
    overlaps = df.loc[overlap_mask, []].copy() if priority else df[overlap_mask].copy()
    masks = _resolve_masks_by_priority(raw_masks, priority)
    final_match_count = sum(mask.astype(int) for mask in masks.values()) if masks else pd.Series(0, index=df.index)
    unclassified = df.loc[final_match_count.eq(0), []].copy()
    return SegmentMatchResult(
        masks=masks,
        overlaps=overlaps,
        unclassified=unclassified,
        resolved_overlap_rows=len(overlaps) if priority else 0,
    )


def segment_dataframe(
    df: pd.DataFrame,
    rules: list[SegmentRule],
    priority: list[str] | None = None,
) -> tuple[dict[str, pd.DataFrame], SegmentMatchResult]:
    result = apply_segment_rules(df, rules, priority=priority)
    segments = {
        segment_id: df[mask].copy()
        for segment_id, mask in result.masks.items()
    }
    return segments, result
