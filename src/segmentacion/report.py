from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path

import pandas as pd


@dataclass
class SegmentSummary:
    segment_id: str
    label: str
    c1_rows: int
    c1_documents: int
    c2_documents: int
    c2_no_migra_documents: int
    c1_outline_contracts: int
    c2_outline_contracts: int
    c2_no_migra_outline_contracts: int
    c2_rows: int
    c2_no_migra_rows: int
    reconciliation_ok: bool
    doc_class_counts: dict[str, int] = field(default_factory=dict)
    source_file_counts: dict[str, int] = field(default_factory=dict)

    @property
    def excluded_pct(self) -> float:
        if self.c1_rows == 0:
            return 0.0
        return round(self.c2_no_migra_rows / self.c1_rows * 100, 2)


def _count_documents(df: pd.DataFrame) -> int:
    if "purchase_document" not in df.columns or df.empty:
        return 0
    return int(df["purchase_document"].nunique(dropna=True))


def _count_outline_contracts(df: pd.DataFrame) -> int:
    if "outline_contract" not in df.columns or df.empty:
        return 0
    values = df["outline_contract"].dropna().astype(str).str.strip()
    values = values[values != ""]
    return int(values.nunique())


def _value_counts(df: pd.DataFrame, column: str) -> dict[str, int]:
    if column not in df.columns or df.empty:
        return {}
    values = df[column].fillna("SIN_VALOR").astype(str).str.strip()
    values = values.where(values != "", other="SIN_VALOR")
    return {str(key): int(value) for key, value in values.value_counts().sort_index().items()}


def build_segment_summary(
    segment_id: str,
    label: str,
    c1: pd.DataFrame,
    c2: pd.DataFrame,
    c2_no_migra: pd.DataFrame,
) -> SegmentSummary:
    return SegmentSummary(
        segment_id=segment_id,
        label=label,
        c1_rows=len(c1),
        c1_documents=_count_documents(c1),
        c2_documents=_count_documents(c2),
        c2_no_migra_documents=_count_documents(c2_no_migra),
        c1_outline_contracts=_count_outline_contracts(c1),
        c2_outline_contracts=_count_outline_contracts(c2),
        c2_no_migra_outline_contracts=_count_outline_contracts(c2_no_migra),
        c2_rows=len(c2),
        c2_no_migra_rows=len(c2_no_migra),
        reconciliation_ok=len(c1) == len(c2) + len(c2_no_migra),
        doc_class_counts=_value_counts(c1, "purchase_doc_class"),
        source_file_counts=_value_counts(c1, "_source_file"),
    )


def _dict_lines(values: dict[str, int]) -> list[str]:
    if not values:
        return ["  - Sin datos"]
    return [f"  - {key}: {value:,}" for key, value in values.items()]


def write_summary_markdown(
    output_path: Path,
    operation: str,
    total_rows: int,
    rows_read: int,
    source_files: list[str],
    segmentation_criteria_count: int,
    overlap_rows: int,
    unclassified_rows: int,
    summaries: list[SegmentSummary],
) -> Path:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    classified_rows = sum(summary.c1_rows for summary in summaries)
    lines: list[str] = [
        f"# Resumen segmentación {operation}",
        "",
        "## Auditoría global",
        f"- Filas leídas desde OC: {rows_read:,}",
        f"- Filas limpias cargadas: {total_rows:,}",
        f"- Filas clasificadas en segmentos: {classified_rows:,}",
        f"- Filas no clasificadas: {unclassified_rows:,}",
        f"- Filas con solapamiento entre segmentos: {overlap_rows:,}",
        f"- Criterios de segmentación cargados: {segmentation_criteria_count:,}",
        "- Archivos fuente:",
        *[f"  - {name}" for name in source_files],
        "",
    ]

    for summary in summaries:
        status = "OK" if summary.reconciliation_ok else "FALLA"
        lines.extend([
            f"## {summary.label}",
            f"- C1 filas: {summary.c1_rows:,}",
            f"- Documentos únicos C1: {summary.c1_documents:,}",
            f"- Documentos únicos C2: {summary.c2_documents:,}",
            f"- Documentos únicos C2_NO_MIGRA: {summary.c2_no_migra_documents:,}",
            f"- Outline contracts únicos C1: {summary.c1_outline_contracts:,}",
            f"- Outline contracts únicos C2: {summary.c2_outline_contracts:,}",
            f"- Outline contracts únicos C2_NO_MIGRA: {summary.c2_no_migra_outline_contracts:,}",
            f"- C2 migra: {summary.c2_rows:,}",
            f"- C2 no migra: {summary.c2_no_migra_rows:,}",
            f"- Porcentaje no migra: {summary.excluded_pct:.2f}%",
            f"- Reconciliación C1 = C2 + C2_NO_MIGRA: {status}",
            "- Cl. documento compras:",
            *_dict_lines(summary.doc_class_counts),
            "- Archivos fuente:",
            *_dict_lines(summary.source_file_counts),
            "",
        ])

    output_path.write_text("\n".join(lines).rstrip() + "\n", encoding="utf-8")
    return output_path
