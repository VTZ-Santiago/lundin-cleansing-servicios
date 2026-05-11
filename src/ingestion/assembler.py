from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path

import pandas as pd

from src.config.settings import Settings
from src.lineage.records import FieldLineageRecord, StageManifest
from src.lineage.report import LineageReport
from src.schema.field_map_mlcc import (
    MLCC_DTYPE_COERCIONS,
    MLCC_RAW_TO_CANONICAL,
    normalize_header,
)

_FIELD_MAPS: dict[str, tuple[dict, dict]] = {
    "MLCC": (MLCC_RAW_TO_CANONICAL, MLCC_DTYPE_COERCIONS),
}


@dataclass
class AssemblyStats:
    """Estadísticas de limpieza del proceso de ensamblado."""
    files_loaded: list[str] = field(default_factory=list)
    raw_rows_per_file: dict[str, int] = field(default_factory=dict)
    empty_rows_dropped: int = 0
    ffill_rows_recovered: int = 0
    total_raw: int = 0
    total_clean: int = 0
    unique_documents: int = 0
    unique_doc_pos_pairs: int = 0
    duplicate_doc_pos: int = 0


def _map_headers(
    df_raw: pd.DataFrame,
    raw_to_canonical: dict[str, str],
    source_file: str,
) -> tuple[pd.DataFrame, list[FieldLineageRecord]]:
    rename_map: dict[str, str] = {}
    records: list[FieldLineageRecord] = []

    for col in df_raw.columns:
        normalized = normalize_header(str(col))
        if normalized in raw_to_canonical:
            canonical = raw_to_canonical[normalized]
            rename_map[col] = canonical
            null_count = int(df_raw[col].isna().sum())
            total = max(len(df_raw), 1)
            samples = [str(v) for v in df_raw[col].dropna().unique()[:5].tolist()]
            records.append(FieldLineageRecord(
                source_file=source_file,
                raw_name=col,
                canonical_name=canonical,
                mapping_status="MAPPED",
                null_count=null_count,
                null_pct=round(null_count / total * 100, 2),
                unique_count=int(df_raw[col].nunique(dropna=True)),
                sample_values=samples,
            ))
        else:
            raw_prefix = f"_raw__{col}"
            rename_map[col] = raw_prefix
            null_count = int(df_raw[col].isna().sum())
            records.append(FieldLineageRecord(
                source_file=source_file,
                raw_name=col,
                canonical_name=raw_prefix,
                mapping_status="UNMAPPED",
                null_count=null_count,
                null_pct=round(null_count / max(len(df_raw), 1) * 100, 2),
                unique_count=int(df_raw[col].nunique(dropna=True)),
                sample_values=[],
            ))

    return df_raw.rename(columns=rename_map), records


def _apply_dtype_coercions(df: pd.DataFrame, coercions: dict[str, str]) -> pd.DataFrame:
    for col, dtype in coercions.items():
        if col not in df.columns:
            continue
        if dtype == "date":
            df[col] = pd.to_datetime(df[col], errors="coerce")
        elif dtype == "float":
            df[col] = pd.to_numeric(df[col], errors="coerce")
        elif dtype == "str":
            df[col] = df[col].where(df[col].isna(), df[col].astype(str).str.strip())
    return df


class ContractAssembler:
    def __init__(self, settings: Settings):
        self.settings = settings

    def build(
        self, operation: str
    ) -> tuple[pd.DataFrame, LineageReport, list[StageManifest], AssemblyStats]:
        started = datetime.now()
        op = operation.upper()
        if op not in _FIELD_MAPS:
            raise ValueError(f"Unknown operation: {op}. Supported: {list(_FIELD_MAPS)}")

        raw_to_canonical, dtype_coercions = _FIELD_MAPS[op]
        files = self.settings.input_files(op)
        all_frames: list[pd.DataFrame] = []
        all_records: list[FieldLineageRecord] = []
        stats = AssemblyStats()

        for path in files:
            df_raw = pd.read_excel(path, dtype=object)
            df_mapped, records = _map_headers(df_raw, raw_to_canonical, path.name)
            df_mapped = df_mapped.copy()  # defragment before adding new columns

            # Forward-fill purchase_document within this file.
            # SAP exports the contract number only in the first row of each group.
            if "purchase_document" in df_mapped.columns:
                null_before = int(df_mapped["purchase_document"].isna().sum())
                df_mapped["purchase_document"] = df_mapped["purchase_document"].ffill()
                null_after = int(df_mapped["purchase_document"].isna().sum())
                stats.ffill_rows_recovered += (null_before - null_after)

            df_mapped["_source_file"] = path.name
            df_mapped["_operation"] = op

            stats.files_loaded.append(path.name)
            stats.raw_rows_per_file[path.name] = len(df_mapped)
            all_frames.append(df_mapped)
            all_records.extend(records)

        stats.total_raw = sum(len(f) for f in all_frames)

        ingestion_manifest = StageManifest(
            stage_id="INGESTION",
            label="Carga de archivos Excel",
            started_at=started,
            completed_at=datetime.now(),
            rows_in=0,
            rows_out=stats.total_raw,
            columns_in=0,
            columns_out=0,
            notes=f"Archivos cargados: {', '.join(p.name for p in files)}",
        )

        schema_started = datetime.now()
        # outer join fills missing columns with NaN (handles 29 vs 110 column mismatch)
        master = pd.concat(all_frames, ignore_index=True, join="outer")
        master = _apply_dtype_coercions(master, dtype_coercions)

        # Drop rows that are completely empty (no purchase_document AND no position)
        empty_mask = master["purchase_document"].isna() & master["position"].isna()
        stats.empty_rows_dropped = int(empty_mask.sum())
        master = master[~empty_mask].reset_index(drop=True)

        stats.total_clean = len(master)
        stats.unique_documents = int(master["purchase_document"].nunique())
        stats.unique_doc_pos_pairs = int(
            master.dropna(subset=["purchase_document", "position"])
                  .drop_duplicates(subset=["purchase_document", "position"])
                  .shape[0]
        )
        stats.duplicate_doc_pos = stats.total_clean - stats.unique_doc_pos_pairs

        # Build lineage report — deduplicate: files share the same core headers
        report = LineageReport()
        seen: set[tuple[str, str]] = set()
        for r in all_records:
            key = (r.raw_name, r.mapping_status)
            if key not in seen:
                seen.add(key)
                report.add(r)

        # Injected auxiliary columns
        for aux_col in ("_source_file", "_operation"):
            report.add(FieldLineageRecord(
                source_file="(pipeline)",
                raw_name="",
                canonical_name=aux_col,
                mapping_status="INJECTED",
                null_count=0,
                null_pct=0.0,
                unique_count=int(master[aux_col].nunique()),
                sample_values=list(master[aux_col].unique()[:5]),
            ))

        schema_manifest = StageManifest(
            stage_id="SCHEMA_NORM",
            label="Normalizacion de esquema",
            started_at=schema_started,
            completed_at=datetime.now(),
            rows_in=stats.total_raw,
            rows_out=stats.total_clean,
            columns_in=len(master.columns),
            columns_out=len(master.columns),
            notes=(
                f"Mapped: {len(report.mapped())}, "
                f"Unmapped: {len(report.unmapped())}, "
                f"Injected: {len(report.injected())}. "
                f"Vacias descartadas: {stats.empty_rows_dropped}. "
                f"Recuperadas por ffill: {stats.ffill_rows_recovered}."
            ),
        )

        return master, report, [ingestion_manifest, schema_manifest], stats
