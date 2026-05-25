from dataclasses import dataclass, field
from datetime import datetime

import pandas as pd

from src.lineage.records import StageManifest
from src.schema.canonical import CANONICAL_BY_NAME, CANONICAL_NAMES
from src.utils.date_logic import effective_validity_dates


@dataclass
class ColumnProfile:
    canonical_name: str
    column_group: str
    is_critical: bool
    null_count: int
    null_pct: float
    unique_count: int
    total_count: int
    sample_values: list
    completeness_pct: float
    min_val: object = None
    max_val: object = None
    mean_val: object = None

    def to_dict(self) -> dict:
        return {
            "canonical_name": self.canonical_name,
            "column_group": self.column_group,
            "is_critical": self.is_critical,
            "completeness_pct": self.completeness_pct,
            "null_count": self.null_count,
            "unique_count": self.unique_count,
            "total_count": self.total_count,
            "min_val": str(self.min_val) if self.min_val is not None else "",
            "max_val": str(self.max_val) if self.max_val is not None else "",
            "mean_val": str(self.mean_val) if self.mean_val is not None else "",
            "sample_values": " | ".join(str(v) for v in self.sample_values[:5]),
        }


@dataclass
class ProfilingResult:
    operation: str
    profiled_at: datetime
    total_rows: int
    total_columns: int
    column_profiles: list[ColumnProfile]
    effective_validity_by_year: dict[int, int] = field(default_factory=dict)
    vendor_top10: list[tuple[str, int]] = field(default_factory=list)
    plant_distribution: dict[str, int] = field(default_factory=dict)
    source_file_counts: dict[str, int] = field(default_factory=dict)
    purchase_group_distribution: dict[str, int] = field(default_factory=dict)
    purchase_document_count_by_group: dict[str, int] = field(default_factory=dict)
    estimated_value_total: float = 0.0
    estimated_value_by_group: dict[str, float] = field(default_factory=dict)


class DataProfiler:
    def profile(
        self, df: pd.DataFrame, operation: str
    ) -> tuple[ProfilingResult, StageManifest]:
        started = datetime.now()
        present = [c for c in CANONICAL_NAMES if c in df.columns]

        profiles: list[ColumnProfile] = []
        for col in present:
            series = df[col]
            null_count = int(series.isna().sum())
            total = max(len(df), 1)
            null_pct = round(null_count / total * 100, 2)
            completeness = round((1 - null_count / total) * 100, 2)
            unique_count = int(series.nunique(dropna=True))
            samples = [str(v) for v in series.dropna().iloc[:5].tolist()]

            col_def = CANONICAL_BY_NAME.get(col)
            col_group = col_def.column_group if col_def else "metadata"
            is_critical = col_def.is_critical if col_def else False

            min_val = max_val = mean_val = None
            if col_def and col_def.dtype in ("date", "float"):
                non_null = series.dropna()
                if len(non_null):
                    if col_def.dtype == "date":
                        min_val = str(non_null.min().date()) if hasattr(non_null.min(), 'date') else str(non_null.min())
                        max_val = str(non_null.max().date()) if hasattr(non_null.max(), 'date') else str(non_null.max())
                    else:
                        min_val = round(float(non_null.min()), 4)
                        max_val = round(float(non_null.max()), 4)
                        mean_val = round(float(non_null.mean()), 4)

            profiles.append(ColumnProfile(
                canonical_name=col,
                column_group=col_group,
                is_critical=is_critical,
                null_count=null_count,
                null_pct=null_pct,
                unique_count=unique_count,
                total_count=len(df),
                sample_values=samples,
                completeness_pct=completeness,
                min_val=min_val,
                max_val=max_val,
                mean_val=mean_val,
            ))

        # Contract-specific insights
        validity_by_year: dict[int, int] = {}
        if "validity_end" in df.columns or "delivery_date" in df.columns:
            dates = effective_validity_dates(df).dropna()
            validity_by_year = {
                int(k): int(v)
                for k, v in dates.dt.year.value_counts().sort_index().items()
            }

        vendor_top10: list[tuple[str, int]] = []
        if "vendor" in df.columns:
            vendor_top10 = [
                (str(k), int(v))
                for k, v in df["vendor"].value_counts().head(10).items()
            ]

        plant_dist: dict[str, int] = {}
        if "plant_code" in df.columns:
            plant_dist = {
                str(k): int(v)
                for k, v in df["plant_code"].value_counts().items()
            }

        source_counts: dict[str, int] = {}
        if "_source_file" in df.columns:
            source_counts = {
                str(k): int(v)
                for k, v in df["_source_file"].value_counts().items()
            }

        purchase_group_distribution: dict[str, int] = {}
        purchase_document_count_by_group: dict[str, int] = {}
        estimated_value_total = 0.0
        estimated_value_by_group: dict[str, float] = {}
        if "purchase_group" in df.columns:
            purchase_group = df["purchase_group"].fillna("SIN_GRUPO").astype(str).str.strip()
            purchase_group = purchase_group.where(purchase_group != "", other="SIN_GRUPO")
            purchase_group_distribution = {
                str(k): int(v)
                for k, v in purchase_group.value_counts().items()
            }

            if "purchase_document" in df.columns:
                purchase_document_count_by_group = {
                    str(k): int(v)
                    for k, v in (
                        df.assign(_purchase_group=purchase_group)
                        .groupby("_purchase_group")["purchase_document"]
                        .nunique()
                        .items()
                    )
                }

            if "estimated_value" in df.columns:
                amount = pd.to_numeric(df["estimated_value"], errors="coerce")
                estimated_value_total = round(float(amount.sum(skipna=True)), 2)
                estimated_value_by_group = {
                    str(k): round(float(v), 2)
                    for k, v in (
                        df.assign(_purchase_group=purchase_group, _estimated_value=amount)
                        .groupby("_purchase_group")["_estimated_value"]
                        .sum()
                        .sort_values(ascending=False)
                        .items()
                    )
                }
        elif "estimated_value" in df.columns:
            amount = pd.to_numeric(df["estimated_value"], errors="coerce")
            estimated_value_total = round(float(amount.sum(skipna=True)), 2)

        result = ProfilingResult(
            operation=operation,
            profiled_at=started,
            total_rows=len(df),
            total_columns=len(df.columns),
            column_profiles=profiles,
            effective_validity_by_year=validity_by_year,
            vendor_top10=vendor_top10,
            plant_distribution=plant_dist,
            source_file_counts=source_counts,
            purchase_group_distribution=purchase_group_distribution,
            purchase_document_count_by_group=purchase_document_count_by_group,
            estimated_value_total=estimated_value_total,
            estimated_value_by_group=estimated_value_by_group,
        )

        critical_below_95 = sum(
            1 for p in profiles if p.is_critical and p.completeness_pct < 95
        )
        manifest = StageManifest(
            stage_id="PROFILING",
            label="Profiling de datos",
            started_at=started,
            completed_at=datetime.now(),
            rows_in=len(df),
            rows_out=len(df),
            columns_in=len(present),
            columns_out=len(present),
            notes=f"Columnas críticas con <95% completitud: {critical_below_95}",
        )

        return result, manifest
