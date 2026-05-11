from dataclasses import dataclass, field
from datetime import datetime


@dataclass
class FieldLineageRecord:
    source_file: str
    raw_name: str
    canonical_name: str
    mapping_status: str  # MAPPED | UNMAPPED | INJECTED
    null_count: int = 0
    null_pct: float = 0.0
    unique_count: int = 0
    sample_values: list = field(default_factory=list)

    def to_dict(self) -> dict:
        return {
            "source_file": self.source_file,
            "raw_name": self.raw_name,
            "canonical_name": self.canonical_name,
            "mapping_status": self.mapping_status,
            "null_count": self.null_count,
            "null_pct": round(self.null_pct, 2),
            "unique_count": self.unique_count,
            "sample_values": " | ".join(str(v) for v in self.sample_values[:5]),
        }


@dataclass
class IssueRecord:
    stage: str
    severity: str  # ERROR | WARNING | INFO
    code: str
    message: str
    detail: str = ""
    row_count: int = 0

    def to_dict(self) -> dict:
        return {
            "stage": self.stage,
            "severity": self.severity,
            "code": self.code,
            "message": self.message,
            "detail": self.detail,
            "row_count": self.row_count,
        }


@dataclass
class StageManifest:
    stage_id: str
    label: str
    started_at: datetime
    completed_at: datetime
    rows_in: int
    rows_out: int
    columns_in: int
    columns_out: int
    notes: str = ""

    def to_dict(self) -> dict:
        duration = (self.completed_at - self.started_at).total_seconds()
        return {
            "stage_id": self.stage_id,
            "label": self.label,
            "started_at": self.started_at.strftime("%Y-%m-%d %H:%M:%S"),
            "completed_at": self.completed_at.strftime("%Y-%m-%d %H:%M:%S"),
            "duration_s": round(duration, 2),
            "rows_in": self.rows_in,
            "rows_out": self.rows_out,
            "columns_in": self.columns_in,
            "columns_out": self.columns_out,
            "notes": self.notes,
        }
