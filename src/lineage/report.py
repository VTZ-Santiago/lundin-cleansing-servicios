from dataclasses import dataclass, field

import pandas as pd

from src.lineage.records import FieldLineageRecord


@dataclass
class LineageReport:
    records: list[FieldLineageRecord] = field(default_factory=list)

    def add(self, record: FieldLineageRecord) -> None:
        self.records.append(record)

    def mapped(self) -> list[FieldLineageRecord]:
        return [r for r in self.records if r.mapping_status == "MAPPED"]

    def unmapped(self) -> list[FieldLineageRecord]:
        return [r for r in self.records if r.mapping_status == "UNMAPPED"]

    def injected(self) -> list[FieldLineageRecord]:
        return [r for r in self.records if r.mapping_status == "INJECTED"]

    def to_dataframe(self) -> pd.DataFrame:
        return pd.DataFrame([r.to_dict() for r in self.records])
