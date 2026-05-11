from dataclasses import dataclass, field
from pathlib import Path


@dataclass
class Settings:
    project_root: Path = field(
        default_factory=lambda: Path(__file__).resolve().parents[2]
    )

    @property
    def inputs_dir(self) -> Path:
        return self.project_root / "inputs"

    @property
    def outputs_dir(self) -> Path:
        return self.project_root / "outputs"

    @property
    def control_points_dir(self) -> Path:
        return self.outputs_dir / "control_points"

    @property
    def rules_yaml_path(self) -> Path:
        return self.project_root / "src" / "rules" / "rules.yaml"

    def input_files(self, operation: str) -> list[Path]:
        op = operation.upper()
        directory = self.inputs_dir / op
        if not directory.exists():
            raise FileNotFoundError(f"Input directory not found: {directory}")
        files = list(directory.glob("*.XLSX")) + list(directory.glob("*.xlsx"))
        seen: set[str] = set()
        unique: list[Path] = []
        for f in files:
            if f.name.lower() not in seen:
                seen.add(f.name.lower())
                unique.append(f)
        unique.sort(key=lambda p: p.name.lower())
        if not unique:
            raise FileNotFoundError(f"No Excel files found in {directory}")
        return unique

    def ensure_dirs(self) -> None:
        self.control_points_dir.mkdir(parents=True, exist_ok=True)
        (self.project_root / "tmp" / "cache").mkdir(parents=True, exist_ok=True)
