from dataclasses import dataclass, field
from pathlib import Path

from src.config.domains import get_domain_config, normalize_domain_name


@dataclass
class Settings:
    domain: str = "contratos"
    project_root: Path = field(
        default_factory=lambda: Path(__file__).resolve().parents[2]
    )

    def __post_init__(self) -> None:
        self.domain = normalize_domain_name(self.domain)

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
    def domain_outputs_dir(self) -> Path:
        return self.outputs_dir / self.domain_config.output_suffix

    @property
    def domain_control_points_dir(self) -> Path:
        return self.domain_outputs_dir / "control_points"

    @property
    def domain_config(self):
        return get_domain_config(self.domain)

    @property
    def domain_src_dir(self) -> Path:
        return self.project_root / "src" / self.domain_config.package_name

    @property
    def rules_yaml_path(self) -> Path:
        domain_path = self.domain_src_dir / "rules" / "rules.yaml"
        if domain_path.exists():
            return domain_path
        return self.project_root / "src" / "rules" / "rules.yaml"

    def input_dir(self, operation: str) -> Path:
        op = operation.upper()
        domain_dir = self.inputs_dir / op / self.domain_config.input_subdir
        if domain_dir.exists():
            return domain_dir
        return self.inputs_dir / op

    def input_files(self, operation: str) -> list[Path]:
        directory = self.input_dir(operation)
        if not directory.exists():
            raise FileNotFoundError(f"Input directory not found: {directory}")
        files = [
            f for f in list(directory.glob("*.XLSX")) + list(directory.glob("*.xlsx"))
            if not f.name.startswith("~$")
        ]
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
        self.domain_control_points_dir.mkdir(parents=True, exist_ok=True)
        self.domain_outputs_dir.mkdir(parents=True, exist_ok=True)
        (self.project_root / "tmp" / "cache").mkdir(parents=True, exist_ok=True)
