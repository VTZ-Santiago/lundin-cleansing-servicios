from dataclasses import dataclass, field
import pandas as pd


@dataclass
class RuleResult:
    rule_id: str
    mask: pd.Series    # True = row is affected by this rule
    reason: pd.Series  # per-row reason string (empty string if not affected)
    updates: dict | None = None  # optional extra column updates {col: Series}
    issues: list = field(default_factory=list)
    notes: str = ""
