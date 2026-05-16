import importlib
from datetime import datetime
from pathlib import Path

import pandas as pd
import yaml

from src.config.settings import Settings
from src.lineage.records import IssueRecord, StageManifest
from src.rules.base import RuleResult

_GROUP_MODULE_PREFIX_BY_DOMAIN: dict[str, dict[str, str]] = {
    "contratos": {
        "G1_EXCLUSIONS": "src.contratos.rules.group1",
        "G2_RESCUE":     "src.contratos.rules.group2",
        "G3_MARKING":    "src.contratos.rules.group3",
    },
    "ordenes_compra": {
        "G1_EXCLUSIONS": "src.ordenes_compra.rules.group1",
        "G2_RESCUE":     "src.ordenes_compra.rules.group2",
        "G3_MARKING":    "src.ordenes_compra.rules.group3",
    },
}


class RuleEngine:
    def __init__(self, settings: Settings):
        self.settings = settings
        with open(settings.rules_yaml_path, encoding="utf-8") as f:
            self.config = yaml.safe_load(f)

    def apply_group(
        self,
        group_id: str,
        df: pd.DataFrame,
        operation: str,
    ) -> tuple[pd.DataFrame, list[IssueRecord], StageManifest]:
        started = datetime.now()
        op = operation.upper()
        group_config = self.config["groups"].get(group_id)
        if group_config is None:
            raise ValueError(f"Group '{group_id}' not found in rules.yaml")

        result_df = df.copy()
        if "exclusion_reason" not in result_df.columns:
            result_df["exclusion_reason"] = ""
        if "exclusion_severity" not in result_df.columns:
            result_df["exclusion_severity"] = ""
        if "rescue_reason" not in result_df.columns:
            result_df["rescue_reason"] = ""

        all_issues: list[IssueRecord] = []
        rules_applied: list[str] = []

        module_prefix = _GROUP_MODULE_PREFIX_BY_DOMAIN[self.settings.domain].get(group_id)
        if module_prefix is None:
            raise ValueError(f"No module prefix configured for group '{group_id}'")

        for rule_def in group_config.get("rules", []):
            rule_id = rule_def["id"]
            enabled = rule_def.get("enabled", {})
            if not enabled.get(op, False):
                continue

            module_name = f"{module_prefix}.{rule_id.lower()}"
            module = importlib.import_module(module_name)

            rule_result: RuleResult = module.apply(
                result_df, op, rule_def.get("config", {})
            )

            action = rule_def.get("action", "EXCLUDE")
            severity = rule_def.get("severity", "ERROR")

            if action == "EXCLUDE":
                # First rule wins for rows not yet excluded
                not_excluded = result_df["exclusion_reason"] == ""
                update_mask = rule_result.mask & not_excluded
                result_df.loc[update_mask, "exclusion_reason"] = rule_id
                result_df.loc[update_mask, "exclusion_severity"] = severity
                flagged = int(update_mask.sum())

            elif action == "RESCUE":
                # Only rescue rows currently excluded
                excluded_mask = result_df["exclusion_reason"] != ""
                actual_rescue = rule_result.mask & excluded_mask
                result_df.loc[actual_rescue, "rescue_reason"]    = rule_id
                result_df.loc[actual_rescue, "exclusion_reason"] = ""
                result_df.loc[actual_rescue, "exclusion_severity"] = ""
                flagged = int(actual_rescue.sum())

            elif action == "MARK":
                if rule_result.updates:
                    for col_name, col_series in rule_result.updates.items():
                        result_df[col_name] = col_series.values
                flagged = int(rule_result.mask.sum())

            else:
                flagged = int(rule_result.mask.sum())

            rules_applied.append(rule_id)
            if flagged > 0:
                all_issues.append(IssueRecord(
                    stage=group_id,
                    severity=severity,
                    code=rule_id,
                    message=rule_def.get("name", rule_id),
                    detail=f"Filas afectadas: {flagged:,}",
                    row_count=flagged,
                ))

        excluded_total = int((result_df["exclusion_reason"] != "").sum())
        manifest = StageManifest(
            stage_id=group_id,
            label=group_config.get("name", group_id),
            started_at=started,
            completed_at=datetime.now(),
            rows_in=len(df),
            rows_out=len(result_df) - excluded_total,
            columns_in=len(df.columns),
            columns_out=len(result_df.columns),
            notes=(
                f"Reglas aplicadas: {', '.join(rules_applied)}. "
                f"Excluidas: {excluded_total:,}"
            ),
        )

        return result_df, all_issues, manifest
