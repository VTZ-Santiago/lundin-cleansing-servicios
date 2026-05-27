import pandas as pd

from src.rules.base import RuleResult


def apply(df: pd.DataFrame, operation: str, config: dict) -> RuleResult:
    no_movement = df.get("material_no_movement_24m", pd.Series(False, index=df.index)).fillna(False).astype(bool)
    material = df.get("material_key", pd.Series("", index=df.index)).fillna("").astype(str).str.strip()
    mask = no_movement & material.ne("")

    mark = pd.Series("", index=df.index)
    mark.loc[mask] = "SIN_MOVIMIENTO_ULTIMOS_24M"
    return RuleResult(
        "C04",
        mask,
        mask.map({True: "C04", False: ""}),
        updates={"mark_material_no_movement_24m": mark},
    )
