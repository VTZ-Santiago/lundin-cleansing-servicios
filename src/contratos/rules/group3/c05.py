import re
import unicodedata

import pandas as pd

from src.rules.base import RuleResult


def _normalize_text(series: pd.Series) -> pd.Series:
    return (
        series.fillna("")
        .astype(str)
        .map(lambda value: unicodedata.normalize("NFKD", value).encode("ascii", "ignore").decode("ascii"))
        .str.upper()
        .map(lambda value: re.sub(r"\s+", " ", value).strip())
    )


def apply(df: pd.DataFrame, operation: str, config: dict) -> RuleResult:
    source = _normalize_text(df.get("short_text", pd.Series("", index=df.index)))
    master = _normalize_text(df.get("material_master_description", pd.Series("", index=df.index)))
    mask = source.ne("") & master.ne("") & source.ne(master)

    mark = pd.Series("", index=df.index)
    mark.loc[mask] = "DESCRIPCION_CONTRATO_DISTINTA_MAESTRO"
    return RuleResult(
        "C05",
        mask,
        mask.map({True: "C05", False: ""}),
        updates={"mark_material_description_mismatch": mark},
    )
