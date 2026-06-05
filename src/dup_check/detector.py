"""Detect materials appearing in more than one active contract/PO simultaneously."""
import pandas as pd


REPORT_COLUMNS: list[str] = [
    "Operación",
    "Fuente",
    "Centro",
    "Org. Compras",
    "Contrato Marco",
    "Pos. Ctto Sup.",
    "Documento compras",
    "Posición",
    "Material",
    "Texto Breve",
    "Proveedor",
    "Cl Doc Compras",
    "Tipo Doc Compras",
    "Tipo de Posición",
    "Tipo Imputación",
    "Grupo Artículos",
    "Grupo Liberación",
    "PR/SOLPED",
    "Historial Pedido",
    "Ind. Borrado",
    "Inicio Validez",
    "Fecha Término",
    "Fecha Documento",
    "Fecha Entrega",
    "Por Entregar (Qty)",
    "Por Entregar (Valor)",
    "N° Docs Material",
    "Docs con Material",
]


def find_duplicates(df: pd.DataFrame, today: pd.Timestamp | None = None) -> pd.DataFrame:
    """Return all active rows for materials that appear in >1 distinct active document.

    "Active" means:
      1. Ind. Borrado is null or empty (no deletion flag), AND
      2. Fecha Término is null or >= today (not expired).

    Groups by (Operación, Fuente, Material) within the same DataFrame.
    Adds:
      N° Docs Material  – count of distinct active Documento compras for this material
      Docs con Material – pipe-separated sorted list of those documents
    """
    if df.empty or "Material" not in df.columns or "Documento compras" not in df.columns:
        return _empty()

    if today is None:
        today = pd.Timestamp.now().normalize()

    # Active = Ind. Borrado is null or empty string
    if "Ind. Borrado" in df.columns:
        flag = df["Ind. Borrado"].fillna("").astype(str).str.strip()
        active = df[flag == ""].copy()
    else:
        active = df.copy()

    # Active = Fecha Término is null OR >= today (not expired)
    if "Fecha Término" in active.columns:
        validity = pd.to_datetime(active["Fecha Término"], errors="coerce")
        active = active[validity.isna() | (validity >= today)].copy()

    if active.empty:
        return _empty()

    grp_cols = ["Operación", "Fuente", "Material"]
    active = active.copy()
    active["_doc"] = active["Documento compras"].fillna("").astype(str)

    stats = (
        active.groupby(grp_cols)
        .agg(
            n_docs=("_doc", lambda s: s[s != ""].nunique()),
            docs_list=("_doc", lambda s: " | ".join(sorted(s[s != ""].unique()))),
        )
        .reset_index()
        .rename(columns={"n_docs": "N° Docs Material", "docs_list": "Docs con Material"})
    )

    dup_mats = stats[stats["N° Docs Material"] > 1][
        grp_cols + ["N° Docs Material", "Docs con Material"]
    ]

    if dup_mats.empty:
        return _empty()

    result = (
        active.drop(columns=["_doc"])
        .merge(dup_mats, on=grp_cols, how="inner")
    )

    for col in REPORT_COLUMNS:
        if col not in result.columns:
            result[col] = None

    return (
        result[REPORT_COLUMNS]
        .sort_values(["Operación", "Fuente", "Material", "Documento compras", "Posición"])
        .reset_index(drop=True)
    )


def _empty() -> pd.DataFrame:
    return pd.DataFrame(columns=REPORT_COLUMNS)
