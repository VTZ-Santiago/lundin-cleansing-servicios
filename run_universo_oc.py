"""Snapshot del universo de Órdenes de Compra contado sobre las POSICIONES
(órdenes de compra totales), no sobre documentos únicos.

El reporte estadístico oficial agrega los tipos de posición y clases
documentales a nivel documento (tipo/clase dominante por contrato) y, además,
está fechado en mayo. La presentación necesita estos cortes sobre el total de
posiciones y al día. Para eso reconstruimos el universo C1 de OC con el MISMO
loader que usa la segmentación (src/segmentacion/po_loader.load_purchase_orders),
que funciona para MLCC y CCMC y normaliza el código numérico de SAP
("9"→D servicio, "2"→K consignación, "3"→L subcontratación, "0"→Stock).

Salida: outputs/control_points/universo_oc_<OP>.json (consumido por el deck).
"""
from __future__ import annotations

import json
import re
from pathlib import Path

import pandas as pd

from src.segmentacion.po_loader import load_purchase_orders

ROOT = Path(__file__).resolve().parent


# El po_loader normaliza los códigos internos de SAP que usa la segmentación
# (0→Stock, 2→K, 3→L, 9→D), pero deja pasar los que no le interesan. Aquí
# completamos esos residuales (presentes en MLCC) para etiquetarlos en el deck:
# 7 = traslado (V), 1 = tope/límite (P), 5 = terceros (S), 6 = traslado (U).
_PSTYP_LEFTOVER = {"1": "P", "4": "L", "5": "S", "6": "U", "7": "V"}


def _text(series: pd.Series) -> pd.Series:
    return series.fillna("").astype(str).str.strip()


def _counts(series: pd.Series, *, blank_label: str | None, remap: dict[str, str] | None = None) -> dict[str, int]:
    values = _text(series)
    if remap:
        values = values.map(lambda v: remap.get(v, v))
    out: dict[str, int] = {}
    for raw, n in values.value_counts(dropna=False).items():
        key = raw if raw else blank_label
        if key is None:
            continue
        out[key] = out.get(key, 0) + int(n)
    return out


def _periods(df: pd.DataFrame) -> list[str]:
    years: set[int] = set()
    if "_source_file" in df.columns:
        for name in df["_source_file"].dropna().unique():
            for token in re.findall(r"\d{4}", str(name)):
                year = int(token)
                if 2000 <= year <= 2100:
                    years.add(year)
    if not years:
        return []
    return [f"{min(years)}-{max(years)}"]


def _universe_payload(df: pd.DataFrame, operation: str) -> dict:
    docs = _text(df["purchase_document"]) if "purchase_document" in df.columns else pd.Series([], dtype=str)
    docs_nonblank = docs[docs != ""]
    unique_documents = int(docs_nonblank.nunique())

    framework_contracts = po_with_framework = 0
    if "outline_contract" in df.columns:
        outline = _text(df["outline_contract"])
        has_outline = outline != ""
        framework_contracts = int(outline[has_outline].nunique())
        po_with_framework = int(docs[has_outline][docs[has_outline] != ""].nunique())
    po_without_framework = max(unique_documents - po_with_framework, 0)
    pct = round(po_with_framework / unique_documents * 100, 2) if unique_documents else 0.0

    return {
        "operation": operation,
        "total_positions": int(len(df)),
        "unique_documents": unique_documents,
        "framework_contracts": framework_contracts,
        "po_with_framework": po_with_framework,
        "po_without_framework": po_without_framework,
        "pct_po_with_framework": pct,
        "periods": _periods(df),
        "by_position_type": _counts(df["position_type"], blank_label="Sin tipo", remap=_PSTYP_LEFTOVER) if "position_type" in df.columns else {},
        "by_doc_class": _counts(df["purchase_doc_class"], blank_label=None) if "purchase_doc_class" in df.columns else {},
    }


def main() -> None:
    out_dir = ROOT / "outputs" / "control_points"
    out_dir.mkdir(parents=True, exist_ok=True)
    for operation in ("MLCC", "CCMC"):
        result = load_purchase_orders(ROOT / "inputs", operation)
        payload = _universe_payload(result.dataframe, operation)
        path = out_dir / f"universo_oc_{operation}.json"
        path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
        print(
            f"{operation}: posiciones={payload['total_positions']:,} | docs={payload['unique_documents']:,} | "
            f"con marco={payload['po_with_framework']:,} ({payload['pct_po_with_framework']}%)"
        )
        print(f"  tipos de posición (por posición): {payload['by_position_type']}")
        print(f"  -> {path}")


if __name__ == "__main__":
    main()
