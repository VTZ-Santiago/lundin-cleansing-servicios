"""Utility: detect materials assigned to more than one active contract/PO (MLCC & CCMC).

Searches two source types for each operation:
  OC    → inputs/{OP}/ordenes-compra/*.XLSX  (purchase order files)
  ME3L  → inputs/{OP}/ME3L_Ctto*.XLSX        (framework contract report)

A material is flagged when it appears in >1 distinct active Documento compras
(active = Indicador de borrado is empty).

Usage:
    python run_material_dup_check.py                   # all operations
    python run_material_dup_check.py --operation MLCC
    python run_material_dup_check.py --operation CCMC
"""
import argparse
import sys
from datetime import datetime
from pathlib import Path

import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parent))

from src.dup_check.detector import find_duplicates
from src.dup_check.loader import load_me3l, load_oc
from src.dup_check.report import generate_report

_INPUTS  = Path(__file__).resolve().parent / "inputs"
_OUTPUTS = Path(__file__).resolve().parent / "outputs" / "material_dup_check"
_OPERATIONS = ["MLCC", "CCMC"]
_SOURCES: list[tuple[str, callable]] = [("OC", load_oc), ("ME3L", load_me3l)]


def _parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(
        description="Detectar materiales duplicados en múltiples contratos/OC activos."
    )
    p.add_argument(
        "--operation",
        choices=["MLCC", "CCMC", "ambos"],
        default="ambos",
        help="Operación a procesar (default: ambos)",
    )
    return p.parse_args()


def _active_count(df: pd.DataFrame) -> int:
    if df.empty:
        return 0
    if "Ind. Borrado" in df.columns:
        return int((df["Ind. Borrado"].fillna("").astype(str).str.strip() == "").sum())
    return len(df)


def _print_row(op: str, src: str, active: int, dups: pd.DataFrame) -> None:
    mats = int(dups["Material"].nunique()) if not dups.empty else 0
    docs = int(dups["Documento compras"].nunique()) if not dups.empty and "Documento compras" in dups.columns else 0
    flag = "  *** DUPLICADOS ***" if mats > 0 else ""
    print(f"  [{op}] {src:<6}  activos: {active:>8,}  |  "
          f"mats duplicados: {mats:>5,}  |  docs afectados: {docs:>5,}{flag}")


def main() -> None:
    args = _parse_args()
    operations = _OPERATIONS if args.operation == "ambos" else [args.operation.upper()]

    print()
    print("=" * 72)
    print("  DETECCIÓN DE MATERIALES DUPLICADOS EN CONTRATOS/OC ACTIVOS")
    print("=" * 72)

    source_counts: dict[tuple[str, str], int] = {}
    dup_dfs: dict[tuple[str, str], pd.DataFrame] = {}

    for op in operations:
        print(f"\n  [{op}]")
        for src_name, loader_fn in _SOURCES:
            print(f"    Cargando {src_name} ...", end=" ", flush=True)
            raw = loader_fn(_INPUTS, op)
            if raw.empty:
                print("sin datos")
                source_counts[(op, src_name)] = 0
                dup_dfs[(op, src_name)] = pd.DataFrame()
                continue

            print(f"{len(raw):,} registros leídos")
            active = _active_count(raw)
            source_counts[(op, src_name)] = active

            dups = find_duplicates(raw)
            dup_dfs[(op, src_name)] = dups
            _print_row(op, src_name, active, dups)

    print()
    print("=" * 72)
    total_mats = sum(
        int(df["Material"].nunique()) for df in dup_dfs.values() if not df.empty
    )
    print(f"  Total materiales duplicados encontrados: {total_mats:,}")
    print("=" * 72)

    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    out_path = _OUTPUTS / f"reporte_duplicados_materiales_{ts}.xlsx"

    print(f"\n  Generando reporte Excel ...")
    generate_report(source_counts, dup_dfs, out_path)

    print(f"  Guardado en:\n    {out_path}\n")


if __name__ == "__main__":
    main()
