"""Entry point: cross-reference materials from contracts/OC with the master and consumption data."""
import argparse
import sys
from datetime import datetime
from pathlib import Path

import pandas as pd

# Ensure project root is on sys.path when run directly
sys.path.insert(0, str(Path(__file__).resolve().parent))

from src.material_crossref.crossref import build_crossref, build_unique_materials, extract_source_materials
from src.material_crossref.loader import load_consumptions, load_material_master
from src.material_crossref.report import generate_report

_INPUTS_ROOT = Path(__file__).resolve().parent / "inputs"
_OUTPUTS_ROOT = Path(__file__).resolve().parent / "outputs"

_MASTER_FILE = _INPUTS_ROOT / "MLCC" / "MLCC - Análisis Completo.xlsx"
_CONSUMOS_FILE = _INPUTS_ROOT / "MLCC" / "MLCC - Consumos Históricos Marzo 2026.xlsx"

_ALL_DOMAINS = ["contratos", "ordenes_compra"]


def _parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Cruce de materiales (contratos/OC) con maestro MLCC y consumos históricos."
    )
    parser.add_argument(
        "--domain",
        choices=["contratos", "ordenes_compra", "ambos"],
        default="ambos",
        help="Dominio a procesar (default: ambos)",
    )
    return parser.parse_args()


def _print_summary(enriched_df: pd.DataFrame) -> None:
    total = len(enriched_df)
    in_master = enriched_df["in_master"].sum()
    has_consumo = enriched_df["has_consumption"].sum() if "has_consumption" in enriched_df.columns else 0
    is_service = enriched_df["is_service"].sum()

    print("\n" + "=" * 60)
    print("  CRUCE DE MATERIALES — RESUMEN")
    print("=" * 60)
    print(f"  Materiales únicos totales  : {total:>6,}")
    print(f"  Encontrados en maestro     : {int(in_master):>6,}  ({in_master/total*100:.1f}%)" if total else "  N/A")
    print(f"  No encontrados en maestro  : {total - int(in_master):>6,}")
    print(f"  Con consumo histórico      : {int(has_consumo):>6,}  ({has_consumo/total*100:.1f}%)" if total else "  N/A")
    print(f"  Materiales de servicio     : {int(is_service):>6,}  ({is_service/total*100:.1f}%)" if total else "  N/A")
    print()

    for domain in enriched_df["domain"].unique():
        sub = enriched_df[enriched_df["domain"] == domain]
        n = len(sub)
        nm = sub["in_master"].sum()
        print(f"  [{domain}]  {n:,} materiales  |  maestro: {nm/n*100:.1f}%")

    print("=" * 60)


def main() -> None:
    args = _parse_args()
    domains = _ALL_DOMAINS if args.domain == "ambos" else [args.domain]

    # --- Load reference tables ---
    print(f"Cargando maestro de materiales: {_MASTER_FILE.name}")
    master_df = load_material_master(_MASTER_FILE)
    print(f"  >> {len(master_df):,} materiales en maestro")

    print(f"Cargando consumos historicos: {_CONSUMOS_FILE.name}")
    consumos_df = load_consumptions(_CONSUMOS_FILE)
    print(f"  >> {len(consumos_df):,} materiales con datos de consumo")

    # --- Load source data per domain ---
    all_unique: list[pd.DataFrame] = []
    for domain in domains:
        print(f"\nCargando archivos de origen: {domain} ...")
        source_df = extract_source_materials(domain, _INPUTS_ROOT)
        print(f"  >> {len(source_df):,} registros cargados")
        unique_df = build_unique_materials(source_df)
        print(f"  >> {len(unique_df):,} materiales unicos")
        all_unique.append(unique_df)

    combined_unique = pd.concat(all_unique, ignore_index=True)

    # --- Cross-reference ---
    print("\nEjecutando cruce con maestro y consumos ...")
    enriched = build_crossref(combined_unique, master_df, consumos_df)

    # --- Generate report ---
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_path = _OUTPUTS_ROOT / "material_crossref" / f"reporte_materiales_{ts}.xlsx"
    print(f"Generando reporte: {output_path.name}")
    generate_report(enriched, output_path, domains)

    _print_summary(enriched)
    print(f"\nReporte guardado en:\n  {output_path}\n")


if __name__ == "__main__":
    main()
