"""Cross-reference materials from contracts/OC with their operation's master and consumption data.

Supports MLCC and CCMC operations. Each operation has its own materials master and
consumption file. Results are combined into a single Excel report.

Usage:
    python run_material_crossref.py                         # all operations, all domains
    python run_material_crossref.py --operation MLCC        # only MLCC
    python run_material_crossref.py --operation CCMC        # only CCMC
    python run_material_crossref.py --domain ordenes_compra # only OC for all operations
"""
import argparse
import glob
import sys
from datetime import datetime
from pathlib import Path

import pandas as pd

sys.path.insert(0, str(Path(__file__).resolve().parent))

from src.material_crossref.crossref import build_crossref, build_unique_materials, extract_source_materials
from src.material_crossref.loader import load_consumptions, load_material_master
from src.material_crossref.report import generate_report

_INPUTS_ROOT  = Path(__file__).resolve().parent / "inputs"
_OUTPUTS_ROOT = Path(__file__).resolve().parent / "outputs"

# Operations and the domains they support
_OPERATION_DOMAINS: dict[str, list[str]] = {
    "MLCC": ["contratos", "ordenes_compra"],
    "CCMC": ["ordenes_compra"],
}


def _find_file(directory: Path, pattern: str) -> Path:
    """Find a single file matching glob pattern inside directory."""
    matches = list(directory.glob(pattern))
    if not matches:
        raise FileNotFoundError(f"No file matching '{pattern}' in {directory}")
    return sorted(matches)[0]


def _parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Cruce de materiales con maestro y consumos historicos (MLCC y CCMC)."
    )
    parser.add_argument(
        "--operation",
        choices=["MLCC", "CCMC", "ambos"],
        default="ambos",
        help="Operacion a procesar (default: ambos)",
    )
    parser.add_argument(
        "--domain",
        choices=["contratos", "ordenes_compra", "ambos"],
        default="ambos",
        help="Dominio a procesar (default: ambos)",
    )
    return parser.parse_args()


def _print_summary(enriched_df: pd.DataFrame) -> None:
    total      = len(enriched_df)
    in_master  = enriched_df["in_master"].sum()
    has_cons   = enriched_df["has_consumption"].sum() if "has_consumption" in enriched_df.columns else 0
    is_service = enriched_df["is_service"].sum()

    print()
    print("=" * 65)
    print("  CRUCE DE MATERIALES -- RESUMEN")
    print("=" * 65)
    print(f"  Materiales unicos totales  : {total:>7,}")
    if total:
        print(f"  En maestro                 : {int(in_master):>7,}  ({in_master/total*100:.1f}%)")
        print(f"  No encontrados en maestro  : {total - int(in_master):>7,}")
        print(f"  Con consumo historico      : {int(has_cons):>7,}  ({has_cons/total*100:.1f}%)")
        print(f"  Materiales de servicio     : {int(is_service):>7,}  ({is_service/total*100:.1f}%)")

    print()
    for op in enriched_df.get("operation", pd.Series(dtype=str)).unique():
        sub_op = enriched_df[enriched_df["operation"] == op] if "operation" in enriched_df.columns else enriched_df
        print(f"  [{op}]")
        for dom in sub_op["domain"].unique():
            sub = sub_op[sub_op["domain"] == dom]
            n = len(sub)
            nm = sub["in_master"].sum()
            print(f"    {dom:<20} {n:>6,} materiales  |  maestro: {nm/n*100:.1f}%")
    print("=" * 65)


def main() -> None:
    args = _parse_args()

    operations = list(_OPERATION_DOMAINS.keys()) if args.operation == "ambos" else [args.operation.upper()]
    domain_filter = None if args.domain == "ambos" else args.domain

    all_enriched: list[pd.DataFrame] = []
    all_domains_used: list[str] = []

    for operation in operations:
        op_dir = _INPUTS_ROOT / operation
        if not op_dir.exists():
            print(f"AVISO: directorio {op_dir} no encontrado, saltando {operation}.")
            continue

        # Load master
        master_path = _find_file(op_dir, f"{operation}*Completo*.xlsx")
        print(f"\n[{operation}] Cargando maestro: {master_path.name}")
        master_df = load_material_master(master_path, operation)
        print(f"  >> {len(master_df):,} materiales en maestro")

        # Load consumos
        consumos_path = _find_file(op_dir, f"{operation}*Consumos*.xlsx")
        print(f"[{operation}] Cargando consumos: {consumos_path.name}")
        consumos_df = load_consumptions(consumos_path, operation)
        print(f"  >> {len(consumos_df):,} materiales con consumo")

        # Load source files per domain
        domains = _OPERATION_DOMAINS[operation]
        if domain_filter:
            domains = [d for d in domains if d == domain_filter]

        op_unique_frames: list[pd.DataFrame] = []
        for domain in domains:
            source_dir = op_dir / ("contratos" if domain == "contratos" else "ordenes-compra")
            if not source_dir.exists():
                print(f"  AVISO: {source_dir} no existe, saltando.")
                continue
            print(f"[{operation}] Cargando {domain} ...")
            source_df = extract_source_materials(domain, _INPUTS_ROOT, operation)
            print(f"  >> {len(source_df):,} registros | {source_df['material_key'].nunique():,} materiales unicos")
            op_unique_frames.append(build_unique_materials(source_df))
            if domain not in all_domains_used:
                all_domains_used.append(domain)

        if not op_unique_frames:
            print(f"  AVISO: sin datos para {operation}, saltando cruce.")
            continue

        op_unique = pd.concat(op_unique_frames, ignore_index=True)
        print(f"[{operation}] Ejecutando cruce ...")
        enriched = build_crossref(op_unique, master_df, consumos_df)
        all_enriched.append(enriched)

    if not all_enriched:
        print("Sin datos para generar reporte.")
        return

    combined = pd.concat(all_enriched, ignore_index=True)

    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    output_path = _OUTPUTS_ROOT / "material_crossref" / f"reporte_materiales_{ts}.xlsx"
    print(f"\nGenerando reporte: {output_path.name}")
    generate_report(combined, output_path, all_domains_used)

    _print_summary(combined)
    print(f"\nReporte guardado en:\n  {output_path}\n")


if __name__ == "__main__":
    main()
