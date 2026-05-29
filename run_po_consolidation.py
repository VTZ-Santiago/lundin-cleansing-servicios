"""Build compact purchase-order consolidation workbooks by operation.

The output keeps only the headers requested for PO review and uses PO files as
the source of truth. Missing source fields remain blank; contract master files
are not joined into this extract.
"""
import argparse
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from src.ordenes_compra.po_consolidator import build_po_consolidation


ROOT = Path(__file__).resolve().parent
INPUTS_ROOT = ROOT / "inputs"
DEFAULT_OUTPUT_DIR = ROOT / "outputs" / "ordenes_compra" / "consolidado"
LEGACY_OUTPUT = DEFAULT_OUTPUT_DIR / "po_consolidado_2026plus.xlsx"
OPERATIONS = ["MLCC", "CCMC"]


def _parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Consolida ordenes de compra 2026+ en una tabla Excel compacta."
    )
    parser.add_argument(
        "--operation",
        choices=["MLCC", "CCMC", "ambos"],
        default="ambos",
        help="Operacion a procesar (default: ambos).",
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=None,
        help="Ruta de salida para una sola operacion o directorio base cuando se procesan ambas.",
    )
    return parser.parse_args()


def _selected_operations(value: str) -> list[str]:
    if value == "ambos":
        return OPERATIONS
    return [value.upper()]


def _default_output_path(operation: str) -> Path:
    return DEFAULT_OUTPUT_DIR / f"po_consolidado_{operation}_2025plus.xlsx"


def _resolve_output_path(output: Path | None, operation: str, multiple: bool) -> Path:
    if output is None:
        return _default_output_path(operation)
    if not multiple:
        return output
    if output.suffix.lower() == ".xlsx":
        return output.with_name(f"{output.stem}_{operation}{output.suffix}")
    return output / f"po_consolidado_{operation}_2025plus.xlsx"


def _print_summary(result, operation: str) -> None:
    print()
    print("=" * 72)
    print(f"  CONSOLIDADO PO {operation} 2025+ -- RESUMEN")
    print("=" * 72)
    print(f"  Filas leidas      : {result.rows_read:>10,}")
    print(f"  Filas retenidas   : {result.rows_kept:>10,}")
    print(f"  Filas descartadas : {result.rows_discarded:>10,}")
    print()

    for stat in result.file_stats:
        status = "SALTADO" if stat.skipped else "OK"
        print(f"  [{stat.operation}] {stat.source_file} -- {status}")
        print(f"    leidas: {stat.rows_read:,} | retenidas: {stat.rows_kept:,}")
        if stat.skip_reason:
            print(f"    aviso: {stat.skip_reason}")
        if stat.missing_headers:
            print(f"    columnas sin fuente: {', '.join(stat.missing_headers)}")

    print()
    print(f"  Archivo generado: {result.output_path}")
    print("=" * 72)


def main() -> None:
    args = _parse_args()
    operations = _selected_operations(args.operation)
    output_paths = [
        _resolve_output_path(args.output, operation, len(operations) > 1)
        for operation in operations
    ]

    if LEGACY_OUTPUT.exists() and LEGACY_OUTPUT not in output_paths:
        LEGACY_OUTPUT.unlink()

    for operation, output_path in zip(operations, output_paths):
        result = build_po_consolidation(INPUTS_ROOT, output_path, [operation])
        _print_summary(result, operation)


if __name__ == "__main__":
    main()