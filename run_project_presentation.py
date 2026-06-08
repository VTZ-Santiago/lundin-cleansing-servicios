from __future__ import annotations

import argparse
from datetime import date
from pathlib import Path

from src.presentation import generate_project_presentation


def parse_args() -> argparse.Namespace:
    project_root = Path(__file__).resolve().parent
    default_output = project_root / "outputs" / "entregables" / f"presentacion_servicios_{date.today():%Y%m%d}_v1.pptx"
    parser = argparse.ArgumentParser(description="Genera una presentación ejecutiva basada en la plantilla de referencia y los reportes actuales.")
    parser.add_argument(
        "--template",
        type=Path,
        default=project_root / "resources" / "Presentacion_Lundin_300426_v1.pptx",
        help="Ruta a la presentación de referencia usada como base visual.",
    )
    parser.add_argument(
        "--mlcc-summary",
        type=Path,
        default=project_root / "outputs" / "control_points" / "resumen_segmentacion_MLCC.md",
        help="Ruta al resumen markdown de segmentación MLCC.",
    )
    parser.add_argument(
        "--ccmc-summary",
        type=Path,
        default=project_root / "outputs" / "control_points" / "resumen_segmentacion_CCMC.md",
        help="Ruta al resumen markdown de segmentación CCMC.",
    )
    parser.add_argument(
        "--output-root",
        type=Path,
        default=project_root / "outputs",
        help="Carpeta raíz de salidas del proyecto para detectar artefactos vigentes.",
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=default_output,
        help="Ruta del archivo .pptx a generar.",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    output_path = generate_project_presentation(
        template_path=args.template,
        output_path=args.output,
        mlcc_summary_path=args.mlcc_summary,
        ccmc_summary_path=args.ccmc_summary,
        output_root=args.output_root,
    )
    print(output_path)


if __name__ == "__main__":
    main()