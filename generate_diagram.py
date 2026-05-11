import argparse
from pathlib import Path

from src.config.settings import Settings
from src.diagram.flowchart import generate_flowchart


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Generate Lundin contracts pipeline flowchart",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "Ejemplos:\n"
            "  python generate_diagram.py\n"
            "  python generate_diagram.py --format png\n"
            "  python generate_diagram.py --output outputs/mi_diagrama"
        ),
    )
    parser.add_argument("--format", "-f", choices=["svg", "png"], default="svg",
                        help="Formato de salida (default: svg)")
    parser.add_argument("--output", "-o", default=None,
                        help="Ruta de salida sin extensión (default: outputs/diagrama_pipeline)")
    parser.add_argument("--operation", default="MLCC", choices=["MLCC"],
                        help="Operación para mostrar reglas activas (default: MLCC)")
    args = parser.parse_args()

    settings = Settings()
    settings.ensure_dirs()

    output_path = (
        Path(args.output)
        if args.output
        else settings.outputs_dir / "diagrama_pipeline"
    )
    # Add extension for clarity (flowchart.py will handle actual suffix)
    output_path = output_path.with_suffix(f".{args.format}")

    result = generate_flowchart(
        yaml_path=settings.rules_yaml_path,
        output_path=output_path,
        fmt=args.format,
        operation=args.operation,
        control_points_dir=settings.control_points_dir,
    )
    print(f"Diagrama generado: {result}")


if __name__ == "__main__":
    main()
