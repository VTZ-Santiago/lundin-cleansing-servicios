import os
import sys
from pathlib import Path

import yaml

_CP_ORANGE = "#F97316"
_EXCL_RED = "#DC2626"
_MAIN_DARK = "#222222"
_INGESTION_BLUE = "#DBEAFE"
_GROUP_YELLOW = "#FEF3C7"
_START_GRAY = "#E5E7EB"
_CP_LABEL_BG = "#FFF7ED"


def _add_graphviz_to_path() -> None:
    if sys.platform != "win32":
        return

    current = os.environ.get("PATH", "")
    for graphviz_dir in (
        r"C:\Program Files\Graphviz\bin",
        r"C:\Program Files (x86)\Graphviz\bin",
        r"C:\Graphviz\bin",
        r"C:\ProgramData\chocolatey\bin",
    ):
        dot_exe = os.path.join(graphviz_dir, "dot.exe")
        if os.path.isfile(dot_exe) and graphviz_dir not in current:
            os.environ["PATH"] = graphviz_dir + os.pathsep + current
            return


def _read_cp_rows(control_points_dir: Path | None, cp_id: str, operation: str) -> str:
    """Read 'Filas en master' from the Info sheet of an existing control point file."""
    if not control_points_dir or not cp_id:
        return ""
    path = control_points_dir / f"{cp_id}_{operation}.xlsx"
    if not path.exists():
        return ""
    try:
        from openpyxl import load_workbook
        wb = load_workbook(str(path), read_only=True, data_only=True)
        if "Info" not in wb.sheetnames:
            wb.close()
            return ""
        ws = wb["Info"]
        for row in ws.iter_rows(values_only=True):
            if row and row[0] == "Filas en master":
                val = row[1]
                wb.close()
                return f"n={int(val):,}" if val is not None else ""
        wb.close()
    except Exception:
        pass
    return ""


def generate_flowchart(
    yaml_path: Path,
    output_path: Path,
    fmt: str = "svg",
    operation: str = "MLCC",
    control_points_dir: Path | None = None,
) -> Path:
    """Generate a pipeline flowchart SVG/PNG from rules.yaml.

    Falls back to saving a .dot source file if the Graphviz binary is not installed.
    """
    try:
        import graphviz
    except ImportError:
        raise ImportError("Install graphviz: pip install graphviz")

    _add_graphviz_to_path()

    with open(yaml_path, encoding="utf-8") as f:
        config = yaml.safe_load(f)

    op = operation.upper()
    diagram_cfg = config.get("diagram", {})
    title = diagram_cfg.get("title", "Pipeline")
    group_output_labels = diagram_cfg.get("group_output_labels", {})

    g = graphviz.Digraph(
        name="pipeline",
        format=fmt,
        graph_attr={
            "rankdir": "TB",
            "bgcolor": "white",
            "fontname": "Helvetica",
            "label": title,
            "labelloc": "t",
            "fontsize": "16",
            "pad": "0.5",
            "splines": "ortho",
        },
        node_attr={"fontname": "Helvetica", "fontsize": "11"},
        edge_attr={"fontname": "Helvetica", "fontsize": "10"},
    )

    # --- Preamble nodes ---
    preamble_ids: list[str] = []
    for node in diagram_cfg.get("preamble", []):
        nid = node["id"]
        label = node["label"]
        shape = node.get("shape", "box")
        preamble_ids.append(nid)

        if shape == "oval":
            g.node(nid, label, shape="ellipse", style="filled",
                   fillcolor=_START_GRAY, color=_MAIN_DARK)
        elif shape == "box":
            g.node(nid, label, shape="box", style="rounded,filled",
                   fillcolor=_INGESTION_BLUE, color=_MAIN_DARK)
        else:
            # Control-point label node
            cp_id = node.get("control_point", "")
            rows_str = _read_cp_rows(control_points_dir, cp_id, op)
            full_label = label
            if rows_str:
                full_label = f"{label}\n{rows_str}"
            g.node(nid, full_label, shape="box", style="filled",
                   fillcolor=_CP_LABEL_BG, color=_CP_ORANGE,
                   penwidth="2")

    # Connect preamble nodes in sequence
    for a, b in zip(preamble_ids, preamble_ids[1:]):
        g.edge(a, b, color=_MAIN_DARK)

    last_preamble = preamble_ids[-1] if preamble_ids else "start"
    final_outputs = diagram_cfg.get("final_outputs", [])

    # --- Rule groups ---
    for group_id, group_cfg in config.get("groups", {}).items():
        rules = group_cfg.get("rules", [])
        enabled_rules = [r for r in rules if r.get("enabled", {}).get(op, False)]

        # Build group diamond label
        rule_lines = "\n".join(f"• {r['id']}: {r['name']}" for r in enabled_rules)
        group_label = group_cfg.get("name", group_id)
        if rule_lines:
            group_label = f"{group_label}\n{rule_lines}"

        g.node(group_id, group_label, shape="diamond", style="filled",
               fillcolor=_GROUP_YELLOW, color=_MAIN_DARK)
        g.edge(last_preamble, group_id, color=_MAIN_DARK)

        if final_outputs:
            last_preamble = group_id
            continue

        # Survivors output (C2, etc.)
        cp_out = group_cfg.get("control_point_out", "C2")
        survivors_label = group_output_labels.get(group_id, cp_out)
        rows_str = _read_cp_rows(control_points_dir, cp_out, op)
        if rows_str:
            survivors_label = f"{survivors_label}\n{rows_str}"

        g.node(cp_out, survivors_label, shape="box", style="filled",
               fillcolor=_CP_LABEL_BG, color=_CP_ORANGE, penwidth="2")
        g.edge(group_id, cp_out, label="Migra", color=_MAIN_DARK)

        # Excluded output (C2_NO_MIGRA, etc.)
        no_migra_id = f"{cp_out}_NO_MIGRA"
        no_migra_rows = _read_cp_rows(control_points_dir, no_migra_id, op)
        no_migra_label = "NO MIGRA"
        if no_migra_rows:
            no_migra_label = f"NO MIGRA\n{no_migra_rows}"

        g.node(no_migra_id, no_migra_label, shape="box", style="filled",
               fillcolor="#FEE2E2", color=_EXCL_RED, penwidth="2",
               fontcolor=_EXCL_RED)
        g.edge(group_id, no_migra_id, label="No migra",
               color=_EXCL_RED, fontcolor=_EXCL_RED)

        last_preamble = cp_out

    if final_outputs:
        for output in final_outputs:
            output_id = output["id"]
            cp_id = output.get("control_point", output_id)
            output_label = output.get("label", output_id)
            rows_str = _read_cp_rows(control_points_dir, cp_id, op)
            if rows_str:
                output_label = f"{output_label}\n{rows_str}"

            if output.get("kind") == "exclude":
                g.node(output_id, output_label, shape="box", style="filled",
                       fillcolor="#FEE2E2", color=_EXCL_RED, penwidth="2",
                       fontcolor=_EXCL_RED)
                g.edge(last_preamble, output_id,
                       label=output.get("edge_label", "No migra"),
                       color=_EXCL_RED, fontcolor=_EXCL_RED)
            else:
                g.node(output_id, output_label, shape="box", style="filled",
                       fillcolor=_CP_LABEL_BG, color=_CP_ORANGE, penwidth="2")
                g.edge(last_preamble, output_id,
                       label=output.get("edge_label", "Migra"),
                       color=_MAIN_DARK)

    # Render
    out_stem = str(output_path.with_suffix(""))
    dot_path = output_path.with_suffix(".dot")
    dot_path.write_text(g.source, encoding="utf-8")

    try:
        rendered = g.render(filename=out_stem, cleanup=True)
        return Path(rendered)
    except Exception as exc:
        print(f"\n[WARN] Graphviz binary not found ({exc}).")
        print("       Intenté resolver rutas estándar de Graphviz en Windows.")
        print(f"       Fuente DOT guardada en: {dot_path}")
        print(f"       Para renderizar: dot -T{fmt} \"{dot_path}\" -o \"{output_path}\"")
        return dot_path
