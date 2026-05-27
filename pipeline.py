import argparse
import logging
from datetime import datetime

from src.config.domains import DOMAIN_CONFIGS, get_domain_config
from src.config.settings import Settings
from src.ingestion.assembler import DomainAssembler
from src.lineage.records import FieldLineageRecord, StageManifest
from src.material_crossref.enrichment import enrich_material_rows, load_reference_tables
from src.output.control_point import export_control_point
from src.output.stats_report import generate_stats_report
from src.profiling.profiler import DataProfiler
from src.rules.engine import RuleEngine
from src.utils.logger import setup_logger


def run(
    operation: str,
    domain: str = "contratos",
    export_control_points: bool = True,
    generate_report: bool = True,
) -> None:
    settings = Settings(domain=domain)
    settings.ensure_dirs()
    domain_cfg = get_domain_config(settings.domain)

    log = setup_logger(settings.project_root / "logs", f"{domain_cfg.output_suffix}_{operation}")
    log.info("=" * 60)
    log.info("Pipeline iniciado  |  dominio: %s  |  operacion: %s", domain_cfg.cli_name, operation)
    log.info("=" * 60)

    # --- 1. Ingestion + schema normalisation ---
    log.info("[1/7] Ingesta y normalizacion de esquema...")
    assembler = DomainAssembler(settings)
    master, lineage, manifests, stats = assembler.build(operation)

    log.info("  Archivos cargados (%d):", len(stats.files_loaded))
    for fname, n in stats.raw_rows_per_file.items():
        log.info("    %-40s  %6d filas brutas", fname, n)
    log.info("  Total filas brutas:                    %6d", stats.total_raw)
    log.info("  Filas vacias descartadas:              %6d", stats.empty_rows_dropped)
    log.info("  Filas recuperadas por forward-fill:    %6d", stats.ffill_rows_recovered)
    log.info("  ---")
    log.info("  Total filas limpias (C1):              %6d", stats.total_clean)
    log.info("  Documentos unicos:                     %6d", stats.unique_documents)
    log.info("  Pares (documento, posicion) unicos:    %6d", stats.unique_doc_pos_pairs)
    if stats.duplicate_doc_pos > 0:
        log.warning("  DUPLICADOS (doc, posicion):            %6d  <-- revisar", stats.duplicate_doc_pos)
    else:
        log.info("  Duplicados (doc, posicion):            %6d  [OK]", stats.duplicate_doc_pos)
    log.info("  Columnas: %d total  |  mapped: %d  unmapped: %d  injected: %d",
             len(master.columns), len(lineage.mapped()), len(lineage.unmapped()), len(lineage.injected()))

    # --- 2. Material reference enrichment ---
    log.info("[2/7] Enriqueciendo materiales con maestro MLCC y consumos...")
    material_started = datetime.now()
    rows_before_material = len(master)
    cols_before_material = len(master.columns)
    cols_before_set = set(master.columns)
    material_master, consumptions = load_reference_tables(settings.inputs_dir)
    master = enrich_material_rows(master, material_master, consumptions)
    material_cols = [c for c in master.columns if c not in cols_before_set]
    for col in material_cols:
        series = master[col]
        total = max(len(master), 1)
        lineage.add(FieldLineageRecord(
            source_file="(material_crossref)",
            raw_name="",
            canonical_name=col,
            mapping_status="INJECTED",
            null_count=int(series.isna().sum()),
            null_pct=round(float(series.isna().sum()) / total * 100, 2),
            unique_count=int(series.nunique(dropna=True)),
            sample_values=[str(v) for v in series.dropna().unique()[:5].tolist()],
        ))
    manifests.append(StageManifest(
        stage_id="MATERIAL_XREF",
        label="Cruce maestro materiales y consumos",
        started_at=material_started,
        completed_at=datetime.now(),
        rows_in=rows_before_material,
        rows_out=len(master),
        columns_in=cols_before_material,
        columns_out=len(master.columns),
        notes=(
            f"Materiales maestro: {len(material_master):,}. "
            f"Materiales con consumos: {len(consumptions):,}. "
            f"Columnas inyectadas: {', '.join(material_cols)}."
        ),
    ))
    in_master = int(master.get("material_in_master", []).sum()) if "material_in_master" in master.columns else 0
    no_movement = int(master.get("material_no_movement_24m", []).sum()) if "material_no_movement_24m" in master.columns else 0
    log.info("  Cruce materiales: en maestro=%d  sin movimiento 24m=%d", in_master, no_movement)

    # --- 2. Profiling ---
    log.info("[3/7] Profiling...")
    profiler = DataProfiler()
    profiling, prof_manifest = profiler.profile(master, operation)
    manifests.append(prof_manifest)

    critical_low = [(p.canonical_name, p.completeness_pct)
                    for p in profiling.column_profiles if p.is_critical and p.completeness_pct < 95]
    if critical_low:
        for col, pct in critical_low:
            log.warning("  Columna critica con baja completitud: %s  %.1f%%", col, pct)
    else:
        log.info("  Todas las columnas criticas >= 95%% completitud  [OK]")

    # --- 3. Export C1 ---
    log.info("[4/7] Exportando C1...")
    if export_control_points:
        cp1 = export_control_point(
            cp_id="C1",
            description=f"Post-ingesta: {domain_cfg.display_name.lower()} normalizados y perfilados.",
            operation=operation,
            df=master,
            lineage=lineage,
            profiling=profiling,
            manifests=list(manifests),
            issues=[],
            output_dir=settings.domain_control_points_dir,
            include_analysis=True,
            analysis_subject=domain_cfg.display_name,
        )
        log.info("  C1 exportado: %s  (%d filas)", cp1.name, len(master))
    else:
        log.info("  C1 omitido por --skip-control-points  (%d filas)", len(master))

    # --- 4. G1 exclusions ---
    log.info("[5/7] Aplicando reglas G1 (exclusiones de alcance PDF)...")
    engine = RuleEngine(settings)
    df_g1, g1_issues, g1_manifest = engine.apply_group("G1_EXCLUSIONS", master, operation)
    manifests.append(g1_manifest)

    g1_excluded = int((df_g1["exclusion_reason"] != "").sum())
    for issue in g1_issues:
        log.info("  Regla %s (%s): %d filas excluidas", issue.code, issue.message, issue.row_count)
    log.info("  Excluidos post-G1: %d  |  Continuan: %d", g1_excluded, len(df_g1) - g1_excluded)

    # --- 5. G2 rescue ---
    log.info("[6/7] Aplicando reglas G2 (rescates configurados)...")
    df_g2, g2_issues, g2_manifest = engine.apply_group("G2_RESCUE", df_g1, operation)
    manifests.append(g2_manifest)

    rescued_count = int((df_g2["rescue_reason"] != "").sum())
    for issue in g2_issues:
        log.info("  Regla %s (%s): %d filas rescatadas", issue.code, issue.message, issue.row_count)
    if rescued_count > 0:
        log.info("  Rescatados por saldo pendiente positivo en vencidos: %d", rescued_count)
    else:
        log.info("  Sin rescates por saldo pendiente positivo en vencidos  [OK]")

    # --- 6. G3 marking (universe post-G2) ---
    log.info("[7/7] Aplicando reglas G3 (identificacion y marcado PDF)...")
    df_g3, g3_issues, g3_manifest = engine.apply_group("G3_MARKING", df_g2, operation)
    manifests.append(g3_manifest)

    for mark_col in sorted(c for c in df_g3.columns if c.startswith("mark_")):
        mark_count = int(df_g3[mark_col].fillna("").astype(str).str.strip().ne("").sum())
        log.info("  %s: %d", mark_col, mark_count)

    # --- Final split ---
    df_no_migra = df_g3[df_g3["exclusion_reason"] != ""].copy()
    df_migra = df_g3[df_g3["exclusion_reason"] == ""].copy()

    all_issues = g1_issues + g2_issues + g3_issues

    if export_control_points:
        cp2_nm = export_control_point(
            cp_id="C2_NO_MIGRA",
            description=f"Post-G1/G2: {domain_cfg.display_name.lower()} excluidos por reglas activas y no rescatados.",
            operation=operation,
            df=df_no_migra,
            lineage=lineage,
            profiling=profiling,
            manifests=list(manifests),
            issues=g1_issues + g2_issues,
            output_dir=settings.domain_control_points_dir,
            include_analysis=True,
            analysis_subject=domain_cfg.display_name,
        )
        cp3 = export_control_point(
            cp_id="C3",
            description=f"Post-G1/G2/G3: {domain_cfg.display_name.lower()} que continúan, con marcas PDF aplicadas.",
            operation=operation,
            df=df_migra,
            lineage=lineage,
            profiling=profiling,
            manifests=list(manifests),
            issues=all_issues,
            output_dir=settings.domain_control_points_dir,
            include_analysis=True,
            analysis_subject=domain_cfg.display_name,
        )
        cp2_name = cp2_nm.name
        cp3_name = cp3.name
    else:
        cp2_name = "(omitido por --skip-control-points)"
        cp3_name = "(omitido por --skip-control-points)"
        log.info("  Control points finales omitidos por --skip-control-points")

    # --- Summary ---
    reconciliation_ok = len(master) == len(df_no_migra) + len(df_migra)

    log.info("=" * 60)
    log.info("RESUMEN  |  operacion: %s", operation)
    log.info("=" * 60)
    log.info("  C1  (normalizados):       %6d", len(master))
    log.info("  C2_NO_MIGRA (excluidos):  %6d  ->  %s", len(df_no_migra), cp2_name)
    log.info("    (rescatados por G2:     %6d)", rescued_count)
    log.info("  C3  (migran):             %6d  ->  %s", len(df_migra), cp3_name)
    if reconciliation_ok:
        log.info("  Reconciliacion C1 == C2_NO_MIGRA + C3:  OK")
    else:
        log.error("  Reconciliacion FALLIDA: C1=%d != C3=%d + C2_NO_MIGRA=%d",
                  len(master), len(df_migra), len(df_no_migra))
    log.info("=" * 60)

    if not reconciliation_ok:
        raise RuntimeError(
            f"Reconciliation failed: "
            f"C1={len(master)} != C3={len(df_migra)} + C2_NO_MIGRA={len(df_no_migra)}"
        )

    # --- Reporte estadístico independiente (C1 — universo pre-reglas) ---
    is_purchase_order_domain = domain_cfg.package_name == "ordenes_compra"
    if generate_report:
        rpt = generate_stats_report(
            master, operation, settings.domain_outputs_dir,
            dataset_label=f"Universo Completo — C1 (pre-reglas, {len(master):,} filas)",
            entity_label=domain_cfg.display_name,
            file_suffix=domain_cfg.report_suffix,
            include_delivery_date_section=is_purchase_order_domain,
            include_purchase_amount_section=not is_purchase_order_domain,
            include_doc_class_section=is_purchase_order_domain,
            include_framework_section=is_purchase_order_domain,
        )
        log.info("  Reporte estadistico:      %s", rpt.name)
    else:
        log.info("  Reporte estadistico omitido por --skip-stats-report")


def _parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Lundin domain-aware cleansing pipeline",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "Ejemplos:\n"
            "  python pipeline.py\n"
            "  python pipeline.py --domain contratos --operation MLCC\n"
            "  python pipeline.py --domain ordenes-compra --operation MLCC"
        ),
    )
    domain_choices = sorted(config.cli_name for config in DOMAIN_CONFIGS.values())
    parser.add_argument(
        "--domain", "-d",
        default="contratos",
        choices=domain_choices,
        help="Dominio a procesar (default: contratos)",
    )
    parser.add_argument(
        "--operation", "-o",
        default="MLCC",
        choices=["MLCC"],
        help="Operacion a procesar (default: MLCC)",
    )
    parser.add_argument(
        "--skip-control-points",
        action="store_true",
        help="Ejecuta reglas y reconciliacion sin escribir workbooks de control point.",
    )
    parser.add_argument(
        "--skip-stats-report",
        action="store_true",
        help="Omite el reporte estadistico independiente.",
    )
    return parser.parse_args()


if __name__ == "__main__":
    args = _parse_args()
    run(
        args.operation,
        args.domain,
        export_control_points=not args.skip_control_points,
        generate_report=not args.skip_stats_report,
    )
