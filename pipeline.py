import argparse
import logging

from src.config.domains import DOMAIN_CONFIGS, get_domain_config
from src.config.settings import Settings
from src.ingestion.assembler import DomainAssembler
from src.output.control_point import export_control_point
from src.output.stats_report import generate_stats_report
from src.profiling.profiler import DataProfiler
from src.rules.engine import RuleEngine
from src.utils.logger import setup_logger


def run(operation: str, domain: str = "contratos") -> None:
    settings = Settings(domain=domain)
    settings.ensure_dirs()
    domain_cfg = get_domain_config(settings.domain)

    log = setup_logger(settings.project_root / "logs", f"{domain_cfg.output_suffix}_{operation}")
    log.info("=" * 60)
    log.info("Pipeline iniciado  |  dominio: %s  |  operacion: %s", domain_cfg.cli_name, operation)
    log.info("=" * 60)

    # --- 1. Ingestion + schema normalisation ---
    log.info("[1/4] Ingesta y normalizacion de esquema...")
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

    # --- 2. Profiling ---
    log.info("[2/4] Profiling...")
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
    log.info("[3/4] Exportando C1...")
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

    # --- 4. G1 exclusions ---
    log.info("[4/6] Aplicando reglas G1 (primer filtro por vencimiento y tipo)...")
    engine = RuleEngine(settings)
    df_g1, g1_issues, g1_manifest = engine.apply_group("G1_EXCLUSIONS", master, operation)
    manifests.append(g1_manifest)

    g1_excluded = int((df_g1["exclusion_reason"] != "").sum())
    for issue in g1_issues:
        log.info("  Regla %s (%s): %d filas excluidas", issue.code, issue.message, issue.row_count)
    log.info("  Excluidos post-G1: %d  |  Candidatos vigentes/tipo D: %d", g1_excluded, len(df_g1) - g1_excluded)

    # --- 5. G2 rescue ---
    log.info("[5/6] Aplicando reglas G2 (migracion segura por saldo pendiente)...")
    df_g2, g2_issues, g2_manifest = engine.apply_group("G2_RESCUE", df_g1, operation)
    manifests.append(g2_manifest)

    rescued_count = int((df_g2["rescue_reason"] != "").sum())
    for issue in g2_issues:
        log.info("  Regla %s (%s): %d filas rescatadas", issue.code, issue.message, issue.row_count)
    if rescued_count > 0:
        log.info("  Rescatados por saldo pendiente positivo en vencidos: %d", rescued_count)
    else:
        log.info("  Sin rescates por saldo pendiente positivo en vencidos  [OK]")

    # --- 6. G3 marking (solo sobre migrantes) ---
    log.info("[6/6] Aplicando reglas G3 (clasificacion y marcado)...")
    df_migra_pre_g3 = df_g2[df_g2["exclusion_reason"] == ""].copy()
    df_g3, g3_issues, g3_manifest = engine.apply_group("G3_MARKING", df_migra_pre_g3, operation)
    manifests.append(g3_manifest)

    if "mark_position_type" in df_g3.columns:
        pt_dist = df_g3["mark_position_type"].value_counts().to_dict()
        log.info("  mark_position_type: %s", "  ".join(f"{k}={v}" for k, v in sorted(pt_dist.items())))
    if "mark_deletion_flag" in df_g3.columns:
        df_dist = df_g3["mark_deletion_flag"].value_counts().to_dict()
        log.info("  mark_deletion_flag: %s", "  ".join(f"{k}={v}" for k, v in sorted(df_dist.items())))
    if "mark_pending_negative_active" in df_g3.columns:
        neg_count = int((df_g3["mark_pending_negative_active"] != "").sum())
        log.info("  mark_pending_negative_active: %d", neg_count)
    if "mark_validity_2026" in df_g3.columns:
        validity_2026_count = int((df_g3["mark_validity_2026"] != "").sum())
        log.info("  mark_validity_2026: %d", validity_2026_count)
    if "mark_direct_cost_center_service" in df_g3.columns:
        direct_cost_center_count = int((df_g3["mark_direct_cost_center_service"] != "").sum())
        log.info("  mark_direct_cost_center_service: %d", direct_cost_center_count)

    # --- Final split ---
    df_no_migra = df_g2[df_g2["exclusion_reason"] != ""].copy()
    df_migra = df_g3  # migrants: G1 survivors + G2 rescued, marked by G3

    all_issues = g1_issues + g2_issues + g3_issues

    cp2_nm = export_control_point(
        cp_id="C2_NO_MIGRA",
        description=f"Post-G1/G2: {domain_cfg.display_name.lower()} excluidos por primer filtro y no rescatados.",
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
        description=f"Post-G1/G2/G3: {domain_cfg.display_name.lower()} que migran, con rescates y marcas aplicadas.",
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

    # --- Summary ---
    reconciliation_ok = len(master) == len(df_no_migra) + len(df_migra)

    log.info("=" * 60)
    log.info("RESUMEN  |  operacion: %s", operation)
    log.info("=" * 60)
    log.info("  C1  (normalizados):       %6d", len(master))
    log.info("  C2_NO_MIGRA (excluidos):  %6d  ->  %s", len(df_no_migra), cp2_nm.name)
    log.info("    (rescatados por G2:     %6d)", rescued_count)
    log.info("  C3  (migran):             %6d  ->  %s", len(df_migra), cp3.name)
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
    rpt = generate_stats_report(
        master, operation, settings.domain_outputs_dir,
        dataset_label=f"Universo Completo — C1 (pre-reglas, {len(master):,} filas)",
        entity_label=domain_cfg.display_name,
    )
    log.info("  Reporte estadistico:      %s", rpt.name)


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
    return parser.parse_args()


if __name__ == "__main__":
    args = _parse_args()
    run(args.operation, args.domain)
