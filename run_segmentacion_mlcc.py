"""Segment MLCC purchase-order inputs into initial cleansing groups."""
from __future__ import annotations

import argparse
from dataclasses import dataclass
from datetime import datetime
import json
from pathlib import Path
import pickle
import shutil
import sys
from typing import Any

import pandas as pd
import yaml

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT))

from src.lineage.records import FieldLineageRecord, StageManifest
from src.material_crossref.enrichment import enrich_material_rows
from src.material_crossref.loader import load_material_master
from src.ordenes_compra.exchange_rates import load_usd_rates
from src.output.suministros_report import generate_suministros_report
from src.output.control_point import export_control_point
from src.segmentacion.rules.group3.m01 import apply as _apply_m01
from src.output.deliverable import export_deliverable
from src.profiling.profiler import DataProfiler
from src.rules.engine import RuleEngine
from src.segmentacion.characterization import SegmentRule, load_segment_rules
from src.segmentacion.contract_validity import apply_contract_validity, load_vigentes_index
from src.segmentacion.rules.group1.e01 import before_cutoff_mask
from src.segmentacion.po_loader import PurchaseOrderLoadResult, load_purchase_orders
from src.segmentacion.report import build_segment_summary, write_summary_markdown
from src.segmentacion.segments import SEGMENT_LABELS, segment_dataframe
from src.segmentacion.supply_segments import SUPPLY_SEGMENT_COLUMN, classify_supply_segments


DEFAULT_OUTPUT_DIR = ROOT / "outputs" / "control_points"
DEFAULT_ENTREGABLES_DIR = ROOT / "outputs" / "entregables"
DEFAULT_CACHE_DIR = ROOT / "tmp" / "cache" / "segmentacion_mlcc"
LEGACY_OUTPUT_DIR = ROOT / "outputs" / "segmentacion_mlcc"
CACHE_VERSION = "2026-06-10-v1"
LEGACY_ROOT_CONTROL_POINTS = (
    "C1_MLCC.xlsx",
    "C2_MLCC.xlsx",
    "C2_NO_MIGRA_MLCC.xlsx",
    "C3_MLCC.xlsx",
)
RULES_YAML = ROOT / "src" / "segmentacion" / "rules" / "rules.yaml"
CONTROL_POINT_COLUMNS = [
    "plant_code",
    "purchase_requisition",
    "outline_contract",
    "purchase_document",
    "position",
    "material",
    "vendor",
    "short_text",
    "position_type",
    "purchase_doc_class",
    "purchase_group",
    "release_group",
    "pending_delivery_qty",
    "pending_delivery_value",
    "document_date",
    "delivery_date",
    "validity_end",
    "contract_end_date",
    "contract_match_source",
    "deletion_flag",
    "migration_category",
    "delivery_year",
]


@dataclass(frozen=True)
class SegmentRuleSettings:
    domain: str = "segmentacion"
    rules_yaml_path: Path = RULES_YAML


def _first_match(directory: Path, pattern: str) -> Path:
    matches = sorted(directory.glob(pattern))
    if not matches:
        raise FileNotFoundError(f"No file found in {directory} matching {pattern}")
    return matches[0]


def _characterization_path() -> Path:
    return _first_match(ROOT / "tmp", "Car*.xlsx")


def _material_master_path(operation: str) -> Path:
    return _first_match(ROOT / "inputs" / operation.upper(), f"{operation.upper()} - An*lisis Completo.xlsx")


def _po_source_files(operation: str, subdir: str = "ordenes-compra") -> list[Path]:
    directory = ROOT / "inputs" / operation.upper() / subdir
    files = [
        path
        for path in list(directory.glob("*.xlsx")) + list(directory.glob("*.XLSX"))
        if not path.name.startswith("~$")
    ]
    unique = sorted({path.resolve(): path for path in files}.values(), key=lambda path: path.name.lower())
    if not unique:
        raise FileNotFoundError(f"No Excel files found in {directory}")
    return unique


def _path_signature(path: Path) -> dict[str, Any]:
    resolved = path.resolve()
    stat = resolved.stat()
    try:
        display_path = resolved.relative_to(ROOT).as_posix()
    except ValueError:
        display_path = resolved.as_posix()
    return {
        "path": display_path,
        "size": stat.st_size,
        "mtime_ns": stat.st_mtime_ns,
    }


def _cache_paths(cache_name: str) -> tuple[Path, Path]:
    return DEFAULT_CACHE_DIR / f"{cache_name}.json", DEFAULT_CACHE_DIR / f"{cache_name}.pkl"


def _load_cache(cache_name: str, cache_label: str, metadata: dict[str, Any], use_cache: bool) -> Any | None:
    if not use_cache:
        print(f"Cache desactivado: {cache_label}.")
        return None

    metadata_path, payload_path = _cache_paths(cache_name)
    if not metadata_path.exists() or not payload_path.exists():
        print(f"Cache miss: {cache_label}.")
        return None

    try:
        cached_metadata = json.loads(metadata_path.read_text(encoding="utf-8"))
        if cached_metadata != metadata:
            print(f"Cache miss: {cache_label} invalidado por cambios en las fuentes.")
            return None
        with payload_path.open("rb") as handle:
            payload = pickle.load(handle)
    except Exception as exc:  # pragma: no cover - defensive fallback for corrupt cache files
        print(f"Cache miss: {cache_label} no pudo leerse ({exc}).")
        return None

    print(f"Cache hit: {cache_label}.")
    return payload


def _write_cache(cache_name: str, metadata: dict[str, Any], payload: Any, cache_label: str) -> None:
    DEFAULT_CACHE_DIR.mkdir(parents=True, exist_ok=True)
    metadata_path, payload_path = _cache_paths(cache_name)
    payload_tmp = Path(f"{payload_path}.tmp")
    metadata_tmp = Path(f"{metadata_path}.tmp")

    with payload_tmp.open("wb") as handle:
        pickle.dump(payload, handle, protocol=pickle.HIGHEST_PROTOCOL)
    metadata_tmp.write_text(json.dumps(metadata, ensure_ascii=False, indent=2, sort_keys=True), encoding="utf-8")

    payload_tmp.replace(payload_path)
    metadata_tmp.replace(metadata_path)
    print(f"Cache actualizado: {cache_label}.")


def _po_cache_metadata(operation: str, subdir: str = "ordenes-compra") -> dict[str, Any]:
    return {
        "cache_version": CACHE_VERSION,
        "cache_kind": "po_universe",
        "operation": operation.upper(),
        "subdir": subdir,
        "source_files": [_path_signature(path) for path in _po_source_files(operation, subdir)],
    }


def _enrichment_cache_metadata(operation: str) -> dict[str, Any]:
    return {
        "cache_version": CACHE_VERSION,
        "cache_kind": "material_enrichment",
        "operation": operation.upper(),
        "source_files": [_path_signature(path) for path in _po_source_files(operation)],
        "material_master": _path_signature(_material_master_path(operation)),
    }


def _load_purchase_orders_cached(
    operation: str,
    use_cache: bool,
    subdir: str = "ordenes-compra",
    cache_suffix: str = "",
) -> PurchaseOrderLoadResult:
    cache_name = f"po_universe_{operation.lower()}{cache_suffix}"
    cache_label = f"universo OC {operation.upper()} ({subdir})"
    metadata = _po_cache_metadata(operation, subdir)
    cached = _load_cache(cache_name, cache_label, metadata, use_cache)
    if cached is not None:
        return cached

    print(f"Cargando ordenes de compra de todos los periodos ({subdir})...")
    load_result = load_purchase_orders(ROOT / "inputs", operation=operation, subdir=subdir)
    if use_cache:
        _write_cache(cache_name, metadata, load_result, cache_label)
    return load_result


def _load_material_enrichment_cached(
    df: pd.DataFrame,
    operation: str,
    load_result: PurchaseOrderLoadResult,
    use_cache: bool,
) -> tuple[pd.DataFrame, StageManifest]:
    cache_name = f"material_enrichment_{operation.lower()}"
    cache_label = "cruce con Analisis Completo"
    metadata = _enrichment_cache_metadata(operation)
    cached = _load_cache(cache_name, cache_label, metadata, use_cache)
    if cached is not None:
        load_result.lineage = cached["lineage"]
        return cached["dataframe"], cached["manifest"]

    print("Cruzando con Analisis Completo...")
    enriched, manifest = _enrich_with_material_master(df, operation, load_result)
    if use_cache:
        _write_cache(
            cache_name,
            metadata,
            {
                "dataframe": enriched,
                "lineage": load_result.lineage,
                "manifest": manifest,
            },
            cache_label,
        )
    return enriched, manifest


def _is_default_output_dir(output_dir: Path) -> bool:
    return output_dir.resolve() == DEFAULT_OUTPUT_DIR.resolve()


def _cleanup_legacy_outputs(output_dir: Path) -> None:
    if not _is_default_output_dir(output_dir):
        return

    removed: list[str] = []
    for filename in LEGACY_ROOT_CONTROL_POINTS:
        legacy_path = output_dir / filename
        if legacy_path.exists():
            legacy_path.unlink()
            removed.append(legacy_path.relative_to(ROOT).as_posix())

    if LEGACY_OUTPUT_DIR.exists():
        shutil.rmtree(LEGACY_OUTPUT_DIR)
        removed.append(LEGACY_OUTPUT_DIR.relative_to(ROOT).as_posix())

    if removed:
        print("Limpieza legacy completada: " + ", ".join(removed))


def _add_injected_lineage(df: pd.DataFrame, columns: list[str], source: str, load_result: PurchaseOrderLoadResult) -> None:
    total = max(len(df), 1)
    for column in columns:
        if column not in df.columns:
            continue
        series = df[column]
        null_count = int(series.isna().sum())
        load_result.lineage.add(FieldLineageRecord(
            source_file=source,
            raw_name="",
            canonical_name=column,
            mapping_status="INJECTED",
            null_count=null_count,
            null_pct=round(null_count / total * 100, 2),
            unique_count=int(series.nunique(dropna=True)),
            sample_values=[str(value) for value in series.dropna().iloc[:5].tolist()],
        ))


def _enrich_with_material_master(
    df: pd.DataFrame,
    operation: str,
    load_result: PurchaseOrderLoadResult,
) -> tuple[pd.DataFrame, StageManifest]:
    started = datetime.now()
    columns_before = set(df.columns)
    master_path = _material_master_path(operation)
    material_master = load_material_master(master_path, operation=operation)
    empty_consumptions = pd.DataFrame({"material_key": pd.Series(dtype="object")})
    enriched = enrich_material_rows(df, material_master, empty_consumptions)
    new_columns = [column for column in enriched.columns if column not in columns_before]
    _add_injected_lineage(enriched, new_columns, master_path.name, load_result)
    manifest = StageManifest(
        stage_id="MATERIAL_MASTER_XREF",
        label="Cruce con Analisis Completo",
        started_at=started,
        completed_at=datetime.now(),
        rows_in=len(df),
        rows_out=len(enriched),
        columns_in=len(columns_before),
        columns_out=len(enriched.columns),
        notes=f"Maestro: {master_path.name}. Columnas inyectadas: {', '.join(new_columns)}",
    )
    return enriched, manifest


def _segment_manifest(total_rows: int, classified_rows: int, overlap_rows: int, unclassified_rows: int) -> StageManifest:
    now = datetime.now()
    return StageManifest(
        stage_id="SEGMENTATION_C1",
        label="Segmentacion inicial MLCC",
        started_at=now,
        completed_at=now,
        rows_in=total_rows,
        rows_out=classified_rows,
        columns_in=0,
        columns_out=0,
        notes=(
            f"Clasificadas: {classified_rows:,}. "
            f"No clasificadas: {unclassified_rows:,}. "
            f"Solapamientos: {overlap_rows:,}."
        ),
    )


def _drop_rule_conditions(
    rules: list[SegmentRule],
    operation: str,
) -> list[SegmentRule]:
    operation = operation.upper()
    drop_columns_by_segment: dict[str, set[str]] = {}
    if operation == "MLCC":
        drop_columns_by_segment = {"contratos": {"outline_contract"}}
    elif operation == "CCMC":
        drop_columns_by_segment = {
            "contratos": {"plant_code", "purchase_doc_class"},
            "ordenes_servicio": {"plant_code", "purchase_doc_class"},
        }

    adjusted: list[SegmentRule] = []
    for rule in rules:
        drop_columns = drop_columns_by_segment.get(rule.segment_id, set())
        conditions = tuple(condition for condition in rule.conditions if condition.column not in drop_columns)
        if conditions:
            adjusted.append(SegmentRule(rule.segment_id, rule.source_sheet, rule.source_row, conditions))
    return adjusted


def _prune_absent_conditions(
    rules: list[SegmentRule],
    available_columns: set[str],
) -> tuple[list[SegmentRule], set[str]]:
    """Quita condiciones cuya columna no existe en absoluto en el universo.

    Una condición ``present``/``absent`` sobre una columna que la fuente no trae
    es degenerada (present => no matchea nada; absent => matchea todo). En vez de
    anular el universo, se ignora. Solo afecta columnas ENTERAMENTE ausentes; los
    nulos por fila no se tocan. Devuelve (reglas_podadas, columnas_ignoradas).
    """
    dropped: set[str] = set()
    pruned: list[SegmentRule] = []
    for rule in rules:
        conditions = tuple(c for c in rule.conditions if c.column in available_columns)
        dropped.update(c.column for c in rule.conditions if c.column not in available_columns)
        if conditions:
            pruned.append(SegmentRule(rule.segment_id, rule.source_sheet, rule.source_row, conditions))
    return pruned, dropped


def _segment_priority(operation: str) -> list[str] | None:
    # Sin prioridad: las filas que caen en ambos segmentos (p.ej. clase ZADI)
    # se conservan DUPLICADAS en contratos y en ordenes_servicio. Requiere
    # ejecutar con allow_overlap=True para no abortar por el solape.
    return None


def _export_cp(
    cp_id: str,
    description: str,
    operation: str,
    df: pd.DataFrame,
    manifests: list[StageManifest],
    issues,
    output_dir: Path,
    load_result: PurchaseOrderLoadResult,
) -> Path:
    output_dir.mkdir(parents=True, exist_ok=True)
    export_df = df[[column for column in CONTROL_POINT_COLUMNS if column in df.columns]].copy()
    profiling, profiling_manifest = DataProfiler().profile(export_df, operation)
    return export_control_point(
        cp_id=cp_id,
        description=description,
        operation=operation,
        df=export_df,
        lineage=load_result.lineage,
        profiling=profiling,
        manifests=[*manifests, profiling_manifest],
        issues=issues,
        output_dir=output_dir,
        include_analysis=False,
        analysis_subject=f"Segmentacion {operation}",
        master_columns=CONTROL_POINT_COLUMNS,
    )


def _load_e01_config() -> dict:
    with open(RULES_YAML, encoding="utf-8") as f:
        cfg = yaml.safe_load(f)
    for rule in cfg["groups"]["G1_EXCLUSIONS"]["rules"]:
        if rule["id"] == "E01":
            return rule.get("config", {})
    return {}


def _load_m01_config() -> dict:
    with open(RULES_YAML, encoding="utf-8") as f:
        cfg = yaml.safe_load(f)
    for rule in cfg["groups"]["G3_MARKING"]["rules"]:
        if rule["id"] == "M01":
            return rule.get("config", {})
    return {}


def _load_complementaria_report_config() -> dict:
    with open(RULES_YAML, encoding="utf-8") as f:
        cfg = yaml.safe_load(f)
    return cfg.get("complementaria_report", {})


def _vencidos_con_saldo_pendiente(df: pd.DataFrame, e01_config: dict) -> pd.DataFrame:
    """Return rows without a vigente contract that still carry a pending delivery balance.

    These records are already excluded (they land in C2_NO_MIGRA), but the combination
    of an expired/missing contract with outstanding balance may warrant follow-up.
    Uses the same expiry mask as E01 (cruce de vigencia incluido).
    """
    qty_col = e01_config.get("pending_qty_column", "pending_delivery_qty")
    value_col = e01_config.get("pending_value_column", "pending_delivery_value")

    pending_qty = pd.to_numeric(df.get(qty_col, pd.Series(0, index=df.index)), errors="coerce").fillna(0)
    pending_value = pd.to_numeric(df.get(value_col, pd.Series(0, index=df.index)), errors="coerce").fillna(0)

    expired = before_cutoff_mask(df, e01_config)
    has_balance = (pending_qty > 0) | (pending_value > 0)
    return df[expired & has_balance].copy()


def _apply_migration_rules(df: pd.DataFrame, operation: str) -> tuple[pd.DataFrame, list, list[StageManifest]]:
    engine = RuleEngine(SegmentRuleSettings())
    df_g1, g1_issues, g1_manifest = engine.apply_group("G1_EXCLUSIONS", df, operation)
    df_g2, g2_issues, g2_manifest = engine.apply_group("G2_RESCUE", df_g1, operation)
    df_g3, g3_issues, g3_manifest = engine.apply_group("G3_MARKING", df_g2, operation)
    return df_g3, g1_issues + g2_issues + g3_issues, [g1_manifest, g2_manifest, g3_manifest]


def _run_suministros_branch(operation: str, output_dir: Path, use_cache: bool) -> None:
    """Procesa el universo SUMINISTROS (no-D) desde su fuente dedicada.

    Carga `inputs/<OP>/suministros-ordenes-compra`, aplica M01 (vigencia por
    fecha y saldo), rotula el sub-bloque (`supply_segment`) de forma ortogonal a
    la vigencia y genera el reporte de Suministros. No interviene la
    segmentación tipo D, que vive en `ordenes-compra`.
    """
    m01_config = _load_m01_config()
    report_cfg = _load_complementaria_report_config()
    reportes_dir = output_dir.parent / report_cfg.get("output_subdir", "reportes")
    enabled = report_cfg.get("enabled", {})
    if isinstance(enabled, dict):
        enabled = enabled.get(operation, True)
    if not (enabled and m01_config):
        return

    load_result = _load_purchase_orders_cached(
        operation,
        use_cache=use_cache,
        subdir="suministros-ordenes-compra",
        cache_suffix="_suministros",
    )
    df = load_result.dataframe
    print(f"Filas Suministros limpias {operation}: {len(df):,}", flush=True)

    # M01: clasifica vigencia (fecha -> saldo) sobre todo el universo; el tipo D
    # cae en TIPO_D y queda fuera del reporte de Suministros.
    m01_result = _apply_m01(df, operation, m01_config)
    for col, series in m01_result.updates.items():
        df[col] = series.values

    cat_col = m01_config.get("output_column", "migration_category")
    excluded_cat = m01_config.get("category_excluded_type", "TIPO_D")
    if cat_col in df.columns:
        cat_counts = df[cat_col].value_counts()
        print(
            f"Suministros {operation} universo completo: "
            + " | ".join(f"{cat}={cnt:,}" for cat, cnt in cat_counts.items()),
            flush=True,
        )

    # Sub-segmentación (rótulo ortogonal a la vigencia).
    df[SUPPLY_SEGMENT_COLUMN] = classify_supply_segments(df)
    non_d = df[df[cat_col] != excluded_cat] if cat_col in df.columns else df
    seg_counts = non_d[SUPPLY_SEGMENT_COLUMN].value_counts()
    print(
        f"Suministros {operation} no-D por sub-bloque ({len(non_d):,}): "
        + " | ".join(f"{seg}={cnt:,}" for seg, cnt in seg_counts.items()),
        flush=True,
    )

    # Tasas USD (fallback silencioso si no hay red).
    usd_rates: dict = {}
    try:
        _fx_cache = ROOT / "tmp" / "cache" / "exchange_rates.json"
        usd_rates = load_usd_rates(
            df.get("currency", []),
            cache_path=_fx_cache,
            as_of_date=datetime.now().strftime("%Y-%m-%d"),
        )
        print(f"Tasas USD cargadas para: {', '.join(sorted(usd_rates))}", flush=True)
    except Exception as exc:
        print(f"Advertencia: no se pudieron cargar tasas USD ({exc}). Reporte sin columnas USD.", flush=True)

    report_path = generate_suministros_report(
        segment_id=operation.lower(),
        label=f"Suministros {operation}",
        df=df,
        m01_config=m01_config,
        output_dir=reportes_dir,
        usd_rates=usd_rates,
    )
    if report_path:
        print(f"Reporte Suministros: {report_path.relative_to(ROOT).as_posix()}", flush=True)


def run(
    operation: str = "MLCC",
    output_dir: Path = DEFAULT_OUTPUT_DIR,
    export_control_points: bool = True,
    export_entregables: bool = True,
    enrich_materials: bool = True,
    allow_overlap: bool = False,
    use_cache: bool = True,
) -> None:
    operation = operation.upper()
    output_dir.mkdir(parents=True, exist_ok=True)

    print(f"Cargando criterios de segmentacion desde {_characterization_path()}...")
    segment_rules = _drop_rule_conditions(load_segment_rules(_characterization_path()), operation)
    print(f"Criterios de segmentacion: {len(segment_rules):,}")

    load_result = _load_purchase_orders_cached(operation, use_cache=use_cache)
    master = load_result.dataframe
    manifests = list(load_result.manifests)
    print(f"Filas OC limpias: {len(master):,}")

    if enrich_materials:
        master, material_manifest = _load_material_enrichment_cached(master, operation, load_result, use_cache=use_cache)
        manifests.append(material_manifest)

    # --- SUMINISTROS (universo no-D): se procesa aparte, leyendo su propia fuente
    # (inputs/<OP>/suministros-ordenes-compra). La carga `master` de arriba queda
    # reservada para las posiciones tipo D (Contratos y Órdenes de Servicio).
    _run_suministros_branch(operation, output_dir, use_cache=use_cache)

    segment_rules, dropped_cols = _prune_absent_conditions(segment_rules, set(master.columns))
    if dropped_cols:
        print(
            f"Aviso: columnas ausentes en {operation} ordenes-compra, condiciones ignoradas: "
            + ", ".join(sorted(dropped_cols)),
            flush=True,
        )

    segment_priority = _segment_priority(operation)
    print("Segmentando universo OC...", flush=True)
    segments, match_result = segment_dataframe(master, segment_rules, priority=segment_priority)
    classified_rows = sum(len(df) for df in segments.values())
    print(
        f"Segmentacion lista: clasificados={classified_rows:,} | "
        f"solapamientos={len(match_result.overlaps):,} | "
        f"sin_clasificar={len(match_result.unclassified):,}",
        flush=True,
    )
    segmentation_manifest = _segment_manifest(
        total_rows=len(master),
        classified_rows=classified_rows,
        overlap_rows=len(match_result.overlaps),
        unclassified_rows=len(match_result.unclassified),
    )
    manifests.append(segmentation_manifest)

    if len(match_result.overlaps) and not (allow_overlap or segment_priority):
        sample_path = output_dir / f"solapamientos_{operation}.xlsx"
        match_result.overlaps.head(200).to_excel(sample_path, index=False)
        raise RuntimeError(
            f"Hay {len(match_result.overlaps):,} filas con solapamiento entre segmentos. "
            f"Muestra guardada en {sample_path}."
        )

    e01_config = _load_e01_config()
    entregables_dir = output_dir.parent / "entregables"

    vigentes_index = load_vigentes_index(operation, ROOT / "inputs")
    print(
        f"Reporte de contratos vigentes {operation}: {vigentes_index.source_file} | "
        + " | ".join(f"{sheet}: {count:,}" for sheet, count in vigentes_index.sheet_counts.items())
        + f" | claves totales: {vigentes_index.total_keys:,}",
        flush=True,
    )

    summaries = []
    for segment_id in sorted(segments):
        label = SEGMENT_LABELS.get(segment_id, segment_id)
        c1 = segments[segment_id]
        segment_output_dir = output_dir / segment_id
        print(f"Procesando segmento {operation} {label}: C1={len(c1):,}", flush=True)

        c1, cruce_stats = apply_contract_validity(c1, vigentes_index)
        print(
            f"  Cruce vigencia {label}: por marco={cruce_stats['by_contract']:,} | "
            f"por documento={cruce_stats['by_document']:,} | sin cruce={cruce_stats['sin_cruce']:,}",
            flush=True,
        )

        if export_control_points:
            _export_cp(
                "C1",
                f"Segmentacion inicial: {label.lower()} desde ordenes de compra {operation}.",
                operation,
                c1,
                manifests,
                [],
                segment_output_dir,
                load_result,
            )

        ruled, issues, rule_manifests = _apply_migration_rules(c1, operation)
        c2_no_migra = ruled[ruled["exclusion_reason"] != ""].copy()
        c2 = ruled[ruled["exclusion_reason"] == ""].copy()
        print(
            f"  Reglas {operation} {label}: C2={len(c2):,} | "
            f"C2_NO_MIGRA={len(c2_no_migra):,}",
            flush=True,
        )

        if export_control_points:
            cp_manifests = [*manifests, *rule_manifests]
            _export_cp(
                "C2",
                f"Post-exclusion comun: {label.lower()} que migran.",
                operation,
                c2,
                cp_manifests,
                issues,
                segment_output_dir,
                load_result,
            )
            _export_cp(
                "C2_NO_MIGRA",
                f"Post-exclusion comun: {label.lower()} que no migran.",
                operation,
                c2_no_migra,
                cp_manifests,
                issues,
                segment_output_dir,
                load_result,
            )

        if export_entregables:
            export_cols = [col for col in CONTROL_POINT_COLUMNS if col in c2.columns]
            export_deliverable(
                c2[export_cols].copy(),
                entregables_dir / f"OC_migra_{segment_id}_{operation}.xlsx",
                sheet_name="C2_MIGRA",
            )
            vencidos = _vencidos_con_saldo_pendiente(c2_no_migra, e01_config)
            export_deliverable(
                vencidos[export_cols].copy(),
                entregables_dir / f"OC_vencidos_saldo_{segment_id}_{operation}.xlsx",
                sheet_name="VENCIDOS_SALDO_PENDIENTE",
            )
            print(
                f"  Entregables {operation} {label}: "
                f"migra={len(c2):,} | vencidos_saldo={len(vencidos):,}"
            )

        summary = build_segment_summary(segment_id, label, c1, c2, c2_no_migra)
        summaries.append(summary)
        print(
            f"{label}: C1={summary.c1_rows:,} | "
            f"C2={summary.c2_rows:,} | "
            f"C2_NO_MIGRA={summary.c2_no_migra_rows:,} | "
            f"reconciliacion={'OK' if summary.reconciliation_ok else 'FALLA'}"
        )
        if not summary.reconciliation_ok:
            raise RuntimeError(f"Reconciliacion fallida para {label}")

    source_files = [stat.source_file for stat in load_result.file_stats if not stat.skipped]
    summary_path = write_summary_markdown(
        output_dir / f"resumen_segmentacion_{operation}.md",
        operation=operation,
        total_rows=len(master),
        rows_read=load_result.rows_read,
        source_files=source_files,
        segmentation_criteria_count=len(segment_rules),
        overlap_rows=len(match_result.overlaps),
        unclassified_rows=len(match_result.unclassified),
        summaries=summaries,
    )

    if export_control_points:
        _cleanup_legacy_outputs(output_dir)

    print(f"Resumen escrito en: {summary_path}")


def _parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Segmenta operaciones en Contratos y Ordenes de Servicio desde ordenes de compra.",
    )
    parser.add_argument("--operation", "-o", default="MLCC", choices=["MLCC", "CCMC"], help="Operacion a procesar.")
    parser.add_argument("--output-dir", type=Path, default=DEFAULT_OUTPUT_DIR, help="Directorio base de salida.")
    parser.add_argument("--skip-control-points", action="store_true", help="Omite workbooks C1/C2/C2_NO_MIGRA.")
    parser.add_argument("--skip-entregables", action="store_true", help="Omite Excel entregables (C2 migra y vencidos con saldo).")
    parser.add_argument("--skip-material-enrichment", action="store_true", help="Omite cruce con Analisis Completo.")
    parser.add_argument("--allow-overlap", action="store_true", help="No falla si una fila cae en mas de un segmento.")
    parser.add_argument("--no-cache", action="store_true", help="Omite cache persistente y reconstruye el flujo base.")
    return parser.parse_args()


if __name__ == "__main__":
    args = _parse_args()
    run(
        operation=args.operation,
        output_dir=args.output_dir,
        export_control_points=not args.skip_control_points,
        export_entregables=not args.skip_entregables,
        enrich_materials=not args.skip_material_enrichment,
        allow_overlap=args.allow_overlap,
        use_cache=not args.no_cache,
    )
