from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
import re
import unicodedata

import openpyxl


HEADER_ALIASES: dict[str, str] = {
    "planta": "plant_code",
    "pr/solped": "purchase_requisition",
    "contrato marco": "outline_contract",
    "documento compras": "purchase_document",
    "posicion": "position",
    "material": "material",
    "proveedor": "vendor",
    "texto breve": "short_text",
    "tipo de posicion": "position_type",
    "cl docto compras": "purchase_doc_class",
    "por entregar (cantidad)": "pending_delivery_qty",
    "por entregar (valor)": "pending_delivery_value",
    "fecha documento": "document_date",
    "fecha de entrega": "delivery_date",
    "fecha de termino": "validity_end",
}

IGNORED_SEGMENTATION_COLUMNS = {
    "pending_delivery_qty",
    "pending_delivery_value",
    "document_date",
    "delivery_date",
    "validity_end",
}

PRESENCE_TOKENS = {"SI", "SÍ"}
ABSENCE_TOKENS = {"NO"}
SKIP_SECTION = "__skip_section__"


@dataclass(frozen=True)
class SegmentCondition:
    column: str
    operator: str
    value: str = ""


@dataclass(frozen=True)
class SegmentRule:
    segment_id: str
    source_sheet: str
    source_row: int
    conditions: tuple[SegmentCondition, ...]


def normalize_text(value: object) -> str:
    if value is None:
        return ""
    text = str(value).strip()
    if not text:
        return ""
    numeric = _integer_text(text)
    if numeric is not None:
        text = numeric
    text = unicodedata.normalize("NFKD", text)
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    text = re.sub(r"\s+", " ", text)
    text = text.replace(".", " ")
    text = re.sub(r"\s+", " ", text)
    return text.strip().upper()


def normalize_header(value: object) -> str:
    text = normalize_text(value).lower()
    text = re.sub(r"[^0-9a-z/()]+", " ", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text


def _integer_text(text: str) -> str | None:
    candidate = text.replace(",", ".")
    try:
        number = float(candidate)
    except ValueError:
        return None
    if number.is_integer():
        return str(int(number))
    return None


def _canonical_header(value: object) -> str | None:
    return HEADER_ALIASES.get(normalize_header(value))


def _segment_from_title(value: object) -> str | None:
    title = normalize_text(value)
    if not title.startswith("CASERONES"):
        return None
    if "REPARACION" in title:
        return SKIP_SECTION
    if "ORDENES DE SERVICIO" in title:
        return "ordenes_servicio"
    if "CONTRATOS" in title:
        return "contratos"
    return None


def _is_header_row(values: tuple[object, ...]) -> bool:
    canonical = {_canonical_header(value) for value in values if value is not None}
    return {"plant_code", "purchase_document", "purchase_doc_class"}.issubset(canonical)


def _headers_by_index(values: tuple[object, ...]) -> dict[int, str]:
    headers: dict[int, str] = {}
    for idx, value in enumerate(values):
        canonical = _canonical_header(value)
        if canonical is not None:
            headers[idx] = canonical
    return headers


def _row_conditions(values: tuple[object, ...], headers: dict[int, str]) -> dict[str, str]:
    conditions: dict[str, str] = {}
    for idx, column in headers.items():
        if idx >= len(values):
            continue
        value = values[idx]
        if value is None or str(value).strip() == "":
            continue
        conditions[column] = str(value).strip()
    return conditions


def _is_full_template(row: dict[str, str]) -> bool:
    required = {"plant_code", "purchase_document", "position", "purchase_doc_class"}
    return required.issubset(row)


def _compile_conditions(raw_conditions: dict[str, str]) -> tuple[SegmentCondition, ...]:
    compiled: list[SegmentCondition] = []
    for column, value in raw_conditions.items():
        if column in IGNORED_SEGMENTATION_COLUMNS:
            continue
        normalized = normalize_text(value)
        if normalized in PRESENCE_TOKENS:
            compiled.append(SegmentCondition(column=column, operator="present"))
        elif normalized in ABSENCE_TOKENS:
            compiled.append(SegmentCondition(column=column, operator="absent"))
        elif normalized:
            compiled.append(SegmentCondition(column=column, operator="equals", value=normalized))
    return tuple(compiled)


def _drop_section_title_condition(raw_conditions: dict[str, str]) -> dict[str, str]:
    cleaned = dict(raw_conditions)
    plant_value = cleaned.get("plant_code")
    if plant_value is not None and _segment_from_title(plant_value) is not None:
        cleaned.pop("plant_code", None)
    return cleaned


def _append_rule(
    rules: list[SegmentRule],
    segment_id: str | None,
    sheet_name: str,
    source_row: int,
    raw_conditions: dict[str, str],
) -> None:
    if segment_id is None:
        return
    conditions = _compile_conditions(raw_conditions)
    if not conditions:
        return
    rules.append(SegmentRule(segment_id, sheet_name, source_row, conditions))


def load_segment_rules(path: Path) -> list[SegmentRule]:
    wb = openpyxl.load_workbook(path, read_only=False, data_only=True)
    rules: list[SegmentRule] = []
    try:
        for ws in wb.worksheets:
            current_segment: str | None = None
            current_headers: dict[int, str] = {}
            known_headers: dict[int, str] = {}
            last_full_conditions: dict[str, str] | None = None
            pending_partials: list[tuple[int, dict[str, str]]] = []

            for row_idx, row in enumerate(ws.iter_rows(values_only=True), 1):
                values = tuple(row)
                first_value = values[0] if values else None
                section_segment = _segment_from_title(first_value)

                if section_segment is not None:
                    if section_segment == SKIP_SECTION:
                        current_segment = None
                        current_headers = {}
                        last_full_conditions = None
                        pending_partials = []
                        continue
                    current_segment = section_segment
                    current_headers = {}
                    last_full_conditions = None
                    partial = _drop_section_title_condition(_row_conditions(values, known_headers))
                    if partial:
                        pending_partials.append((row_idx, partial))
                    continue

                if _is_header_row(values):
                    current_headers = _headers_by_index(values)
                    known_headers.update(current_headers)
                    continue

                headers = current_headers or known_headers
                if not headers or current_segment is None:
                    continue

                row_conditions = _row_conditions(values, headers)
                if not row_conditions:
                    continue

                if _is_full_template(row_conditions):
                    if pending_partials:
                        for pending_row, pending in pending_partials:
                            merged = dict(row_conditions)
                            merged.update(pending)
                            _append_rule(rules, current_segment, ws.title, pending_row, merged)
                        pending_partials = []
                    _append_rule(rules, current_segment, ws.title, row_idx, row_conditions)
                    last_full_conditions = row_conditions
                    continue

                if last_full_conditions is not None:
                    merged = dict(last_full_conditions)
                    merged.update(row_conditions)
                    _append_rule(rules, current_segment, ws.title, row_idx, merged)
                else:
                    pending_partials.append((row_idx, row_conditions))
    finally:
        wb.close()
    return rules
