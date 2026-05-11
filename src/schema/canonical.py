from dataclasses import dataclass


@dataclass(frozen=True)
class CanonicalColumn:
    canonical_name: str
    description_es: str
    dtype: str        # "str" | "float" | "date"
    nullable: bool
    is_key: bool
    is_critical: bool
    column_group: str  # identity | dates | classification | quantity | pricing | description | metadata


CANONICAL_SCHEMA: list[CanonicalColumn] = [
    CanonicalColumn("position",               "Posición",                            "str",   False, True,  True,  "identity"),
    CanonicalColumn("purchase_document",       "Documento Compras",                   "str",   False, True,  True,  "identity"),
    CanonicalColumn("material",                "Material",                            "str",   True,  False, False, "identity"),
    CanonicalColumn("purchase_doc_class",      "Cl.documento compras",                "str",   True,  False, False, "classification"),
    CanonicalColumn("purchase_doc_type",       "Tipo doc.compras",                    "str",   True,  False, False, "classification"),
    CanonicalColumn("purchase_group",          "Grupo de compras",                    "str",   True,  False, False, "classification"),
    CanonicalColumn("order_history_ref",       "Historial pedido/Docu.orden entrega", "str",   True,  False, False, "identity"),
    CanonicalColumn("document_date",           "Fecha documento",                     "date",  True,  False, False, "dates"),
    CanonicalColumn("vendor",                  "Proveedor/Centro suministrador",      "str",   True,  False, True,  "identity"),
    CanonicalColumn("short_text",              "Texto breve",                         "str",   True,  False, True,  "description"),
    CanonicalColumn("validity_start",          "In.período validez",                  "date",  True,  False, False, "dates"),
    CanonicalColumn("validity_end",            "Fin período validez",                 "date",  True,  False, True,  "dates"),
    CanonicalColumn("article_group",           "Grupo de artículos",                  "str",   True,  False, False, "classification"),
    CanonicalColumn("deletion_flag",           "Indicador de borrado",                "str",   True,  False, False, "classification"),
    CanonicalColumn("position_type",           "Tipo de posición",                    "str",   True,  False, False, "classification"),
    CanonicalColumn("account_assignment_type", "Tipo de imputación",                  "str",   True,  False, False, "classification"),
    CanonicalColumn("plant_code",              "Centro",                              "str",   True,  False, False, "identity"),
    CanonicalColumn("warehouse_code",          "Almacén",                             "str",   True,  False, False, "identity"),
    CanonicalColumn("order_quantity",          "Cantidad de pedido",                  "float", True,  False, False, "quantity"),
    CanonicalColumn("order_uom",               "Unidad medida pedido",                "str",   True,  False, False, "quantity"),
    CanonicalColumn("quantity_uma",            "Cantidad en UMA",                     "float", True,  False, False, "quantity"),
    CanonicalColumn("warehouse_uom",           "Unidad de medida de almacén",         "str",   True,  False, False, "quantity"),
    CanonicalColumn("net_price",               "Precio neto",                         "float", True,  False, False, "pricing"),
    CanonicalColumn("currency",                "Moneda",                              "str",   True,  False, False, "pricing"),
    CanonicalColumn("base_quantity",           "Cantidad base",                       "float", True,  False, False, "quantity"),
    CanonicalColumn("estimated_value",         "Val.prev.(cab.)",                     "float", True,  False, False, "pricing"),
    CanonicalColumn("planned_quantity",        "Cantidad prevista",                   "float", True,  False, False, "quantity"),
    CanonicalColumn("pending_planned_qty",     "Ctd.prev.pendiente",                  "float", True,  False, False, "quantity"),
    CanonicalColumn("position_count",          "Cantidad de posiciones",              "float", True,  False, False, "quantity"),
    CanonicalColumn("_source_file",            "Archivo de origen (inyectada)",       "str",   False, False, False, "metadata"),
    CanonicalColumn("_operation",              "Operación (inyectada)",               "str",   False, False, False, "metadata"),
]

CANONICAL_BY_NAME: dict[str, CanonicalColumn] = {c.canonical_name: c for c in CANONICAL_SCHEMA}
CANONICAL_NAMES: list[str] = [c.canonical_name for c in CANONICAL_SCHEMA]
