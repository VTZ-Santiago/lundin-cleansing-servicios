from src.schema.field_map_mlcc import MLCC_DTYPE_COERCIONS, MLCC_RAW_TO_CANONICAL, normalize_header


MLCC_RAW_TO_CANONICAL = dict(MLCC_RAW_TO_CANONICAL)
MLCC_DTYPE_COERCIONS = dict(MLCC_DTYPE_COERCIONS)
MLCC_RAW_TO_CANONICAL["Contrato marco"] = "framework_contract"
MLCC_RAW_TO_CANONICAL["Contrato Marco"] = "framework_contract"
MLCC_RAW_TO_CANONICAL["Fecha de entrega"] = "delivery_date"
MLCC_RAW_TO_CANONICAL["Fecha entrega"] = "delivery_date"
MLCC_DTYPE_COERCIONS["framework_contract"] = "str"
MLCC_DTYPE_COERCIONS["delivery_date"] = "date"