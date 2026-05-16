from src.schema.field_map_mlcc import MLCC_DTYPE_COERCIONS, MLCC_RAW_TO_CANONICAL, normalize_header


MLCC_RAW_TO_CANONICAL = dict(MLCC_RAW_TO_CANONICAL)
MLCC_RAW_TO_CANONICAL.pop("Tipo de posición", None)
MLCC_RAW_TO_CANONICAL.pop("Tipo de posicion", None)
MLCC_RAW_TO_CANONICAL["Tipo de posición.1"] = "position_type"
MLCC_RAW_TO_CANONICAL["Tipo de posicion.1"] = "position_type"