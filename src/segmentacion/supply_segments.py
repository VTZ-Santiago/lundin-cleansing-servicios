"""Sub-segmentación de SUMINISTROS (posiciones no-D).

Rotula cada fila no-D en un sub-bloque, ortogonal a la vigencia (M01). Usa el
tipo de posición (letra canónica) y la presencia de contrato marco; no usa
fecha ni saldo (esos definen la categoría de migración M01).

Equivalencia de tipos (núm SAP / CCMC / MLCC):
  3  L  L  subcontratación      -> Reparación
  2  K  C  consignación         -> Consignación
  7  U  V  traslado/transporte  -> Traslado/Transporte
  0  '' '' estándar (vacío)     -> Stock (con marco) / Cargo Directo (sin marco)
  9  D  D  servicio             -> excluido del universo Suministros (TIPO_D)

Las letras no colisionan entre operaciones (MLCC: V/C/L/P/blank · CCMC:
K/L/blank), por lo que un único conjunto por sub-bloque sirve para ambas.
"""
from __future__ import annotations

import pandas as pd

from src.segmentacion.segments import _normalized_series, _present_mask


SUPPLY_SEGMENT_COLUMN = "supply_segment"

SEG_REPARACION = "Reparación"
SEG_CONSIGNACION = "Consignación"
SEG_TRASLADO = "Traslado/Transporte"
SEG_STOCK = "Stock"
SEG_CARGO_DIRECTO = "Cargo Directo"
SEG_OTROS = "Otros"

# Orden de presentación (de mayor a menor especificidad por tipo).
SUPPLY_SEGMENT_ORDER = [
    SEG_REPARACION,
    SEG_CONSIGNACION,
    SEG_TRASLADO,
    SEG_STOCK,
    SEG_CARGO_DIRECTO,
    SEG_OTROS,
]

# Letra canónica de tipo de posición -> sub-bloque (los tipos "vacío" se
# resuelven aparte por presencia de marco).
_TYPE_TO_SEGMENT = {
    "L": SEG_REPARACION,
    "C": SEG_CONSIGNACION,
    "K": SEG_CONSIGNACION,
    "V": SEG_TRASLADO,
    "U": SEG_TRASLADO,
}


def classify_supply_segments(df: pd.DataFrame) -> pd.Series:
    """Return a Series (index alineado a df) con el sub-bloque de cada fila.

    La partición es exhaustiva y sin solape sobre el universo no-D: cada fila
    cae en exactamente un sub-bloque.
    """
    pos = _normalized_series(df, "position_type")
    has_marco = _present_mask(_normalized_series(df, "outline_contract"))

    segment = pd.Series(SEG_OTROS, index=df.index, dtype="object")

    for letter, name in _TYPE_TO_SEGMENT.items():
        segment[pos.eq(letter)] = name

    # Tipo vacío/estándar: Stock (con marco) vs Cargo Directo (sin marco).
    is_blank = pos.eq("") | pos.isin({"0"})
    segment[is_blank & has_marco] = SEG_STOCK
    segment[is_blank & ~has_marco] = SEG_CARGO_DIRECTO

    return segment
