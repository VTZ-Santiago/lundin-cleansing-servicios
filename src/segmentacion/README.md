# `src/segmentacion` — Segmentación MLCC/CCMC desde órdenes de compra

Entrada principal: `run_segmentacion_mlcc.py` (raíz). Aquí viven el loader, los
criterios de caracterización, el cruce de vigencia y los dos motores de
segmentación: el de **tipo D** (Contratos / Órdenes de Servicio) y el de
**Suministros** (no-D).

## Dos universos, dos fuentes

| Universo | Fuente | Procesamiento |
|---|---|---|
| **Tipo D** (Contratos + Órdenes de Servicio) | `inputs/<OP>/ordenes-compra` | enriquecimiento de materiales → `segment_dataframe` con los criterios D de la caracterización → cruce de vigencia (`contract_validity`) → E01 → C2 / C2_NO_MIGRA |
| **Suministros** (no-D) | `inputs/<OP>/suministros-ordenes-compra` | M01 (vigencia por fecha→saldo) → `classify_supply_segments` (sub-bloque) → reporte `suministros_<op>_*.xlsx` |

El loader (`po_loader.load_purchase_orders`) recibe `subdir` para elegir la
fuente; el mapeo de headers cubre los formatos ES (MLCC) e EN (CCMC) y normaliza
el tipo de posición a la letra canónica. Cada universo se cachea por separado
(`po_universe_<op>` y `po_universe_<op>_suministros`).

## Sub-segmentación de Suministros (`supply_segments.py`)

`classify_supply_segments(df)` rotula cada fila no-D en un sub-bloque **sin
solape y exhaustivo**, usando solo el tipo de posición y la presencia de
contrato marco (no usa fecha ni saldo — eso es vigencia):

| Sub-bloque | Criterio (MLCC + CCMC) |
|---|---|
| **Reparación** | tipo `L` (subcontratación) |
| **Consignación** | tipo `C` (MLCC) / `K` (CCMC) |
| **Traslado/Transporte** | tipo `V` (MLCC) / `U` (CCMC) |
| **Stock** | tipo vacío/estándar **con** contrato marco |
| **Cargo Directo** | tipo vacío/estándar **sin** contrato marco |
| **Otros** | residual (p. ej. `P` límite) |

Las letras no colisionan entre operaciones, así que un único conjunto por
sub-bloque sirve para ambas.

### Por qué no se usa el motor `segment_dataframe` para Suministros

Los bloques no-D de la caracterización (`tmp/Caracterización de Contratos.xlsx`)
se solapan entre sí (la misma fila "V·marco·ZADI" aparece en Consignación,
Stock y Cargo Directo) y el parser `characterization.load_segment_rules` además
**salta** explícitamente el bloque de reparación. Por eso la rotulación de
Suministros es una clasificación dedicada y deliberadamente independiente.

## Vigencia

- **Tipo D**: la vigencia **no** sale de la base OC; viene del cruce con los
  reportes de contratos vigentes (`contract_validity.apply_contract_validity`),
  por contrato marco y, en su defecto, por documento de compra.
- **Suministros**: la vigencia es M01 (`rules/group3/m01.py`) — fecha efectiva
  (`delivery_date`, fallback `validity_end`) contra el corte, y luego el saldo
  pendiente decide si un vencido igual migra (`NO_VIGENTE_CON_SALDO`).
