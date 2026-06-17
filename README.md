# lundin-cleansing-servicios

Pipeline de limpieza, identificación, segmentación y control de datos para **convenios/contratos**, **órdenes de compra** y segmentos derivados desde OC MLCC/CCMC.

La rama `all-rules` ajusta las reglas al alcance descrito en `resources/Aumento alcance servicio Data cleansing vf2.pdf`. El alcance activo se limita a los dominios con inputs disponibles en este repositorio:

- `contratos`: convenios/contratos SAP en `inputs/MLCC/contratos`.
- `ordenes-compra`: órdenes de compra **tipo D** (Contratos y Órdenes de Servicio) en `inputs/<operacion>/ordenes-compra`.
- `suministros-ordenes-compra`: órdenes de compra **no-D** (Suministros) en `inputs/<operacion>/suministros-ordenes-compra`.

Los puntos del PDF sobre órdenes de servicio, HES, PR/SOLPED y reservas quedan documentados como alcance no activo porque no existe un dominio de entrada separado para esos paquetes de datos.

## Nuevo flujo: segmentación MLCC/CCMC desde órdenes

El flujo `run_segmentacion_mlcc.py` trabaja con **dos universos independientes** que se leen de **fuentes separadas**:

1. **Tipo D — Contratos y Órdenes de Servicio.** Se cargan todos los Excel de `inputs/<operacion>/ordenes-compra` y se segmentan en:
   - `contratos`: órdenes que cumplen los criterios de contratos definidos en `tmp/Caracterización de Contratos.xlsx`.
   - `ordenes_servicio`: órdenes de servicio definidas en el mismo archivo.
2. **Suministros (no-D).** Se cargan todos los Excel de `inputs/<operacion>/suministros-ordenes-compra` y se procesan aparte (M01 de vigencia + sub-segmentación + reporte). Las filas tipo D que vengan en esa fuente caen en `TIPO_D` y quedan fuera del universo de Suministros.

### Sub-segmentación de Suministros

Cada posición no-D recibe un **sub-bloque** (`supply_segment`) usando el tipo de posición (letra canónica) y la presencia de contrato marco — **ortogonal** a la vigencia. La equivalencia de tipos entre operaciones (el `.xlsx` de caracterización manda en la rotulación):

| Núm SAP | CCMC | MLCC | Significado | Sub-bloque |
|---|---|---|---|---|
| 3 | L | L | subcontratación | **Reparación** |
| 2 | K | C | consignación | **Consignación** |
| 7 | U | V | traslado/transporte | **Traslado/Transporte** |
| 0 | (vacío) | (vacío) | estándar | **Stock** (con marco) / **Cargo Directo** (sin marco) |
| 9 | D | D | servicio | excluido → `TIPO_D` |

La vigencia (M01) se aplica **primero** (por fecha y luego por saldo → VIGENTE / NO_VIGENTE_CON_SALDO que migra / NO_VIGENTE_SIN_SALDO que no migra) y la rotulación del sub-bloque se concilia encima: dentro de un mismo sub-bloque puede haber filas vigentes y no vigentes. Ver `src/segmentacion/README.md`.

La segmentación `C1` ignora fechas y saldos por entregar. En MLCC, el segmento `contratos` no exige `outline_contract` y los solapamientos se resuelven priorizando `ordenes_servicio`; en CCMC, `contratos` sí exige `outline_contract` porque el campo está disponible y es discriminante. Para CCMC se usan headers en inglés como `Outline Agreement`, `Plant`, `Purchasing Doc. Type`, `Item Category.1`, `Release group` y `Deletion Indicator`.

Luego aplica una exclusión común `E01` que envía a `C2_NO_MIGRA` las filas con fecha efectiva `delivery_date` con fallback a `validity_end` menor o igual a `2026-06-30`, o sin saldo pendiente por entregar, o con `deletion_flag`/indicador de borrado informado. Las filas restantes quedan en `C2`.

El código histórico en `src/contratos` y `src/ordenes_compra` se mantiene intacto; el nuevo flujo vive en `src/segmentacion`.

El flujo mantiene un cache persistente en `tmp/cache/segmentacion_mlcc` para reutilizar tanto el universo consolidado de OC como el cruce con `MLCC - Análisis Completo.xlsx`. Si se necesita reconstruir desde cero, usar `--no-cache`.

Los control points del flujo segmentado muestran en la hoja **Master** las columnas presentes en `tmp/Caracterización de Contratos.xlsx`, usando `outline_contract` como nombre canónico, más `release_group` y `deletion_flag`.

## Dependencia crítica: cruce de materiales

Las reglas actuales dependen del maestro de materiales y consumos históricos:

- `inputs/MLCC/MLCC - Análisis Completo.xlsx`
- `inputs/MLCC/MLCC - Consumos Históricos Marzo 2026.xlsx`

El pipeline cruza cada fila normalizada contra esos archivos antes de ejecutar reglas. El enriquecimiento inyecta columnas como `material_key`, `material_planning_type`, `material_in_master`, `material_no_movement_24m`, `material_is_critical`, `material_master_description`, `material_frequency_24m` y totales de consumo.

El módulo standalone `src/material_crossref` sigue disponible para generar el reporte de cobertura de materiales.

## Flujo del pipeline

```
DATOS RECIBIDOS
        ↓
INGESTA + NORMALIZACIÓN
        ↓
CRUCE MAESTRO MATERIALES + CONSUMOS
        ↓
[C1] DATOS NORMALIZADOS Y ENRIQUECIDOS
        ↓
G1 — Eliminaciones/conclusiones activas
        ↓
G2 — Rescates configurados
        ↓
G3 — Identificación y marcado PDF
        ↓                         ↓
[C3] REGISTROS CONTINÚAN   [C2_NO_MIGRA] REGISTROS ELIMINADOS/CONCLUIDOS
```

**Invariante de reconciliación:** `C1 = C3 + C2_NO_MIGRA`.

G3 se aplica sobre todo el universo post-G2, no solo sobre registros que continúan. Esto permite auditar marcas en `C2_NO_MIGRA` cuando existan exclusiones activas.

## Reglas activas: contratos / convenios

| ID | Acción | Regla PDF | Implementación |
|---|---|---|---|
| C01 | EXCLUDE | Eliminar materiales cuyo tipo de planificación corresponda a `ZZ` | Usa `material_planning_type` desde el maestro de materiales. |
| C02 | MARK | Identificar materiales duplicados dentro del mismo contrato | Marca filas donde `(purchase_document, material_key)` aparece más de una vez. |
| C03 | MARK | Detectar materiales presentes en más de un contrato vigente | Considera vigente si la fecha efectiva es posterior a `2025-12-31` o está vacía. |
| C04 | MARK | Identificar ítems sin movimientos en los últimos 24 meses | Usa `material_frequency_24m <= 0` desde el maestro. |
| C05 | MARK | Identificar diferencias entre descripción del maestro y contrato | Compara `short_text` con `material_master_description` normalizado. |
| C07 | MARK | Identificar proveedor/material repetido en contratos diferentes | Marca pares `(material_key, vendor)` presentes en más de un contrato. |
| C09 | MARK | Identificar contratos vencidos con saldo sin consumo | Fecha efectiva `<= 2025-12-31`, saldo pendiente positivo y sin movimiento 24m. |

## Reglas activas: órdenes de compra

| ID | Acción | Regla PDF | Implementación |
|---|---|---|---|
| P01 | MARK | Identificar OC con fechas de entrega vencidas provenientes de PR | Marca `delivery_date < 2026-05-25`, cantidad/valor pendiente por entregar y `purchase_requisition` no vacío. |
| P05 | MARK | Identificar OC de cargo directo con entrega vencida no entregada | Marca fecha vencida con saldo pendiente y `account_assignment_type == K`. |

## Reglas desactivadas o fuera de alcance

| ID / tema | Dominio | Motivo |
|---|---|---|
| C06 / P02 Incoterm | contratos / órdenes | Los inputs MLCC actuales no traen campo Incoterm. |
| C08 contratos próximos a vencer | contratos | El PDF deja el período a definir; 4 meses es una estimación, no un criterio cerrado. |
| C10 convenios vigentes no utilizados | contratos | No hay campo directo de utilización de convenio; requiere criterio de negocio. |
| C11 PIR estándar y Supply Option | contratos | No existe campo Supply Option y la tarea implica acción SAP posterior. |
| P03 concluir OC con más de 6 meses de atraso | órdenes | Requiere validación previa de Lundin y criterio formal para excluir materiales críticos. |
| P04 concluir OC de materiales reparados con más de 12 meses | órdenes | Requiere validación previa de Lundin y definición confiable de material reparado. |
| P06 clasificar cargo directo en Operación/Capex/reparables/garantías | órdenes | Los inputs no contienen una clasificación confiable para esas categorías. |
| Órdenes de servicio / HES | no activo | No existe dominio de entrada separado. |
| PR/SOLPED | no activo | No existe dominio de entrada PR/SOLPED; las referencias disponibles en OC se usan solo para P01. |
| Reservas | no activo | No existe dominio de entrada de reservas. |

Las reglas desactivadas quedan en los YAML con `enabled: {MLCC: false}` cuando pertenecen a un dominio existente. No se importan ni se ejecutan.

## Estructura relevante

```
lundin-cleansing-servicios/
├── pipeline.py
├── run_material_crossref.py
├── src/
│   ├── contratos/rules/              # Reglas de convenios
│   ├── ordenes_compra/rules/         # Reglas de órdenes de compra
│   ├── material_crossref/            # Cruce maestro materiales + consumos
│   ├── schema/                       # Esquema canónico y mapping MLCC
│   ├── ingestion/assembler.py
│   ├── output/control_point.py
│   └── output/stats_report.py
├── inputs/MLCC/                      # Archivos Excel de entrada (gitignored)
└── outputs/                          # Control points y reportes (gitignored)
```

## Uso

```powershell
# Contratos / convenios
venv\Scripts\python.exe pipeline.py --domain contratos --operation MLCC

# Órdenes de compra
venv\Scripts\python.exe pipeline.py --domain ordenes-compra --operation MLCC

# Verificación rápida de órdenes sin escribir workbooks pesados
venv\Scripts\python.exe pipeline.py --domain ordenes-compra --operation MLCC --skip-control-points --skip-stats-report

# Segmentación MLCC desde todas las órdenes de compra
venv\Scripts\python.exe run_segmentacion_mlcc.py --operation MLCC

# Segmentación CCMC desde todas las órdenes de compra
venv\Scripts\python.exe run_segmentacion_mlcc.py --operation CCMC

# Verificación rápida de segmentación sin escribir workbooks pesados
venv\Scripts\python.exe run_segmentacion_mlcc.py --operation MLCC --skip-control-points

# Verificación rápida de segmentación sin cruce material
venv\Scripts\python.exe run_segmentacion_mlcc.py --operation CCMC --skip-control-points --skip-material-enrichment

# Fuerza reconstrucción del universo base y del cruce, sin reutilizar cache persistente
venv\Scripts\python.exe run_segmentacion_mlcc.py --operation MLCC --no-cache

# Reporte standalone de cruce de materiales
venv\Scripts\python.exe run_material_crossref.py --domain ambos

# Verificación rápida de sintaxis
venv\Scripts\python.exe -m compileall src pipeline.py generate_diagram.py run_material_crossref.py

# Diagramas
venv\Scripts\python.exe generate_diagram.py --domain contratos --format png
venv\Scripts\python.exe generate_diagram.py --domain ordenes-compra --format png
```

Si un reporte estadístico está abierto en Excel y Windows bloquea el archivo destino, el pipeline escribe una copia con timestamp para no detener la ejecución.

Para validaciones de reglas sobre dominios grandes se pueden usar `--skip-control-points` y `--skip-stats-report`. Esos flags no cambian la lógica de reglas ni la reconciliación; solo omiten escrituras pesadas de salida.

## Outputs principales

| Ruta | Descripción |
|---|---|
| `outputs/contratos/control_points/C1_MLCC.xlsx` | Convenios normalizados y enriquecidos con materiales. |
| `outputs/contratos/control_points/C2_NO_MIGRA_MLCC.xlsx` | Convenios eliminados por reglas EXCLUDE activas. |
| `outputs/contratos/control_points/C3_MLCC.xlsx` | Convenios que continúan, con marcas PDF. |
| `outputs/ordenes_compra/control_points/C1_MLCC.xlsx` | Órdenes normalizadas y enriquecidas con materiales. |
| `outputs/ordenes_compra/control_points/C2_NO_MIGRA_MLCC.xlsx` | Órdenes concluidas por reglas EXCLUDE activas. |
| `outputs/ordenes_compra/control_points/C3_MLCC.xlsx` | Órdenes que continúan, con marcas PDF. |
| `outputs/control_points/contratos/C1_<operacion>.xlsx` | Segmento contratos pre-exclusión desde OC. |
| `outputs/control_points/contratos/C2_<operacion>.xlsx` | Segmento contratos que migra post-exclusión común. |
| `outputs/control_points/contratos/C2_NO_MIGRA_<operacion>.xlsx` | Segmento contratos excluido por fecha efectiva `<= 2026-06-30`, saldo pendiente o indicador de borrado. |
| `outputs/control_points/ordenes_servicio/C1_<operacion>.xlsx` | Segmento órdenes de servicio pre-exclusión desde OC. |
| `outputs/control_points/ordenes_servicio/C2_<operacion>.xlsx` | Segmento órdenes de servicio que migra post-exclusión común. |
| `outputs/control_points/ordenes_servicio/C2_NO_MIGRA_<operacion>.xlsx` | Segmento órdenes de servicio excluido por fecha efectiva `<= 2026-06-30`, saldo pendiente o indicador de borrado. |
| `outputs/control_points/resumen_segmentacion_<operacion>.md` | Conteos de filas, documentos únicos, outline contracts únicos y reconciliación del flujo segmentado. |
| `outputs/material_crossref/reporte_materiales_*.xlsx` | Reporte standalone de cobertura de materiales. |

Cada control point incluye hojas **Info**, **Master**, **Field Map**, **Issues**, **Stages** y **Profiling**. La hoja **Master** muestra columnas canónicas, columnas `mark_*` y columnas `material_*` inyectadas por el cruce.
