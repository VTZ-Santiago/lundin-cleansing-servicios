# lundin-cleansing-servicios

Pipeline de limpieza y estandarización de datos de **contratos/servicios** para MLCC.

Basado en la arquitectura de `lundin-cleansing-all-rules`, adaptado al dominio de contratos.

---

## Flujo del pipeline

```
DATOS RECIBIDOS (4 archivos Excel MLCC)
        ↓
INGESTA + NORMALIZACIÓN + PROFILING
        ↓
[C1] CONTRATOS NORMALIZADOS
        ↓
G1 — Primer filtro (R01: vencimiento, R02: tipo contrato/posición)
        ↓
G2 — Migración segura (R03: vencidos con saldo pendiente positivo)
        ↓
G3 — Clasificación y marcado (R04/R05 + marcas técnicas)
        ↓                        ↓
[C3] CONTRATOS MIGRAN     [C2_NO_MIGRA] CONTRATOS NO MIGRAN
```

**Invariante de reconciliación:** `C1 = C3 + C2_NO_MIGRA`

---

## Estructura

```
lundin-cleansing-servicios/
├── pipeline.py             # Punto de entrada
├── generate_diagram.py     # Genera diagrama SVG
├── requirements.txt
├── src/
│   ├── config/settings.py          # Paths y configuración
│   ├── schema/canonical.py         # Esquema canónico (29 columnas + 2 auxiliares)
│   ├── schema/field_map_mlcc.py    # Mapeo de headers español → canónico
│   ├── ingestion/assembler.py      # Carga y normalización de archivos Excel
│   ├── rules/rules.yaml            # Config YAML de reglas (fuente de verdad)
│   ├── rules/engine.py             # Motor de reglas (carga dinámica)
│   ├── rules/group1/               # G1: exclusiones por vencimiento/tipo
│   ├── rules/group2/               # G2: rescate / migración segura
│   ├── rules/group3/               # G3: clasificación y marcado
│   ├── profiling/profiler.py       # Estadísticas por columna
│   ├── output/control_point.py     # Exportación de puntos de control (Excel)
│   ├── output/stats_report.py      # Reporte estadístico independiente
│   ├── diagram/flowchart.py        # Generación de diagrama SVG
│   └── lineage/                    # Trazabilidad de columnas
├── inputs/MLCC/                    # Archivos Excel de entrada (gitignored)
└── outputs/control_points/         # Resultados del pipeline
```

---

## Setup

```powershell
# Crear entorno virtual (si no existe)
python -m venv venv

# Instalar dependencias
venv\Scripts\python.exe -m pip install -r requirements.txt

# (Opcional) Para generar el diagrama como imagen renderizada:
winget install graphviz
```

---

## Uso

```powershell
# Ejecutar pipeline completo
venv\Scripts\python.exe pipeline.py

# Operación específica
venv\Scripts\python.exe pipeline.py --operation MLCC

# Generar diagrama SVG
venv\Scripts\python.exe generate_diagram.py

# Generar diagrama PNG
venv\Scripts\python.exe generate_diagram.py --format png
```

---

## Puntos de control (outputs)

| Archivo | Descripción |
|---|---|
| `C1_MLCC.xlsx` | Post-ingesta: todos los contratos normalizados |
| `C2_NO_MIGRA_MLCC.xlsx` | Post-G1/G2: contratos excluidos por primer filtro y no rescatados |
| `C3_MLCC.xlsx` | Post-G1/G2/G3: contratos que migran, con rescates y marcas aplicadas |
| `reporte_estadistico_MLCC.xlsx` | Resumen estadístico C1 con vigencia, tipo, monto y grupo de compras |

Cada archivo tiene hojas: **Info**, **Master**, **Field Map**, **Issues**, **Stages**, **Profiling**.

Los resúmenes de control incluyen fecha de vencimiento, tipo de contrato/posición, monto en USD y grupo de compras.

---

## Reglas activas

| ID | Nombre | Grupo | Acción |
|---|---|---|---|
| R01 | Fecha de vencimiento en 2025 o antes | G1_EXCLUSIONS | EXCLUDE |
| R02 | Tipo de contrato/posición distinto de D | G1_EXCLUSIONS | EXCLUDE |
| R03 | Valores pendientes > 0 en contratos vencidos | G2_RESCUE | RESCUE |
| R04 | Valores pendientes < 0 en contratos no vencidos | G3_MARKING | MARK |
| R05 | Contratos que vencen en 2026 | G3_MARKING | MARK |
| R06 | Monto del contrato | reporting.profiling_summary | Resumen |
| M01 | Clasificación por tipo de posición | G3_MARKING | MARK |
| M02 | Clasificación por indicador de borrado | G3_MARKING | MARK |

Configuración en `src/rules/rules.yaml`. Para agregar una nueva regla:
1. Crear `src/rules/groupN/rNN.py` con una función `apply(df, operation, config) -> RuleResult`
2. Añadir la entrada en `rules.yaml` con `enabled`, `priority`, `action` y `config`

---

## Esquema canónico

El pipeline normaliza los 29 headers en español a nombres canónicos en inglés.
La columna clave para la regla de exclusión es `validity_end` (← `Fin período validez`).

Grain key de MLCC: (`purchase_document`, `position`)

Las 81 columnas extra del archivo 2025-2026 se preservan con prefijo `_raw__`
(NaN en filas de archivos anteriores).
