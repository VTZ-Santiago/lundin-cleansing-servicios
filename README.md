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
G1 — Exclusiones (R01: vencidos ≤ 2025-12-31)
        ↓                        ↓
[C2] CONTRATOS VIGENTES   [C2_NO_MIGRA] EXCLUIDOS
```

**Invariante de reconciliación:** `C1 = C2 + C2_NO_MIGRA`

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
│   ├── rules/group1/r01.py         # R01: contratos vencidos
│   ├── profiling/profiler.py       # Estadísticas por columna
│   ├── output/control_point.py     # Exportación de puntos de control (Excel)
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
| `C2_MLCC.xlsx` | Post-G1: contratos vigentes (migran) |
| `C2_NO_MIGRA_MLCC.xlsx` | Post-G1: contratos excluidos (vencidos ≤ 2025) |

Cada archivo tiene hojas: **Info**, **Master**, **Field Map**, **Issues**, **Stages**, **Profiling**.

---

## Reglas activas

| ID | Nombre | Grupo | Acción |
|---|---|---|---|
| R01 | Contrato vencido en 2025 o antes | G1_EXCLUSIONS | EXCLUDE |

Configuración en `src/rules/rules.yaml`. Para agregar una nueva regla:
1. Crear `src/rules/group1/rNN.py` con una función `apply(df, operation, config) -> RuleResult`
2. Añadir la entrada en `rules.yaml` con `enabled`, `priority`, `action` y `config`

---

## Esquema canónico

El pipeline normaliza los 29 headers en español a nombres canónicos en inglés.
La columna clave para la regla de exclusión es `validity_end` (← `Fin período validez`).

Grain key de MLCC: (`purchase_document`, `position`)

Las 81 columnas extra del archivo 2025-2026 se preservan con prefijo `_raw__`
(NaN en filas de archivos anteriores).
