# lundin-cleansing-servicios

Pipeline de limpieza y estandarización de datos de **contratos** y **órdenes de compra** para MLCC.

Basado en la arquitectura de `lundin-cleansing-all-rules`, adaptado al dominio de contratos.

---

## Dominios soportados

- `contratos`: usa los archivos en [inputs/MLCC/contratos](inputs/MLCC/contratos) y conserva el flujo de reglas original de contratos.
- `ordenes-compra`: usa los archivos en [inputs/MLCC/ordenes-compra](inputs/MLCC/ordenes-compra) y aplica reglas equivalentes para PO, incluyendo una marca adicional para cargos directos a centro de costo vía `Tipo de imputación`.

Los paquetes Python asociados son [src/contratos](src/contratos) y [src/ordenes_compra](src/ordenes_compra). El selector operativo es `--domain contratos|ordenes-compra`.

## Flujo del pipeline

```
DATOS RECIBIDOS (4 archivos Excel MLCC)
        ↓
INGESTA + NORMALIZACIÓN + PROFILING
        ↓
[C1] DATOS NORMALIZADOS
        ↓
G1 — Primer filtro (R01: vencimiento, R02: tipo contrato/posición)
        ↓
G2 — Migración segura (R03: vencidos con saldo pendiente positivo)
        ↓
G3 — Clasificación y marcado (R04/R05 + marcas técnicas + regla PO de imputación)
        ↓                        ↓
[C3] REGISTROS MIGRAN     [C2_NO_MIGRA] REGISTROS NO MIGRAN
```

**Invariante de reconciliación:** `C1 = C3 + C2_NO_MIGRA`

Para la segmentación por vigencia del flujo, la fecha efectiva usa `validity_end` (← `Fin período validez`) y, cuando viene vacía, cae en `delivery_date` (← `Fecha de entrega`).

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
│   ├── contratos/                  # Dominio contratos: reglas y mapping MLCC
│   ├── ordenes_compra/             # Dominio órdenes de compra: reglas y mapping MLCC
│   ├── rules/engine.py             # Motor de reglas (carga dinámica por dominio)
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
# Ejecutar pipeline completo para contratos
venv\Scripts\python.exe pipeline.py

# Contratos explícito
venv\Scripts\python.exe pipeline.py --domain contratos --operation MLCC

# Órdenes de compra
venv\Scripts\python.exe pipeline.py --domain ordenes-compra --operation MLCC

# Generar diagrama SVG para contratos
venv\Scripts\python.exe generate_diagram.py --domain contratos

# Generar diagrama PNG para órdenes de compra
venv\Scripts\python.exe generate_diagram.py --domain ordenes-compra --format png

# Verificar compilación rápida
venv\Scripts\python.exe -m compileall src pipeline.py generate_diagram.py
```

---

## Puntos de control (outputs)

| Ruta | Descripción |
|---|---|
| `outputs/contratos/control_points/C1_MLCC.xlsx` | Post-ingesta contratos |
| `outputs/contratos/control_points/C2_NO_MIGRA_MLCC.xlsx` | Contratos no migran |
| `outputs/contratos/control_points/C3_MLCC.xlsx` | Contratos migran |
| `outputs/contratos/reporte_estadistico_MLCC-contratos.xlsx` | Resumen estadístico contratos |
| `outputs/ordenes_compra/control_points/C1_MLCC.xlsx` | Post-ingesta órdenes de compra |
| `outputs/ordenes_compra/control_points/C2_NO_MIGRA_MLCC.xlsx` | Órdenes no migran |
| `outputs/ordenes_compra/control_points/C3_MLCC.xlsx` | Órdenes migran |
| `outputs/ordenes_compra/reporte_estadistico_MLCC-PO.xlsx` | Resumen estadístico órdenes de compra, con clasificación por documento y contrato marco |

Cada archivo tiene hojas: **Info**, **Master**, **Field Map**, **Issues**, **Stages**, **Profiling**. Los diagramas se generan por dominio en `outputs/<dominio>/diagrama_pipeline.*`.

El reporte estadístico de contratos conserva volumen, grupo de compras, monto, tipo de posición, vigencia y borrado. El reporte estadístico de PO reemplaza el bloque de grupo/monto por clasificación `Cl.documento compras`, conserva la vigencia por validez, agrega un bloque adicional por `Fecha de entrega`, incluye resumen de `Contrato marco` y una hoja adicional `Contrato Marco` con el conteo de PO asociadas a cada contrato marco.

---

## Reglas activas

| ID | Nombre | Dominio | Grupo | Acción |
|---|---|---|---|
| R01 | Fecha de vencimiento en 2025 o antes | contratos / ordenes-compra | G1_EXCLUSIONS | EXCLUDE |
| R02 | Tipo de posición distinto de D | contratos / ordenes-compra | G1_EXCLUSIONS | EXCLUDE |
| R03 | Valores pendientes > 0 en registros vencidos | contratos / ordenes-compra | G2_RESCUE | RESCUE |
| R04 | Valores pendientes < 0 en registros no vencidos | contratos / ordenes-compra | G3_MARKING | MARK |
| R05 | Registros que vencen en 2026 | contratos / ordenes-compra | G3_MARKING | MARK |
| R06 | Servicios con cargo directo a centro de costo (`K`) | ordenes-compra | G3_MARKING | MARK |
| M01 | Clasificación por tipo de posición | contratos / ordenes-compra | G3_MARKING | MARK |
| M02 | Clasificación por indicador de borrado | contratos / ordenes-compra | G3_MARKING | MARK |

Configuración por dominio en [src/contratos/rules/rules.yaml](src/contratos/rules/rules.yaml) y [src/ordenes_compra/rules/rules.yaml](src/ordenes_compra/rules/rules.yaml). Para agregar una nueva regla:
1. Crear el módulo en el dominio correspondiente, por ejemplo `src/ordenes_compra/rules/group3/rNN.py`
2. Añadir la entrada en el `rules.yaml` del dominio con `enabled`, `priority`, `action` y `config`

---

## Esquema canónico

El pipeline normaliza los 29 headers en español a nombres canónicos en inglés.
La fecha clave para las reglas de vigencia es `validity_end` (← `Fin período validez`), con fallback a `delivery_date` (← `Fecha de entrega`) cuando el primer campo viene vacío.

Grain key de MLCC: (`purchase_document`, `position`)

Las 81 columnas extra del archivo 2025-2026 se preservan con prefijo `_raw__`
(NaN en filas de archivos anteriores).
