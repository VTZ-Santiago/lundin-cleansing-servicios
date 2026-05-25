import pandas as pd
import openpyxl
from pathlib import Path

# Archivos a analizar
files_to_check = {
    "MLCC - Análisis Completo": r"C:\Users\rojas\vantaz\lundin\vtz-github\lundin-cleansing-servicios\inputs\MLCC\MLCC - Análisis Completo.xlsx",
    "Contrato 2013-2017": r"C:\Users\rojas\vantaz\lundin\vtz-github\lundin-cleansing-servicios\inputs\MLCC\contratos\Contrato 2013-2017.XLSX",
    "OC 2010-2016": r"C:\Users\rojas\vantaz\lundin\vtz-github\lundin-cleansing-servicios\inputs\MLCC\ordenes-compra\01.01.2010- 31.12.2016.XLSX",
}

for name, filepath in files_to_check.items():
    try:
        print(f"\n{'='*80}")
        print(f"ARCHIVO: {name}")
        print(f"RUTA: {filepath}")
        print(f"{'='*80}")
        
        # Obtener nombres de hojas
        wb = openpyxl.load_workbook(filepath, read_only=True, data_only=True)
        sheets = wb.sheetnames
        print(f"HOJAS: {sheets}")
        wb.close()
        
        # Leer primera hoja
        df = pd.read_excel(filepath, sheet_name=0, nrows=5)
        print(f"\nPRIMERAS 5 FILAS - Columnas: {list(df.columns)}")
        print(f"Total columnas: {len(df.columns)}")
        print(f"\nTipos de datos:")
        print(df.dtypes)
        print(f"\nPrimeros valores:")
        print(df.head())
        
    except Exception as e:
        print(f"ERROR: {str(e)}")
