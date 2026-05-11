import pandas as pd
import os

file_path = "03_OTROS/PROGRAMA DEPOSITO.xlsm"

def check_structure(sheet_name):
    print(f"\n--- Estructura de {sheet_name} (Primeras 20 filas) ---")
    try:
        df = pd.read_excel(file_path, sheet_name=sheet_name, header=None, engine='openpyxl')
        print(df.head(20).to_string())
    except Exception as e:
        print(f"Error: {e}")

check_structure("RECUENTO TOTAL")
