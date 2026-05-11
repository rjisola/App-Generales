import pandas as pd
import os

file_path = "03_OTROS/PROGRAMA DEPOSITO.xlsm"

def check_raw_data():
    print(f"\n--- Inspección de Columnas A, B, C, D en CALCULAR HORAS ---")
    try:
        df = pd.read_excel(file_path, sheet_name="CALCULAR HORAS", header=None, engine='openpyxl')
        for r in range(45, min(80, len(df))):
            row = df.iloc[r, :5]
            print(f"Fila {r+1}: A='{row[0]}' | B='{row[1]}' | C='{row[2]}' | D='{row[3]}' | E='{row[4]}'")
    except Exception as e:
        print(f"Error: {e}")

check_raw_data()
