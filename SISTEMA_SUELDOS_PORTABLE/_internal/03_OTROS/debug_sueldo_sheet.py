import pandas as pd
import os

file_path = "03_OTROS/PROGRAMA DEPOSITO.xlsm"

def check_sueldo_sheet():
    print(f"\n--- Analizando Hoja: SUELDO_ALQ_GASTOS ---")
    try:
        df = pd.read_excel(file_path, sheet_name="SUELDO_ALQ_GASTOS", header=None, engine='openpyxl')
        # Leer desde la fila 9 (índice 8)
        for r in range(8, 20):
            row_data = []
            # J=9, K=10, L=11
            for col_idx in [9, 10, 11]:
                val = df.iloc[r, col_idx]
                col_letter = chr(65 + col_idx)
                row_data.append(f"{col_letter}: {val}")
            print(f"Fila {r+1}: " + " | ".join(row_data))
    except Exception as e:
        print(f"Error: {e}")

check_sueldo_sheet()
