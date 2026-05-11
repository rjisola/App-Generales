import pandas as pd
import os

file_path = "03_OTROS/PROGRAMA DEPOSITO.xlsm"

def check_sheet(sheet_name, start_row, cols_to_check):
    print(f"\n--- Analizando Hoja: {sheet_name} ---")
    try:
        df = pd.read_excel(file_path, sheet_name=sheet_name, header=None, engine='openpyxl')
        # Buscar la primera fila con datos en la columna D (índice 3)
        for r in range(start_row, len(df)):
            if pd.notna(df.iloc[r, 3]):
                row_data = []
                for col_idx in cols_to_check:
                    val = df.iloc[r, col_idx]
                    col_letter = chr(65 + col_idx) if col_idx < 26 else f"A{chr(65 + col_idx - 26)}"
                    row_data.append(f"{col_letter}: {val}")
                print(f"Fila {r+1}: " + " | ".join(row_data))
                if r > start_row + 10: break # Limitar salida
    except Exception as e:
        print(f"Error al leer hoja {sheet_name}: {e}")

# RECUENTO TOTAL: D(3), E(4), J(9), K(10)
check_sheet("RECUENTO TOTAL", 0, [3, 4, 9, 10])
