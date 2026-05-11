import openpyxl
import os

file_path = "03_OTROS/PROGRAMA DEPOSITO.xlsm"

def inspect_columns():
    print(f"\n--- Inspección de Columnas en RECUENTO TOTAL (Archivo Original) ---")
    try:
        wb = openpyxl.load_workbook(file_path, data_only=True)
        ws = wb["RECUENTO TOTAL"]
        
        # Leer las primeras 10 filas de datos (empezando desde la 2)
        for r in range(1, 11):
            row_values = []
            for c in range(1, 16): # Columnas A a O
                val = ws.cell(row=r, column=c).value
                col_letter = chr(64 + c) if c <= 26 else "A" + chr(64 + c - 26)
                row_values.append(f"{col_letter}: {val}")
            print(f"Fila {r}: " + " | ".join(row_values))
        wb.close()
    except Exception as e:
        print(f"Error: {e}")

inspect_columns()
