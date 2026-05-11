import pandas as pd
import os

file_path = "03_OTROS/PROGRAMA DEPOSITO.xlsm"

def check_names():
    print(f"\n--- Analizando Nombres en CALCULAR HORAS ---")
    try:
        # Cargamos sin encabezado para ver las filas reales
        df = pd.read_excel(file_path, sheet_name="CALCULAR HORAS", header=None, engine='openpyxl')
        print(f"Total de filas leídas: {len(df)}")
        # Listar desde fila 9 (índice 8)
        count = 0
        for i in range(8, len(df)):
            name = df.iloc[i, 0] # Columna A
            print(f"Fila {i+1}: {name}")
            if pd.notna(name) and str(name).strip() != "":
                count += 1
        print(f"\nTotal de nombres no vacíos detectados: {count}")
    except Exception as e:
        print(f"Error: {e}")

check_names()
