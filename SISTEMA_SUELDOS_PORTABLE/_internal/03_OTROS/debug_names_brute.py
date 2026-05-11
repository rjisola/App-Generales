import openpyxl
import os

file_path = "03_OTROS/PROGRAMA DEPOSITO.xlsm"

def check_names_openpyxl():
    print(f"\n--- Analizando Nombres con OPENPYXL (Fuerza Bruta) ---")
    try:
        wb = openpyxl.load_workbook(file_path, data_only=True)
        ws = wb["CALCULAR HORAS"]
        count = 0
        # Leer hasta la fila 100 para estar seguros
        for r in range(9, 101):
            name = ws.cell(row=r, column=1).value
            if name:
                print(f"Fila {r}: {name}")
                count += 1
            else:
                # Si hay una celda vacía, imprimimos que está vacía
                # print(f"Fila {r}: VACIA")
                pass
        print(f"\nTotal de nombres detectados: {count}")
        wb.close()
    except Exception as e:
        print(f"Error: {e}")

check_names_openpyxl()
