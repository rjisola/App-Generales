import pandas as pd
import json
import os
import sys
import traceback
from data_loader import load_structured_data, load_rate_config
from logic_payroll import process_payroll_for_employee

def run_mass_audit():
    config_path = r"c:\Users\rjiso\OneDrive\Escritorio\PROCESAR ARCHIVO SUELDOS\03_OTROS\config.json"
    
    try:
        # 1. Cargar Configuración y Datos de Entrada
        with open(config_path, 'r') as f:
            config = json.load(f)
        
        file_path = r"C:\Users\rjiso\OneDrive\Escritorio\PROGRAMA DEPOSITO.xlsm"
        print(f"Cargando datos desde: {file_path}...")
        
        df_employees, day_definitions = load_structured_data(file_path, config)
        rates = load_rate_config(file_path)
        
        # 2. Cargar Totales de Referencia (RECUENTO TOTAL)
        # En PROGRAMA DEPOSITO.xlsm (Original), el nombre está en Col 4 (index 3) 
        # y el TOTAL parece estar en Col 8 (index 7).
        df_ref = pd.read_excel(file_path, sheet_name='RECUENTO TOTAL', header=None)
        
        # Crear un mapa de referencia: Nombre -> Total Excel
        ref_map = {}
        for i in range(len(df_ref)):
            name_raw = str(df_ref.iloc[i, 3]).upper().strip()
            total_ref = df_ref.iloc[i, 7] # Col 8 (H) en el archivo original
            if pd.notna(total_ref) and isinstance(total_ref, (int, float)) and name_raw != 'NOMBRE Y APELLIDO':
                ref_map[name_raw] = float(total_ref)

        print(f"Referencias encontradas en Excel: {len(ref_map)}")
        
        # 3. Procesar Auditoría
        results = []
        matches = 0
        mismatches = 0
        
        print("\nPROCESANDO AUDITORÍA...")
        print(f"{'NOMBRE':<30} | {'CAT':<10} | {'PYTHON':<12} | {'EXCEL':<12} | {'DIF'}")
        print("-" * 90)
        
        for _, row in df_employees.iterrows():
            name = str(row.get('NOMBRE Y APELLIDO', '')).upper().strip()
            color = str(row.get('Cat_Color_Name', 'GRIS')).upper()
            
            # Calcular con Python (Misma lógica que B-PROCESARSUELDOS.pyw)
            res_py = process_payroll_for_employee(
                row.to_dict(), config, day_definitions, rates
            )
            # Para paridad con RECUENTO TOTAL, usamos el Total Quincena
            total_py = res_py['Total Quincena']
            
            # Buscar Referencia en RECUENTO TOTAL
            total_ex = ref_map.get(name, 0.0)
            
            diff = abs(total_py - total_ex)
            # Tolerancia de 2 pesos por redondeos
            status = "OK" if diff < 2.0 else "MISMATCH"
            
            if status == "OK": matches += 1
            else: mismatches += 1
            
            print(f"{name[:30]:<30} | {color:<10} | {total_py:12.2f} | {total_ex:12.2f} | {total_py - total_ex:10.2f} [{status}]")
            
            results.append({
                'Nombre': name,
                'Color': color,
                'Python': total_py,
                'Excel': total_ex,
                'Diff': total_py - total_ex,
                'Status': status
            })

        print("-" * 80)
        print(f"RESULTADO FINAL:")
        print(f"  - Coincidencias: {matches}")
        print(f"  - Discrepancias: {mismatches}")
        print(f"  - Total Procesados: {len(df_employees)}")
        
        # Guardar resultados para análisis si es necesario
        with open('audit_results_final.json', 'w') as f:
            json.dump(results, f, indent=2)

    except Exception:
        traceback.print_exc()

if __name__ == "__main__":
    run_mass_audit()
