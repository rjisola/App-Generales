import pandas as pd
import json
import traceback
from data_loader import load_structured_data, load_rate_config
from logic_payroll import process_payroll_for_employee

def debug_payroll_values():
    config_path = r"c:\Users\rjiso\OneDrive\Escritorio\PROCESAR ARCHIVO SUELDOS\03_OTROS\config.json"
    try:
        with open(config_path, 'r') as f:
            config = json.load(f)
        
        file_path = config['file_paths']['input_excel']
        df_main, days = load_structured_data(file_path, config)
        rates = load_rate_config(file_path)
        
        print(f"Tarifas cargadas: {rates['job_title_rates']}")
        
        for _, row in df_main.iterrows():
            nom = str(row.get('NOMBRE Y APELLIDO', '')).upper()
            if "ROJAS" in nom:
                # Verificar qué categoría tiene asignada el empleado
                cat_emp = str(row.get('CATEGORÍA', '')).upper().strip()
                print(f"\nEmpleado: {nom}")
                print(f"Categoría leída del Excel: '{cat_emp}'")
                
                res = process_payroll_for_employee(row.to_dict(), config, rates, days, file_path=file_path)
                
                print(f"Resultado Logic:")
                print(f"  - Horas 50%: {res['Horas al 50%']}")
                print(f"  - Total Extras: {res['Total Extras']}")
                
    except Exception:
        traceback.print_exc()

if __name__ == "__main__":
    debug_payroll_values()
