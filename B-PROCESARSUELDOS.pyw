import sys
import os
import json
import pandas as pd

# Asegurar que el directorio del script esté en sys.path para imports locales
script_dir = os.path.dirname(os.path.abspath(__file__))
if script_dir not in sys.path:
    sys.path.insert(0, script_dir)

from data_loader import load_structured_data, load_rate_config
from logic_payroll import process_payroll_for_employee
from logic_accountant import process_accountant_summary_for_employee
from excel_format_writer import write_payroll_to_excel, verify_output_file
from receipt_font_formatter import apply_font_to_receipts


def main_console():
    """
    Función principal que orquesta todo el proceso en modo consola.
    """
    # 1. Cargar Configuración
    try:
        config_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'config.json')
        with open(config_path, 'r', encoding='utf-8') as f:
            config = json.load(f)
        print("Archivo de configuración 'config.json' cargado exitosamente.")
    except FileNotFoundError:
        print("Error Crítico: No se encontró el archivo 'config.json'. Asegúrese de que exista en el mismo directorio.")
        return
    except json.JSONDecodeError:
        print("Error Crítico: El archivo 'config.json' tiene un formato inválido.")
        return

    # 2. Cargar Datos y Tarifas del Archivo Excel
    # Verificar si se pasó el argumento --file y --font
    file_path = None
    font_name = "Arial" # Default font
    
    if '--font' in sys.argv:
        try:
            font_index = sys.argv.index('--font') + 1
            if font_index < len(sys.argv):
                font_name = sys.argv[font_index]
                print(f"INFO: Usando fuente: {font_name}")
        except ValueError:
            pass

    if '--file' in sys.argv:
        try:
            file_index = sys.argv.index('--file') + 1
            if file_index < len(sys.argv):
                provided_path = sys.argv[file_index]
                if os.path.exists(provided_path):
                    file_path = os.path.abspath(provided_path)
                    print(f"INFO: Usando archivo proporcionado: {file_path}")
                else:
                    print(f"ADVERTENCIA: El archivo proporcionado '{provided_path}' no existe.")
        except ValueError:
            pass

    # Si no se pasó --file, buscar archivo .xlsm en el directorio actual primero
    if not file_path:
        for file in os.listdir('.'):
            if file.endswith('.xlsm'):
                file_path = os.path.abspath(file)
                print(f"INFO: Archivo .xlsm encontrado localmente: {file}")
                break

    # Si no se encuentra, usar la ruta de config.json
    if not file_path:
        print("INFO: No se encontró archivo .xlsm local, usando la ruta de config.json.")
        file_path = config['file_paths']['input_excel']
    
    # Generar ruta de salida en la misma carpeta que el archivo de entrada
    input_dir = os.path.dirname(file_path)
    input_filename = os.path.basename(file_path)
    input_name, input_ext = os.path.splitext(input_filename)
    output_file = os.path.join(input_dir, f"{input_name}_MODIFICADO{input_ext}")
    
    # a) Cargar la hoja principal de empleados con estructura de días y feriados
    header_row = 8 
    data_start_row = 9
    holiday_row = 7 # Fila 7 en Excel para marcadores de feriados
    sheet_name = 'CALCULAR HORAS' 
    try:
        employee_df, day_definitions = load_structured_data(file_path, sheet_name, data_start_row, header_row, holiday_row)
    except PermissionError as e:
        print(f"ERROR CRÍTICO: {e}")
        return
    
    if employee_df is None:
        print("Finalizando el programa debido a un error al cargar los datos.")
        return

    # b) Cargar las tarifas adicionales
    rate_config = load_rate_config(file_path)
    if rate_config is None:
        print("Continuando sin configuración de tarifas adicionales.")
        rate_config = {}


    # 3. Abrir el archivo Excel para procesamiento
    # Necesitamos dos cargas:
    # wb_cache: data_only=True para leer valores de fórmulas si los hubiera (aunque aquí son inputs manuales, es más seguro)
    # wb_styles: data_only=False para leer COLORES de celdas (CRÍTICO para el contador)
    print("\nAbriendo archivo Excel para procesamiento...")
    import openpyxl
    try:
        wb_cache = openpyxl.load_workbook(file_path, data_only=True, keep_vba=True)
        wb_styles = openpyxl.load_workbook(file_path, data_only=False, keep_vba=True)
        print("✓ Archivo Excel cargado (cache y estilos)")
    except PermissionError:
        print(f"ERROR CRÍTICO: El archivo '{os.path.basename(file_path)}' está abierto por otro programa.")
        print("Por favor, cierre el archivo y vuelva a intentarlo.")
        return
    except Exception as e:
        print(f"ERROR CRÍTICO al abrir el archivo: {e}")
        return

    # 4. Procesar los datos para cada empleado
    payroll_results = []
    accountant_results = []
    
    # Iteramos sobre cada fila del DataFrame
    # Usamos enumerate para tener un índice secuencial confiable (0, 1, 2...)
    for i, (index, employee_row) in enumerate(employee_df.iterrows()):
        # Convertimos la fila a un diccionario para un manejo más fácil
        employee_data = employee_row.to_dict()
        
        # Obtener fila real en Excel desde la columna preservada
        current_excel_row = int(employee_data.get('Excel_Row_Index', data_start_row + i))
        
        # Saltamos filas que no tengan un nombre de empleado
        if pd.isna(employee_data.get('NOMBRE Y APELLIDO')) or str(employee_data.get('NOMBRE Y APELLIDO')).strip() == '':
            continue
        
        # --- Llamadas a la lógica de negocio ---
        
        # a) Procesar la nómina (pasando wb_cache para optimización)
        payroll_result = process_payroll_for_employee(employee_data, config, rate_config, day_definitions, file_path, wb_cache)
        payroll_result['Excel_Row_Index'] = current_excel_row # Añadir fila para el writer
        payroll_results.append(payroll_result)
        
        # b) Procesar el resumen del contador
        # Pasamos wb_styles y el número de fila para poder leer colores de celdas
        accountant_result = process_accountant_summary_for_employee(
            employee_data, config, day_definitions, rate_config, 
            wb_styles=wb_styles, row_idx=current_excel_row
        )
        accountant_result['Excel_Row_Index'] = current_excel_row # Añadir fila para el writer
        accountant_results.append(accountant_result)
    
    # Cerrar workbooks
    wb_cache.close()
    wb_styles.close()
    print(f"✓ Procesados {len(payroll_results)} empleados")

    # 5. Escribir los resultados al archivo Excel preservando formato
    # CORRECCIÓN: Usar el mismo archivo que se procesó como base, no el de config.json
    input_file = file_path 
    # output_file ya fue generado dinámicamente arriba
    
    print("\n" + "="*80)
    print("ESCRIBIENDO RESULTADOS AL ARCHIVO EXCEL")
    print("="*80)
    
    success = write_payroll_to_excel(
        input_file=input_file,
        output_file=output_file,
        payroll_results=payroll_results,
        accountant_results=accountant_results
    )
    
    if success:
        verify_output_file(output_file)
        
        # 6. Aplicar fuente a los recibos
        print(f"\nAplicando fuente '{font_name}' a los recibos...")
        apply_font_to_receipts(output_file, font_name)
    
    print("\n" + "="*80)
    print("PROCESO FINALIZADO")
    print("="*80)
    print(f"Se procesaron {len(payroll_results)} empleados.")
    print(f"\n📄 ARCHIVO GENERADO:")
    print(f"  ✓ {output_file}")


def main_gui():
    """
    Lanza la interfaz gráfica de usuario.
    """
    import tkinter as tk
    from gui_modern import PayrollModernGUI
    
    root = tk.Tk()
    app = PayrollModernGUI(root)
    root.mainloop()


if __name__ == '__main__':
    # Si se pasa el argumento --console, ejecutar en modo consola
    # De lo contrario, lanzar la GUI
    if '--console' in sys.argv:
        print("🖥️  Ejecutando en modo CONSOLA")
        print("="*80)
        main_console()
    else:
        print("🎨 Lanzando interfaz gráfica...")
        main_gui()
