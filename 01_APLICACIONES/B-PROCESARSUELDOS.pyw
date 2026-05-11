import sys
import os
import json
import pandas as pd
import openpyxl

# Asegurar que el directorio del script esté en sys.path para imports locales
# Asegurar que el directorio del script esté en sys.path para imports locales
script_dir = os.path.dirname(os.path.abspath(__file__))
if script_dir not in sys.path:
    sys.path.insert(0, script_dir)

# Agregar 03_OTROS
root_dir = os.path.dirname(script_dir)
others_dir = os.path.join(root_dir, "03_OTROS")
if others_dir not in sys.path:
    sys.path.insert(0, others_dir)

from data_loader import load_structured_data, load_rate_config, safe_openpyxl_load
from logic_payroll import process_payroll_for_employee
from logic_accountant import process_accountant_summary_for_employee
from excel_format_writer import write_payroll_to_excel, verify_output_file
from receipt_font_formatter import apply_font_to_receipts
from icon_loader import set_window_icon


def main_console():
    """
    Función principal que orquesta todo el proceso en modo consola.
    """
    # 1. Cargar Configuración
    try:
        # Rutas actualizadas tras reorganización
        parent_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        config_path = os.path.join(parent_dir, '03_OTROS', 'config.json')
        
        with open(config_path, 'r', encoding='utf-8') as f:
            config = json.load(f)
        print(f"Archivo de configuración '{os.path.basename(config_path)}' cargado exitosamente.")
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

    # Si no se pasó --file, usar la ruta de config.json como prioridad
    if not file_path:
        config_path = config.get('file_paths', {}).get('input_excel')
        if config_path and os.path.exists(config_path):
            file_path = os.path.abspath(config_path)
            print(f"INFO: Usando archivo desde config.json: {file_path}")
        else:
            # 1. Intentar localmente (en 01_APLICACIONES)
            for file in os.listdir('.'):
                if file.endswith('.xlsm'):
                    file_path = os.path.abspath(file)
                    print(f"INFO: Archivo .xlsm encontrado localmente: {file}")
                    break
            
            # 2. Intentar en 02_CARPETAS/Datos
            if not file_path:
                datos_dir = os.path.join(parent_dir, '02_CARPETAS', 'Datos')
                if os.path.exists(datos_dir):
                    for file in os.listdir(datos_dir):
                        if file.endswith('.xlsm') and "DEPOSITO" in file.upper():
                            file_path = os.path.join(datos_dir, file)
                            print(f"INFO: Archivo .xlsm encontrado en Datos: {file}")
                            break
    
    # Generar ruta de salida (Forzamos al Escritorio para facilidad del usuario)
    desktop_path = os.path.join(os.environ['USERPROFILE'], 'OneDrive', 'Escritorio')
    if not os.path.exists(desktop_path): # Fallback si no hay OneDrive
        desktop_path = os.path.join(os.environ['USERPROFILE'], 'Desktop')
    
    output_file = os.path.join(desktop_path, "PROGRAMA DEPOSITO_MODIFICADO.xlsm")
    
    # a) Cargar la hoja principal de empleados con estructura de días y feriados
    header_row = 8 
    data_start_row = 9
    holiday_row = 7 # Fila 7 en Excel para marcadores de feriados
    sheet_name = 'CALCULAR HORAS' 
    try:
        employee_df, day_definitions = load_structured_data(file_path, config, data_start_row, header_row, holiday_row)
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

    # 3. Abrir el archivo Excel para procesamiento masivo
    print("\nAbriendo archivo Excel para procesamiento...")
    try:
        wb_cache = safe_openpyxl_load(file_path, data_only=True, keep_vba=True)
        wb_styles = safe_openpyxl_load(file_path, data_only=False, keep_vba=True)
        # La apertura se maneja de forma segura con copias temporales en los cargadores
    except Exception as e:
        print(f"ERROR CRÍTICO al abrir el archivo: {e}")
        return

    # 4. Procesar los datos para cada empleado de forma optimizada
    payroll_results = []
    accountant_results = []
    
    try:
        # Optimización: Convertimos todo el DataFrame a una lista de diccionarios de una vez.
        # Es considerablemente más rápido que usar .iterrows() fila por fila en Pandas.
        employees_data = employee_df.to_dict(orient='records')
        
        for i, employee_data in enumerate(employees_data):
            # Obtener fila real en Excel desde la columna preservada o calcularla
            raw_row_idx = employee_data.get('Excel_Row_Index')
            if pd.isna(raw_row_idx):
                current_excel_row = data_start_row + i
            else:
                try:
                    current_excel_row = int(raw_row_idx)
                except (ValueError, TypeError):
                    current_excel_row = data_start_row + i
            
            # Validación rápida para saltar filas vacías sin instanciar lógicas pesadas
            nombre = employee_data.get('NOMBRE Y APELLIDO')
            if pd.isna(nombre) or str(nombre).strip() == '':
                continue
            
            try:
                # --- Llamadas a la lógica de negocio ---
                
                # a) Procesar la nómina
                # Se pasa wb_styles + row_idx para que AMARILLO pueda leer colores de celda
                # y aplicar multiplicadores diferenciados por proyecto (Quilmes ×1.2 / Papelera ×1.344)
                payroll_result = process_payroll_for_employee(
                    employee_data, config, day_definitions, rate_config,
                    wb_styles=wb_styles, row_idx=current_excel_row
                )
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
            except Exception as e:
                print(f"ERROR procesando empleado '{nombre}' en fila {current_excel_row}: {e}")
                import traceback
                traceback.print_exc()
                continue
            
        print(f"Procesados {len(payroll_results)} empleados")
        
    except Exception as e:
        print(f"ERROR DURANTE EL PROCESAMIENTO DE NÓMINA: {e}")
        return
    
    finally:
        # Gestión de Memoria: Aseguramos que los workbooks se cierren siempre, 
        # liberando recursos RAM fundamentales tras iterar miles de filas.
        wb_cache.close()
        wb_styles.close()

    # 5. Escribir los resultados al archivo Excel preservando formato
    input_file = file_path 
    
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

        # 7. AVISO FINAL AL USUARIO (Pop-up)
        try:
            import tkinter as tk
            from tkinter import messagebox
            root = tk.Tk()
            root.withdraw() # Ocultar ventana principal de tk
            root.attributes("-topmost", True) # Asegurar que aparezca arriba de todo
            messagebox.showinfo("Proceso Finalizado", f"El archivo se generó con éxito en:\n\n{output_file}")
            root.destroy()
        except:
            pass # Si falla el GUI por alguna razón, no interrumpir el cierre normal
    
    print("\n" + "="*80)
    print("PROCESO FINALIZADO")
    print("="*80)
    print(f"Se procesaron exitosamente {len(payroll_results)} empleados.")
    print(f"\nARCHIVO GENERADO:")
    print(f"  - {output_file}")


def main_gui():
    """
    Lanza la interfaz gráfica de usuario.
    """
    import customtkinter as ctk
    from gui_modern import PayrollModernGUI
    
    root = ctk.CTk()
    set_window_icon(root, 'payroll')
    app = PayrollModernGUI(root)
    root.mainloop()


if __name__ == '__main__':
    # Si se pasa el argumento --console, ejecutar en modo consola
    # De lo contrario, lanzar la GUI
    if '--console' in sys.argv:
        print("Ejecutando en modo CONSOLA")
        print("="*80)
        main_console()
    else:
        print("Lanzando interfaz gráfica...")
        main_gui()