import os
import shutil
import sys

# Definir archivos y carpetas esenciales que NO se deben mover
ARCHIVOS_A_CONSERVAR = {
    # Launcher Principal
    '1-LAUNCHER.pyw',
    
    # Aplicaciones Activas
    '10-CALCULAR_HORAS.pyw',
    '15-PASAR_HORAS_DEPOSITO.pyw',
    'A-GENERAR_RECIBOS_CONTROL.pyw',
    'B-PROCESARSUELDOS.pyw',
    'BuscarRecibosPDF.pyw',
    'imprimir_sobres_gui.pyw',
    'Asistente_Aguinaldo_UNIFICADO.pyw',
    
    # Módulos y Librerías
    'modern_gui_components.py',
    'gui_modern.py',
    'icon_loader.py',
    'backup_manager.py',
    'data_loader.py',
    'logic_payroll.py',
    'logic_accountant.py',
    'logic_cleaning.py',
    'excel_format_writer.py',
    'receipt_font_formatter.py',
    'calculators.py',
    
    # Configuración y Scripts Utiles
    'config.json',
    'requirements.txt',
    'instalar_dependencias.bat',
    'EJECUTAR_LIMPIO.bat',
    'cleanup_files.py', # Este mismo script
    'sync_github.py',   # Script de sincronización
    'SYNC_GITHUB.bat',  # Ejecutable de sincronización
    '.gitignore',       # Archivo de configuración git
    
    # Archivos de Datos / Templates Excel Importantes
    '2-VALOR_HORAS_SUELDOS.xlsx',
    'Antiguedad.xlsx',
    'indice.xlsx',
    'fechasIngreso.csv',
    'Imprimir Sobres.xlsm',
    'LISTADO CARJOR detalle personal.xlsm',
    'Ordenes Actualizada.xlsm',
    'PROGRAMA DEPOSITO.xlsm',
}

CARPETAS_A_CONSERVAR = {
    'launcher_icons',
    'app_icons',
    'Datos',
    'Macros Excel',
    'backups',
    '.github',
    '__pycache__',
    '.git',
    '.gemini' # Por si acaso
}

DIRECTORIO_BORRAR = 'borrar'

def main():
    root_dir = os.getcwd()
    borrar_path = os.path.join(root_dir, DIRECTORIO_BORRAR)
    
    if not os.path.exists(borrar_path):
        os.makedirs(borrar_path)
        print(f"Creado directorio: {DIRECTORIO_BORRAR}")
    
    moved_count = 0
    
    for item in os.listdir(root_dir):
        # Ignorar archivo de script y directorio de destino
        if item == DIRECTORIO_BORRAR:
            continue
            
        full_path = os.path.join(root_dir, item)
        
        # Determinar si se conserva
        keep = False
        if item in ARCHIVOS_A_CONSERVAR:
            keep = True
        elif item in CARPETAS_A_CONSERVAR:
            keep = True
        elif os.path.isdir(full_path) and item in CARPETAS_A_CONSERVAR:
            keep = True
            
        if not keep:
            try:
                dest_path = os.path.join(borrar_path, item)
                print(f"Moviendo: {item} -> {DIRECTORIO_BORRAR}/")
                shutil.move(full_path, dest_path)
                moved_count += 1
            except Exception as e:
                print(f"Error moviendo {item}: {e}")
                
    print(f"\nProceso completado.")
    print(f"Archivos/Carpetas movidos a '{DIRECTORIO_BORRAR}': {moved_count}")
    print(f"Archivos conservados en raíz: {len(ARCHIVOS_A_CONSERVAR)}")
    input("\nPresione Enter para salir...")

if __name__ == '__main__':
    main()
