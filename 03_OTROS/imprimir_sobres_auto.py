import win32com.client
import os
import sys

def imprimir_sobres_masivo():
    # --- CONFIGURACIÓN (AJUSTA ESTO SI ES NECESARIO) ---
    ruta_excel = r"C:\Users\rjiso\OneDrive\Escritorio\PROCESAR ARCHIVO SUELDOS\Imprimir Sobres.xlsm"
    
    # Nombres exactos de las hojas en tu Excel
    hoja_lista_nombre = "Hoja1"  # <--- CAMBIAR si tu hoja de lista se llama diferente (ej: "Hoja1", "Empleados")
    hoja_sobre_nombre = "ImpresionSobres"      # <--- CAMBIAR si tu hoja de sobre se llama diferente
    
    # Coordenadas
    columna_legajos = 1        # Columna A = 1, B = 2, etc. (Donde están los legajos en la lista)
    fila_inicio = 2            # Fila donde empieza el primer legajo
    celda_destino_sobre = "B2" # Celda en la hoja 'Sobre' donde se escribe el legajo
    # ---------------------------------------------------

    if not os.path.exists(ruta_excel):
        print(f"Error: No se encuentra el archivo en:\n{ruta_excel}")
        input("Presiona Enter para salir...")
        return

    print(f"Iniciando Excel y abriendo: {os.path.basename(ruta_excel)}...")
    
    try:
        # Iniciar Excel
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = True # Lo hacemos visible para que veas el proceso
        
        # Abrir libro
        wb = excel.Workbooks.Open(ruta_excel)
        
        # Seleccionar hojas
        try:
            ws_lista = wb.Sheets(hoja_lista_nombre)
            ws_sobre = wb.Sheets(hoja_sobre_nombre)
        except Exception:
            print(f"\nERROR: No se encontraron las hojas '{hoja_lista_nombre}' o '{hoja_sobre_nombre}'.")
            print("Las hojas disponibles son:")
            for s in wb.Sheets:
                print(f" - {s.Name}")
            print("\nPor favor, edita este script (clic derecho -> Edit) y pon los nombres correctos en la sección CONFIGURACIÓN.")
            return

        # Encontrar última fila con datos en la columna de legajos
        # -4162 es la constante xlUp de Excel
        ultima_fila = ws_lista.Cells(ws_lista.Rows.Count, columna_legajos).End(-4162).Row
        
        cantidad = ultima_fila - fila_inicio + 1
        print(f"\nSe encontraron {cantidad} legajos para imprimir (hasta fila {ultima_fila}).")
        
        confirm = input("¿Deseas comenzar la impresión? (s/n): ")
        if confirm.lower() != 's':
            print("Cancelado.")
            return

        # Bucle de impresión
        for fila in range(fila_inicio, ultima_fila + 1):
            legajo = ws_lista.Cells(fila, columna_legajos).Value
            
            if legajo:
                print(f"Imprimiendo legajo: {legajo}")
                
                # 1. Copiar legajo al sobre
                ws_sobre.Range(celda_destino_sobre).Value = legajo
                
                # 2. Imprimir
                ws_sobre.PrintOut()
        
        print("\nProceso finalizado exitosamente.")

    except Exception as e:
        print(f"\nOcurrió un error: {e}")
    
if __name__ == "__main__":
    imprimir_sobres_masivo()
    input("\nPresiona Enter para cerrar...")