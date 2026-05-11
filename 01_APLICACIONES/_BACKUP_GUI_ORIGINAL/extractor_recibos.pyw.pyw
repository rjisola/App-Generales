import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pdfplumber
import pandas as pd
import re
import os

def seleccionar_pdf():
    archivo = filedialog.askopenfilename(
        title="Seleccionar recibos en PDF",
        filetypes=[("Archivos PDF", "*.pdf")]
    )
    if archivo:
        entrada_pdf.delete(0, tk.END)
        entrada_pdf.insert(0, archivo)

def procesar_y_exportar():
    ruta_pdf = entrada_pdf.get()
    concepto_buscado = entrada_concepto.get().strip()
    filtro_seleccionado = combo_filtro.get()

    if not ruta_pdf or not concepto_buscado:
        messagebox.showwarning("Faltan datos", "Por favor, selecciona un PDF e ingresa el concepto a buscar.")
        return

    if not os.path.exists(ruta_pdf):
        messagebox.showerror("Error", "El archivo PDF seleccionado no existe.")
        return

    try:
        datos_extraidos = []
        
        # Abrir el PDF
        with pdfplumber.open(ruta_pdf) as pdf:
            for i, pagina in enumerate(pdf.pages):
                texto = pagina.extract_text()
                if not texto:
                    continue
                
                texto_minuscula = texto.lower()

                # --- LÓGICA DE FILTRADO DE HOJAS ---
                if filtro_seleccionado == "Solo hojas con 'Firma del empleado'":
                    if "firma del empleado" not in texto_minuscula:
                        continue # Salta a la siguiente página si no encuentra la frase
                elif filtro_seleccionado == "Solo hojas con 'Firma de empleador'":
                    # Buscamos ambas variantes por si hay diferencias de redacción
                    if "firma de empleador" not in texto_minuscula and "firma del empleador" not in texto_minuscula:
                        continue

                # 1. Buscar Legajo y Nombre
                match_empleado = re.search(r'(\d{2,4})\s+([A-Z\s,]+?)\s+(?:DOCUMENTO|CUIL|LIQUIDACION)', texto)
                
                if match_empleado:
                    legajo = match_empleado.group(1)
                    nombre = match_empleado.group(2).strip()
                else:
                    legajo = f"No encontrado (Pág {i+1})"
                    nombre = "No encontrado"

                # 2. Buscar el concepto solicitado y su importe
                importe = "0,00"
                lineas = texto.split('\n')
                
                for linea in lineas:
                    if concepto_buscado.lower() in linea.lower():
                        # Busca el formato de número de importe
                        montos = re.findall(r'\d{1,3}(?:\.\d{3})*(?:,\d{2})', linea)
                        if montos:
                            importe = montos[-1]
                        break

                if legajo and "No encontrado" not in legajo:
                    datos_extraidos.append({
                        "Legajo": legajo,
                        "Apellido y Nombre": nombre,
                        concepto_buscado: importe
                    })

        # Generar el Excel si hay datos
        if datos_extraidos:
            df = pd.DataFrame(datos_extraidos)
            
            ruta_guardado = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Archivos de Excel", "*.xlsx")],
                title="Guardar archivo Excel como..."
            )
            
            if ruta_guardado:
                df.to_excel(ruta_guardado, index=False)
                messagebox.showinfo("Éxito", f"¡Extracción completada!\nSe procesaron {len(datos_extraidos)} registros.\nArchivo guardado en:\n{ruta_guardado}")
        else:
            messagebox.showinfo("Sin resultados", "No se encontraron datos coincidentes aplicando los filtros seleccionados.")

    except Exception as e:
        messagebox.showerror("Error inesperado", f"Ocurrió un problema durante la extracción:\n{e}")

# --- Configuración de la Ventana (Interfaz Gráfica) ---
ventana = tk.Tk()
ventana.title("Extractor de Recibos a Excel")
ventana.geometry("500x320") # Ventana un poco más alta para que entre la nueva opción
ventana.configure(padx=20, pady=20)

# 1. Etiqueta y campo para PDF
tk.Label(ventana, text="1. Archivo PDF de Recibos:").pack(anchor="w")
frame_pdf = tk.Frame(ventana)
frame_pdf.pack(fill="x", pady=5)
entrada_pdf = tk.Entry(frame_pdf, width=50)
entrada_pdf.pack(side="left", fill="x", expand=True, padx=(0, 10))
btn_buscar = tk.Button(frame_pdf, text="Buscar", command=seleccionar_pdf)
btn_buscar.pack(side="right")

# 2. Etiqueta y campo para Filtro de Hojas
tk.Label(ventana, text="2. ¿Qué hojas procesar?").pack(anchor="w", pady=(15, 0))
opciones_filtro = [
    "Todas las hojas",
    "Solo hojas con 'Firma del empleado'",
    "Solo hojas con 'Firma de empleador'"
]
combo_filtro = ttk.Combobox(ventana, values=opciones_filtro, state="readonly")
combo_filtro.current(0) # Selecciona la primera opción por defecto
combo_filtro.pack(fill="x", pady=5)

# 3. Etiqueta y campo para Concepto
tk.Label(ventana, text="3. Nombre del Concepto a buscar (Ej: Ajuste Ac):").pack(anchor="w", pady=(15, 0))
entrada_concepto = tk.Entry(ventana, width=55)
entrada_concepto.pack(fill="x", pady=5)

# 4. Botón Ejecutar
btn_ejecutar = tk.Button(ventana, text="Extraer y Generar Excel", bg="#4CAF50", fg="white", font=("Arial", 10, "bold"), command=procesar_y_exportar)
btn_ejecutar.pack(pady=20, fill="x")

# Iniciar la aplicación
ventana.mainloop()