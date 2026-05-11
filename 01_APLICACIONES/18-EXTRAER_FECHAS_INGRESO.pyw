# -*- coding: utf-8 -*-
import os
import sys
import threading
import csv
import re
import unicodedata
from datetime import datetime
import tkinter as tk
from tkinter import filedialog, messagebox
import fitz  # PyMuPDF
import pandas as pd
import openpyxl
import customtkinter as ctk
from PIL import Image

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

import modern_gui_components as mgc

def load_ctk_icon(icon_name, size=(64, 64)):
    icon_path = os.path.join(script_dir, "..", "02_CARPETAS", "app_icons", f"{icon_name}.png")
    if os.path.exists(icon_path):
        try:
            img = Image.open(icon_path)
            return ctk.CTkImage(light_image=img, dark_image=img, size=size)
        except:
            return None
    return None

class AntiguedadPDFApp:
    def __init__(self, root):
        self.root = root
        self.root.title("📅 Antigüedad desde PDF")
        self.root.geometry("900x700")
        self.root.resizable(True, True)
        self.root.configure(fg_color=mgc.COLORS['bg_primary'])
        
        mgc.center_window(self.root, 900, 700)
        
        # Iconos
        self.icon_calendar = load_ctk_icon('calendar', (64, 64))
        self.icon_check = load_ctk_icon('check', (24, 24))
        self.icon_excel = load_ctk_icon('excel_icon', (24, 24))
        
        self.pdf_path = tk.StringVar()
        self.index_path = tk.StringVar()
        self.setup_ui()

    def setup_ui(self):
        self.status_frame, self.status_var = mgc.create_status_bar(self.root, "Listo")
        main_container = mgc.create_main_container(self.root)
        mgc.create_header(main_container, "Antigüedad (desde PDF)", 
                         "Extrae datos de recibos PDF y calcula años de antigüedad automáticamente",
                         icon_image=self.icon_calendar)
        
        # Card 1: Selección de PDF
        card_outer, card_inner = mgc.create_card(main_container, "1. Seleccionar PDF de Recibos", padding=20)
        card_outer.pack(fill=tk.X, pady=(0, 10))
        mgc.create_file_selector(card_inner, "Archivo PDF:", self.pdf_path, 
                                self.seleccionar_pdf, icon="📄").pack(fill=tk.X)
        
        # Card 2: Selección de Índice (Opcional)
        card_idx_outer, card_idx_inner = mgc.create_card(main_container, "2. Seleccionar Índice para Ordenar (Opcional)", padding=20)
        card_idx_outer.pack(fill=tk.X, pady=(0, 10))
        mgc.create_file_selector(card_idx_inner, "Archivo Excel:", self.index_path, 
                                self.seleccionar_index, icon="📊").pack(fill=tk.X)
        
        # Card de acción
        card3_outer, card3_inner = mgc.create_card(main_container, padding=15)
        card3_outer.pack(fill=tk.X, pady=(0, 15))
        button_container = ctk.CTkFrame(card3_inner, fg_color="transparent")
        button_container.pack()
        self.btn_procesar = mgc.create_large_button(button_container, "GENERAR EXCEL DE ANTIGÜEDAD",
                                                     self.procesar, color='green',
                                                     icon_image=self.icon_check)
        self.btn_procesar.pack()
        
        # Log de actividad
        card_log_outer, card_log_inner = mgc.create_card(main_container, "Registro de Actividad", padding=10)
        card_log_outer.pack(fill=tk.BOTH, expand=True)
        self.log_text = tk.Text(card_log_inner, height=10, font=('Consolas', 9), 
                                 bg='#0d1117', fg='#c9d1d9', state='disabled', relief=tk.FLAT)
        self.log_text.pack(fill=tk.BOTH, expand=True)

    def log(self, msg):
        self.log_text.configure(state='normal')
        self.log_text.insert(tk.END, f"[{datetime.now().strftime('%H:%M:%S')}] {msg}\n")
        self.log_text.see(tk.END)
        self.log_text.configure(state='disabled')
        self.root.update_idletasks()

    def seleccionar_pdf(self):
        path = filedialog.askopenfilename(title="Seleccionar PDF", filetypes=[("Archivos PDF", "*.pdf")])
        if path:
            self.pdf_path.set(path)
            self.log(f"PDF seleccionado: {os.path.basename(path)}")

    def seleccionar_index(self):
        initial_dir = r"C:\Users\rjiso\OneDrive\Escritorio\carjorExcelJS\API\public\ACOMODAR_PDF"
        if not os.path.exists(initial_dir):
            initial_dir = os.path.expanduser("~/Desktop")
            
        path = filedialog.askopenfilename(title="Seleccionar Índice Excel", 
                                          initialdir=initial_dir,
                                          filetypes=[("Archivos Excel", "*.xlsx *.xls")])
        if path:
            self.index_path.set(path)
            self.log(f"Índice seleccionado: {os.path.basename(path)}")

    def procesar(self):
        pdf = self.pdf_path.get()
        if not pdf:
            messagebox.showwarning("Aviso", "Por favor seleccione un archivo PDF.")
            return
        self.btn_procesar.configure(state='disabled')
        threading.Thread(target=self.run_logic, args=(pdf,), daemon=True).start()

    def run_logic(self, pdf_path):
        try:
            self.log("Iniciando extracción de texto...")
            doc = fitz.open(pdf_path)
            full_text = ""
            for page in doc:
                full_text += page.get_text()
            doc.close()
            
            self.log("Analizando estructura de recibos...")
            empleados_raw = full_text.split("FIRMA DEL EMPLEADO")
            
            datos_empleados = []
            vistos = set()
            date_pattern = re.compile(r'\d{2}/\d{2}/\d{2,4}')
            
            for i, bloque in enumerate(empleados_raw):
                if not bloque.strip(): continue
                lineas = [l.strip() for l in bloque.strip().split('\n') if l.strip()]
                
                try:
                    legajo = "N/A"
                    nombre = "N/A"
                    fecha_ingreso_raw = "N/A"
                    
                    for j, l in enumerate(lineas):
                        text_upper = l.upper()
                        if text_upper == "LEGAJO" and j > 0:
                            legajo = lineas[j-1]
                        elif text_upper == "APELLIDO Y NOMBRE" and j > 0:
                            nombre = lineas[j-1]
                        elif text_upper == "FECHA DE INGRESO" and j > 0:
                            fecha_ingreso_raw = lineas[j-1]
                    
                    legajo_clean = "".join(filter(str.isdigit, legajo))
                    
                    if legajo_clean and legajo_clean not in vistos:
                        antiguedad = 0
                        fecha_formateada = "N/A"
                        try:
                            fecha_norm = re.sub(r'\s+', '', fecha_ingreso_raw)
                            match_f = date_pattern.search(fecha_norm)
                            if match_f:
                                fecha_norm = match_f.group()
                            
                            for fmt in ('%d/%m/%Y', '%d/%m/%y', '%Y-%m-%d'):
                                try:
                                    fi = datetime.strptime(fecha_norm, fmt).date()
                                    hoy = datetime.now().date()
                                    if fi.year < 100:
                                        if fi.year > (hoy.year % 100) + 1:
                                            fi = fi.replace(year=1900 + fi.year)
                                        else:
                                            fi = fi.replace(year=2000 + fi.year)
                                    antiguedad = hoy.year - fi.year - ((hoy.month, hoy.day) < (fi.month, fi.day))
                                    fecha_formateada = fi.strftime('%d/%m/%Y')
                                    break
                                except: continue
                        except: pass
                            
                        datos_empleados.append({
                            'Legajo': int(legajo_clean) if legajo_clean.isdigit() else legajo_clean,
                            'Nombre y Apellido': nombre.strip().title(),
                            'Fecha de Ingreso': fecha_formateada,
                            'Antiguedad (anios)': antiguedad
                        })
                        vistos.add(legajo_clean)
                        self.log(f"Extraído: {nombre.strip().title()} (Leg: {legajo_clean})")
                except Exception as e:
                    self.log(f"Error bloque {i}: {e}")
            
            if not datos_empleados:
                self.root.after(0, lambda: messagebox.showwarning("Sin resultados", "No se detectó información válida."))
                return

            # Ordenamiento
            idx_file = self.index_path.get()
            if idx_file and os.path.exists(idx_file):
                self.log(f"Ordenando según índice: {os.path.basename(idx_file)}")
                try:
                    df_idx = pd.read_excel(idx_file, header=None)
                    orden_legajos = []
                    for val in df_idx[0].tolist():
                        if pd.isna(val): continue
                        s_val = "".join(filter(str.isdigit, str(val)))
                        if s_val: orden_legajos.append(s_val)
                    pesos = {leg: i for i, leg in enumerate(orden_legajos)}
                    datos_empleados.sort(key=lambda x: pesos.get(str(x['Legajo']), 999999))
                except Exception as e:
                    self.log(f"Error al ordenar: {e}")
            
            # Guardado final en EXCEL
            output_dir_init = os.path.dirname(pdf_path)
            excel_path = filedialog.asksaveasfilename(
                title="Guardar Planilla de Antigüedad",
                initialdir=output_dir_init,
                initialfile="Antiguedad.xlsx",
                defaultextension=".xlsx",
                filetypes=[("Archivos de Excel", "*.xlsx")]
            )
            
            if not excel_path:
                self.log("Guardado cancelado.")
                return

            df = pd.DataFrame(datos_empleados)
            df.to_excel(excel_path, index=False, sheet_name='Antiguedad')
            
            # Auto-ajuste de columnas
            wb = openpyxl.load_workbook(excel_path)
            ws = wb.active
            for col in ws.columns:
                max_length = 0
                column = col[0].column_letter
                for cell in col:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except: pass
                ws.column_dimensions[column].width = max_length + 2
            wb.save(excel_path)

            self.log(f"¡Éxito! Archivo generado con {len(datos_empleados)} registros.")
            self.root.after(0, lambda: messagebox.showinfo("Éxito", f"Se ha generado la planilla de antigüedad correctamente.\nRegistros: {len(datos_empleados)}"))
            
            if messagebox.askyesno("Abrir archivo", "¿Desea abrir el archivo Excel generado?"):
                os.startfile(excel_path)

        except Exception as e:
            self.log(f"ERROR: {e}")
            self.root.after(0, lambda: messagebox.showerror("Error", str(e)))
        finally:
            self.root.after(0, lambda: self.btn_procesar.configure(state='normal'))

if __name__ == "__main__":
    ctk.set_appearance_mode("Light")
    root = ctk.CTk()
    app = AntiguedadPDFApp(root)
    root.mainloop()
