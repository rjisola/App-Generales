# -*- coding: utf-8 -*-
"""
🖨️ Generador de Sobres - CLON EXACTO DE VERSIÓN JS
Reconstrucción visual total basada en la captura de pantalla:
- Paso 1 y Paso 2 con layout exacto.
- Selectores de columna C y D.
- Checkboxes individuales en lista scrollable.
- Buscador con lupa.
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
import sys
import pythoncom
import win32com.client
import traceback
import threading
import tempfile
from reportlab.pdfgen import canvas
from reportlab.lib.units import mm

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

class EnvelopePrinterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Impresión de Sobres - Asistente de Sueldos")
        self.root.geometry("900x700")
        self.root.resizable(True, True)
        self.root.configure(bg=mgc.COLORS['bg_primary'])
        
        mgc.center_window(self.root, 900, 700)

        if _has_icon_loader:
            set_window_icon(self.root, 'printer')

        # Variables de Datos
        self.file_path = tk.StringVar(value="Ningún archivo seleccionado")
        self.full_data = []         # Todos los registros leídos
        self.filtered_data = []     # Registros filtrados
        self.checkbox_vars = []     # Lista de BooleanVars para los checkboxes de la tabla
        self.row_widgets = []       # Widgets de filas en la tabla
        
        # Variables de Configuración (como en JS)
        self.sheet_names = []
        self.selected_sheet = tk.StringVar()
        self.filter_efectivo = tk.BooleanVar(value=True)
        self.col_legajo_idx = tk.StringVar(value="2") # Col C (Ind 2)
        self.col_nombre_idx = tk.StringVar(value="3") # Col D (Ind 3)
        self.search_var = tk.StringVar()
        self.search_var.trace_add("write", lambda *args: self.apply_filters())
        
        # Opciones para selects de columnas
        self.col_options = [f"Col {chr(65+i)} (Ind {i})" for i in range(26)] + [f"Col AA (Ind 26)", f"Col AB (Ind 27)"]
        
        # Iconos
        self.icon_sobres = load_icon('printer', (64, 64))
        self.icon_excel = load_icon('excel', (24, 24))
        self.icon_printer = load_icon('printer', (24, 24))
        
        # Contenedor principal con scrollbar
        self.main_container = mgc.create_main_container(self.root)
        
        self.create_widgets()

        # Barra de estado
        self.status_frame, self.status_var = mgc.create_status_bar(self.root, "Listo. Selecciona un archivo.")

    def create_widgets(self):
        # --- HEADER ---
        mgc.create_header(self.main_container, "Impresión de Sobres", 
                         "Selecciona un archivo Excel y genera los sobres C5",
                         icon_image=self.icon_sobres)

        # --- SECCIÓN 1: ARCHIVO Y CONFIGURACIÓN (Dos columnas) ---
        form_row = ctk.CTkFrame(self.main_container, fg_color="transparent")
        form_row.pack(fill=tk.X, pady=(0, 15))

        # Card 1: Archivo
        card1_outer, card1_inner = mgc.create_card(form_row, "1. Cargar Archivo", padding=15)
        card1_outer.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 10))
        
        btn_file = mgc.create_button(card1_inner, "Elegir archivo", self.select_file,
                                    color='purple', icon_image=self.icon_excel)
        btn_file.pack(pady=(0, 10))
        
        self.lbl_filename = ctk.CTkLabel(card1_inner, textvariable=self.file_path, 
                                        font=mgc.FONTS['small'], text_color=mgc.COLORS['text_secondary'],
                                        wraplength=350)
        self.lbl_filename.pack()

        # Card 2: Configuración
        card_cfg_outer, card_cfg_inner = mgc.create_card(form_row, "Configuración", padding=15)
        card_cfg_outer.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        ctk.CTkLabel(card_cfg_inner, text="Hoja Excel:", font=mgc.FONTS['small']).pack(anchor='w')
        self.sheet_select = ctk.CTkComboBox(card_cfg_inner, values=self.sheet_names, variable=self.selected_sheet, 
                                            command=lambda v: self.load_data(v), height=28)
        self.sheet_select.pack(fill=tk.X, pady=(0, 8))

        self.cb_efectivo = ctk.CTkCheckBox(card_cfg_inner, text='Filtrar solo "EFECTIVO"', 
                                           variable=self.filter_efectivo, command=self.apply_filters,
                                           font=mgc.FONTS['small'])
        self.cb_efectivo.pack(anchor='w')

        # --- SECCIÓN 2: SELECCIÓN Y BÚSQUEDA ---
        self.card2_outer, card2_inner = mgc.create_card(self.main_container, "2. Seleccionar Destinatarios", padding=15)
        self.card2_outer.pack(fill=tk.BOTH, expand=True, pady=(0, 15))

        # Barra de herramientas (Buscador + Selectores de Columna)
        toolbar = ctk.CTkFrame(card2_inner, fg_color="transparent")
        toolbar.pack(fill=tk.X, pady=(0, 10))

        self.search_entry = ctk.CTkEntry(toolbar, textvariable=self.search_var, 
                                         placeholder_text="🔍 Buscar por nombre o legajo...", 
                                         height=32, font=mgc.FONTS['normal'])
        self.search_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))

        # Selectores de columna compactos
        ctk.CTkLabel(toolbar, text="Cols (L|N):", font=mgc.FONTS['small']).pack(side=tk.LEFT, padx=2)
        self.legajo_select = ctk.CTkComboBox(toolbar, values=self.col_options, width=110, height=28,
                                             command=lambda v: self.set_col('legajo', v))
        self.legajo_select.set("Col C (Ind 2)")
        self.legajo_select.pack(side=tk.LEFT, padx=2)

        self.nombre_select = ctk.CTkComboBox(toolbar, values=self.col_options, width=110, height=28,
                                              command=lambda v: self.set_col('nombre', v))
        self.nombre_select.set("Col D (Ind 3)")
        self.nombre_select.pack(side=tk.LEFT, padx=2)

        # TABLA CON SCROLL
        table_container = ctk.CTkFrame(card2_inner, fg_color=mgc.COLORS['bg_primary'], 
                                       border_color=mgc.COLORS['border'], border_width=1, corner_radius=8)
        table_container.pack(fill=tk.BOTH, expand=True)
        
        # Cabecera
        header_table = ctk.CTkFrame(table_container, fg_color=mgc.COLORS['bg_card'], height=35, corner_radius=0)
        header_table.pack(fill=tk.X)
        header_table.pack_propagate(False)
        
        self.all_var = tk.BooleanVar(value=True)
        self.cb_all = ctk.CTkCheckBox(header_table, text="", variable=self.all_var, width=20, 
                                       command=self.toggle_all, checkbox_width=16, checkbox_height=16)
        self.cb_all.pack(side=tk.LEFT, padx=(10, 10))
        
        ctk.CTkLabel(header_table, text="Legajo", font=mgc.FONTS['heading'], width=80, anchor='w').pack(side=tk.LEFT)
        ctk.CTkLabel(header_table, text="Nombre Completo", font=mgc.FONTS['heading'], width=300, anchor='w').pack(side=tk.LEFT, padx=(10,0))
        ctk.CTkLabel(header_table, text="Ref (Col O)", font=mgc.FONTS['heading'], anchor='w').pack(side=tk.LEFT, fill=tk.X, expand=True, padx=10)

        # Cuerpo
        self.scroll_frame = ctk.CTkScrollableFrame(table_container, fg_color="transparent", corner_radius=0)
        self.scroll_frame.pack(fill=tk.BOTH, expand=True)

        # Pie (Contador)
        self.lbl_selected_count = ctk.CTkLabel(card2_inner, text="0 seleccionados", font=mgc.FONTS['small'], text_color=mgc.COLORS['text_secondary'])
        self.lbl_selected_count.pack(anchor='w', pady=(5, 0))

        # --- BOTÓN DE ACCIÓN ---
        self.btn_generate = mgc.create_large_button(self.main_container, "GENERAR PDF DE SOBRES", 
                                                    self.generate_pdf, color='blue',
                                                    icon_image=self.icon_printer)
        self.btn_generate.pack(fill=tk.X, pady=(0, 5))

    # --- Lógica de Negocio ---

    def select_file(self):
        path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx *.xlsm *.xls")])
        if path:
            self.file_path.set(os.path.basename(path))
            self.full_path = path
            self.extract_sheets(path)

    def extract_sheets(self, path):
        self.status_var.set("Cargando hojas...")
        threading.Thread(target=self._extract_sheets_thread, args=(path,), daemon=True).start()

    def _extract_sheets_thread(self, path):
        excel = None
        try:
            pythoncom.CoInitialize()
            excel = win32com.client.Dispatch("Excel.Application")
            wb = excel.Workbooks.Open(os.path.abspath(path), ReadOnly=True)
            names = [s.Name for s in wb.Sheets]
            wb.Close(False)
            excel.Quit()
            
            self.root.after(0, lambda: self._on_sheets_loaded(names))
        except Exception as e:
            self.root.after(0, lambda: self.status_var.set(f"Error: {e}"))

    def _on_sheets_loaded(self, names):
        self.sheet_names = names
        self.sheet_select.configure(values=names)
        
        # Auto selección JS
        target = next((n for n in names if n.upper().strip() == "RECUENTO TOTAL (2)"), names[0])
        self.selected_sheet.set(target)
        self.load_data(target)

    def set_col(self, type, value):
        idx = value.split("Ind ")[1].split(")")[0]
        if type == 'legajo': self.col_legajo_idx.set(idx)
        else: self.col_nombre_idx.set(idx)
        self.apply_filters()

    def load_data(self, sheet_name):
        self.status_var.set(f"Cargando datos de {sheet_name}...")
        threading.Thread(target=self._load_data_thread, args=(sheet_name,), daemon=True).start()

    def _load_data_thread(self, sheet_name):
        excel = None
        try:
            pythoncom.CoInitialize()
            excel = win32com.client.Dispatch("Excel.Application")
            wb = excel.Workbooks.Open(os.path.abspath(self.full_path), ReadOnly=True)
            ws = wb.Sheets(sheet_name)
            
            # Leer rango amplio (Cols A:O)
            last_row = ws.Cells(ws.Rows.Count, 3).End(-4162).Row # xlUp
            if last_row < 2: last_row = 1000
            
            vals = ws.Range(ws.Cells(1, 1), ws.Cells(last_row, 15)).Value
            wb.Close(False)
            excel.Quit()
            
            data = []
            if vals:
                for row in vals:
                    # Guardamos la fila entera para poder re-mapear columnas dinámicamente
                    data.append(row)
            
            self.full_data = data
            self.root.after(0, self.apply_filters)
        except Exception as e:
            self.root.after(0, lambda: self.status_var.set(f"Error: {e}"))

    def apply_filters(self):
        if not self.full_data: return
        
        leg_idx = int(self.col_legajo_idx.get())
        nom_idx = int(self.col_nombre_idx.get())
        ref_idx = 14 # Col O fija segun JS
        
        query = self.search_var.get().lower()
        only_efectivo = self.filter_efectivo.get()
        
        new_list = []
        for row in self.full_data:
            leg = str(row[leg_idx]) if row[leg_idx] is not None else ""
            if not leg or "legajo" in leg.lower(): continue # Saltar vacíos y headers
            
            # Normalizar legajo si es float (ej: 1048.0 -> 1048)
            if "." in leg and leg.replace(".", "").isdigit():
                leg = str(int(float(leg)))

            nom = str(row[nom_idx]).upper() if row[nom_idx] is not None else ""
            ref = str(row[ref_idx]) if len(row) > 14 and row[ref_idx] is not None else ""
            
            if only_efectivo and "efectivo" not in ref.lower(): continue
            if query and (query not in leg.lower() and query not in nom.lower()): continue
            
            new_list.append({'legajo': leg, 'nombre': nom, 'ref': ref})
            
        self.filtered_data = new_list
        self.render_list()

    def render_list(self):
        # Limpiar widgets actuales
        for w in self.row_widgets:
            w.destroy()
        self.row_widgets = []
        self.checkbox_vars = []
        
        for i, item in enumerate(self.filtered_data):
            row_frame = ctk.CTkFrame(self.scroll_frame, fg_color=mgc.COLORS['bg_card'], corner_radius=0)
            row_frame.pack(fill=tk.X)
            self.row_widgets.append(row_frame)
            
            # Efecto hover
            def on_enter(e, f=row_frame): f.configure(fg_color=mgc.COLORS['bg_input'])
            def on_leave(e, f=row_frame): f.configure(fg_color=mgc.COLORS['bg_card'])
            row_frame.bind("<Enter>", on_enter)
            row_frame.bind("<Leave>", on_leave)
            
            var = tk.BooleanVar(value=self.all_var.get())
            cb = ctk.CTkCheckBox(row_frame, text="", variable=var, width=20, 
                                 checkbox_width=18, checkbox_height=18, command=self.update_selection_count)
            cb.pack(side=tk.LEFT, padx=(10, 10), pady=8)
            self.checkbox_vars.append(var)
            
            ctk.CTkLabel(row_frame, text=item['legajo'], width=150, anchor='w', text_color=mgc.COLORS['text_primary']).pack(side=tk.LEFT)
            ctk.CTkLabel(row_frame, text=item['nombre'], width=400, anchor='w', text_color=mgc.COLORS['text_primary']).pack(side=tk.LEFT)
            ctk.CTkLabel(row_frame, text=item['ref'], anchor='w', text_color=mgc.COLORS['text_secondary'], font=('Segoe UI', 11)).pack(side=tk.LEFT, fill=tk.X)
            
            # Separador sutil abajo
            tk.Frame(row_frame, height=1, bg=mgc.COLORS['border']).place(relx=0, rely=0.99, relwidth=1)

        self.update_selection_count()
        self.status_var.set(f"Lista actualizada: {len(self.filtered_data)} registros.")

    def toggle_all(self):
        state = self.all_var.get()
        for v in self.checkbox_vars:
            v.set(state)
        self.update_selection_count()

    def update_selection_count(self):
        selected = sum(1 for v in self.checkbox_vars if v.get())
        self.lbl_selected_count.configure(text=f"{selected} seleccionados")

    def generate_pdf(self):
        destinatarios = []
        for i, var in enumerate(self.checkbox_vars):
            if var.get():
                destinatarios.append(self.filtered_data[i])
        
        if not destinatarios:
            messagebox.showwarning("Atención", "Selecciona al menos una persona.")
            return

        self.status_var.set(f"Generando PDF para {len(destinatarios)} sobres...")
        
        try:
            temp_dir = tempfile.gettempdir()
            pdf_path = os.path.join(temp_dir, f"sobres_{os.getpid()}.pdf")
            
            # Dimensiones: 160mm x 257mm
            width_pts = 160 * mm
            height_pts = 257 * mm
            
            c = canvas.Canvas(pdf_path, pagesize=(width_pts, height_pts))
            for p in destinatarios:
                texto = f"({p['legajo']}) {p['nombre']}"
                c.setFont("Helvetica-Bold", 14)
                y_pos = height_pts - (30 * mm) # 3cm
                c.drawCentredString(width_pts / 2.0, y_pos, texto)
                c.showPage()
            c.save()
            
            self.status_var.set("✅ PDF generado con éxito.")
            os.startfile(pdf_path)
            
        except Exception as e:
            messagebox.showerror("Error", str(e))

if __name__ == "__main__":
    try:
        root = ctk.CTk()
        app = EnvelopePrinterApp(root)
        root.mainloop()
    except Exception as e:
        print(e)
        traceback.print_exc()
