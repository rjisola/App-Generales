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
script_dir = os.path.dirname(os.path.abspath(__file__))
if script_dir not in sys.path:
    sys.path.insert(0, script_dir)

# También agregar 03_OTROS (donde se encuentran los submódulos del sistema)
others_dir = os.path.abspath(os.path.join(script_dir, "..", "03_OTROS"))
if others_dir not in sys.path:
    sys.path.insert(0, others_dir)

# Importar componentes modernos
try:
    import modern_gui_components as mgc
    import customtkinter as ctk
    from icon_loader import set_window_icon, load_icon
    _has_icon_loader = True
except ImportError:
    import tkinter.messagebox as messagebox
    messagebox.showerror("Error", "Faltan módulos necesarios (modern_gui_components.py y customtkinter)")
    sys.exit(1)

class EnvelopePrinterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Impresión de Sobres - Asistente de Sueldos")
        self.root.geometry("1000x900")
        self.root.resizable(True, True)
        self.root.configure(bg=mgc.COLORS['bg_primary'])
        
        mgc.center_window(self.root, 1000, 900)

        if _has_icon_loader:
            set_window_icon(self.root, 'bonus_white')

        # Variables de Datos
        self.file_path = tk.StringVar(value="")
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
        self.icon_printer = load_icon('printer', (24, 24))
        
        # Contenedor principal con scrollbar
        self.main_frame = mgc.create_main_container(self.root)
        
        self.create_widgets()

        # Barra de estado
        self.status_frame, self.status_var = mgc.create_status_bar(self.root, "Listo. Selecciona un archivo.")

    def create_widgets(self):
        # --- HEADER (Azul como en captura) ---
        header_frame = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        header_frame.pack(fill=tk.X, pady=(0, 10))
        
        title_label = ctk.CTkLabel(header_frame, text="Impresión de Sobres C5",
                                   font=('Segoe UI', 32, 'bold'), text_color="#60a5fa")
        title_label.pack()
        
        subtitle_label = ctk.CTkLabel(header_frame, text="Selecciona un archivo Excel, elige las personas y genera los sobres.",
                                      font=('Segoe UI', 14), text_color="#94a3b8")
        subtitle_label.pack()

        # --- PASO 1: Cargar Archivo ---
        card1_outer, card1_inner = mgc.create_card(self.main_frame, "1. Cargar Archivo", padding=15)
        card1_outer.pack(fill=tk.X, pady=(0, 20))
        
        file_row = ctk.CTkFrame(card1_inner, fg_color="transparent")
        file_row.pack(fill=tk.X)
        
        btn_file = ctk.CTkButton(file_row, text="Elegir archivo", command=self.select_file,
                                 width=120, height=32, fg_color="#1e2a3a", text_color="#f1f5f9",
                                 border_width=1, border_color="#3b82f6", hover_color="#1e3a5f")
        btn_file.pack(side=tk.LEFT, padx=(0, 10))
        
        self.lbl_filename = ctk.CTkLabel(file_row, textvariable=self.file_path, font=mgc.FONTS['normal'], text_color="#94a3b8")
        self.lbl_filename.pack(side=tk.LEFT)

        # --- PASO 2: Seleccionar Destinatarios ---
        self.card2_outer, card2_inner = mgc.create_card(self.main_frame, "2. Seleccionar Destinatarios", padding=20)
        self.card2_outer.pack(fill=tk.BOTH, expand=True, pady=(0, 20))
        
        # Sub-Card de configuración (fondo celeste muy claro como en JS)
        config_frame = ctk.CTkFrame(card2_inner, fg_color="#0d1526", border_color="#1e3a5f", border_width=1, corner_radius=8)
        config_frame.pack(fill=tk.X, pady=(0, 15), padx=2)
        
        inner_config = ctk.CTkFrame(config_frame, fg_color="transparent")
        inner_config.pack(fill=tk.X, padx=15, pady=15)
        
        # FILA 1: Hoja Detectada | Filtro Efectivo
        row1 = ctk.CTkFrame(inner_config, fg_color="transparent")
        row1.pack(fill=tk.X, pady=(0, 10))
        
        col1 = ctk.CTkFrame(row1, fg_color="transparent")
        col1.pack(side=tk.LEFT, fill=tk.X, expand=True)
        ctk.CTkLabel(col1, text="Hoja Detectada:", font=mgc.FONTS['heading'], anchor='w').pack(fill=tk.X)
        self.sheet_select = ctk.CTkComboBox(col1, values=self.sheet_names, variable=self.selected_sheet, 
                                            command=lambda v: self.load_data(v), width=350)
        self.sheet_select.pack(fill=tk.X, pady=(2, 0))
        
        col2 = ctk.CTkFrame(row1, fg_color="transparent")
        col2.pack(side=tk.LEFT, padx=(20, 0), anchor='s')
        self.cb_efectivo = ctk.CTkCheckBox(col2, text='Filtrar solo "EFECTIVO" (Referencia en Col O)', 
                                           variable=self.filter_efectivo, command=self.apply_filters,
                                           font=('Segoe UI', 13, 'bold'), text_color="#1e293b")
        self.cb_efectivo.pack(pady=5)

        # FILA 2: Columna Legajo | Columna Nombre
        row2 = ctk.CTkFrame(inner_config, fg_color="transparent")
        row2.pack(fill=tk.X)
        
        col_l = ctk.CTkFrame(row2, fg_color="transparent")
        col_l.pack(side=tk.LEFT, fill=tk.X, expand=True)
        ctk.CTkLabel(col_l, text="Columna Legajo: (Default: C)", font=mgc.FONTS['heading'], anchor='w').pack(fill=tk.X)
        self.legajo_select = ctk.CTkComboBox(col_l, values=self.col_options, 
                                             command=lambda v: self.set_col('legajo', v))
        self.legajo_select.set("Col C (Ind 2)")
        self.legajo_select.pack(fill=tk.X, pady=(2, 0))
        
        col_n = ctk.CTkFrame(row2, fg_color="transparent")
        col_n.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(20, 0))
        ctk.CTkLabel(col_n, text="Columna Nombre: (Default: D)", font=mgc.FONTS['heading'], anchor='w').pack(fill=tk.X)
        self.nombre_select = ctk.CTkComboBox(col_n, values=self.col_options,
                                              command=lambda v: self.set_col('nombre', v))
        self.nombre_select.set("Col D (Ind 3)")
        self.nombre_select.pack(fill=tk.X, pady=(2, 0))

        # BUSCADOR
        self.search_entry = ctk.CTkEntry(card2_inner, textvariable=self.search_var, 
                                         placeholder_text="🔍 Buscar por nombre o legajo...", 
                                         height=35, font=mgc.FONTS['normal'])
        self.search_entry.pack(fill=tk.X, pady=(0, 10))

        # TABLA DE REGISTROS (Clonación de tabla web)
        table_outer = ctk.CTkFrame(card2_inner, fg_color="#0d1526", border_color="#1e2a3a", border_width=1, corner_radius=0)
        table_outer.pack(fill=tk.BOTH, expand=True)
        
        # Cabecera de Tabla
        self.table_header = ctk.CTkFrame(table_outer, fg_color="#111827", height=40, corner_radius=0)
        self.table_header.pack(fill=tk.X)
        self.table_header.pack_propagate(False)
        
        self.all_var = tk.BooleanVar(value=True)
        self.cb_all = ctk.CTkCheckBox(self.table_header, text="", variable=self.all_var, width=20, 
                                       command=self.toggle_all, checkbox_width=18, checkbox_height=18)
        self.cb_all.pack(side=tk.LEFT, padx=(10, 10))
        
        ctk.CTkLabel(self.table_header, text="Legajo", font=mgc.FONTS['heading'], width=150, anchor='w').pack(side=tk.LEFT)
        ctk.CTkLabel(self.table_header, text="Nombre Completo", font=mgc.FONTS['heading'], width=400, anchor='w').pack(side=tk.LEFT)
        ctk.CTkLabel(self.table_header, text="Ref (Col O)", font=mgc.FONTS['heading'], anchor='w').pack(side=tk.LEFT, fill=tk.X, expand=True)

        # Cuerpo con Scroll
        self.scroll_frame = ctk.CTkScrollableFrame(table_outer, fg_color="#0a0e1a", corner_radius=0)
        self.scroll_frame.pack(fill=tk.BOTH, expand=True)

        # Contador
        self.lbl_selected_count = ctk.CTkLabel(card2_inner, text="0 seleccionados", font=mgc.FONTS['small'], text_color="#94a3b8")
        self.lbl_selected_count.pack(anchor='w', pady=(5, 0))

        # BOTÓN FINAL (Azul Vibrante)
        self.btn_generate = ctk.CTkButton(self.main_frame, text="🖨️ GENERAR PDF DE SOBRES", 
                                          command=self.generate_pdf, font=('Segoe UI', 16, 'bold'),
                                          fg_color="#2563eb", hover_color="#1d4ed8", height=50, corner_radius=8)
        self.btn_generate.pack(fill=tk.X, pady=10)

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
            row_frame = ctk.CTkFrame(self.scroll_frame, fg_color="#0d1120", corner_radius=0)
            row_frame.pack(fill=tk.X)
            self.row_widgets.append(row_frame)
            
            # Efecto hover
            def on_enter(e, f=row_frame): f.configure(fg_color="#1e2a3a")
            def on_leave(e, f=row_frame): f.configure(fg_color="#0d1120")
            row_frame.bind("<Enter>", on_enter)
            row_frame.bind("<Leave>", on_leave)
            
            var = tk.BooleanVar(value=self.all_var.get())
            cb = ctk.CTkCheckBox(row_frame, text="", variable=var, width=20, 
                                 checkbox_width=18, checkbox_height=18, command=self.update_selection_count)
            cb.pack(side=tk.LEFT, padx=(10, 10), pady=8)
            self.checkbox_vars.append(var)
            
            ctk.CTkLabel(row_frame, text=item['legajo'], width=150, anchor='w', text_color="#f1f5f9").pack(side=tk.LEFT)
            ctk.CTkLabel(row_frame, text=item['nombre'], width=400, anchor='w', text_color="#f1f5f9").pack(side=tk.LEFT)
            ctk.CTkLabel(row_frame, text=item['ref'], anchor='w', text_color="#94a3b8", font=('Segoe UI', 11)).pack(side=tk.LEFT, fill=tk.X)
            
            # Separador sutil abajo
            tk.Frame(row_frame, height=1, bg="#1e2a3a").place(relx=0, rely=0.99, relwidth=1)

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
        root = tk.Tk()
        app = EnvelopePrinterApp(root)
        root.mainloop()
    except Exception as e:
        print(e)
        traceback.print_exc()
