
import xlsxwriter
import os
import sys
import re
import tkinter as tk
import customtkinter as ctk
from tkinter import filedialog, messagebox
import time
import threading

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

def safe_filename(nombre):
    """Convierte el nombre a un nombre de archivo válido en mayúsculas."""
    nombre = nombre.upper().strip()
    nombre = re.sub(r'[<>:"/\\|?*]', '', nombre)
    return nombre


class EPPGeneratorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("📝 Generador de Formulario EPP")
        self.root.geometry("900x700")
        self.root.configure(bg=mgc.COLORS['bg_primary'])
        mgc.center_window(self.root, 900, 700)

        self.index_file_path = tk.StringVar()
        self.setup_modern_ui()

    def setup_modern_ui(self):
        main_container = mgc.create_main_container(self.root)
        
        inner_frame = tk.Frame(main_container, bg=mgc.COLORS['bg_primary'])
        inner_frame.pack(fill=tk.BOTH, expand=True, padx=25, pady=8)

        mgc.create_header(
            inner_frame, 
            "Formulario EPP", 
            "Generación masiva de formularios de entrega de elementos de protección personal", 
            icon="📝"
        )

        # ─── Sección: Archivo Índice (OBLIGATORIO para batch) ───
        idx_outer, idx_inner = mgc.create_card(inner_frame, "📂  Archivo Índice de Personal", padding=10)
        idx_outer.pack(fill=tk.X, pady=(0, 8))

        tk.Entry(idx_inner, textvariable=self.index_file_path,
                 font=mgc.FONTS['small'], bd=1, relief=tk.SOLID,
                 state='readonly', fg=mgc.COLORS['text_secondary']
                 ).grid(row=0, column=0, sticky='we', padx=(0, 8), pady=(0, 4))
        mgc.create_button(idx_inner, "Seleccionar", self.select_index_file,
                          color='purple', icon="📂").grid(row=0, column=1, pady=(0, 4))

        # Rango de filas
        rng_frame = tk.Frame(idx_inner, bg=mgc.COLORS['bg_card'])
        rng_frame.grid(row=1, column=0, columnspan=2, sticky='we')
        tk.Label(rng_frame, text="Desde fila:", font=mgc.FONTS['small'],
                 bg=mgc.COLORS['bg_card'], fg=mgc.COLORS['text_secondary']).pack(side=tk.LEFT, padx=(0, 4))
        self.idx_desde = tk.IntVar(value=2)
        tk.Spinbox(rng_frame, from_=2, to=500, width=4,
                   textvariable=self.idx_desde, font=mgc.FONTS['normal'],
                   bd=1, relief=tk.SOLID).pack(side=tk.LEFT)
        tk.Label(rng_frame, text="  Hasta:", font=mgc.FONTS['small'],
                 bg=mgc.COLORS['bg_card'], fg=mgc.COLORS['text_secondary']).pack(side=tk.LEFT, padx=(10, 4))
        self.idx_hasta = tk.IntVar(value=999)
        tk.Spinbox(rng_frame, from_=2, to=999, width=4,
                   textvariable=self.idx_hasta, font=mgc.FONTS['normal'],
                   bd=1, relief=tk.SOLID).pack(side=tk.LEFT)
        tk.Label(rng_frame, text="  (999 = hasta el final)",
                 font=mgc.FONTS['tiny'], bg=mgc.COLORS['bg_card'],
                 fg=mgc.COLORS['gray']).pack(side=tk.LEFT, padx=6)
        idx_inner.columnconfigure(0, weight=1)

        # ─── Sección: Datos por defecto (cuando NO hay índice) ───
        card_outer, card_inner = mgc.create_card(
            inner_frame, "✏️  Datos por Defecto (sin índice)", padding=10)
        card_outer.pack(fill=tk.X, pady=(0, 8))

        self.vars = {
            'nombre':        tk.StringVar(value='AGUILAR IVAN ARTURO'),
            'dni':           tk.StringVar(value='20277430534'),
            'proyecto':      tk.StringVar(value='CARJOR DEPOSITO'),
            'cargo':         tk.StringVar(value='OPERARIO'),
            'jefe':          tk.StringVar(value='TORCHIANA AGOSTINA'),
            'fecha_entrega': tk.StringVar(value='05-01-2026')
        }
        for i, (key, label) in enumerate([
            ('nombre',        'Nombre:'),
            ('dni',           'DNI / CUIL:'),
            ('proyecto',      'Área / Proyecto:'),
            ('cargo',         'Cargo:'),
            ('jefe',          'Jefe Inmediato:'),
            ('fecha_entrega', 'Fecha (DD-MM-AAAA):')
        ]):
            tk.Label(card_inner, text=label, font=mgc.FONTS['small'],
                     bg=mgc.COLORS['bg_card'], fg=mgc.COLORS['text_secondary'],
                     width=18, anchor='w').grid(row=i, column=0, sticky='w', pady=2)
            tk.Entry(card_inner, textvariable=self.vars[key],
                     font=mgc.FONTS['normal'], bd=1, relief=tk.SOLID
                     ).grid(row=i, column=1, sticky='we', padx=(8, 0), pady=2)
        card_inner.columnconfigure(1, weight=1)

        # ─── Progreso ───
        self.prog_frame, self.progress_bar, self.progress_text = mgc.create_progress_section(inner_frame)
        self.prog_frame.pack(fill=tk.X, pady=6)

        # ─── Formato de salida ───
        fmt_outer, fmt_inner = mgc.create_card(inner_frame, "📄  Seleccionar Formulario", padding=10)
        fmt_outer.pack(fill=tk.X, pady=(0, 8))
        self.formato_var = tk.StringVar(value='excel')

        tk.Radiobutton(fmt_inner,
                       text="Formulario 1 — Excel (.xlsx)  [réplica generada por la app]",
                       variable=self.formato_var, value='excel',
                       font=mgc.FONTS['normal'], bg=mgc.COLORS['bg_card'],
                       fg=mgc.COLORS['text_primary']
                       ).grid(row=0, column=0, columnspan=3, sticky='w', padx=4, pady=3)

        tk.Radiobutton(fmt_inner,
                       text="Formulario 2 — Word (.docx) [copia fiel del EPP MASTER]",
                       variable=self.formato_var, value='word',
                       font=mgc.FONTS['normal'], bg=mgc.COLORS['bg_card'],
                       fg=mgc.COLORS['text_primary']
                       ).grid(row=1, column=0, columnspan=3, sticky='w', padx=4, pady=3)

        # Template Word en row=2 (sin conflicto)
        tk.Label(fmt_inner, text="Template Word:", font=mgc.FONTS['tiny'],
                 bg=mgc.COLORS['bg_card'], fg=mgc.COLORS['gray']
                 ).grid(row=2, column=0, sticky='w', padx=4, pady=(4, 0))
        self.word_template_var = tk.StringVar(value=DEFAULT_WORD_TEMPLATE)
        tk.Entry(fmt_inner, textvariable=self.word_template_var, font=mgc.FONTS['tiny'],
                 bd=1, relief=tk.SOLID, fg=mgc.COLORS['text_secondary']
                 ).grid(row=2, column=1, sticky='we', padx=(0, 6), pady=(4, 0))
        mgc.create_button(fmt_inner, "…", self.select_word_template,
                          color='gray').grid(row=2, column=2, padx=(0, 4), pady=(4, 0))
        fmt_inner.columnconfigure(1, weight=1)

        # ─── Carpeta de Destino ───
        dest_outer, dest_inner = mgc.create_card(inner_frame, "📁  Carpeta de Destino", padding=8)
        dest_outer.pack(fill=tk.X, pady=(0, 8))
        self.output_dir_var = tk.StringVar(value=os.path.expanduser("~\\Desktop"))
        tk.Entry(dest_inner, textvariable=self.output_dir_var, font=mgc.FONTS['small'],
                 bd=1, relief=tk.SOLID).grid(row=0, column=0, sticky='we', padx=(0, 8))
        mgc.create_button(dest_inner, "Elegir…", self.select_output_dir,
                          color='orange', icon="📁").grid(row=0, column=1)
        dest_inner.columnconfigure(0, weight=1)

        tk.Label(dest_inner,
                 text="Los archivos se llamarán: NOMBRE COMPLETO.EPP.xlsx (.docx si Word)",
                 font=mgc.FONTS['tiny'], bg=mgc.COLORS['bg_card'],
                 fg=mgc.COLORS['gray']).grid(row=1, column=0, columnspan=2, sticky='w', pady=(4, 0))

        # ─── Botones ───
        btn_frame = tk.Frame(inner_frame, bg=mgc.COLORS['bg_primary'])
        btn_frame.pack(fill=tk.X, pady=(4, 12))

        self.btn_generar_uno = mgc.create_large_button(
            btn_frame, "  Generar (datos actuales)", self.start_one,
            color='blue', icon="📄")
        self.btn_generar_uno.pack(fill=tk.X, pady=(0, 6))

        self.btn_generar_todos = mgc.create_large_button(
            btn_frame, "  Generar TODOS desde Índice", self.start_batch,
            color='purple', icon="🚀")
        self.btn_generar_todos.pack(fill=tk.X)

        _, self.status_var = mgc.create_status_bar(self.root, "Listo para generar")

    # ────────────────────────────────────────────────────────
    # SELECTORES
    # ────────────────────────────────────────────────────────
    def select_index_file(self):
        path = filedialog.askopenfilename(
            title="Seleccionar archivo índice de personal",
            filetypes=[("Excel", "*.xlsx *.xls *.xlsm"), ("Todos", "*.*")]
        )
        if path:
            self.index_file_path.set(path)
            self.status_var.set(f"Índice: {os.path.basename(path)}")

    def select_output_dir(self):
        path = filedialog.askdirectory(title="Seleccionar carpeta de destino")
        if path:
            self.output_dir_var.set(path)

    def select_word_template(self):
        path = filedialog.askopenfilename(
            title="Seleccionar template Word EPP maestro",
            filetypes=[("Word", "*.docx *.doc"), ("Todos", "*.*")]
        )
        if path:
            self.word_template_var.set(path)

    # ────────────────────────────────────────────────────────
    # LANZADORES
    # ────────────────────────────────────────────────────────
    def _disable_buttons(self):
        if mgc:
            mgc.disable_button(self.btn_generar_uno)
            mgc.disable_button(self.btn_generar_todos)

    def _enable_buttons(self):
        if mgc:
            mgc.enable_button(self.btn_generar_uno, 'blue')
            mgc.enable_button(self.btn_generar_todos, 'purple')

    def start_one(self):
        """Genera un único formulario con los datos actuales de los campos."""
        self._disable_buttons()
        t = threading.Thread(target=self._run_one)
        t.daemon = True
        t.start()

    def start_batch(self):
        """Lee todas las filas del índice y genera un EPP por persona."""
        if not self.index_file_path.get():
            messagebox.showwarning("Sin índice", "Primero selecciona el archivo índice de personal.")
            return
        if not HAS_OPENPYXL:
            messagebox.showerror("Dependencia", "Instalá openpyxl:\n  pip install openpyxl")
            return
        self._disable_buttons()
        t = threading.Thread(target=self._run_batch)
        t.daemon = True
        t.start()

    # ────────────────────────────────────────────────────────
    # EJECUCIÓN: MODO INDIVIDUAL
    # ────────────────────────────────────────────────────────
    def _run_one(self):
        try:
            datos = {k: v.get() for k, v in self.vars.items()}
            nombre_archivo = safe_filename(datos['nombre']) + '.EPP'
            fmt = self.formato_var.get()
            out = os.path.join(self.output_dir_var.get(), nombre_archivo + '.xlsx')  # SIEMPRE xlsx

            self.root.after(0, self.update_progress, 20, f"Generando {nombre_archivo}.xlsx…")
            time.sleep(0.2)

            if fmt == 'word':
                self.generar_excel_formulario2(out, datos)
            else:
                self.generar_excel_logic(out, datos)

            self.root.after(0, self.update_progress, 100, "¡Completado!")
            self.root.after(0, self.status_var.set, f"✓ Guardado: {nombre_archivo}.xlsx")
            if messagebox.askyesno("Éxito", f"Formulario generado:\n{out}\n\n¿Abrirlo ahora?"):
                os.startfile(out)
        except Exception as e:
            messagebox.showerror("Error", str(e))
        finally:
            self.root.after(0, self._enable_buttons)

    # ────────────────────────────────────────────────────────
    # EJECUCIÓN: MODO BATCH (TODOS desde ÍNDICE)
    # ────────────────────────────────────────────────────────
    def _run_batch(self):
        try:
            wb = openpyxl.load_workbook(self.index_file_path.get(), data_only=True)
            ws = wb.active

            desde  = self.idx_desde.get()
            hasta  = self.idx_hasta.get()
            fmt    = self.formato_var.get()
            outdir = self.output_dir_var.get()  # siempre .xlsx

            # Recopilar filas válidas
            filas = []
            for r in range(desde, ws.max_row + 1):
                if r > hasta:
                    break
                def cel(col):
                    v = ws.cell(row=r, column=col).value
                    return str(v).strip().upper() if v is not None else ''
                nombre = cel(1)
                if not nombre:
                    break  # fin de datos
                filas.append({
                    'nombre':        nombre,
                    'dni':           cel(2),
                    'proyecto':      cel(3),
                    'cargo':         cel(4),
                    'jefe':          cel(5),
                    'fecha_entrega': cel(6) or self.vars['fecha_entrega'].get()
                })

            total = len(filas)
            if total == 0:
                messagebox.showinfo("Sin datos", f"No se encontraron filas con datos desde la fila {desde}.")
                return

            self.root.after(0, self.update_progress, 0, f"Procesando {total} persona(s)…")
            generados = []
            errores   = []

            for i, datos in enumerate(filas, 1):
                pct  = int(i / total * 100)
                nombre_archivo = safe_filename(datos['nombre']) + '.EPP'
                out  = os.path.join(outdir, nombre_archivo + '.xlsx')  # SIEMPRE xlsx
                self.root.after(0, self.update_progress, pct,
                                f"[{i}/{total}] {nombre_archivo}.xlsx")
                try:
                    if fmt == 'word':
                        self.generar_excel_formulario2(out, datos)
                    else:
                        self.generar_excel_logic(out, datos)
                    generados.append(nombre_archivo + '.xlsx')
                except Exception as e:
                    errores.append(f"{datos['nombre']}: {e}")

            self.root.after(0, self.update_progress, 100, f"✓ {len(generados)}/{total} generados")
            self.root.after(0, self.status_var.set,
                            f"✓ {len(generados)} EPP generados en: {outdir}")

            resumen = f"✅ Generados: {len(generados)} archivos\n📁 Destino: {outdir}"
            if errores:
                resumen += f"\n\n⚠️ Errores ({len(errores)}):\n" + "\n".join(errores)
            messagebox.showinfo("Proceso completo", resumen)

            if len(generados) > 0:
                os.startfile(outdir)

        except Exception as e:
            messagebox.showerror("Error en batch", str(e))
        finally:
            self.root.after(0, self._enable_buttons)

    def update_progress(self, value, text):
        self.progress_bar['value'] = value
        self.progress_text.set(text)
        self.root.update_idletasks()

    # ────────────────────────────────────────────────────────
    # FORMULARIO 2 — Réplica Resolución 299/11 en Excel
    # ────────────────────────────────────────────────────────
    def generar_excel_formulario2(self, nombre_archivo, datos):
        """Replica la estructura del EPP MASTER.docx (Res 299/11) en un archivo xlsx."""
        wb = xlsxwriter.Workbook(nombre_archivo)
        ws = wb.add_worksheet("EPP Res 299-11")

        # ── Formatos ──
        bold       = wb.add_format({'bold': True, 'font_size': 10, 'valign': 'vcenter'})
        center_b   = wb.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter',
                                    'border': 1, 'text_wrap': True, 'font_size': 10})
        center_b_s = wb.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter',
                                    'border': 1, 'text_wrap': True, 'font_size': 9,
                                    'font_color': '#000080'})
        label_f    = wb.add_format({'bold': True, 'font_size': 10, 'valign': 'vcenter',
                                    'border': 1, 'text_wrap': True})
        data_f     = wb.add_format({'font_size': 10, 'valign': 'vcenter',
                                    'border': 1, 'text_wrap': True, 'bold': True})
        row_f      = wb.add_format({'font_size': 10, 'valign': 'vcenter', 'border': 1})
        note_f     = wb.add_format({'font_size': 8, 'italic': True, 'text_wrap': True,
                                    'valign': 'top', 'border': 1})

        # ── Anchos de columna (A-G = 7 columnas) ──
        ws.set_column('A:A', 28)  # Producto
        ws.set_column('B:B', 18)  # Tipo/Modelo
        ws.set_column('C:C', 16)  # Marca
        ws.set_column('D:D', 14)  # Certif
        ws.set_column('E:E', 10)  # Cantidad
        ws.set_column('F:F', 16)  # Fecha
        ws.set_column('G:G', 22)  # Firma

        # ── Fila 0: título principal ──
        ws.set_row(0, 22)
        ws.merge_range('A1:G1',
                       'Resolución 299/11, Anexo I — FORMULARIO ENTREGA DE ELEMENTOS DE PROTECCIÓN PERSONAL',
                       center_b)

        # ── Filas 1-2: empresa ──
        ws.set_row(1, 18); ws.set_row(2, 18)
        ws.merge_range('A2:D2', 'Razón Social: CARJOR SRL', label_f)
        ws.merge_range('E2:G2', 'C.U.I.T.: 30-70921165-6', label_f)
        ws.merge_range('A3:B3', 'Dirección: Independencia 685 Entrepiso', label_f)
        ws.write('C3', 'Localidad: ZÁRATE', label_f)
        ws.write('D3', 'C.P: 2800', label_f)
        ws.merge_range('E3:G3', 'Provincia: Bs. As', label_f)

        # ── Fila 3-4: datos del trabajador ──
        ws.set_row(3, 20); ws.set_row(4, 20)
        ws.merge_range('A4:E4',
                       f"Nombre y Apellido del Trabajador: {datos.get('nombre', '')}",
                       data_f)
        ws.merge_range('F4:G4', f"DNI: {datos.get('dni', '')}", data_f)
        ws.merge_range('A5:C5',
                       f"Cargo: {datos.get('cargo', '')}",
                       data_f)
        ws.merge_range('D5:G5',
                       f"Área / Proyecto: {datos.get('proyecto', '')}",
                       data_f)

        # ── Fila 5: cabecera tabla ──
        ws.set_row(5, 28)
        for col, txt in enumerate(['Producto', 'Tipo // Modelo', 'Marca',
                                   'Posee Certif. SI/NO', 'Cantidad',
                                   'Fecha de Entrega', 'Firma del Trabajador']):
            ws.write(5, col, txt, center_b_s)

        # ── Filas de ítems EPP (datos del Word maestro como referencia) ──
        ITEMS = [
            ('CAMISA',                  'GRAFA',                 'PAMPERO',        'SI', ''),
            ('PANTALON',                'GRAFA',                 'PAMPERO',        'SI', ''),
            ('ZAPATOS DE SEGURIDAD',    'FUNCIONAL',             'MACSI',          'SI', ''),
            ('CASCO DE PROTECCION',     'MILENIUM CLASS S/V',    'LIBUS',          'SI', ''),
            ('CHALECO REFLECTIVO',      'ESTAMPADO',             'ESTAMPADOS H',   'SI', ''),
            ('ANTEOJOS DE SEGURIDAD',   'TRANSPARENTES',         'LIBUS',          'SI', ''),
            ('PROTECTORES AUDITIVOS',   'QUANTUM DISPENSER',     'LIBUS',          'SI', ''),
            ('GUANTES DE SEGURIDAD',    'VAQUETA',               'DEPASCALE',      'SI', ''),
            ('',                        '',                      '',               '',   ''),
            ('',                        '',                      '',               '',   ''),
        ]
        fecha = datos.get('fecha_entrega', '')
        for r, (prod, tipo, marca, cert, _) in enumerate(ITEMS, start=6):
            ws.set_row(r, 22)
            ws.write(r, 0, prod,   row_f)
            ws.write(r, 1, tipo,   row_f)
            ws.write(r, 2, marca,  row_f)
            ws.write(r, 3, cert,   row_f)
            ws.write(r, 4, '',     row_f)  # cantidad (a completar)
            ws.write(r, 5, fecha if prod else '', row_f)
            ws.write(r, 6, '',     row_f)  # firma

        # ── Nota al pie ──
        nota_row = 6 + len(ITEMS) + 1
        ws.set_row(nota_row, 55)
        ws.merge_range(nota_row, 0, nota_row, 6,
            "Nota: Estos elementos de protección personal son de uso exclusivo para ejercer las "
            "funciones asignadas dentro de la empresa y por ello al finalizar el proyecto deberán "
            "ser devueltos a oficina principal en el estado que se encuentren.",
            note_f)

        ws.set_paper(9)
        ws.set_margins(left=0.3, right=0.3, top=0.4, bottom=0.4)
        ws.fit_to_pages(1, 1)
        wb.close()

    # ────────────────────────────────────────────────────────
    # GENERACIÓN EXCEL (réplica con xlsxwriter)
    # ────────────────────────────────────────────────────────
    def generar_excel_logic(self, nombre_archivo, datos_empleado):
        workbook = xlsxwriter.Workbook(nombre_archivo)
        sheet = workbook.add_worksheet("Entrega de EPP")

        border_full        = workbook.add_format({'border': 1, 'valign': 'vcenter'})
        header_main_fmt    = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'text_wrap': True, 'font_size': 11})
        header_table_fmt   = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'text_wrap': True, 'font_size': 8, 'font_color': '#000080'})
        vertical_fmt       = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'rotation': 90, 'border': 1, 'font_size': 7, 'text_wrap': True, 'font_color': '#000080'})
        vertical_bold_fmt  = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'rotation': 90, 'border': 1, 'font_size': 9, 'text_wrap': True})
        horizontal_sub_fmt = workbook.add_format({'bold': True, 'align': 'center', 'valign': 'vcenter', 'border': 1, 'font_size': 7, 'text_wrap': True, 'font_color': '#000080'})
        label_fmt          = workbook.add_format({'bold': True, 'font_size': 10, 'valign': 'vcenter'})
        label_right_fmt    = workbook.add_format({'bold': True, 'font_size': 10, 'valign': 'vcenter', 'align': 'right'})
        data_underline_fmt = workbook.add_format({'bottom': 1, 'font_size': 11, 'valign': 'vcenter', 'bold': True})
        note_fmt           = workbook.add_format({'font_size': 7, 'text_wrap': True, 'align': 'left', 'valign': 'top'})
        border_center_fmt  = workbook.add_format({'border': 1, 'align': 'center', 'valign': 'vcenter'})
        header_sub_fmt     = workbook.add_format({'border': 1, 'align': 'center', 'valign': 'vcenter', 'font_size': 10})
        data_table_fmt     = workbook.add_format({'border': 1, 'align': 'center', 'valign': 'vcenter', 'font_color': '#000080'})

        sheet.set_column('A:A', 15); sheet.set_column('B:E', 3.8)
        sheet.set_column('F:I', 9.5); sheet.set_column('J:O', 3.8)
        sheet.set_column('P:P', 10); sheet.set_column('Q:R', 15)
        sheet.set_row(0, 20); sheet.set_row(1, 20)
        sheet.set_row(6, 20); sheet.set_row(8, 20)
        sheet.set_row(9, 45); sheet.set_row(10, 30); sheet.set_row(11, 85)

        sheet.merge_range('A1:C2', '', border_full)
        sheet.merge_range('D1:R1', 'FORMULARIO DE ENTREGA DE ELEMENTOS DE PROTECCIÓN PERSONAL Y DOTACIONES', header_main_fmt)
        sheet.merge_range('D2:L2', 'CODIGO: 0001', header_sub_fmt)
        sheet.merge_range('M2:O2', '', border_full)
        sheet.merge_range('P2:R2', 'Versión: 1', header_sub_fmt)

        sheet.merge_range('J6:N6', 'ENTREGA DE EPP', label_right_fmt)
        sheet.merge_range('J8:N8', 'DEVOLUCION DE EPP', label_right_fmt)
        sheet.write('O6', 'X', border_center_fmt)
        sheet.write('O8', '',  border_center_fmt)

        row_map = {'nombre': 3, 'dni': 4, 'proyecto': 5, 'cargo': 6, 'jefe': 7}
        labels  = {'nombre': 'Nombre Empleado:', 'dni': 'No de Identificación:',
                   'proyecto': 'Area y/o Proyecto:', 'cargo': 'Cargo:', 'jefe': 'Jefe inmediato:'}
        for key, row in row_map.items():
            sheet.merge_range(row, 0, row, 2, labels[key], label_fmt)
            sheet.merge_range(row, 3, row, 8, datos_empleado.get(key, ''), data_underline_fmt)

        sheet.merge_range('A10:A12', 'FECHA DE ENTREGA', header_table_fmt)
        sheet.merge_range('B10:I10', 'ELEMENTOS DE PROTECCION PERSONAL', header_table_fmt)
        sheet.merge_range('J10:O10',
            'DOTACIÓN DE INVIERNO - ELEMENTO DE IDENTIFICACIÓN VISUAL - ELEMENTO DE PROTECCIÓN PERSONAL CONTRA CAÍDAS',
            header_table_fmt)
        sheet.write('P10', 'ID Personal', header_table_fmt)
        sheet.merge_range('Q10:Q12', 'OBSERVACIONES', header_table_fmt)
        sheet.merge_range('R10:R12', 'FIRMA DEL TRABAJADOR', vertical_bold_fmt)

        sheet.merge_range('B11:B12', 'CASCO', vertical_fmt)
        sheet.merge_range('C11:C12', 'PROTECCION AUDITIVA', vertical_fmt)
        sheet.merge_range('D11:D12', 'PROTECCION VISUAL', vertical_fmt)
        sheet.merge_range('E11:E12', 'PROTECCION RESPIRATORIA', vertical_fmt)
        sheet.merge_range('F11:G11', 'PROTECCION EXTREMIDADES INFERIORES', header_table_fmt)
        sheet.merge_range('H11:I11', 'PROTECCION EXTREMIDADES SUPERIORES', header_table_fmt)
        sheet.write('F12', 'BOTAS DE SEGURIDAD', horizontal_sub_fmt)
        sheet.write('G12', 'BOTAS DIELECTRICAS O AISLANTES', horizontal_sub_fmt)
        sheet.write('H12', 'GUANTE CUERO (VAQUETA)', horizontal_sub_fmt)
        sheet.write('I12', 'GUANTE AISLANTE DE ELECTRICIDAD', horizontal_sub_fmt)
        sheet.merge_range('J11:J12', 'IMPERMEABLE / CHAQUETA IMPERMEABLE', vertical_fmt)
        sheet.merge_range('K11:K12', 'BOTAS DE CAUCHO TIPO ING. ROYAL', vertical_fmt)
        sheet.merge_range('L11:L12', 'CHAQUETA CON CINTAS REFLECTIVAS DE IDENTIFICACION EMPRESARIAL', vertical_fmt)
        sheet.merge_range('M11:M12', 'CHALECO CON CINTAS REFLECTIVAS', vertical_fmt)
        sheet.merge_range('N11:N12', 'OVEROL', vertical_fmt)
        sheet.merge_range('O11:O12', 'ARNES DE SEGURIDAD', vertical_fmt)
        sheet.merge_range('P11:P12', 'CARNET DE IDENTIFICACION', vertical_fmt)

        for r in range(12, 18):
            sheet.set_row(r, 30)
            for c in range(18):
                sheet.write(r, c, '', border_full)

        if datos_empleado.get('fecha_entrega'):
            sheet.write('A13', datos_empleado['fecha_entrega'], data_table_fmt)
            sheet.write('B13', 'X', data_table_fmt)
            sheet.write('C13', 'X', data_table_fmt)
            sheet.write('D13', 'X', data_table_fmt)
            sheet.write('F13', 'X', data_table_fmt)
            sheet.write('H13', 'X', data_table_fmt)

        nota_texto = (
            "Nota:\n"
            "Estos elementos de protección personal son de uso exclusivo para ejercer las funciones asignadas "
            "dentro de la empresa y por ello al finalizar el proyecto deberán ser devueltos a oficina principal "
            "en el estado que se encuentren. (Incluye carnet personal de identificación).\n"
            "Estos elementos deben ser usados en todo momento mientras el colaborador se encuentre en la obra y/o proyecto"
        )
        sheet.set_row(28, 60)
        sheet.merge_range('A29:R29', nota_texto, note_fmt)
        sheet.set_paper(9)
        sheet.set_margins(left=0.2, right=0.2, top=0.3, bottom=0.3)
        sheet.fit_to_pages(1, 1)
        workbook.close()

    def setup_basic_ui(self):
        self.root.title("Generador EPP")
        tk.Label(self.root, text="Interfaz básica: falta modern_gui_components.py").pack()


if __name__ == "__main__":
    root = ctk.CTk()
    app = EPPGeneratorApp(root)
    root.mainloop()
