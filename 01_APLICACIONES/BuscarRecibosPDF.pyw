import os
import fitz  # Reemplaza PyPDF2
import tkinter as tk
import customtkinter as ctk
from tkinter import ttk, filedialog, messagebox
import tkinter.scrolledtext as scrolledtext
import sys
import pandas as pd
import unicodedata
import zipfile
import io
import threading
import re

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
from icon_loader import set_window_icon, load_icon

def normalize_text(text):
    """Normaliza texto para comparación (elimina acentos, puntuación, mayúsculas y espacios extra)."""
    if not isinstance(text, str):
        text = str(text)
    # Eliminar acentos
    text = ''.join(
        c for c in unicodedata.normalize('NFD', text)
        if unicodedata.category(c) != 'Mn'
    )
    # Reemplazar puntuación por espacios
    text = re.sub(r'[^\w\s]', ' ', text)
    # Convertir a minúsculas y colapsar espacios/saltos de línea
    return ' '.join(text.lower().split())

def check_signature(text_norm, filter_type):
    """Valida la firma de forma flexible ignorando espacios entre letras."""
    if filter_type == 'ambas':
        return True
    
    # Eliminar espacios para buscar palabras pegadas (caso de PDFs con espaciado ancho)
    text_clean = text_norm.replace(" ", "")
    
    if filter_type == 'empleado':
        # Debe tener 'firma' y 'empleado' pero no ser 'empleador'
        # El empleador contiene al empleado, así que buscamos la 'r' o 'dor' final
        has_empleado = 'firmadelempleado' in text_clean or ('firma' in text_clean and 'empleado' in text_clean)
        is_empleador = 'empleador' in text_clean
        return has_empleado and not is_empleador
        
    if filter_type == 'empleador':
        return 'firmadelempleador' in text_clean or ('firma' in text_clean and 'empleador' in text_clean)
    
    return True

def get_search_variations(full_name_norm):
    """Genera variaciones de búsqueda para un nombre ya normalizado (ej: Apellido Nombre1)."""
    parts = full_name_norm.split()
    variations = [full_name_norm]
    
    if len(parts) >= 2:
        # 1. Apellido + Primer Nombre (para casos truncados o simplificados)
        variations.append(f"{parts[0]} {parts[1]}")
        
        # 2. Invertir: Nombre1 + Apellido (Maneja casos donde el PDF tiene Nombre Apellido)
        # Asumiendo que parts[0] es el Apellido
        # Si tiene 2 partes: [Apellido, Nombre] -> [Nombre, Apellido]
        # Si tiene 3 partes: [Apellido, Nombre1, Nombre2] -> [Nombre1, Apellido]
        variations.append(f"{parts[1]} {parts[0]}")
        
        # 3. Nombre1 Nombre2 + Apellido
        if len(parts) >= 3:
            variations.append(f"{' '.join(parts[1:])} {parts[0]}")
            
    return list(set(variations)) # Eliminar duplicados

class SeparadorRecibosApp:
    def __init__(self, root):
        self.root = root
        self.root.title("🔍 Gestor de Recibos PDF")
        self.root.geometry("900x700")
        self.root.resizable(True, True)
        self.root.configure(bg=mgc.COLORS['bg_primary'])
        
        mgc.center_window(self.root, 900, 700)
        
        # Establecer icono de ventana
        set_window_icon(self.root, 'search')
        
        # Cargar iconos PNG
        self.icon_search = load_icon('search', (64, 64))
        self.icon_pdf = load_icon('pdf', (24, 24))
        self.icon_folder = load_icon('folder', (24, 24))
        self.icon_excel = load_icon('excel', (24, 24))
        self.icon_warning = load_icon('warning', (24, 24))
        self.icon_check = load_icon('check', (24, 24))
        self.icon_zip = load_icon('folder', (24, 24)) # Using generic folder icon instead of zip

        # Estado
        self.archivos_pdf_seleccionados = []
        self.archivo_indice_path = None
        self.open_pdf_readers = []
        self.var_filtro_firma = tk.StringVar(value="ambas") # 'ambas', 'empleado', 'empleador'
        self.var_optimizar_auto = tk.BooleanVar(value=True) # Nueva: Optimización automática
        self.pdf_data_cache = None # Caché persistente en memoria: { "path": [(page_num, text), ...] }

        # Interfaz Principal
        self.setup_ui()

    def setup_ui(self):
        # Barra de estado inferior (Crear primero)
        self.status_frame, self.status_var = mgc.create_status_bar(self.root, "Listo para comenzar")

        # Contenedor principal con scroll
        self.scroll_container = mgc.create_main_container(self.root, padding=0)

        # Frame Principal
        main_frame = tk.Frame(self.scroll_container, bg=mgc.COLORS['bg_primary'])
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=15)
        
        # Header
        mgc.create_header(main_frame, "Buscador de Recibos PDF", 
                         "Separa y organiza recibos de sueldo individual o masivamente", 
                         icon_image=self.icon_search)

        # Contenedor Superior (PDFs Comunes)
        card_pdf_outer, card_pdf_inner = mgc.create_card(main_frame, "1. Archivos PDF Fuente (Recibos Mezclados)", padding=10)
        card_pdf_outer.pack(fill=tk.X, pady=(0, 10))

        # Lista de PDFs y Botones
        pdf_frame = tk.Frame(card_pdf_inner, bg=mgc.COLORS['bg_card'])
        pdf_frame.pack(fill=tk.X)
        
        self.listbox_archivos = tk.Listbox(pdf_frame, selectmode=tk.MULTIPLE, height=4, 
                                          font=('Consolas', 9), bg='white', relief=tk.GROOVE, bd=1)
        self.listbox_archivos.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        scrollbar = ttk.Scrollbar(pdf_frame, orient="vertical", command=self.listbox_archivos.yview)
        scrollbar.pack(side=tk.LEFT, fill=tk.Y)
        self.listbox_archivos.configure(yscrollcommand=scrollbar.set)

        btn_files_frame = tk.Frame(card_pdf_inner, bg=mgc.COLORS['bg_card'])
        btn_files_frame.pack(fill=tk.X, pady=(5, 0))
        
        mgc.create_button(btn_files_frame, "Añadir PDFs", self.seleccionar_archivos_pdf_gui, 
                         color='blue', icon_image=self.icon_folder, padx=10, pady=5).pack(side=tk.LEFT, padx=2)
        mgc.create_button(btn_files_frame, "Limpiar Lista", self.eliminar_archivos_seleccionados, 
                         color='red', icon_image=self.icon_warning, padx=10, pady=5).pack(side=tk.LEFT, padx=2)

        self.update_file_listbox()

        # 2. Filtro de Firmas (Nuevo)
        card_filter_outer, card_filter_inner = mgc.create_card(main_frame, "2. Filtro de Versión (Firma)", padding=10)
        card_filter_outer.pack(fill=tk.X, pady=(0, 10))
        
        filter_frame = tk.Frame(card_filter_inner, bg=mgc.COLORS['bg_card'])
        filter_frame.pack(fill=tk.X)
        
        opts = [("Todas (Ambas)", "ambas"), 
                ("Solo Firma Empleado", "empleado"), 
                ("Solo Firma Empleador", "empleador")]
        
        for text, val in opts:
            tk.Radiobutton(filter_frame, text=text, variable=self.var_filtro_firma, value=val,
                           bg=mgc.COLORS['bg_card'], font=mgc.FONTS['normal'],
                           activebackground=mgc.COLORS['bg_card'], cursor='hand2').pack(side=tk.LEFT, padx=15)

        # Tabs para Modos (Manual vs Lote)
        self.notebook = ttk.Notebook(main_frame)
        self.notebook.pack(fill=tk.BOTH, expand=True, pady=(10, 0))

        # Tab 1: Búsqueda Manual
        self.tab_manual = tk.Frame(self.notebook, bg=mgc.COLORS['bg_primary'])
        self.notebook.add(self.tab_manual, text="  👤 Búsqueda Manual (Individual)  ")
        self.setup_manual_tab()

        # Tab 2: Procesamiento por Lote
        self.tab_lote = tk.Frame(self.notebook, bg=mgc.COLORS['bg_primary'])
        self.notebook.add(self.tab_lote, text="  📑 Procesamiento por Lote (Excel)  ")
        self.setup_lote_tab()

        # Tab 3: Utilidades (Compresión) (Nueva)
        self.tab_util = tk.Frame(self.notebook, bg=mgc.COLORS['bg_primary'])
        self.notebook.add(self.tab_util, text="  ⚡ Optimizar PDF  ")
        self.setup_util_tab()

        # Log de Estado (Reducido para que quepa todo)
        card_log_outer, card_log_inner = mgc.create_card(main_frame, "Registro de Actividad", padding=10)
        card_log_outer.pack(side=tk.BOTTOM, fill=tk.X, pady=(10, 0)) # No expandir

        self.status_text = scrolledtext.ScrolledText(card_log_inner, height=4, font=('Consolas', 9), 
                                                    bg='#f8f9fa', state='disabled', relief=tk.FLAT)
        self.status_text.pack(fill=tk.BOTH, expand=True)

    def setup_manual_tab(self):
        container = tk.Frame(self.tab_manual, bg=mgc.COLORS['bg_primary'], padx=20, pady=20)
        container.pack(fill=tk.BOTH, expand=True)

        tk.Label(container, text="Ingrese el Nombre y/o Apellido del empleado a buscar:", 
                bg=mgc.COLORS['bg_primary'], fg=mgc.COLORS['text_primary'], font=mgc.FONTS['normal']).pack(anchor='w', pady=(0, 5))
        
        self.entry_nombre_manual = tk.Entry(container, font=mgc.FONTS['normal'], relief=tk.GROOVE, bd=1,
                                            bg=mgc.COLORS['bg_input'], fg=mgc.COLORS['text_primary'], insertbackground=mgc.COLORS['text_primary'])
        self.entry_nombre_manual.pack(fill=tk.X, pady=(0, 20))

        self.btn_manual = mgc.create_large_button(container, "BUSCAR Y GUARDAR PDF", 
                                                 self.procesar_manual, color='green', 
                                                 icon_image=self.icon_check)
        self.btn_manual.pack()

    def setup_lote_tab(self):
        container = tk.Frame(self.tab_lote, bg=mgc.COLORS['bg_primary'], padx=20, pady=10)
        container.pack(fill=tk.BOTH, expand=True)

        # Selector de Archivo Excel
        tk.Label(container, text="Seleccione el archivo Excel índice (Nombres en Columna B):", 
                bg=mgc.COLORS['bg_primary'], fg=mgc.COLORS['text_primary'], font=mgc.FONTS['normal']).pack(anchor='w', pady=(0, 5))

        excel_frame = tk.Frame(container, bg=mgc.COLORS['bg_primary'])
        excel_frame.pack(fill=tk.X, pady=(0, 15))

        self.var_excel_path = tk.StringVar()
        entry_excel = tk.Entry(excel_frame, textvariable=self.var_excel_path, state='readonly', 
                              font=mgc.FONTS['normal'], relief=tk.GROOVE, bd=1,
                              readonlybackground=mgc.COLORS['bg_input'], fg=mgc.COLORS['text_primary'])
        entry_excel.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))

        mgc.create_button(excel_frame, "Examinar...", self.seleccionar_indice, 
                         color='purple', icon_image=self.icon_excel, padx=15, pady=5).pack(side=tk.LEFT)

        # Opción de Optimización Automática
        tk.Checkbutton(container, text="Optimizar tamaño automáticamente (Sin pérdida de calidad)", 
                       variable=self.var_optimizar_auto, bg=mgc.COLORS['bg_primary'], 
                       fg=mgc.COLORS['text_primary'], font=mgc.FONTS['normal'], 
                       activebackground=mgc.COLORS['bg_primary'], selectcolor=mgc.COLORS['bg_input']).pack(anchor='w', pady=(0, 15))

        # Botón de Procesar
        self.btn_lote = mgc.create_large_button(container, "GENERAR ZIP CON RECIBOS", 
                                               self.procesar_lote, color='blue', 
                                               icon_image=self.icon_check)
        self.btn_lote.pack(pady=(0, 10))

        # Sección de Progreso (Nueva)
        self.progress_container, self.progress_bar, self.progress_label_var = mgc.create_progress_section(container)
        self.progress_container.pack(fill=tk.X)

    def setup_util_tab(self):
        container = tk.Frame(self.tab_util, bg=mgc.COLORS['bg_primary'], padx=20, pady=20)
        container.pack(fill=tk.BOTH, expand=True)

        tk.Label(container, text="Reduce el tamaño de cualquier PDF optimizando su estructura interna:", 
                bg=mgc.COLORS['bg_primary'], fg=mgc.COLORS['text_primary'], font=mgc.FONTS['normal']).pack(anchor='w', pady=(0, 10))

        # Selector de PDF a comprimir
        self.var_pdf_to_opt = tk.StringVar()
        opt_frame = tk.Frame(container, bg=mgc.COLORS['bg_primary'])
        opt_frame.pack(fill=tk.X, pady=(0, 20))

        tk.Entry(opt_frame, textvariable=self.var_pdf_to_opt, state='readonly', 
                 font=mgc.FONTS['normal'], relief=tk.GROOVE, bd=1,
                 readonlybackground=mgc.COLORS['bg_input'], fg=mgc.COLORS['text_primary']).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))

        mgc.create_button(opt_frame, "Seleccionar PDF...", self._seleccionar_pdf_para_optimizar, 
                         color='purple', icon_image=self.icon_folder, padx=15, pady=5).pack(side=tk.LEFT)

        self.btn_opt = mgc.create_large_button(container, "OPTIMIZAR Y GUARDAR COMO...", 
                                              self._ejecutar_optimizacion_manual, color='orange', 
                                              icon_image=self.icon_check)
        self.btn_opt.pack()
        
        info_label = tk.Label(container, text="Nota: Esta optimización es 'sin pérdida'. Elimina datos redundantes y\nreorganiza el archivo para que pese menos sin afectar la calidad visual.", 
                             bg=mgc.COLORS['bg_primary'], font=mgc.FONTS['small'], fg=mgc.COLORS['text_secondary'], justify=tk.LEFT)
        info_label.pack(pady=20)

    def _seleccionar_pdf_para_optimizar(self):
        file = filedialog.askopenfilename(title="Seleccionar PDF para optimizar", filetypes=[("Archivo PDF", "*.pdf")])
        if file:
            self.var_pdf_to_opt.set(file)

    def _ejecutar_optimizacion_manual(self):
        input_path = self.var_pdf_to_opt.get()
        if not input_path:
            return messagebox.showwarning("Falta Archivo", "Seleccione un PDF primero.")
        
        output_path = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("Archivo PDF", "*.pdf")],
            initialfile=os.path.basename(input_path).replace(".pdf", "_reducido.pdf"),
            title="Guardar PDF Optimizado"
        )
        
        if output_path:
            try:
                self.log(f"Optimizando: {os.path.basename(input_path)}...")
                self._reducir_pdf_internal(input_path, output_path)
                
                size_in = os.path.getsize(input_path) / (1024*1024)
                size_out = os.path.getsize(output_path) / (1024*1024)
                reduction = ((size_in - size_out) / size_in) * 100
                
                self.log(f"¡Éxito! Tamaño final: {size_out:.2f} MB (Reducción: {reduction:.1f}%)")
                messagebox.showinfo("Éxito", f"Archivo optimizado guardado.\nReducción: {reduction:.1f}%")
            except Exception as e:
                self.log(f"Error en optimización: {e}")
                messagebox.showerror("Error", str(e))

    def _reducir_pdf_internal(self, input_path_or_stream, output_path_or_stream):
        """Lógica central de optimización con pikepdf"""
        import pikepdf
        with pikepdf.open(input_path_or_stream) as pdf:
            pdf.save(
                output_path_or_stream, 
                linearize=True, 
                compress_streams=True,
                object_stream_mode=pikepdf.ObjectStreamMode.generate
            )

    # --- Funciones de Archivos ---

    def seleccionar_archivos_pdf_gui(self):
        nuevos = filedialog.askopenfilenames(title="Añadir PDFs de Recibos", filetypes=[("Archivos PDF", "*.pdf")])
        if nuevos:
            for f in nuevos:
                if f not in self.archivos_pdf_seleccionados:
                    self.archivos_pdf_seleccionados.append(f)
            self.update_file_listbox()
            self.log(f"Agregados {len(nuevos)} archivos PDF.")

    def eliminar_archivos_seleccionados(self):
        self.archivos_pdf_seleccionados = []
        self.pdf_data_cache = None # Limpiar cache si cambian los archivos
        self.update_file_listbox()
        self.log("Lista de PDFs limpiada.")

    def update_file_listbox(self):
        self.listbox_archivos.delete(0, tk.END)
        if not self.archivos_pdf_seleccionados:
            self.listbox_archivos.insert(tk.END, "Ningún PDF seleccionado...")
            self.listbox_archivos.itemconfig(0, fg='gray')
        else:
            for f in self.archivos_pdf_seleccionados:
                self.listbox_archivos.insert(tk.END, f"📄 {os.path.basename(f)}")

    def seleccionar_indice(self):
        path = filedialog.askopenfilename(title="Seleccionar Índice Excel", 
                                         filetypes=[("Excel Files", "*.xlsx;*.xls")])
        if path:
            self.var_excel_path.set(path)
            self.archivo_indice_path = path
            self.log(f"Índice seleccionado: {os.path.basename(path)}")

    # --- Hilos y Seguridad ---

    def safe_log(self, msj):
        """Escribe en el log de forma segura para hilos."""
        self.root.after(0, lambda: self.log(msj))

    def safe_update_status(self, msj):
        """Actualiza la barra de estado de forma segura."""
        self.root.after(0, lambda: self.status_var.set(msj))

    # --- Lógica de Procesamiento ---

    def log(self, msj):
        self.status_text.configure(state="normal")
        self.status_text.insert(tk.END, f"[{pd.Timestamp.now().strftime('%H:%M:%S')}] {msj}\n")
        self.status_text.see(tk.END)
        self.status_text.configure(state="disabled")
        self.root.update_idletasks()

    def buscar_paginas(self, nombre_busqueda):
        """Busca páginas coincidentes en todos los PDFs cargados usando 3 pasadas por página."""
        paginas_encontradas = []
        nombre_norm = normalize_text(nombre_busqueda)
        nombre_agresivo = nombre_norm.replace(" ", "")
        parts = nombre_norm.split()
        filtro = self.var_filtro_firma.get()
        variaciones = get_search_variations(nombre_norm)
        
        # 1. Obtener lista de páginas y texto (usar caché si existe)
        pdf_data_list = []
        if self.pdf_data_cache:
            for path, pages in self.pdf_data_cache.items():
                pdf_data_list.extend(pages)
        else:
            self.safe_update_status("Analizando PDFs...")
            for archivo_pdf in self.archivos_pdf_seleccionados:
                try:
                    doc = fitz.open(archivo_pdf)
                    for n_pag in range(len(doc)):
                        pdf_data_list.append((archivo_pdf, n_pag, normalize_text(doc[n_pag].get_text() or "")))
                    doc.close()
                except: pass
        
        # 2. Aplicar lógica de 3 pasadas
        for path, page_num, texto in pdf_data_list:
            if not check_signature(texto, filtro): continue
            
            encontrado = False
            # P1: Variaciones
            for v in variaciones:
                if v in texto:
                    encontrado = True
                    break
            # P2: Partes
            if not encontrado and len(parts) >= 2:
                if all(p in texto for p in parts[:2]):
                    encontrado = True
            # P3: Agresiva (sin espacios)
            if not encontrado:
                texto_sin_espacios = texto.replace(" ", "")
                if nombre_agresivo in texto_sin_espacios:
                    encontrado = True
                else:
                    for v in variaciones:
                        if v.replace(" ", "") in texto_sin_espacios:
                            encontrado = True
                            break
            
            if encontrado:
                paginas_encontradas.append((path, page_num))
        
        self.safe_update_status("Listo")
        return paginas_encontradas

    def procesar_manual(self):
        nombre = self.entry_nombre_manual.get().strip()
        if not nombre:
            return messagebox.showwarning("Falta Datos", "Ingrese un nombre para buscar.")
        if not self.archivos_pdf_seleccionados:
            return messagebox.showwarning("Falta Datos", "Seleccione al menos un PDF fuente.")

        threading.Thread(target=self._run_manual_thread, args=(nombre,), daemon=True).start()

    def _run_manual_thread(self, nombre):
        self.log(f"Iniciando búsqueda manual para: {nombre}")
        self.root.after(0, lambda: mgc.disable_button(self.btn_manual))
        
        try:
            paginas = self.buscar_paginas(nombre)
            
            if not paginas:
                self.root.after(0, lambda: messagebox.showinfo("Sin Resultados", f"No se encontraron recibos para '{nombre}'."))
                self.safe_log("Búsqueda finalizada sin resultados.")
            else:
                self.safe_log(f"¡Encontradas {len(paginas)} páginas! Solicitando guardado...")
                
                # Diálogos de guardado deben ser en el hilo principal
                def save_dialog():
                    output_path = filedialog.asksaveasfilename(
                        defaultextension=".pdf",
                        filetypes=[("Archivos PDF", "*.pdf")],
                        initialfile=f"Recibos_{nombre.replace(' ', '_')}.pdf",
                        title="Guardar PDF Unificado"
                    )
                    if output_path:
                        writer = fitz.open()
                        for p_path, p_num in paginas:
                            src_doc = fitz.open(p_path)
                            writer.insert_pdf(src_doc, from_page=p_num, to_page=p_num)
                            src_doc.close()
                        
                        pdf_bytes = writer.write()
                        writer.close()
                        temp_pdf = io.BytesIO(pdf_bytes)
                        temp_pdf.seek(0)
                        
                        if self.var_optimizar_auto.get():
                            self.log("Aplicando optimización sin pérdida...")
                            self._reducir_pdf_internal(temp_pdf, output_path)
                        else:
                            with open(output_path, 'wb') as f:
                                f.write(pdf_bytes)
                                
                        self.log(f"PDF Guardado exitosamente en: {output_path}")
                        messagebox.showinfo("Éxito", f"Archivo guardado correctamente.")
                    else:
                        self.log("Guardado cancelado por el usuario.")
                
                self.root.after(0, save_dialog)

        except Exception as e:
            self.safe_log(f"ERROR CRÍTICO: {e}")
            self.root.after(0, lambda: messagebox.showerror("Error", str(e)))
        finally:
            self.root.after(0, lambda: mgc.enable_button(self.btn_manual, 'green'))

    def procesar_lote(self):
        if not self.archivo_indice_path:
            return messagebox.showwarning("Falta Índice", "Seleccione un archivo Excel índice.")
        if not self.archivos_pdf_seleccionados:
            return messagebox.showwarning("Falta Fuente", "Seleccione al menos un PDF fuente.")

        # Preguntar ruta de guardado antes de empezar el hilo
        output_zip_path = filedialog.asksaveasfilename(
            defaultextension=".zip",
            filetypes=[("Archivo ZIP", "*.zip")],
            initialfile="Recibos_Procesados_Lote.zip",
            title="Guardar Lote Completo (Seleccione ubicación)"
        )
        
        if not output_zip_path:
            return

        threading.Thread(target=self._run_lote_thread, args=(output_zip_path,), daemon=True).start()

    def _run_lote_thread(self, output_zip_path):
        self.safe_log("Iniciando procesamiento por LOTE...")
        self.root.after(0, lambda: mgc.disable_button(self.btn_lote))
        
        files_handles = []
        try:
            # 1. Leer Nombres del Excel
            self.safe_log("Leyendo archivo índice...")
            df = pd.read_excel(self.archivo_indice_path, header=None)
            nombres_raw = df.iloc[:, 1].dropna().astype(str).tolist()
            nombres = [n.strip() for n in nombres_raw if n.strip() and n.lower() != 'nombre']
            self.safe_log(f"Se encontraron {len(nombres)} nombres.")
            
            # Configurar barra de progreso
            total_pasos = len(nombres)
            self.root.after(0, lambda: self.progress_bar.configure(maximum=total_pasos, value=0))
            self.root.after(0, lambda: self.progress_label_var.set("Preparando..."))

            # 2. Pre-cargar TEXTO (Fase de Caché con Progreso)
            if not self.pdf_data_cache:
                self.safe_log("Analizando PDFs para extracción de texto...")
                self.pdf_data_cache = {}
                
                # Contar páginas totales para la barra
                paginas_totales = 0
                for path in self.archivos_pdf_seleccionados:
                    try:
                        doc = fitz.open(path)
                        paginas_totales += len(doc)
                        doc.close()
                    except: pass
                
                self.root.after(0, lambda: [self.progress_bar.configure(maximum=paginas_totales, value=0),
                                            self.progress_label_var.set("Iniciando análisis de PDFs...")])
                
                progreso_paginas = 0
                for path in self.archivos_pdf_seleccionados:
                    try:
                        doc = fitz.open(path)
                        fname = os.path.basename(path)
                        self.safe_log(f"  -> Extrayendo texto: {fname}")
                        
                        paginas_este_pdf = []
                        for n_pag in range(len(doc)):
                            progreso_paginas += 1
                            self.root.after(0, lambda p=progreso_paginas: [self.progress_bar.configure(value=p),
                                                                            self.progress_label_var.set(f"Analizando página {p} de {paginas_totales}...")])
                            paginas_este_pdf.append((n_pag, normalize_text(doc[n_pag].get_text() or "")))
                        
                        self.pdf_data_cache[path] = paginas_este_pdf
                        doc.close()
                    except: pass
            else:
                self.safe_log("Usando caché de texto existente (Omitiendo extracción).")

            # Preparar la lista plana de datos para el proceso (Consolidado)
            pdf_data_list = []
            for path, pages in self.pdf_data_cache.items():
                for n_pag, txt in pages:
                    pdf_data_list.append((path, n_pag, txt))

            # 3. Procesar cada nombre (Fase de Búsqueda con Progreso)
            zip_buffer = io.BytesIO()
            archivos_generados = 0
            filtro = self.var_filtro_firma.get()
            self.safe_log(f"Generando archivos individuales (Filtro: {filtro})...")
            
            self.root.after(0, lambda: [self.progress_bar.configure(maximum=len(nombres), value=0),
                                        self.progress_label_var.set("Iniciando búsqueda de empleados...")])
            
            with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
                for i, nombre in enumerate(nombres):
                    self.safe_update_status(f"Procesando {i+1}/{len(nombres)}: {nombre}")
                    self.root.after(0, lambda idx=i: [self.progress_bar.configure(value=idx+1), 
                                                      self.progress_label_var.set(f"Empleado {idx+1} de {len(nombres)}: {nombre}")])
                    
                    nombre_norm = normalize_text(nombre)
                    nombre_agresivo = nombre_norm.replace(" ", "")
                    parts = nombre_norm.split()
                    variaciones = get_search_variations(nombre_norm)

                    writer = fitz.open()
                    paginas_agregadas = 0
                    pasadas_usadas = set()

                    for path, page_num, page_text in pdf_data_list:
                        if not check_signature(page_text, filtro): continue
                        
                        encontrado_en_pag = 0
                        # 1. Pasada Exacta
                        for v in variaciones:
                            if v in page_text:
                                encontrado_en_pag = 1
                                break
                        
                        # 2. Pasada por partes
                        if not encontrado_en_pag and len(parts) >= 2:
                            if parts[0] in page_text and parts[1] in page_text:
                                encontrado_en_pag = 2
                        
                        # 3. Pasada Agresiva
                        if not encontrado_en_pag:
                            texto_sin_espacios = page_text.replace(" ", "")
                            if nombre_agresivo in texto_sin_espacios:
                                encontrado_en_pag = 3
                            else:
                                for v in variaciones:
                                    if v.replace(" ", "") in texto_sin_espacios:
                                        encontrado_en_pag = 3
                                        break
                        
                        if encontrado_en_pag:
                            src_doc = fitz.open(path)
                            writer.insert_pdf(src_doc, from_page=page_num, to_page=page_num)
                            src_doc.close()
                            paginas_agregadas += 1
                            pasadas_usadas.add(encontrado_en_pag)
                    
                    if paginas_agregadas > 0:
                        pdf_bytes = writer.write()
                        writer.close()
                        
                        final_data = pdf_bytes
                        
                        # Optimizar si la opción está activa
                        if self.var_optimizar_auto.get():
                            opt_stream = io.BytesIO()
                            temp_in = io.BytesIO(pdf_bytes)
                            try:
                                self._reducir_pdf_internal(temp_in, opt_stream)
                                final_data = opt_stream.getvalue()
                            except: pass # Fallback al original si falla pikepdf
                        
                        pdf_filename = f"{nombre.replace(' ', '_')}.pdf"
                        zip_file.writestr(pdf_filename, final_data)
                        archivos_generados += 1
                        pasadas_str = ",".join(map(str, sorted(list(pasadas_usadas))))
                        self.safe_log(f"✔ Generado (Pasadas {pasadas_str}): {pdf_filename} ({paginas_agregadas} págs)")
            
            # 4. Guardar ZIP final
            if archivos_generados > 0:
                with open(output_zip_path, 'wb') as f:
                    f.write(zip_buffer.getvalue())
                self.safe_log(f"ARCHIVO ZIP GUARDADO: {output_zip_path}")
                self.root.after(0, lambda: messagebox.showinfo("Proceso Completado", f"Se generaron {archivos_generados} archivos en el ZIP."))
            else:
                self.safe_log("⚠ No se encontró ningún recibo coincidente.")
                self.root.after(0, lambda: messagebox.showwarning("Sin Resultados", "No se encontró ningún recibo coincidente."))

        except Exception as e:
            self.safe_log(f"ERROR EN LOTE: {e}")
            self.root.after(0, lambda: messagebox.showerror("Error", str(e)))
        finally:
            for f in files_handles:
                try: f.close()
                except: pass
            self.root.after(0, lambda: mgc.enable_button(self.btn_lote, 'blue'))
            self.root.after(0, lambda: self.progress_label_var.set("Completado"))
            self.safe_update_status("Listo")


if __name__ == "__main__":
    root = ctk.CTk()
    app = SeparadorRecibosApp(root)
    root.mainloop()
