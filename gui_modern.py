import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
import json
import os
import time
from datetime import datetime
import pandas as pd
from data_loader import load_structured_data, load_rate_config
from logic_payroll import process_payroll_for_employee
from logic_accountant import process_accountant_summary_for_employee
from excel_format_writer import write_payroll_to_excel, verify_output_file
import logic_cleaning
from receipt_font_formatter import apply_font_to_receipts
from icon_loader import set_window_icon, load_icon


class PayrollModernGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("💼 Sistema de Procesamiento de Nómina")
        
        # Dimensiones fijas (consistente con otros GUIs)
        window_width = 900
        window_height = 680
        
        # Centrar ventana
        screen_width = root.winfo_screenwidth()
        screen_height = root.winfo_screenheight()
        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2
        
        self.root.geometry(f"{window_width}x{window_height}+{x}+{y}")
        self.root.resizable(True, True)
        
        # Establecer icono de ventana
        set_window_icon(self.root, 'payroll')
        
        # Cargar iconos PNG
        self.icon_folder = load_icon('folder', (32, 32))
        self.icon_chart = load_icon('chart', (32, 32))
        self.icon_settings = load_icon('settings', (24, 24))
        self.icon_warning = load_icon('warning', (24, 24))
        self.icon_document = load_icon('document', (32, 32))
        self.icon_check = load_icon('check', (24, 24))
        self.icon_info = load_icon('info', (24, 24))
        self.icon_export = load_icon('export', (24, 24))
        
        # Variables
        self.processing = False
        self.input_file = None
        self.output_file = None
        self.start_time = None
        
        # Fuente para recibos
        self.receipt_font = tk.StringVar(value="Arial")
        self.font_options = ["Arial", "Calibri", "Courier New", "Microsoft Sans Serif"]
        
        # Paleta mejorada
        self.colors = {
            'bg': '#f5f7fa',
            'card_bg': '#ffffff',
            'primary': '#3b82f6',
            'primary_hover': '#2563eb',
            'success': '#10b981',
            'success_hover': '#059669',
            'purple': '#8b5cf6',
            'purple_hover': '#7c3aed',
            'orange': '#f59e0b',
            'red': '#ef4444',
            'red_hover': '#dc2626',
            'gray': '#6b7280',
            'gray_hover': '#4b5563',
            'text': '#1f2937',
            'text_secondary': '#6b7280',
            'border': '#e5e7eb'
        }
        
        self.root.configure(bg=self.colors['bg'])
        
        self.create_widgets()
        self.load_default_file()
    
    def create_card(self, parent, relief='raised'):
        """Crea tarjeta con estilo"""
        card = tk.Frame(parent, bg=self.colors['card_bg'], 
                       relief=relief, bd=1,
                       highlightbackground=self.colors['border'],
                       highlightthickness=1)
        return card
    
    def create_hover_button(self, parent, text, bg_color, hover_color, command=None, icon_image=None, **kwargs):
        """Crea botón con efectos hover y soporte para iconos PNG"""
        # Si hay icono PNG, usarlo
        if icon_image:
            btn = tk.Button(parent, text=text, image=icon_image, compound=tk.LEFT,
                           command=command,
                           bg=bg_color, fg='white',
                           font=kwargs.get('font', ('Segoe UI', 10, 'bold')),
                           relief='flat', bd=0,
                           padx=kwargs.get('padx', 20),
                           pady=kwargs.get('pady', 10),
                           cursor='hand2',
                           state=kwargs.get('state', 'normal'))
            btn.image = icon_image  # Mantener referencia
        else:
            btn = tk.Button(parent, text=text, command=command,
                           bg=bg_color, fg='white',
                           font=kwargs.get('font', ('Segoe UI', 10, 'bold')),
                           relief='flat', bd=0,
                           padx=kwargs.get('padx', 20),
                           pady=kwargs.get('pady', 10),
                           cursor='hand2',
                           state=kwargs.get('state', 'normal'))
        
        def on_enter(e):
            if btn['state'] == 'normal':
                btn.config(bg=hover_color)
        
        def on_leave(e):
            btn.config(bg=bg_color)
        
        btn.bind('<Enter>', on_enter)
        btn.bind('<Leave>', on_leave)
        
        return btn
    
    def create_widgets(self):
        # Frame principal con padding mínimo
        main_frame = tk.Frame(self.root, bg=self.colors['bg'])
        main_frame.pack(fill='both', expand=True, padx=10, pady=8)
        
        # Título compacto
        title = tk.Label(main_frame, 
                        text="💼 Sistema de Procesamiento de Nómina",
                        font=('Segoe UI', 14, 'bold'),
                        bg=self.colors['bg'],
                        fg=self.colors['text'])
        title.pack(pady=(0, 1))
        
        subtitle = tk.Label(main_frame,
                           text="Gestión automatizada de sueldos y horas",
                           font=('Segoe UI', 8),
                           bg=self.colors['bg'],
                           fg=self.colors['text_secondary'])
        subtitle.pack(pady=(0, 5))
        
        # Botones de acción (Mover aquí para anclar al fondo)
        self.create_action_buttons(main_frame)
        
        # Grid principal
        content_frame = tk.Frame(main_frame, bg=self.colors['bg'])
        content_frame.pack(fill='both', expand=True)
        
        # Panel izquierdo MUCHO más ancho, log MUY angosto
        content_frame.columnconfigure(0, weight=5)
        content_frame.columnconfigure(1, weight=2)
        # Fila superior más pequeña
        content_frame.rowconfigure(0, weight=3)
        content_frame.rowconfigure(1, weight=4)
        
        # === TARJETA 1: Archivo ===
        self.create_file_card(content_frame)
        
        # === TARJETA 2: Progreso ===
        self.create_progress_card(content_frame)
        
        # === TARJETA 3: Panel Izquierdo (Estadísticas + Herramientas) ===
        self.create_left_panel(content_frame)
        
        # === TARJETA 4: Log ===
        self.create_log_card(content_frame)
    
    def create_file_card(self, parent):
        card = self.create_card(parent)
        card.grid(row=0, column=0, sticky='nsew', padx=(0, 5), pady=(0, 5))
        
        inner = tk.Frame(card, bg=self.colors['card_bg'])
        inner.pack(fill='both', expand=True, padx=10, pady=10)
        
        # Icono pequeño
        icon = tk.Label(inner, image=self.icon_folder,
                       bg=self.colors['card_bg'])
        icon.image = self.icon_folder
        icon.pack(pady=(0, 5))
        
        title = tk.Label(inner, text="Archivo de Entrada",
                        font=('Segoe UI', 10, 'bold'),
                        bg=self.colors['card_bg'],
                        fg=self.colors['text'])
        title.pack()
        
        self.file_label = tk.Label(inner, text="Ningún archivo",
                                   font=('Segoe UI', 9),
                                   bg=self.colors['card_bg'],
                                   fg=self.colors['text_secondary'],
                                   wraplength=250)
        self.file_label.pack(pady=(5, 10))
        
        browse_btn = self.create_hover_button(
            inner, "📂 Examinar",
            self.colors['purple'], self.colors['purple_hover'],
            command=self.browse_file
        )
        browse_btn.pack()
    
    def create_progress_card(self, parent):
        card = self.create_card(parent)
        card.grid(row=0, column=1, sticky='nsew', pady=(0, 5))
        
        inner = tk.Frame(card, bg=self.colors['card_bg'])
        inner.pack(fill='both', expand=True, padx=10, pady=10)
        
        # Icono pequeño
        icon = tk.Label(inner, image=self.icon_chart,
                       bg=self.colors['card_bg'])
        icon.image = self.icon_chart
        icon.pack(pady=(0, 5))
        
        title = tk.Label(inner, text="Progreso",
                        font=('Segoe UI', 10, 'bold'),
                        bg=self.colors['card_bg'],
                        fg=self.colors['text'])
        title.pack()
        
        # Barra de progreso
        style = ttk.Style()
        style.theme_use('clam')
        style.configure("Modern.Horizontal.TProgressbar",
                       troughcolor=self.colors['border'],
                       background=self.colors['primary'],
                       bordercolor=self.colors['border'],
                       lightcolor=self.colors['primary'],
                       darkcolor=self.colors['primary'],
                       thickness=12)
        
        self.progress_bar = ttk.Progressbar(inner, length=300,
                                           mode='determinate',
                                           style="Modern.Horizontal.TProgressbar")
        self.progress_bar.pack(pady=(10, 5))
        
        self.progress_percent = tk.Label(inner, text="0%",
                                         font=('Segoe UI', 12, 'bold'),
                                         bg=self.colors['card_bg'],
                                         fg=self.colors['primary'])
        self.progress_percent.pack()
        
        self.progress_label = tk.Label(inner, text="Esperando...",
                                       font=('Segoe UI', 9),
                                       bg=self.colors['card_bg'],
                                       fg=self.colors['text_secondary'],
                                       wraplength=280)
        self.progress_label.pack(pady=(5, 0))
        
        self.time_label = tk.Label(inner, text="",
                                   font=('Segoe UI', 8),
                                   bg=self.colors['card_bg'],
                                   fg=self.colors['text_secondary'])
        self.time_label.pack()
    
    def create_left_panel(self, parent):
        card = self.create_card(parent)
        card.grid(row=1, column=0, sticky='nsew', padx=(0, 5))
        
        inner = tk.Frame(card, bg=self.colors['card_bg'])
        inner.pack(fill='both', expand=True, padx=10, pady=10)
        
        # === Estadísticas ===
        stats_header = tk.Frame(inner, bg=self.colors['card_bg'])
        stats_header.pack(fill='x', pady=(0, 5))
        
        icon_stats = tk.Label(stats_header, image=self.icon_chart,
                             bg=self.colors['card_bg'])
        icon_stats.image = self.icon_chart
        icon_stats.pack(side='left', padx=(0, 5))
        
        stats_info = tk.Frame(stats_header, bg=self.colors['card_bg'])
        stats_info.pack(side='left', fill='both', expand=True)
        
        stats_title = tk.Label(stats_info, text="Estadísticas",
                              font=('Segoe UI', 9, 'bold'),
                              bg=self.colors['card_bg'],
                              fg=self.colors['text'])
        stats_title.pack(anchor='w')
        
        self.stats_label = tk.Label(stats_info, text="👥 0 emp. | 💰 $0.00",
                                    font=('Segoe UI', 9, 'bold'),
                                    bg=self.colors['card_bg'],
                                    fg=self.colors['success'])
        self.stats_label.pack(anchor='w')
        
        # Separador
        sep1 = tk.Frame(inner, height=1, bg=self.colors['border'])
        sep1.pack(fill='x', pady=5)
        
        # === Fuente Recibos ===
        font_header = tk.Frame(inner, bg=self.colors['card_bg'])
        font_header.pack(fill='x', pady=(0, 3))
        
        icon_font = tk.Label(font_header, image=self.icon_settings,
                            bg=self.colors['card_bg'])
        icon_font.image = self.icon_settings
        icon_font.pack(side='left', padx=(0, 5))
        
        font_title = tk.Label(font_header, text="Fuente Recibos",
                             font=('Segoe UI', 9, 'bold'),
                             bg=self.colors['card_bg'],
                             fg=self.colors['text'])
        font_title.pack(side='left')
        
        font_dropdown = ttk.Combobox(inner, textvariable=self.receipt_font,
                                    values=self.font_options, state='readonly',
                                    font=('Segoe UI', 8))
        font_dropdown.pack(fill='x', pady=(3, 5))
        
        # Separador
        sep2 = tk.Frame(inner, height=1, bg=self.colors['border'])
        sep2.pack(fill='x', pady=5)
        
        # === Herramientas ===
        tools_header = tk.Frame(inner, bg=self.colors['card_bg'])
        tools_header.pack(fill='x', pady=(0, 5))
        
        icon_tools = tk.Label(tools_header, image=self.icon_warning,
                             bg=self.colors['card_bg'])
        icon_tools.image = self.icon_warning
        icon_tools.pack(side='left', padx=(0, 5))
        
        tools_title = tk.Label(tools_header, text="Herramientas",
                              font=('Segoe UI', 10, 'bold'),
                              bg=self.colors['card_bg'],
                              fg=self.colors['text'])
        tools_title.pack(side='left')
        
        # Botón de borrado general (más pequeño)
        gen_btn = tk.Button(tools_header, image=self.icon_warning,
                           command=self.run_general_cleanup,
                           bg='#dc2626', fg='white',
                           relief='flat', bd=0,
                           padx=8, pady=2,
                           cursor='hand2')
        gen_btn.image = self.icon_warning
        gen_btn.bind('<Enter>', lambda e: gen_btn.config(bg='#b91c1c'))
        gen_btn.bind('<Leave>', lambda e: gen_btn.config(bg='#dc2626'))
        gen_btn.pack(side='right')
        
        # Botones de herramientas en grid 2x2
        tools_grid = tk.Frame(inner, bg=self.colors['card_bg'])
        tools_grid.pack(fill='x', pady=(5, 0))
        
        # Fila 1
        row1 = tk.Frame(tools_grid, bg=self.colors['card_bg'])
        row1.pack(fill='x', pady=(0, 3))
        
        btn1 = tk.Button(row1, text="Contador",
                        command=self.run_clean_contador,
                        font=('Segoe UI', 8),
                        bg=self.colors['red'], fg='white',
                        relief='flat', bd=0, padx=3, pady=4,
                        cursor='hand2')
        btn1.pack(side='left', fill='x', expand=True, padx=(0, 2))
        
        btn2 = tk.Button(row1, text="Recuento",
                        command=self.run_clean_recuento,
                        font=('Segoe UI', 8),
                        bg=self.colors['red'], fg='white',
                        relief='flat', bd=0, padx=3, pady=4,
                        cursor='hand2')
        btn2.pack(side='left', fill='x', expand=True)
        
        # Fila 2
        row2 = tk.Frame(tools_grid, bg=self.colors['card_bg'])
        row2.pack(fill='x')
        
        btn3 = tk.Button(row2, text="Imprimir",
                        command=self.run_clean_imprimir,
                        font=('Segoe UI', 8),
                        bg=self.colors['red'], fg='white',
                        relief='flat', bd=0, padx=3, pady=4,
                        cursor='hand2')
        btn3.pack(side='left', fill='x', expand=True, padx=(0, 2))
        
        btn4 = tk.Button(row2, text="Valores",
                        command=self.run_clean_values,
                        font=('Segoe UI', 8),
                        bg=self.colors['red'], fg='white',
                        relief='flat', bd=0, padx=3, pady=4,
                        cursor='hand2')
        btn4.pack(side='left', fill='x', expand=True)
    
    def create_log_card(self, parent):
        card = self.create_card(parent)
        card.grid(row=1, column=1, sticky='nsew')
        
        inner = tk.Frame(card, bg=self.colors['card_bg'])
        inner.pack(fill='both', expand=True, padx=10, pady=10)
        
        # Icono
        icon = tk.Label(inner, image=self.icon_document,
                       bg=self.colors['card_bg'])
        icon.image = self.icon_document
        icon.pack(pady=(0, 5))
        
        title = tk.Label(inner, text="Log de Actividad",
                        font=('Segoe UI', 10, 'bold'),
                        bg=self.colors['card_bg'],
                        fg=self.colors['text'])
        title.pack()
        
        # Log text
        log_frame = tk.Frame(inner, bg='#f9fafb', bd=1,
                            relief='solid',
                            highlightbackground=self.colors['border'])
        log_frame.pack(fill='both', expand=True, pady=(10, 0))
        
        scrollbar = tk.Scrollbar(log_frame)
        scrollbar.pack(side='right', fill='y')
        
        # Reducir altura del log
        self.log_text = tk.Text(log_frame, height=5,
                               yscrollcommand=scrollbar.set,
                               font=('Consolas', 9),
                               bg='#f9fafb',
                               fg=self.colors['text'],
                               relief='flat',
                               padx=10, pady=10,
                               wrap='word')
        self.log_text.pack(side='left', fill='both', expand=True)
        scrollbar.config(command=self.log_text.yview)
        
        # Tags para colores
        self.log_text.tag_config('success', foreground=self.colors['success'])
        self.log_text.tag_config('error', foreground=self.colors['red'])
        self.log_text.tag_config('warning', foreground=self.colors['orange'])
        self.log_text.tag_config('info', foreground=self.colors['primary'])
        self.log_text.tag_config('normal', foreground=self.colors['text_secondary'])
    
    def create_action_buttons(self, parent):
        button_frame = tk.Frame(parent, bg=self.colors['bg'])
        button_frame.pack(side='bottom', fill='x', pady=(8, 0))
        
        center_frame = tk.Frame(button_frame, bg=self.colors['bg'])
        center_frame.pack()
        
        self.process_btn = self.create_hover_button(
            center_frame, "PROCESAR NÓMINA",
            self.colors['success'], self.colors['success_hover'],
            command=self.start_process,
            font=('Segoe UI', 12, 'bold'),
            padx=40, pady=15,
            icon_image=self.icon_check
        )
        self.process_btn.pack(side='left', padx=5)
        
        clear_btn = self.create_hover_button(
            center_frame, "Limpiar Log",
            self.colors['gray'], self.colors['gray_hover'],
            command=self.clear_log,
            font=('Segoe UI', 11),
            padx=25, pady=15,
            icon_image=self.icon_info
        )
        clear_btn.pack(side='left', padx=5)
        
        self.open_btn = self.create_hover_button(
            center_frame, "Abrir Resultado",
            self.colors['primary'], self.colors['primary_hover'],
            command=self.open_output,
            font=('Segoe UI', 11),
            padx=25, pady=15,
            state='disabled',
            icon_image=self.icon_export
        )
        self.open_btn.pack(side='left', padx=5)
        
        del_btn = self.create_hover_button(
            center_frame, "Borrar Modificado",
            self.colors['red'], self.colors['red_hover'],
            command=self.delete_output_file,
            font=('Segoe UI', 11),
            padx=25, pady=15,
            icon_image=self.icon_warning
        )
        del_btn.pack(side='left', padx=5)
    
    # ===== FUNCIONES DE LÓGICA (mantener todas las existentes) =====
    
    def load_default_file(self):
        default_file = None
        for file in os.listdir('.'):
            if file.endswith('.xlsm'):
                default_file = os.path.abspath(file)
                self.log(f"ℹ️ Archivo local encontrado: {os.path.basename(default_file)}", 'info')
                break
        
        if not default_file:
            self.log("ℹ️ No se encontró archivo .xlsm local, buscando en config.json...", 'info')
            try:
                with open('config.json', 'r', encoding='utf-8') as f:
                    config = json.load(f)
                default_file = config['file_paths']['input_excel']
            except Exception as e:
                self.log(f"⚠ No se pudo cargar el archivo desde config.json: {e}", 'warning')
                return
        
        if default_file and os.path.exists(default_file):
            self.input_file = default_file
            filename = os.path.basename(default_file)
            self.file_label.config(text=f"✓ {filename}", fg=self.colors['success'])
            self.log(f"✓ Archivo por defecto cargado: {filename}", 'success')
        else:
            self.log("⚠ No se pudo encontrar o cargar el archivo por defecto.", 'warning')
    
    def browse_file(self):
        filename = filedialog.askopenfilename(
            title="Seleccionar archivo Excel",
            filetypes=[("Excel files", "*.xlsx *.xlsm"), ("All files", "*.*")]
        )
        if filename:
            self.input_file = filename
            basename = os.path.basename(filename)
            self.file_label.config(text=f"✓ {basename}", fg=self.colors['success'])
            self.log(f"✓ Archivo seleccionado: {basename}", 'success')
    
    def log(self, message, tag='normal'):
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.insert(tk.END, f"[{timestamp}] {message}\n", tag)
        self.log_text.see(tk.END)
        self.root.update()
    
    def clear_log(self):
        self.log_text.delete(1.0, tk.END)
        self.log("🗑 Log limpiado", 'info')
    
    def update_progress(self, percent, message):
        self.progress_bar['value'] = percent
        self.progress_percent.config(text=f"{int(percent)}%")
        self.progress_label.config(text=message)
        
        if self.start_time and percent > 0:
            elapsed = time.time() - self.start_time
            if percent < 100:
                total_estimated = (elapsed / percent) * 100
                remaining = total_estimated - elapsed
                self.time_label.config(text=f"⏱ {int(remaining)}s restantes")
            else:
                self.time_label.config(text=f"✓ Completado en {int(elapsed)}s")
        
        self.root.update()
    
    def start_process(self):
        if self.processing:
            return
        
        if not self.input_file or not os.path.exists(self.input_file):
            messagebox.showerror("Error", "Por favor seleccione un archivo de entrada válido")
            return
        
        thread = threading.Thread(target=self.process)
        thread.daemon = True
        thread.start()
    
    def process(self):
        self.processing = True
        self.start_time = time.time()
        self.process_btn.config(state='disabled', bg='#9ca3af')
        self.open_btn.config(state='disabled')
        
        try:
            self.log("="*50, 'info')
            self.log("🚀 INICIANDO PROCESAMIENTO", 'info')
            self.log("="*50, 'info')
            
            self.update_progress(5, "Cargando configuración...")
            
            with open('config.json', 'r', encoding='utf-8') as f:
                config = json.load(f)
            self.log("✓ Configuración cargada", 'success')
            
            self.update_progress(10, "Cargando datos...")
            
            header_row = 8
            data_start_row = 9
            holiday_row = 7
            sheet_name = 'CALCULAR HORAS'
            
            employee_df, day_definitions = load_structured_data(
                self.input_file, sheet_name, data_start_row, header_row, holiday_row
            )
            
            if employee_df is None:
                raise Exception("Error al cargar datos")
            
            total_employees = len(employee_df)
            self.log(f"✓ {total_employees} empleados cargados", 'success')
            
            self.update_progress(20, "Cargando tarifas...")
            
            rate_config = load_rate_config(self.input_file)
            if rate_config is None:
                rate_config = {}
            self.log("✓ Tarifas cargadas", 'success')
            
            self.update_progress(25, "Abriendo Excel...")
            
            import openpyxl
            wb_cache = openpyxl.load_workbook(self.input_file, data_only=True, keep_vba=True)
            wb_styles = openpyxl.load_workbook(self.input_file, data_only=False, keep_vba=True)
            self.log("✓ Excel cargado (Cache + Estilos)", 'success')
            
            payroll_results = []
            accountant_results = []
            
            for index, employee_row in employee_df.iterrows():
                employee_data = employee_row.to_dict()
                
                current_excel_row = int(employee_data.get('Excel_Row_Index', data_start_row + index))
                
                if pd.isna(employee_data.get('NOMBRE Y APELLIDO')) or \
                   str(employee_data.get('NOMBRE Y APELLIDO')).strip() == '':
                    continue
                
                progress = 25 + (55 * (index + 1) / total_employees)
                employee_name = employee_data.get('NOMBRE Y APELLIDO', 'Desconocido')
                self.update_progress(progress, f"Procesando: {employee_name[:20]}...")
                
                payroll_result = process_payroll_for_employee(
                    employee_data, config, rate_config, day_definitions, self.input_file, wb_cache
                )
                payroll_result['Excel_Row_Index'] = current_excel_row
                payroll_results.append(payroll_result)
                
                accountant_result = process_accountant_summary_for_employee(
                    employee_data, config, day_definitions, rate_config,
                    wb_styles=wb_styles, row_idx=current_excel_row
                )
                accountant_result['Excel_Row_Index'] = current_excel_row
                accountant_results.append(accountant_result)
                
                if (index + 1) % 10 == 0:
                    self.log(f"  ✓ {index + 1}/{total_employees} procesados", 'normal')
            
            wb_cache.close()
            wb_styles.close()
            self.log(f"✓ {len(payroll_results)} empleados procesados", 'success')
            
            total_amount = sum(r['TOTAL CALCULADO'] for r in payroll_results)
            self.stats_label.config(
                text=f"👥 {len(payroll_results)} empleados | 💰 ${total_amount:,.2f}"
            )
            
            self.update_progress(85, "Escribiendo resultados...")
            
            input_dir = os.path.dirname(self.input_file)
            input_filename = os.path.basename(self.input_file)
            input_name, input_ext = os.path.splitext(input_filename)
            output_file = os.path.join(input_dir, f"{input_name}_MODIFICADO{input_ext}")
            
            success = write_payroll_to_excel(
                self.input_file, output_file, payroll_results, accountant_results
            )
            
            if not success:
                raise Exception("Falló la escritura del archivo de salida. Verifique si el archivo está dañado o protegido.")

            if success:
                self.output_file = output_file
                self.log(f"✓ Archivo generado: {os.path.basename(output_file)}", 'success')
                verify_output_file(output_file)
                
                selected_font = self.receipt_font.get()
                self.update_progress(90, f"Aplicando fuente {selected_font}...")
                font_applied = apply_font_to_receipts(output_file, selected_font)
                if font_applied:
                    self.log(f"✓ Fuente '{selected_font}' aplicada a recibos", 'success')
                else:
                    self.log(f"⚠ No se pudo aplicar la fuente", 'warning')
            
            self.update_progress(100, "✓ Completado")
            
            self.log("="*50, 'success')
            self.log("✓✓✓ PROCESO COMPLETADO ✓✓✓", 'success')
            self.log("="*50, 'success')
            
            self.open_btn.config(state='normal')
            
            messagebox.showinfo("Éxito", 
                              f"✓ Procesamiento completado\n\n"
                              f"Empleados: {len(payroll_results)}\n"
                              f"Total: ${total_amount:,.2f}")
        
        except PermissionError as e:
            self.log(f"🔒 ARCHIVO BLOQUEADO: {str(e)}", 'warning')
            self.update_progress(0, "🔒 Archivo Abierto")
            messagebox.showwarning("Archivo Bloqueado", 
                                 f"⚠️ El archivo Excel está abierto en otro programa.\n\n"
                                 f"Por favor, cierre el archivo y vuelva a intentarlo.\n\n"
                                 f"Detalle: {str(e)}")
        except Exception as e:
            self.log(f"✗ ERROR: {str(e)}", 'error')
            self.update_progress(0, "❌ Error")
            messagebox.showerror("Error", f"Error:\n\n{str(e)}")
        
        finally:
            self.processing = False
            self.process_btn.config(state='normal', bg=self.colors['success'])
    
    def open_output(self):
        try:
            if self.output_file and os.path.exists(self.output_file):
                os.startfile(self.output_file)
            else:
                messagebox.showwarning("Advertencia", "El archivo no existe o no se ha procesado aún")
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo abrir:\n{str(e)}")
    
    def delete_output_file(self):
        target_file = self.output_file
        
        if not target_file and self.input_file:
            input_dir = os.path.dirname(self.input_file)
            input_filename = os.path.basename(self.input_file)
            input_name, input_ext = os.path.splitext(input_filename)
            target_file = os.path.join(input_dir, f"{input_name}_MODIFICADO{input_ext}")
            self.log(f"ℹ️ Buscando archivo: {target_file}", 'info')
        
        if not target_file:
            messagebox.showwarning("Advertencia", "No se ha identificado ningún archivo modificado para borrar.")
            return
        
        if not os.path.exists(target_file):
            self.log(f"⚠ Archivo no encontrado: {target_file}", 'warning')
            messagebox.showinfo("Información", f"El archivo ya no existe:\n{os.path.basename(target_file)}")
            return
        
        self.log(f"📁 Archivo encontrado: {os.path.basename(target_file)}", 'info')
        
        if not messagebox.askyesno("Confirmar Borrado", 
                                 f"¿Está seguro de que desea ELIMINAR PERMANENTEMENTE el archivo generado?\n\n"
                                 f"Archivo: {os.path.basename(target_file)}\n"
                                 f"Ruta: {target_file}"):
            return
        
        try:
            os.remove(target_file)
            self.log(f"🗑 Archivo eliminado: {os.path.basename(target_file)}", 'success')
            messagebox.showinfo("Éxito", "Archivo eliminado correctamente.")
            self.output_file = None
            self.open_btn.config(state='disabled')
        except PermissionError:
            self.log(f"✗ Error: El archivo está abierto en otra aplicación", 'error')
            messagebox.showerror("Error", f"No se pudo eliminar el archivo.\n\nEl archivo está abierto en Excel u otra aplicación.\nCierre el archivo e intente nuevamente.")
        except Exception as e:
            self.log(f"✗ Error al eliminar archivo: {str(e)}", 'error')
            messagebox.showerror("Error", f"No se pudo eliminar el archivo.\nAsegúrese de que no esté abierto.\n\nError: {str(e)}")
    
    def run_cleaning_task(self, func, task_name):
        if not self.input_file or not os.path.exists(self.input_file):
            messagebox.showerror("Error", "Por favor seleccione un archivo primero.")
            return
        
        if not messagebox.askyesno("Confirmar", f"¿Está seguro de que desea ejecutar: {task_name}?\nEsto modificará el archivo: {os.path.basename(self.input_file)}"):
            return
        
        self.log(f"⏳ Iniciando: {task_name}...", 'info')
        
        def task():
            try:
                success, msg = func(self.input_file)
                if success:
                    self.log(f"✓ {msg}", 'success')
                    messagebox.showinfo("Éxito", msg)
                else:
                    self.log(f"✗ {msg}", 'error')
                    messagebox.showerror("Error", msg)
            except Exception as e:
                self.log(f"✗ Error crítico: {str(e)}", 'error')
                messagebox.showerror("Error Crítico", str(e))
        
        threading.Thread(target=task, daemon=True).start()
    
    def run_clean_contador(self):
        self.run_cleaning_task(logic_cleaning.borrar_envio_contador, "Limpiar Envío Contador")
    
    def run_clean_recuento(self):
        self.run_cleaning_task(logic_cleaning.vaciar_recuento_total, "Vaciar Recuento Total")
    
    def run_clean_imprimir(self):
        self.run_cleaning_task(logic_cleaning.vaciar_imprimir_totales, "Vaciar Imprimir Totales")
    
    def run_clean_values(self):
        self.run_cleaning_task(logic_cleaning.limpiar_valores_calcular_horas, "Limpiar Valores (CALCULAR HORAS)")
    
    def run_general_cleanup(self):
        if not self.input_file or not os.path.exists(self.input_file):
            messagebox.showerror("Error", "Por favor seleccione un archivo primero.")
            return
        
        if not messagebox.askyesno("Confirmar Borrado General", 
                                 f"⚠️ PELIGRO: ¿Está seguro de realizar un BORRADO GENERAL?\n\n"
                                 f"Esto ejecutará TODAS las tareas de limpieza en el archivo:\n{os.path.basename(self.input_file)}\n\n"
                                 "1. Limpiar 'ENVIO CONTADOR'\n"
                                 "2. Vaciar 'RECUENTO TOTAL'\n"
                                 "3. Vaciar 'IMPRIMIR TOTALES'\n"
                                 "4. Limpiar Valores (CALCULAR HORAS)"):
            return
        
        self.log(f"🔥 INICIANDO BORRADO GENERAL OPTIMIZADO...", 'warning')
        
        def task():
            try:
                success, msg = logic_cleaning.ejecutar_borrado_general_optimizado(self.input_file)
                
                if success:
                    self.log(f"✓ Operación completada", 'success')
                    for line in msg.split('\n'):
                        if line.strip():
                            self.log(f"  {line}", 'normal')
                    messagebox.showinfo("Éxito", msg)
                else:
                    self.log(f"✗ Error: {msg}", 'error')
                    messagebox.showerror("Error", msg)
            except Exception as e:
                self.log(f"✗ Error crítico: {str(e)}", 'error')
                messagebox.showerror("Error Crítico", str(e))
        
        threading.Thread(target=task, daemon=True).start()


if __name__ == '__main__':
    root = tk.Tk()
    app = PayrollModernGUI(root)
    root.mainloop()
