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

import customtkinter as ctk

class PayrollModernGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("💼 Sistema de Procesamiento de Nómina")
        self.root.geometry("900x700")
        self.root.resizable(True, True)
        
        import modern_gui_components as mgc
        self.mgc = mgc
        mgc.center_window(self.root, 900, 700)
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
        
        # Variables lógicas
        self.processing = False
        self.input_file = None
        self.output_file = None
        self.start_time = None
        
        # Variables GUI
        self.receipt_font = tk.StringVar(value="Arial")
        self.font_options = ["Arial", "Calibri", "Courier New", "Microsoft Sans Serif"]
        self.colors = self.mgc.COLORS # Alias de compatibilidad
        
        # UI
        self.create_widgets()
        self.load_default_file()
    
    def create_widgets(self):
        # Contenedor principal con scroll
        self.scroll_container = self.mgc.create_main_container(self.root, padding=0)

        # Frame general
        main_frame = ctk.CTkFrame(self.scroll_container, fg_color="transparent")
        main_frame.pack(fill='both', expand=True, padx=15, pady=15)
        
        # Header (Reusa el helper de mgc)
        self.mgc.create_header(main_frame, "Sistema de Procesamiento de Nómina", "Gestión automatizada de sueldos y horas", icon="💼")
        
        # Botones de acción inferiores
        self.create_action_buttons(main_frame)
        
        # Grid central para tarjetas
        content_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        content_frame.pack(fill='both', expand=True, pady=(10, 5))
        
        content_frame.columnconfigure(0, weight=5)
        content_frame.columnconfigure(1, weight=3)
        content_frame.rowconfigure(0, weight=3)
        content_frame.rowconfigure(1, weight=4)
        
        self.create_file_card(content_frame)
        self.create_progress_card(content_frame)
        self.create_left_panel(content_frame)
        self.create_log_card(content_frame)
        
    def create_file_card(self, parent):
        card_outer, inner = self.mgc.create_card(parent)
        card_outer.grid(row=0, column=0, sticky='nsew', padx=(0, 8), pady=(0, 8))
        
        icon = ctk.CTkLabel(inner, text="", image=self.icon_folder)
        icon.pack(pady=(0, 5))
        
        title = ctk.CTkLabel(inner, text="Archivo de Entrada", font=self.mgc.FONTS['heading'], text_color=self.colors['text_primary'])
        title.pack()
        
        self.file_label = ctk.CTkLabel(inner, text="Ningún archivo", font=self.mgc.FONTS['normal'], text_color=self.colors['text_secondary'])
        self.file_label.pack(pady=(5, 12))
        
        browse_btn = self.mgc.create_button(inner, "Examinar", self.browse_file, color='purple', icon="📂")
        browse_btn.pack()
        
    def create_progress_card(self, parent):
        card_outer, inner = self.mgc.create_card(parent)
        card_outer.grid(row=0, column=1, sticky='nsew', pady=(0, 8))
        
        icon = ctk.CTkLabel(inner, text="", image=self.icon_chart)
        icon.pack(pady=(0, 5))
        
        title = ctk.CTkLabel(inner, text="Progreso", font=self.mgc.FONTS['heading'], text_color=self.colors['text_primary'])
        title.pack()
        
        self.progress_bar = ctk.CTkProgressBar(inner, width=220, height=12, corner_radius=6, progress_color=self.colors['blue'])
        self.progress_bar.set(0) # CTkProgressBar usa .set(valor de 0 a 1)
        self.progress_bar.pack(pady=(12, 8))
        
        self.progress_percent = ctk.CTkLabel(inner, text="0%", font=self.mgc.FONTS['title'], text_color=self.colors['blue'])
        self.progress_percent.pack()
        
        self.progress_label = ctk.CTkLabel(inner, text="Esperando...", font=self.mgc.FONTS['normal'], text_color=self.colors['text_secondary'])
        self.progress_label.pack(pady=(2, 0))
        
        self.time_label = ctk.CTkLabel(inner, text="", font=self.mgc.FONTS['small'], text_color=self.colors['text_secondary'])
        self.time_label.pack()
        
    def create_left_panel(self, parent):
        card_outer, inner = self.mgc.create_card(parent)
        card_outer.grid(row=1, column=0, sticky='nsew', padx=(0, 8))
        
        # 1. Stats
        stats_frame = ctk.CTkFrame(inner, fg_color="transparent")
        stats_frame.pack(fill='x', pady=(0, 10))
        ctk.CTkLabel(stats_frame, text="", image=self.icon_chart).pack(side='left', padx=(0, 10))
        
        stats_info = ctk.CTkFrame(stats_frame, fg_color="transparent")
        stats_info.pack(side='left', fill='both', expand=True)
        ctk.CTkLabel(stats_info, text="Estadísticas", font=self.mgc.FONTS['heading'], text_color=self.colors['text_primary'], anchor='w').pack(fill='x')
        self.stats_label = ctk.CTkLabel(stats_info, text="👥 0 emp. | 💰 $0.00", font=self.mgc.FONTS['heading'], text_color=self.colors['green'], anchor='w')
        self.stats_label.pack(fill='x')
        
        # Separator
        ctk.CTkFrame(inner, height=1, fg_color=self.colors['border']).pack(fill='x', pady=8)
        
        # 2. Font
        font_frame = ctk.CTkFrame(inner, fg_color="transparent")
        font_frame.pack(fill='x', pady=8)
        ctk.CTkLabel(font_frame, text="", image=self.icon_settings).pack(side='left', padx=(0, 10))
        ctk.CTkLabel(font_frame, text="Fuente Recibos", font=self.mgc.FONTS['heading'], text_color=self.colors['text_primary']).pack(side='left')
        
        font_dropdown = ctk.CTkOptionMenu(inner, variable=self.receipt_font, values=self.font_options, fg_color=self.colors['bg_primary'], text_color=self.colors['text_primary'], button_color=self.colors['border'], font=self.mgc.FONTS['normal'])
        font_dropdown.pack(fill='x', pady=(2, 10))
        
        # Separator
        ctk.CTkFrame(inner, height=1, fg_color=self.colors['border']).pack(fill='x', pady=8)
        
        # 3. Herramientas
        tools_frame = ctk.CTkFrame(inner, fg_color="transparent")
        tools_frame.pack(fill='x', pady=(5, 10))
        ctk.CTkLabel(tools_frame, text="", image=self.icon_warning).pack(side='left', padx=(0, 10))
        ctk.CTkLabel(tools_frame, text="Herramientas", font=self.mgc.FONTS['heading'], text_color=self.colors['text_primary']).pack(side='left')
        
        gen_btn = self.mgc.create_button(tools_frame, "Borrado General", self.run_general_cleanup, color='red', icon="⚠️")
        gen_btn.pack(side='right')
        
        tools_grid = ctk.CTkFrame(inner, fg_color="transparent")
        tools_grid.pack(fill='x')
        tools_grid.columnconfigure((0,1), weight=1)
        
        self.mgc.create_button(tools_grid, "Envío Contador", self.run_clean_contador, color='orange').grid(row=0, column=0, padx=4, pady=4, sticky="we")
        self.mgc.create_button(tools_grid, "Recuento Total", self.run_clean_recuento, color='orange').grid(row=0, column=1, padx=4, pady=4, sticky="we")
        self.mgc.create_button(tools_grid, "Imprimir Totales", self.run_clean_imprimir, color='orange').grid(row=1, column=0, padx=4, pady=4, sticky="we")
        self.mgc.create_button(tools_grid, "Limpiar Valores", self.run_clean_values, color='orange').grid(row=1, column=1, padx=4, pady=4, sticky="we")
        
    def create_log_card(self, parent):
        card_outer, inner = self.mgc.create_card(parent)
        card_outer.grid(row=1, column=1, sticky='nsew')
        
        icon = ctk.CTkLabel(inner, text="", image=self.icon_document)
        icon.pack(pady=(0, 5))
        ctk.CTkLabel(inner, text="Log de Actividad", font=self.mgc.FONTS['heading'], text_color=self.colors['text_primary']).pack()
        
        self.log_text = ctk.CTkTextbox(inner, font=('Consolas', 11), fg_color='#f8fafc', text_color=self.colors['text_primary'], corner_radius=8, border_width=1, border_color=self.colors['border'])
        self.log_text.pack(fill='both', expand=True, pady=(10, 0))
        
        self.log_text.tag_config('success', foreground=self.colors['green'])
        self.log_text.tag_config('error', foreground=self.colors['red'])
        self.log_text.tag_config('warning', foreground=self.colors['orange'])
        self.log_text.tag_config('info', foreground=self.colors['blue'])
        self.log_text.tag_config('normal', foreground=self.colors['text_secondary'])
        
    def create_action_buttons(self, parent):
        # Frame en la parte inferior del TODO
        bottom_frame = ctk.CTkFrame(parent, fg_color="transparent")
        bottom_frame.pack(side='bottom', fill='x', pady=(15, 0))
        
        center = ctk.CTkFrame(bottom_frame, fg_color="transparent")
        center.pack()
        
        self.process_btn = self.mgc.create_large_button(center, "PROCESAR NÓMINA", self.start_process, color='green', icon="✓")
        self.process_btn.pack(side='left', padx=6)
        
        self.clear_btn = self.mgc.create_button(center, "Limpiar Log", self.clear_log, color='gray', icon="ℹ️")
        self.clear_btn.pack(side='left', padx=6)
        
        self.open_btn = self.mgc.create_button(center, "Abrir Resultado", self.open_output, color='blue', icon="📄")
        self.open_btn.pack(side='left', padx=6)
        
        # Corregido: .config() no existe en CustomTkinter, es .configure()
        self.mgc.disable_button(self.open_btn) 
        
        self.del_btn = self.mgc.create_button(center, "Borrar Modificado", self.delete_output_file, color='red', icon="🗑️")
        self.del_btn.pack(side='left', padx=6)    

    # ===== FUNCIONES DE LÓGICA =====
    
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
            # Corregido: fg -> text_color, .config() -> .configure()
            self.file_label.configure(text=f"✓ {filename}", text_color=self.colors['success'])
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
            # Corregido: fg -> text_color, .config() -> .configure()
            self.file_label.configure(text=f"✓ {basename}", text_color=self.colors['success'])
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
        # Corregido: CustomTkinter usa .set() con valores de 0.0 a 1.0
        self.progress_bar.set(percent / 100.0) 
        
        # Corregido: .config() -> .configure()
        self.progress_percent.configure(text=f"{int(percent)}%")
        self.progress_label.configure(text=message)
        
        # Blindaje para evitar el error si time_label no se instancia a tiempo
        if hasattr(self, 'time_label'):
            if self.start_time and percent > 0:
                elapsed = time.time() - self.start_time
                if percent < 100:
                    total_estimated = (elapsed / percent) * 100
                    remaining = total_estimated - elapsed
                    self.time_label.configure(text=f"⏱ {int(remaining)}s restantes")
                else:
                    self.time_label.configure(text=f"✓ Completado en {int(elapsed)}s")
            else:
                self.time_label.configure(text="")
                
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
        
        # Corregido: .config() -> .configure() y bg -> fg_color
        self.process_btn.configure(state='disabled', fg_color='#9ca3af')
        self.open_btn.configure(state='disabled')
        
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
            # Corregido: .config() -> .configure()
            self.stats_label.configure(
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
            
            # Corregido: .config() -> .configure()
            self.open_btn.configure(state='normal')
            
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
            # Corregido: .config() -> .configure() y bg -> fg_color
            self.process_btn.configure(state='normal', fg_color=self.colors['success'])
    
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
            # Corregido: .config() -> .configure()
            self.open_btn.configure(state='disabled')
        except PermissionError:
            self.log(f"✗ Error: El archivo está abierto en otra aplicación", 'error')
            messagebox.showerror("Error", f"No se pudo eliminar el archivo.\n\