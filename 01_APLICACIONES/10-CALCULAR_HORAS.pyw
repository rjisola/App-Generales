import sys
import tkinter as tk
from tkinter import ttk, messagebox
import os

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

import customtkinter as ctk
import modern_gui_components as mgc
from icon_loader import set_window_icon, load_icon

# Intentar importar la lógica de calculadoras existente
try:
    import calculators
except ImportError as e:
    calculators_error = str(e)
    calculators = None
else:
    calculators_error = None

# ==============================================================================
class ResultsWindow:
    """Ventana secundaria para mostrar resultados de cálculos con diseño de Reporte Ejecutivo."""
    
    # Paleta Premium (Modo Claro)
    COLOR_HABER_BG = "#dcfce7"       # Verde muy claro
    COLOR_HABER_FG = "#065f46"       # Verde oscuro
    COLOR_DEDUCCION_BG = "#fee2e2"   # Rojo muy claro
    COLOR_DEDUCCION_FG = "#991b1b"   # Rojo oscuro
    COLOR_BRUTO_BG = "#dbeafe"       # Azul muy claro
    COLOR_BRUTO_FG = "#1e40af"       # Azul oscuro
    COLOR_NETO_BG = "#fef3c7"        # Ambar muy claro
    COLOR_NETO_FG = "#92400e"        # Ambar oscuro
    
    def __init__(self, parent, title, result_dict, calc_type="UOCRA/NASA"):
        self.window = ctk.CTkToplevel(parent)
        self.window.title(f"📊 {title}")
        self.window.geometry("900x700")
        self.window.resizable(False, False)
        self.window.configure(fg_color=mgc.COLORS['bg_primary'])
        
        # Forzar que la ventana esté al frente
        self.window.attributes("-topmost", True)
        
        mgc.center_window(self.window, 900, 700)
        
        # Establecer icono de ventana
        set_window_icon(self.window, 'results')
        
        # Cargar iconos PNG
        self.icon_results = load_icon('results', (64, 64))
        self.icon_warning = load_icon('warning', (24, 24))
        
        # Contenedor principal con scroll
        main_frame = mgc.create_main_container(self.window)
        
        # Header de reporte
        mgc.create_header(main_frame, "Reporte de Haberes", f"Liquidación Estimada: {calc_type}", icon_image=self.icon_results)
        
        # Verificar si hay error
        if "error" in result_dict:
            card_outer, card_inner = mgc.create_card(main_frame, "⚠️ Error en el Cálculo", padding=20)
            card_outer.pack(fill=tk.BOTH, expand=True, pady=20)
            error_label = ctk.CTkLabel(card_inner, 
                                      text=result_dict['error'],
                                      text_color=mgc.COLORS['red'],
                                      font=mgc.FONTS['subtitle'],
                                      wraplength=800)
            error_label.pack(pady=40)
        else:
            # Metadata sutil
            meta_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
            meta_frame.pack(fill=tk.X, pady=(0, 10))
            
            val_ref = result_dict.get('tasa_horaria', result_dict.get('sueldo_quincenal', '0,00'))
            ctk.CTkLabel(meta_frame, text=f"UNIDAD REF: ${val_ref}  |  CONVENIO: {calc_type}", 
                         font=mgc.FONTS['tiny'], text_color=mgc.COLORS['text_secondary']).pack(side=tk.LEFT)
            
            # Determinar tipo
            is_uecara = "sueldo_quincenal" in result_dict
            
            # Layout de Columnas (Haberes vs Deducciones)
            cols_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
            cols_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
            
            if is_uecara:
                self._build_uecara_layout(cols_frame, result_dict)
            else:
                self._build_uocra_layout(cols_frame, result_dict)
            
            # Resumen Final (Bruto, NR, Deduc, NETO)
            self._build_summary(main_frame, result_dict, is_uecara)
        
        # Botones de acción
        btn_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        btn_frame.pack(fill=tk.X, pady=(10, 0))
        
        mgc.create_button(btn_frame, "VOLVER", self.window.destroy, color='gray', width=150).pack(side=tk.RIGHT, padx=5)
        
        # Barra de estado interna
        mgc.create_status_bar(self.window, "✓ Resultados generados con éxito")
    
    def _build_uocra_layout(self, parent, data):
        """Build items for UOCRA/NASA."""
        # Columna HABERES
        h_card, h_inner = mgc.create_card(parent, "✅ HABERES REMUNERATIVOS", padding=15)
        h_card.configure(fg_color=self.COLOR_HABER_BG)
        h_card.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 10))
        
        items_h = [
            ("Horas Normales", data.get("importe_normal")),
            ("Horas 50%", data.get("importe_50")),
            ("Horas 100%", data.get("importe_100")),
            ("Feriados", data.get("importe_feriado")),
            ("Presentismo", data.get("importe_presentismo")),
            ("Adicional", data.get("importe_adicional")),
            ("Vacaciones", data.get("importe_vacaciones")),
            ("Enfermedad / ART", data.get("importe_enfermedad") if data.get("importe_enfermedad") != "0,00" else data.get("importe_art")),
            ("Conv. Especial", data.get("importe_horas_conv_especial")),
            ("Bonos Rem.", data.get("importe_bono") if data.get("bono_tipo") == "Remunerativo" else "0,00")
        ]
        self._add_rows(h_inner, items_h, self.COLOR_HABER_FG)
        
        # Columna DEDUCCIONES
        d_card, d_inner = mgc.create_card(parent, "❌ DEDUCCIONES DE LEY", padding=15)
        d_card.configure(fg_color=self.COLOR_DEDUCCION_BG)
        d_card.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        items_d = [
            ("Jubilación (11%)", data.get("jubilacion")),
            ("L. 19032 / O. Social", data.get("ley19032")),
            ("Obra Social s/Bruto", data.get("obra_social_bruto")),
            ("Cuota Sindical", data.get("sindicato_bruto")),
            ("Seguro de Vida", data.get("seguro_vida")),
            ("Ret. Ganancias", data.get("ret_ganancias")),
            ("Ret. Judicial", data.get("importe_ret_judicial"))
        ]
        self._add_rows(d_inner, items_d, self.COLOR_DEDUCCION_FG)

    def _build_uecara_layout(self, parent, data):
        """Build items for UECARA."""
        h_card, h_inner = mgc.create_card(parent, "✅ HABERES MUESTRA", padding=15)
        h_card.configure(fg_color=self.COLOR_HABER_BG)
        h_card.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 10))
        
        items_h = [
            ("Sueldo Quincenal", data.get("sueldo_quincenal")),
            ("Antigüedad", data.get("importe_antiguedad")),
            ("Título Profesional", data.get("importe_titulo")),
            ("Presentismo (10%)", data.get("importe_presentismo")),
            ("Feriados", data.get("importe_feriados")),
            ("Descuento DNT", data.get("descuento_dnt"))
        ]
        self._add_rows(h_inner, items_h, self.COLOR_HABER_FG)
        
        d_card, d_inner = mgc.create_card(parent, "❌ DEDUCCIONES", padding=15)
        d_card.configure(fg_color=self.COLOR_DEDUCCION_BG)
        d_card.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        items_d = [
            ("Jubilación (11%)", data.get("jubilacion")),
            ("Obra Social", data.get("obra_social")),
            ("Cuota Sindical", data.get("sindicato")),
            ("Ret. Ganancias", data.get("ganancias_importe")),
            ("Ret. Judicial", data.get("importe_ret_judicial"))
        ]
        self._add_rows(d_inner, items_d, self.COLOR_DEDUCCION_FG)

    def _add_rows(self, parent, items, fg_color):
        for label, value in items:
            if not value or value == "0,00" or value == "-0,00": continue
            
            row = ctk.CTkFrame(parent, fg_color="transparent")
            row.pack(fill=tk.X, pady=1)
            
            ctk.CTkLabel(row, text=label, font=mgc.FONTS['small'], text_color=fg_color).pack(side=tk.LEFT)
            ctk.CTkLabel(row, text=f"${value}", font=mgc.FONTS['small'], text_color=fg_color).pack(side=tk.RIGHT)

    def _build_summary(self, parent, data, is_uecara):
        sum_card, sum_inner = mgc.create_card(parent, "", padding=10)
        sum_card.pack(fill=tk.X, pady=(10, 0))
        
        # Sub-totales
        bruto = data.get("bruto" if is_uecara else "total_bruto")
        self._create_summary_line(sum_inner, "💰 TOTAL BRUTO SUJETO A DESCUENTOS:", bruto, self.COLOR_BRUTO_FG, self.COLOR_BRUTO_BG)
        
        no_rem = data.get("haberes_sin_descuento", "0,00")
        if no_rem != "0,00":
            self._create_summary_line(sum_inner, "➕ CONCEPTOS NO REMUNERATIVOS / BONOS:", no_rem, "#fbbf24", "#451a03")
            
        deduc = data.get("total_deducciones")
        self._create_summary_line(sum_inner, "📉 TOTAL DEDUCCIONES APLICADAS:", deduc, self.COLOR_DEDUCCION_FG, self.COLOR_DEDUCCION_BG)
        
        # NETO FINAL - EL PROTAGONISTA
        neto_val = data.get("neto_redondeado" if is_uecara else "total_final")
        neto_frame = ctk.CTkFrame(sum_inner, fg_color=self.COLOR_NETO_BG, corner_radius=10, border_width=2, border_color=self.COLOR_NETO_FG)
        neto_frame.pack(fill=tk.X, pady=(10, 5))
        
        ctk.CTkLabel(neto_frame, text="💵 NETO FINAL A COBRAR:", font=mgc.FONTS['subtitle'], text_color=self.COLOR_NETO_FG).pack(side=tk.LEFT, padx=20, pady=15)
        ctk.CTkLabel(neto_frame, text=f"$ {neto_val}", font=('Segoe UI', 32, 'bold'), text_color=self.COLOR_NETO_FG).pack(side=tk.RIGHT, padx=20, pady=15)
        
        # Neto en letras
        ctk.CTkLabel(sum_inner, text=f"SON: {data.get('neto_en_letras', '').upper()}", 
                     font=mgc.FONTS['tiny'], text_color=mgc.COLORS['text_secondary']).pack(pady=(5, 0))

    def _create_summary_line(self, parent, label, value, fg, bg):
        line = ctk.CTkFrame(parent, fg_color=bg, corner_radius=6, height=35)
        line.pack(fill=tk.X, pady=2)
        line.pack_propagate(False)
        ctk.CTkLabel(line, text=label, font=mgc.FONTS['small'], text_color=fg).pack(side=tk.LEFT, padx=10)
        ctk.CTkLabel(line, text=f"${value}", font=mgc.FONTS['heading'], text_color=fg).pack(side=tk.RIGHT, padx=10)
# ==============================================================================
# VENTANA PRINCIPAL UNIFICADA CON PESTAÑAS
# ==============================================================================

class UnifiedCalculatorApp:
    """Aplicación unificada con pestañas para UOCRA/NASA y UECARA."""
    
    def __init__(self, root):
        self.root = root
        self.root.title("🧮 Calculadora de Sueldos")
        self.root.geometry("900x700")
        self.root.resizable(False, False)
        self.root.configure(bg=mgc.COLORS['bg_primary'])
        
        mgc.center_window(self.root, 900, 700)
        
        # Establecer icono de ventana
        set_window_icon(self.root, 'calculator')
        
        # Iconos
        self.icon_calculator = load_icon('calculator', (64, 64))
        self.icon_check = load_icon('check', (24, 24))
        self.icon_folder = load_icon('folder', (20, 20))
        
        # --- LÓGICA DE CARGA DE DATOS UNIVERSAL ---
        if not calculators:
            msg = f"No se pudo importar el módulo 'calculators.py'.\n\nDetalle: {calculators_error}" if calculators_error else "No se encontró el módulo 'calculators.py'."
            messagebox.showerror("Error Crítico", msg)
            self.root.destroy()
            return

        # Intentar carga automática (fallback tradicional)
        self.excel_data = calculators._load_excel_data()
        
        # Si falla la carga automática, pedir al usuario el archivo
        if not self.excel_data:
            messagebox.showinfo("Configuración Inicial", "No se encontró el archivo de tarifas por defecto.\nPor favor, seleccione el archivo '2-VALOR_HORAS_SUELDOS.xlsx'.")
            custom_path = filedialog.askopenfilename(
                title="Seleccionar Archivo de Tarifas (Excel)",
                filetypes=[("Excel", "*.xlsx")]
            )
            if custom_path:
                calculators.set_custom_excel_path(custom_path)
                self.excel_data = calculators._load_excel_data()
            
        if not self.excel_data:
            messagebox.showerror("Error", "No se pudieron cargar los datos de tarifas.\nLa aplicación no puede continuar sin esta información.")
            self.root.destroy()
            return
        
        # Contenedor principal con scroll
        main_frame = mgc.create_main_container(self.root)
        
        # Header moderno con iconos (estilo launcher)
        mgc.create_header(main_frame, "Calculadora de Sueldos", 
                         "Cálculo rápido de haberes para UOCRA, NASA y UECARA", 
                         icon_image=self.icon_calculator)
        
        # Separador inicial
        tk.Frame(main_frame, height=1, bg=mgc.COLORS['border']).pack(fill=tk.X, pady=(0, 10))
        
        # Notebook con pestañas — USANDO CTKTABVIEW PARA UN LOOK MODERNO
        self.tabview = ctk.CTkTabview(main_frame, 
                                     fg_color="transparent",
                                     segmented_button_fg_color=mgc.COLORS['bg_card'],
                                     segmented_button_selected_color=mgc.COLORS['blue'],
                                     segmented_button_selected_hover_color=mgc.COLORS['accent_blue'],
                                     segmented_button_unselected_color=mgc.COLORS['bg_card'],
                                     segmented_button_unselected_hover_color=mgc.COLORS['border'],
                                     text_color=mgc.COLORS['text_primary'])
        self.tabview.pack(fill=tk.BOTH, expand=True, pady=(0, 5))
        
        # Pestaña UOCRA/NASA
        self.tab_uocra = self.tabview.add("  🧮 UOCRA / NASA  ")
        self.setup_uocra_tab()
        
        # Pestaña UECARA
        self.tab_uecara = self.tabview.add("  💼 UECARA  ")
        self.setup_uecara_tab()
        
        # Barra de estado
        self.status_frame, self.status_var = mgc.create_status_bar(self.root, "Listo para calcular")
    
    def setup_uocra_tab(self):
        """Configura la pestaña UOCRA/NASA con diseño compacto."""
        # Card de inputs con padding reducido
        card_outer, card_inner = mgc.create_card(self.tab_uocra, "Datos de Entrada", padding=8)
        card_outer.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Grid de configuración
        f_grid = tk.Frame(card_inner, bg=mgc.COLORS['bg_card'])
        f_grid.pack(fill=tk.BOTH, expand=True)
        
        # Configurar columnas para mejor distribución
        for i in range(6):
            f_grid.columnconfigure(i, weight=1)
        
        # Fila 0: Convenio y Categoría
        mgc.ctk.CTkLabel(f_grid, text="Convenio:", 
                         font=mgc.FONTS['small'], text_color=mgc.COLORS['text_primary']).grid(row=0, column=0, sticky='w', padx=3, pady=2)
        self.cb_convenio = mgc.ctk.CTkComboBox(f_grid, values=["UOCRA", "QUILMES", "NASA"], state="readonly", width=120, font=mgc.FONTS['small'],
                                               command=self.update_categories_uocra)
        self.cb_convenio.set("UOCRA")
        self.cb_convenio.grid(row=0, column=1, sticky='w', padx=3, pady=2)
        
        mgc.ctk.CTkLabel(f_grid, text="Categoría:", 
                         font=mgc.FONTS['small'], text_color=mgc.COLORS['text_primary']).grid(row=0, column=2, sticky='w', padx=3, pady=2)
        self.cb_categoria = mgc.ctk.CTkComboBox(f_grid, state="readonly", width=250, font=mgc.FONTS['small'])
        self.cb_categoria.grid(row=0, column=3, columnspan=3, sticky='w', padx=3, pady=2)
        
        # Separador
        sep1 = tk.Frame(f_grid, bg=mgc.COLORS['purple'], height=1)
        sep1.grid(row=1, column=0, columnspan=6, sticky='ew', pady=(5, 3))
        
        self.entries_uocra = {}
        
        # Sección: Horas Extras
        lbl_extras = mgc.ctk.CTkLabel(f_grid, text="⏰ Horas Extras", 
                             font=mgc.FONTS['heading'], text_color=mgc.COLORS['purple'])
        lbl_extras.grid(row=2, column=0, columnspan=6, sticky='w', padx=3, pady=(2, 1))
        
        self.entries_uocra['horas_normales'] = self.create_compact_entry(f_grid, "Hs Normales:", 3, 0)
        self.entries_uocra['horas_50'] = self.create_compact_entry(f_grid, "Hs 50%:", 3, 2)
        self.entries_uocra['horas_100'] = self.create_compact_entry(f_grid, "Hs 100%:", 3, 4)
        
        self.entries_uocra['horas_feriado'] = self.create_compact_entry(f_grid, "Hs Feriado:", 4, 0)
        self.entries_uocra['horas_altura'] = self.create_compact_entry(f_grid, "Hs Altura:", 4, 2)
        self.entries_uocra['horas_hormigon'] = self.create_compact_entry(f_grid, "Hs Hormigón:", 4, 4)
        
        # Separador
        sep2 = tk.Frame(f_grid, bg=mgc.COLORS['orange'], height=1)
        sep2.grid(row=5, column=0, columnspan=6, sticky='ew', pady=(4, 3))
        
        # Sección: Ausencias y Licencias
        lbl_ausencias = mgc.ctk.CTkLabel(f_grid, text="🏥 Ausencias y Licencias", 
                                        font=mgc.FONTS['heading'], text_color=mgc.COLORS['orange'])
        lbl_ausencias.grid(row=6, column=0, columnspan=6, sticky='w', padx=3, pady=(2, 1))
        
        self.entries_uocra['horas_enfermedad'] = self.create_compact_entry(f_grid, "Hs Enfermedad:", 7, 0)
        self.entries_uocra['horas_art'] = self.create_compact_entry(f_grid, "Hs ART:", 7, 2)
        self.entries_uocra['dias_vacaciones'] = self.create_compact_entry(f_grid, "Días Vacaciones:", 7, 4)
        
        self.entries_uocra['horas_quincena_anterior'] = self.create_compact_entry(f_grid, "Hs Quinc. Ant.:", 8, 0)
        self.entries_uocra['horas_conv_especial'] = self.create_compact_entry(f_grid, "Hs Conv. Esp.:", 8, 2)
        self.entries_uocra['porc_conv_especial'] = self.create_compact_entry(f_grid, "% Conv. Esp.:", 8, 4)
        
        # Separador
        sep3 = tk.Frame(f_grid, bg=mgc.COLORS['red'], height=1)
        sep3.grid(row=9, column=0, columnspan=6, sticky='ew', pady=(4, 3))
        
        # Sección: Parámetros Adicionales
        lbl_params = mgc.ctk.CTkLabel(f_grid, text="💰 Parámetros Adicionales", 
                                     font=mgc.FONTS['heading'], text_color=mgc.COLORS['red'])
        lbl_params.grid(row=10, column=0, columnspan=6, sticky='w', padx=3, pady=(2, 1))
        
        # Checkbox Presentismo (Nuevo)
        self.var_pres_uocra = tk.StringVar(value="Presentismo")
        # Checkbox Presentismo (Nuevo)
        self.var_pres_uocra = tk.StringVar(value="Presentismo")
        self.chk_pres_uocra = mgc.ctk.CTkCheckBox(f_grid, text="Aplicar Presentismo", variable=self.var_pres_uocra,
                                 onvalue="Presentismo", offvalue="Sin Presentismo",
                                 font=mgc.FONTS['small'], text_color=mgc.COLORS['text_primary'],
                                 fg_color=mgc.COLORS['blue'], hover_color=mgc.COLORS['accent_blue'])
        self.chk_pres_uocra.grid(row=11, column=0, columnspan=2, sticky='w', padx=3, pady=2)
        
        # Checkbox Seguro de Vida (Nuevo)
        self.var_seguro_uocra = tk.StringVar(value="Seg Vida")
        self.chk_seguro_uocra = mgc.ctk.CTkCheckBox(f_grid, text="Seguro de Vida", variable=self.var_seguro_uocra,
                                 onvalue="Seg Vida", offvalue="Sin Seg Vida",
                                 font=mgc.FONTS['small'], text_color=mgc.COLORS['text_primary'],
                                 fg_color=mgc.COLORS['purple'], hover_color=mgc.COLORS['purple'])
        self.chk_seguro_uocra.grid(row=11, column=2, columnspan=2, sticky='w', padx=3, pady=2)

        self.entries_uocra['porc_presentismo'] = self.create_compact_entry(f_grid, "% Presentismo:", 11, 4, default="20")
        self.entries_uocra['porc_adicional'] = self.create_compact_entry(f_grid, "% Adicional:", 12, 0, default="0")
        
        # Fila 12: Bono y Ret. Judicial
        self.entries_uocra['monto_bono'] = self.create_compact_entry(f_grid, "Monto Bono:", 12, 2, default="0")
        self.entries_uocra['porc_ret_judicial'] = self.create_compact_entry(f_grid, "% Ret. Judicial:", 12, 4, default="0")
        
        mgc.ctk.CTkLabel(f_grid, text="Tipo Bono:", 
                         font=mgc.FONTS['small'], text_color=mgc.COLORS['text_primary']).grid(row=13, column=0, sticky='w', padx=3, pady=2)
        self.cb_bono_tipo = mgc.ctk.CTkComboBox(f_grid, values=["No Remunerativo (aporta OS)", "No Remunerativo", "Remunerativo"], 
                                                state="readonly", width=230, font=mgc.FONTS['small'])
        self.cb_bono_tipo.set("No Remunerativo (aporta OS)")
        self.cb_bono_tipo.grid(row=13, column=1, columnspan=3, sticky='w', padx=3, pady=2)
        
        # Botón CALCULAR
        btn_frame = tk.Frame(self.tab_uocra, bg=mgc.COLORS['bg_primary'])
        btn_frame.pack(fill=tk.X, pady=(10, 0), padx=5)
        
        self.btn_calc_uocra = mgc.create_large_button(btn_frame, "CALCULAR", self.calculate_uocra, color='green', icon_image=self.icon_check)
        self.btn_calc_uocra.pack(pady=5)
        
        # Cargar categorías iniciales
        self.update_categories_uocra()
    
    def setup_uecara_tab(self):
        """Configura la pestaña UECARA con diseño compacto."""
        # Card de inputs
        card_outer, card_inner = mgc.create_card(self.tab_uecara, "Datos de Entrada", padding=8)
        card_outer.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        f_grid = tk.Frame(card_inner, bg=mgc.COLORS['bg_card'])
        f_grid.pack(fill=tk.X)
        
        mgc.ctk.CTkLabel(f_grid, text="Categoría:", 
                         font=mgc.FONTS['small'], text_color=mgc.COLORS['text_primary']).grid(row=0, column=0, sticky='w', padx=3, pady=3)
        self.cb_cat_uecara = mgc.ctk.CTkComboBox(f_grid, state="readonly", width=250, font=mgc.FONTS['small'])
        self.cb_cat_uecara.grid(row=0, column=1, columnspan=3, sticky='w', padx=3, pady=3)
        
        # Separador
        sep1 = tk.Frame(f_grid, bg=mgc.COLORS['purple'], height=1)
        sep1.grid(row=1, column=0, columnspan=4, sticky='ew', pady=(8, 5))
        
        # Sección: Días y Licencias
        lbl_dias = mgc.ctk.CTkLabel(f_grid, text="📅 Días y Licencias", 
                           font=mgc.FONTS['heading'], text_color=mgc.COLORS['purple'])
        lbl_dias.grid(row=2, column=0, columnspan=4, sticky='w', padx=3, pady=(3, 2))
        
        self.entries_uecara = {}
        self.entries_uecara['dnt'] = self.create_compact_entry(f_grid, "Días No Trabajados:", 3, 0)
        self.entries_uecara['feriados'] = self.create_compact_entry(f_grid, "Feriados:", 3, 2)
        
        # Separador
        sep2 = tk.Frame(f_grid, bg=mgc.COLORS['blue'], height=1)
        sep2.grid(row=4, column=0, columnspan=4, sticky='ew', pady=(6, 5))
        
        # Sección: Antigüedad y Estudios
        lbl_antig = mgc.ctk.CTkLabel(f_grid, text="🎓 Antigüedad y Estudios", 
                            font=mgc.FONTS['heading'], text_color=mgc.COLORS['blue'])
        lbl_antig.grid(row=5, column=0, columnspan=4, sticky='w', padx=3, pady=(3, 2))
        
        self.entries_uecara['anios_antiguedad'] = self.create_compact_entry(f_grid, "Años Antigüedad:", 6, 0)
        
        mgc.ctk.CTkLabel(f_grid, text="Título:", 
                         font=mgc.FONTS['small'], text_color=mgc.COLORS['text_primary']).grid(row=6, column=2, sticky='w', padx=3, pady=3)
        self.cb_titulo = mgc.ctk.CTkComboBox(f_grid, values=["Sin Título", "Título Secundario", "Título Técnico", "Título Universitario"], 
                                             state="readonly", width=200, font=mgc.FONTS['small'])
        self.cb_titulo.set("Sin Título")
        self.cb_titulo.grid(row=6, column=3, sticky='w', padx=3, pady=3)
        
        # Separador
        sep3 = tk.Frame(f_grid, bg=mgc.COLORS['green'], height=1)
        sep3.grid(row=7, column=0, columnspan=4, sticky='ew', pady=(6, 5))
        
        # Sección: Presentismo
        lbl_pres = tk.Label(f_grid, text="✓ Presentismo", bg=mgc.COLORS['bg_card'], 
                           font=('Segoe UI', 8, 'bold'), fg=mgc.COLORS['green'])
        lbl_pres.grid(row=8, column=0, columnspan=4, sticky='w', padx=3, pady=(3, 2))
        
        self.var_pres_uecara = tk.StringVar(value="Con Presentismo")
        self.chk_pres_uecara = mgc.ctk.CTkCheckBox(f_grid, text="Aplicar Presentismo", variable=self.var_pres_uecara,
                                 onvalue="Con Presentismo", offvalue="Sin Presentismo",
                                 font=mgc.FONTS['small'], text_color=mgc.COLORS['text_primary'],
                                 fg_color=mgc.COLORS['green'], hover_color=mgc.COLORS['success'])
        self.chk_pres_uecara.grid(row=9, column=0, columnspan=2, sticky='w', pady=3, padx=3)
        
        # Separador
        sep4 = tk.Frame(f_grid, bg=mgc.COLORS['orange'], height=1)
        sep4.grid(row=10, column=0, columnspan=4, sticky='ew', pady=(6, 5))
        
        # Sección: Bono
        lbl_bono = mgc.ctk.CTkLabel(f_grid, text="💰 Bono", 
                           font=mgc.FONTS['heading'], text_color=mgc.COLORS['orange'])
        lbl_bono.grid(row=11, column=0, columnspan=4, sticky='w', padx=3, pady=(3, 2))
        
        mgc.ctk.CTkLabel(f_grid, text="Monto Bono:", 
                         font=mgc.FONTS['small'], text_color=mgc.COLORS['text_primary']).grid(row=12, column=0, sticky='w', padx=3, pady=3)
        self.entries_uecara['bono'] = mgc.ctk.CTkEntry(f_grid, width=100, font=mgc.FONTS['small'],
                                                      fg_color=mgc.COLORS['bg_input'], border_color=mgc.COLORS['border'])
        self.entries_uecara['bono'].insert(0, "0")
        self.entries_uecara['bono'].grid(row=12, column=1, sticky='w', padx=3, pady=3)
        
        self.cb_bono_tipo_uecara = mgc.ctk.CTkComboBox(f_grid, values=["Aporta Obra Social", "Sin Retenciones", "Con Retenciones"], 
                                                       state="readonly", width=200, font=mgc.FONTS['small'])
        self.cb_bono_tipo_uecara.set("Aporta Obra Social")
        self.cb_bono_tipo_uecara.grid(row=12, column=2, columnspan=2, sticky='w', padx=3, pady=3)
        
        # Separador retención judicial
        sep5 = tk.Frame(f_grid, bg=mgc.COLORS['red'], height=1)
        sep5.grid(row=13, column=0, columnspan=4, sticky='ew', pady=(6, 5))
        
        # Retención Judicial
        lbl_judicial = mgc.ctk.CTkLabel(f_grid, text="⚖️ Retención Judicial", 
                                font=mgc.FONTS['heading'], text_color=mgc.COLORS['red'])
        lbl_judicial.grid(row=14, column=0, columnspan=4, sticky='w', padx=3, pady=(3, 2))
        
        self.entries_uecara['porc_ret_judicial'] = self.create_compact_entry(f_grid, "% Ret. Judicial:", 15, 0, default="0")
        
        # Botón CALCULAR
        btn_frame = tk.Frame(self.tab_uecara, bg=mgc.COLORS['bg_primary'])
        btn_frame.pack(fill=tk.X, pady=(10, 0), padx=5)
        
        self.btn_calc_uecara = mgc.create_large_button(btn_frame, "CALCULAR", self.calculate_uecara, color='purple', icon_image=self.icon_check)
        self.btn_calc_uecara.pack(pady=5)
        
        # Cargar categorías
        cats = self.excel_data["uecara_data"]["categorias"]
        self.cb_cat_uecara.configure(values=cats)
        if cats: self.cb_cat_uecara.set(cats[0])
    
    def create_compact_entry(self, parent, label_text, row, col, default="0"):
        """Crea un campo de entrada compacto con tema oscuro."""
        mgc.ctk.CTkLabel(parent, text=label_text,
                 font=mgc.FONTS['small'], text_color=mgc.COLORS['text_primary']).grid(
            row=row, column=col, sticky='w', padx=3, pady=2)
        entry = mgc.ctk.CTkEntry(parent, width=100, font=mgc.FONTS['small'],
                         fg_color=mgc.COLORS['bg_input'], border_color=mgc.COLORS['border'],
                         border_width=1, corner_radius=6,
                         text_color=mgc.COLORS['text_primary'])
        entry.insert(0, default)
        entry.grid(row=row, column=col+1, sticky='w', padx=3, pady=2)
        return entry
    
    def update_categories_uocra(self, choice=None):
        convenio = self.cb_convenio.get()
        if convenio in self.excel_data["hourly_rates"]:
            cats = list(self.excel_data["hourly_rates"][convenio].keys())
            self.cb_categoria.configure(values=cats)
            if cats: self.cb_categoria.set(cats[0])
        
        # Autocompletar Adicional 3% para NASA
        if convenio == "NASA":
            self.entries_uocra['porc_adicional'].delete(0, tk.END)
            self.entries_uocra['porc_adicional'].insert(0, "3")
        else:
            self.entries_uocra['porc_adicional'].delete(0, tk.END)
            self.entries_uocra['porc_adicional'].insert(0, "0")
    
    def calculate_uocra(self):
        print("DEBUG: Botón CALCULAR presionado (UOCRA)")
        try:
            data = {
                'convenio': self.cb_convenio.get(),
                'categoria': self.cb_categoria.get(),
                'bono_tipo': self.cb_bono_tipo.get(),
                'presentismo': self.var_pres_uocra.get(),
                'seguro_vida_opcion': self.var_seguro_uocra.get()
            }
            for key, entry in self.entries_uocra.items():
                data[key] = entry.get()
            
            print(f"DEBUG: Calculando para {data['convenio']} - {data['categoria']}")
            result = calculators.calculate_payroll_uocra_quilmes_nasa(data)
            
            calc_type = f"{self.cb_convenio.get()} - {self.cb_categoria.get()}"
            print("DEBUG: Creando ventana de resultados...")
            ResultsWindow(self.root, "Resultados del Cálculo", result, calc_type)
            self.status_var.set(f"✓ Cálculo completado para {calc_type}")
            
        except Exception as e:
            print(f"DEBUG: ERROR en calculate_uocra: {e}")
            messagebox.showerror("Error de Cálculo", f"Ocurrió un error al procesar los datos:\n{str(e)}")
    
    def calculate_uecara(self):
        print("DEBUG: Botón CALCULAR presionado (UECARA)")
        try:
            data = {
                'categoria': self.cb_cat_uecara.get(),
                'titulo': self.cb_titulo.get(),
                'presentismo_opcion': self.var_pres_uecara.get(),
                'bono_tipo': self.cb_bono_tipo_uecara.get()
            }
            for key, entry in self.entries_uecara.items():
                data[key] = entry.get()
            
            result = calculators.calculate_uecara(data)
            calc_type = f"UECARA - {self.cb_cat_uecara.get()}"
            ResultsWindow(self.root, "Resultados del Cálculo", result, calc_type)
            self.status_var.set(f"✓ Cálculo completado para {calc_type}")
        except Exception as e:
            messagebox.showerror("Error de Cálculo", f"Ocurrió un error al procesar los datos:\n{str(e)}")

# ==============================================================================
# MAIN ENTRY POINT
# ==============================================================================

def main():
    # Lanzar GUI unificada usando ctk.CTk para máxima compatibilidad
    root = ctk.CTk()
    app = UnifiedCalculatorApp(root)
    root.mainloop()

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        try:
            import ctypes
            ctypes.windll.user32.MessageBoxW(0, f"Error fatal en Calculadora de Sueldos:\n{str(e)}", "Error de Inicio", 0x10)
        except:
            pass
        sys.stderr.write(f"Error fatal: {e}\n")