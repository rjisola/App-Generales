import sys
import argparse
import tkinter as tk
from tkinter import ttk, messagebox
from decimal import Decimal, ROUND_HALF_UP
import os

# Asegurar que el directorio del script esté en sys.path para imports locales
script_dir = os.path.dirname(os.path.abspath(__file__))
if script_dir not in sys.path:
    sys.path.insert(0, script_dir)

# Importar componentes modernos
import modern_gui_components as mgc
from icon_loader import set_window_icon, load_icon

# Intentar importar la lógica de calculadoras existente
try:
    import calculators
except ImportError:
    calculators = None

# ==============================================================================
# LÓGICA DE CÁLCULO CLI (Mantenida para compatibilidad)
# ==============================================================================
from num2words import num2words

def to_decimal(value):
    if value is None:
        return Decimal('0')
    try:
        s = str(value).strip().replace(',', '.')
        return Decimal(s)
    except Exception:
        return Decimal('0')

def format_currency_decimal(dec_value: Decimal):
    try:
        q = dec_value.quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)
        return f"{float(q):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return "0,00"

def calculate_uocra_nasa_cli(args):
    # ... (Lógica original CLI mantenida, ver abajo) ...
    pass # Se implementará completa si se usa CLI

# ==============================================================================
# VENTANA SECUNDARIA DE RESULTADOS
# ==============================================================================

class ResultsWindow:
    """Ventana secundaria para mostrar resultados de cálculos con diseño profesional."""
    
    # Paleta de colores
    COLOR_HABER_BG = "#E8F5E9"
    COLOR_HABER_FG = "#2E7D32"
    COLOR_DEDUCCION_BG = "#FFEBEE"
    COLOR_DEDUCCION_FG = "#C62828"
    COLOR_BRUTO_BG = "#E3F2FD"
    COLOR_BRUTO_FG = "#1565C0"
    COLOR_NETO_BG = "#FFF9C4"
    COLOR_NETO_FG = "#F57F17"
    
    def __init__(self, parent, title, result_dict, calc_type="UOCRA/NASA"):
        self.window = tk.Toplevel(parent)
        self.window.title(f"📊 {title}")
        self.window.geometry("900x650")
        self.window.resizable(False, False)
        self.window.configure(bg=mgc.COLORS['bg_primary'])
        
        mgc.center_window(self.window, 900, 650)
        
        # Establecer icono de ventana
        set_window_icon(self.window, 'results')
        
        # Cargar iconos PNG
        self.icon_results = load_icon('results', (64, 64))
        self.icon_warning = load_icon('warning', (24, 24))
        
        # Contenedor principal
        main_frame = tk.Frame(self.window, bg=mgc.COLORS['bg_primary'])
        main_frame.pack(fill=tk.BOTH, expand=True, padx=15, pady=10)
        
        # Header compacto
        mgc.create_header(main_frame, title, f"Detalle - {calc_type}", icon_image=self.icon_results)
        
        # Verificar si hay error
        if "error" in result_dict:
            card_outer, card_inner = mgc.create_card(main_frame, "Error", padding=20)
            card_outer.pack(fill=tk.BOTH, expand=True, pady=(10, 10))
            error_label = tk.Label(card_inner, 
                                  text=f"❌ {result_dict['error']}",
                                  bg=mgc.COLORS['bg_card'],
                                  fg=mgc.COLORS['red'],
                                  font=mgc.FONTS['subtitle'],
                                  wraplength=800,
                                  justify=tk.LEFT)
            error_label.pack(pady=20)
        else:
            # Info simple (sin tarjeta)
            info_text = f"INFO: {calc_type}  |  Valor Hora/Mes: ${result_dict.get('tasa_horaria', result_dict.get('sueldo_quincenal', '0,00'))}"
            tk.Label(main_frame, text=info_text, bg=mgc.COLORS['bg_primary'], 
                    font=('Segoe UI', 9), fg=mgc.COLORS['text_secondary']).pack(pady=(2, 5))
            
            # Determinar si es UECARA o UOCRA/NASA/QUILMES
            is_uecara = "sueldo_quincenal" in result_dict
            
            if is_uecara:
                self._build_uecara_layout(main_frame, result_dict, calc_type)
            else:
                self._build_uocra_layout(main_frame, result_dict, calc_type)
        
        # Botón cerrar
        btn_frame = tk.Frame(main_frame, bg=mgc.COLORS['bg_primary'])
        btn_frame.pack(fill=tk.X, pady=(5, 0))
        
        btn_close = mgc.create_button(btn_frame, "CERRAR", self.window.destroy, color='red', icon_image=self.icon_warning)
        btn_close.pack()
        
        # Barra de estado
        status_frame, status_var = mgc.create_status_bar(self.window, f"✓ Cálculo completado")
    
    def _build_uocra_layout(self, parent, data, calc_type):
        """Layout para UOCRA/NASA/QUILMES."""
        # Contenedor de 2 columnas
        cols_frame = tk.Frame(parent, bg=mgc.COLORS['bg_primary'])
        cols_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 2))
        
        # Columna HABERES (izquierda)
        haberes_frame = tk.Frame(cols_frame, bg=self.COLOR_HABER_BG, relief=tk.RIDGE, bd=2)
        haberes_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 5))
        
        tk.Label(haberes_frame, text="✓ HABERES", bg=self.COLOR_HABER_BG, 
                fg=self.COLOR_HABER_FG, font=('Segoe UI', 9, 'bold')).pack(pady=2)
        
        haberes_items = [
            ("Horas Normales", data.get("importe_normal", "0,00")),
            ("Horas 50%", data.get("importe_50", "0,00")),
            ("Horas 100%", data.get("importe_100", "0,00")),
            ("Horas Altura", data.get("importe_altura", "0,00")),
            ("Horas Hormigón", data.get("importe_hormigon", "0,00")),
            ("Presentismo", data.get("importe_presentismo", "0,00")),
            ("Adicional", data.get("importe_adicional", "0,00")),
            ("Feriados", data.get("importe_feriado", "0,00")),
            ("Vacaciones", data.get("importe_vacaciones", "0,00")),
            ("Enfermedad", data.get("importe_enfermedad", "0,00")),
            ("ART", data.get("importe_art", "0,00")),
            ("Quincena Ant.", data.get("importe_quincena_anterior", "0,00")),
            ("Conv. Especial", data.get("importe_horas_conv_especial", "0,00")),
            ("Adic. Conv. Esp.", data.get("importe_adic_conv_especial", "0,00")),
            ("Importe Bono", data.get("importe_bono", "0,00")),
        ]
        
        self._add_items_to_column(haberes_frame, haberes_items, self.COLOR_HABER_BG, self.COLOR_HABER_FG)
        
        # Columna DEDUCCIONES (derecha)
        deduc_frame = tk.Frame(cols_frame, bg=self.COLOR_DEDUCCION_BG, relief=tk.RIDGE, bd=2)
        deduc_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        tk.Label(deduc_frame, text="✗ DEDUCCIONES", bg=self.COLOR_DEDUCCION_BG, 
                fg=self.COLOR_DEDUCCION_FG, font=('Segoe UI', 9, 'bold')).pack(pady=2)
        
        deduc_items = [
            ("Jubilación (11%)", data.get("jubilacion", "0,00")),
            ("Ley 19032 (3%)", data.get("ley19032", "0,00")),
            ("Obra Social (3%)", data.get("obra_social_bruto", "0,00")),
            ("Sindicato (2.5%)", data.get("sindicato_bruto", "0,00")),
            ("OS Bono", data.get("obra_social_bono", "0,00")),
            ("Seguro Vida", data.get("seguro_vida", "0,00")),
            ("Ret. Ganancias", data.get("ret_ganancias", "0,00")),
            ("Ret. Judicial", data.get("importe_ret_judicial", "0,00")),
        ]
        
        self._add_items_to_column(deduc_frame, deduc_items, self.COLOR_DEDUCCION_BG, self.COLOR_DEDUCCION_FG)
        
        # Resumen final
        self._build_summary(parent, data)
    
    def _build_uecara_layout(self, parent, data, calc_type):
        """Layout para UECARA."""
        # Contenedor de 2 columnas
        cols_frame = tk.Frame(parent, bg=mgc.COLORS['bg_primary'])
        cols_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 2))
        
        # HABERES
        haberes_frame = tk.Frame(cols_frame, bg=self.COLOR_HABER_BG, relief=tk.RIDGE, bd=2)
        haberes_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 5))
        
        tk.Label(haberes_frame, text="✓ HABERES", bg=self.COLOR_HABER_BG, 
                fg=self.COLOR_HABER_FG, font=('Segoe UI', 9, 'bold')).pack(pady=2)
        
        haberes_items = [
            ("Sueldo Quincenal", data.get("sueldo_quincenal", "0,00")),
            ("Descuento DNT", "-" + data.get("descuento_dnt", "0,00")),
            ("Feriados", data.get("importe_feriados", "0,00")),
            ("Antigüedad", data.get("importe_antiguedad", "0,00")),
            ("Título", data.get("importe_titulo", "0,00")),
            ("Presentismo", data.get("importe_presentismo", "0,00")),
            ("Ajuste Redondeo", data.get("ajuste_redondeo", "0,00")),
            ("Importe Bono", data.get("importe_bono", "0,00")),
        ]
        
        self._add_items_to_column(haberes_frame, haberes_items, self.COLOR_HABER_BG, self.COLOR_HABER_FG)
        
        # DEDUCCIONES
        deduc_frame = tk.Frame(cols_frame, bg=self.COLOR_DEDUCCION_BG, relief=tk.RIDGE, bd=2)
        deduc_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        tk.Label(deduc_frame, text="✗ DEDUCCIONES", bg=self.COLOR_DEDUCCION_BG, 
                fg=self.COLOR_DEDUCCION_FG, font=('Segoe UI', 9, 'bold')).pack(pady=2)
        
        deduc_items = [
            ("Jubilación (11%)", data.get("jubilacion", "0,00")),
            ("Ley 19032 (3%)", data.get("ley19032", "0,00")),
            ("Obra Social (3%)", data.get("obra_social", "0,00")),
            ("Sindicato (2.5%)", data.get("sindicato", "0,00")),
            ("OS Bono", data.get("os_bono", "0,00")),
            ("Ret. Ganancias", data.get("ganancias_importe", "0,00")),
        ]
        
        self._add_items_to_column(deduc_frame, deduc_items, self.COLOR_DEDUCCION_BG, self.COLOR_DEDUCCION_FG)
        
        # Resumen
        self._build_summary(parent, data, is_uecara=True)
    
    def _add_items_to_column(self, parent, items, bg_color, fg_color):
        """Agrega items a una columna, solo si valor != 0,00."""
        for label, value in items:
            # Convertir valor a float para comparar
            try:
                val_clean = value.replace(".", "").replace(",", ".")
                if val_clean.startswith("-"):
                    val_float = float(val_clean)
                else:
                    val_float = float(val_clean)
                if abs(val_float) < 0.01:  # Ignorar valores cercanos a 0
                    continue
            except:
                pass  # Si no se puede convertir, mostrar igual
            
            item_frame = tk.Frame(parent, bg=bg_color)
            item_frame.pack(fill=tk.X, padx=5, pady=0) # Compactado pady=0
            
            tk.Label(item_frame, text=label, bg=bg_color, fg=fg_color, 
                    font=('Segoe UI', 8), anchor='w').pack(side=tk.LEFT, fill=tk.X, expand=True) # Fuente 8
            tk.Label(item_frame, text=f"${value}", bg=bg_color, fg=fg_color, 
                    font=('Segoe UI', 8, 'bold'), anchor='e').pack(side=tk.RIGHT) # Fuente 8
    
    def _build_summary(self, parent, data, is_uecara=False):
        """Construye el resumen final con bruto, deducciones y neto."""
        summary_card, summary_inner = mgc.create_card(parent, "", padding=5) # Compactado padding=5
        summary_card.pack(fill=tk.X, pady=(0, 5))
        
        # Total Bruto
        bruto_val = data.get("bruto" if is_uecara else "total_bruto", "0,00")
        bruto_frame = tk.Frame(summary_inner, bg=self.COLOR_BRUTO_BG, relief=tk.SOLID, bd=1)
        bruto_frame.pack(fill=tk.X, pady=1) # Compactado
        tk.Label(bruto_frame, text="💰 TOTAL BRUTO (REM):", bg=self.COLOR_BRUTO_BG, 
                fg=self.COLOR_BRUTO_FG, font=('Segoe UI', 9, 'bold'), anchor='w').pack(side=tk.LEFT, padx=10, pady=2) # Compactado
        tk.Label(bruto_frame, text=f"${bruto_val}", bg=self.COLOR_BRUTO_BG, 
                fg=self.COLOR_BRUTO_FG, font=('Segoe UI', 10, 'bold'), anchor='e').pack(side=tk.RIGHT, padx=10, pady=2) # Compactado
        
        # Total No Remunerativo (Nuevo)
        no_rem_val = data.get("haberes_sin_descuento", "0,00")
        if no_rem_val != "0,00":
            nr_frame = tk.Frame(summary_inner, bg="#FFF3E0", relief=tk.SOLID, bd=1)
            nr_frame.pack(fill=tk.X, pady=1) # Compactado
            tk.Label(nr_frame, text="➕ NO REMUNERATIVO:", bg="#FFF3E0", 
                    fg="#E65100", font=('Segoe UI', 9, 'bold'), anchor='w').pack(side=tk.LEFT, padx=10, pady=2) # Compactado
            tk.Label(nr_frame, text=f"${no_rem_val}", bg="#FFF3E0", 
                    fg="#E65100", font=('Segoe UI', 10, 'bold'), anchor='e').pack(side=tk.RIGHT, padx=10, pady=2) # Compactado

        # Total Deducciones
        deduc_val = data.get("total_deducciones", "0,00")
        deduc_frame = tk.Frame(summary_inner, bg=self.COLOR_DEDUCCION_BG, relief=tk.SOLID, bd=1)
        deduc_frame.pack(fill=tk.X, pady=1) # Compactado
        tk.Label(deduc_frame, text="📉 TOTAL DEDUCCIONES:", bg=self.COLOR_DEDUCCION_BG, 
                fg=self.COLOR_DEDUCCION_FG, font=('Segoe UI', 9, 'bold'), anchor='w').pack(side=tk.LEFT, padx=10, pady=2) # Compactado
        tk.Label(deduc_frame, text=f"${deduc_val}", bg=self.COLOR_DEDUCCION_BG, 
                fg=self.COLOR_DEDUCCION_FG, font=('Segoe UI', 10, 'bold'), anchor='e').pack(side=tk.RIGHT, padx=10, pady=2) # Compactado
        
        # Separador
        tk.Frame(summary_inner, bg=mgc.COLORS['text_secondary'], height=1).pack(fill=tk.X, pady=3)
        
        # NETO (destacado)
        neto_val = data.get("neto_redondeado" if is_uecara else "total_final", "0,00")
        neto_frame = tk.Frame(summary_inner, bg=self.COLOR_NETO_BG, relief=tk.RAISED, bd=3)
        neto_frame.pack(fill=tk.X, pady=1) # Compactado
        tk.Label(neto_frame, text="💵 NETO A COBRAR:", bg=self.COLOR_NETO_BG, 
                fg=self.COLOR_NETO_FG, font=('Segoe UI', 11, 'bold'), anchor='w').pack(side=tk.LEFT, padx=10, pady=5) # Compactado
        tk.Label(neto_frame, text=f"${neto_val}", bg=self.COLOR_NETO_BG, 
                fg=self.COLOR_NETO_FG, font=('Segoe UI', 13, 'bold'), anchor='e').pack(side=tk.RIGHT, padx=10, pady=5) # Compactado
        
        # Neto en letras
        neto_letras = data.get("neto_en_letras", "")
        if neto_letras:
            tk.Label(summary_inner, text=neto_letras, bg=mgc.COLORS['bg_card'], 
                    fg=mgc.COLORS['text_secondary'], font=('Segoe UI', 8, 'italic'), 
                    wraplength=800).pack(pady=(1, 0))

# ==============================================================================
# VENTANA SECUNDARIA - CALCULADORA UOCRA/NASA
# ==============================================================================

# ==============================================================================
# VENTANA PRINCIPAL UNIFICADA CON PESTAÑAS
# ==============================================================================

class UnifiedCalculatorApp:
    """Aplicación unificada con pestañas para UOCRA/NASA y UECARA."""
    
    def __init__(self, root):
        self.root = root
        self.root.title("🧮 Calculadora de Sueldos")
        self.root.geometry("900x650")
        self.root.resizable(False, False)
        self.root.configure(bg=mgc.COLORS['bg_primary'])
        
        mgc.center_window(self.root, 900, 650)
        
        # Establecer icono de ventana
        set_window_icon(self.root, 'calculator')
        
        # Cargar iconos PNG
        self.icon_calculator = load_icon('calculator', (64, 64))
        self.icon_check = load_icon('check', (24, 24))
        
        # Cargar datos del Excel
        if not calculators:
            messagebox.showerror("Error", "No se encontró el módulo 'calculators.py'.")
            self.root.destroy()
            return
        
        self.excel_data = calculators._load_excel_data()
        if not self.excel_data:
            messagebox.showerror("Error", "No se pudieron cargar los datos del Excel '2-VALOR_HORAS_SUELDOS.xlsx'.\nAsegúrese de que existe en la carpeta 'Datos'.")
            self.root.destroy()
            return
        
        # Contenedor principal
        main_frame = tk.Frame(self.root, bg=mgc.COLORS['bg_primary'])
        main_frame.pack(fill=tk.BOTH, expand=True, padx=15, pady=10)
        
        # Header compacto con iconos
        header_frame = tk.Frame(main_frame, bg=mgc.COLORS['bg_primary'])
        header_frame.pack(fill=tk.X, pady=(0, 5))
        
        title_label = tk.Label(header_frame, text="🧮 Calculadora de Sueldos", 
                              font=('Segoe UI', 14, 'bold'), bg=mgc.COLORS['bg_primary'], 
                              fg=mgc.COLORS['text_primary'])
        title_label.pack()
        
        # Notebook con pestañas
        style = ttk.Style()
        style.configure("TNotebook", background=mgc.COLORS['bg_primary'], borderwidth=0)
        style.configure("TNotebook.Tab", font=('Segoe UI', 10, 'bold'), padding=[15, 8])
        
        self.notebook = ttk.Notebook(main_frame)
        self.notebook.pack(fill=tk.BOTH, expand=True, pady=(0, 5))
        
        # Pestaña UOCRA/NASA
        self.tab_uocra = tk.Frame(self.notebook, bg=mgc.COLORS['bg_primary'])
        self.notebook.add(self.tab_uocra, text="  🧮 UOCRA / NASA  ")
        self.setup_uocra_tab()
        
        # Pestaña UECARA
        self.tab_uecara = tk.Frame(self.notebook, bg=mgc.COLORS['bg_primary'])
        self.notebook.add(self.tab_uecara, text="  💼 UECARA  ")
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
        tk.Label(f_grid, text="Convenio:", bg=mgc.COLORS['bg_card'], font=('Segoe UI', 9)).grid(row=0, column=0, sticky='w', padx=3, pady=2)
        self.cb_convenio = ttk.Combobox(f_grid, values=["UOCRA", "QUILMES", "NASA"], state="readonly", width=10, font=('Segoe UI', 9))
        self.cb_convenio.current(0)
        self.cb_convenio.grid(row=0, column=1, sticky='w', padx=3, pady=2)
        self.cb_convenio.bind("<<ComboboxSelected>>", self.update_categories_uocra)
        
        tk.Label(f_grid, text="Categoría:", bg=mgc.COLORS['bg_card'], font=('Segoe UI', 9)).grid(row=0, column=2, sticky='w', padx=3, pady=2)
        self.cb_categoria = ttk.Combobox(f_grid, state="readonly", width=22, font=('Segoe UI', 9))
        self.cb_categoria.grid(row=0, column=3, columnspan=3, sticky='w', padx=3, pady=2)
        
        # Separador
        sep1 = tk.Frame(f_grid, bg=mgc.COLORS['purple'], height=1)
        sep1.grid(row=1, column=0, columnspan=6, sticky='ew', pady=(5, 3))
        
        self.entries_uocra = {}
        
        # Sección: Horas Extras
        lbl_extras = tk.Label(f_grid, text="⏰ Horas Extras", bg=mgc.COLORS['bg_card'], 
                             font=('Segoe UI', 8, 'bold'), fg=mgc.COLORS['purple'])
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
        lbl_ausencias = tk.Label(f_grid, text="🏥 Ausencias y Licencias", bg=mgc.COLORS['bg_card'], 
                                font=('Segoe UI', 8, 'bold'), fg=mgc.COLORS['orange'])
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
        lbl_params = tk.Label(f_grid, text="💰 Parámetros Adicionales", bg=mgc.COLORS['bg_card'], 
                             font=('Segoe UI', 8, 'bold'), fg=mgc.COLORS['red'])
        lbl_params.grid(row=10, column=0, columnspan=6, sticky='w', padx=3, pady=(2, 1))
        
        # Checkbox Presentismo (Nuevo)
        self.var_pres_uocra = tk.StringVar(value="Presentismo")
        tk.Checkbutton(f_grid, text="Aplicar Presentismo", variable=self.var_pres_uocra, 
                      onvalue="Presentismo", offvalue="Sin Presentismo", 
                      bg=mgc.COLORS['bg_card'], font=('Segoe UI', 9), 
                      activebackground=mgc.COLORS['bg_card']).grid(row=11, column=0, columnspan=2, sticky='w', padx=3, pady=2)

        self.entries_uocra['porc_presentismo'] = self.create_compact_entry(f_grid, "% Presentismo:", 11, 2, default="20")
        self.entries_uocra['porc_adicional'] = self.create_compact_entry(f_grid, "% Adicional:", 11, 4, default="0")
        
        # Fila 12: Bono
        self.entries_uocra['monto_bono'] = self.create_compact_entry(f_grid, "Monto Bono:", 12, 0, default="0")
        
        tk.Label(f_grid, text="Tipo Bono:", bg=mgc.COLORS['bg_card'], font=('Segoe UI', 9)).grid(row=12, column=2, sticky='w', padx=3, pady=2)
        self.cb_bono_tipo = ttk.Combobox(f_grid, values=["No Remunerativo (aporta OS)", "No Remunerativo", "Remunerativo"], 
                                        state="readonly", width=22, font=('Segoe UI', 9))
        self.cb_bono_tipo.current(0)
        self.cb_bono_tipo.grid(row=12, column=3, columnspan=3, sticky='w', padx=3, pady=2)
        
        # Botón CALCULAR
        btn_frame = tk.Frame(self.tab_uocra, bg=mgc.COLORS['bg_primary'])
        btn_frame.pack(fill=tk.X, pady=(3, 0), padx=5)
        
        btn_calc = mgc.create_large_button(btn_frame, "CALCULAR", self.calculate_uocra, color='blue', icon_image=self.icon_check)
        btn_calc.pack()
        
        # Cargar categorías iniciales
        self.update_categories_uocra()
    
    def setup_uecara_tab(self):
        """Configura la pestaña UECARA con diseño compacto."""
        # Card de inputs
        card_outer, card_inner = mgc.create_card(self.tab_uecara, "Datos de Entrada", padding=8)
        card_outer.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        f_grid = tk.Frame(card_inner, bg=mgc.COLORS['bg_card'])
        f_grid.pack(fill=tk.X)
        
        tk.Label(f_grid, text="Categoría:", bg=mgc.COLORS['bg_card'], font=('Segoe UI', 9)).grid(row=0, column=0, sticky='w', padx=3, pady=3)
        self.cb_cat_uecara = ttk.Combobox(f_grid, state="readonly", width=25, font=('Segoe UI', 9))
        self.cb_cat_uecara.grid(row=0, column=1, columnspan=3, sticky='w', padx=3, pady=3)
        
        # Separador
        sep1 = tk.Frame(f_grid, bg=mgc.COLORS['purple'], height=1)
        sep1.grid(row=1, column=0, columnspan=4, sticky='ew', pady=(8, 5))
        
        # Sección: Días y Licencias
        lbl_dias = tk.Label(f_grid, text="📅 Días y Licencias", bg=mgc.COLORS['bg_card'], 
                           font=('Segoe UI', 8, 'bold'), fg=mgc.COLORS['purple'])
        lbl_dias.grid(row=2, column=0, columnspan=4, sticky='w', padx=3, pady=(3, 2))
        
        self.entries_uecara = {}
        self.entries_uecara['dnt'] = self.create_compact_entry(f_grid, "Días No Trabajados:", 3, 0)
        self.entries_uecara['feriados'] = self.create_compact_entry(f_grid, "Feriados:", 3, 2)
        
        # Separador
        sep2 = tk.Frame(f_grid, bg=mgc.COLORS['blue'], height=1)
        sep2.grid(row=4, column=0, columnspan=4, sticky='ew', pady=(6, 5))
        
        # Sección: Antigüedad y Estudios
        lbl_antig = tk.Label(f_grid, text="🎓 Antigüedad y Estudios", bg=mgc.COLORS['bg_card'], 
                            font=('Segoe UI', 8, 'bold'), fg=mgc.COLORS['blue'])
        lbl_antig.grid(row=5, column=0, columnspan=4, sticky='w', padx=3, pady=(3, 2))
        
        self.entries_uecara['anios_antiguedad'] = self.create_compact_entry(f_grid, "Años Antigüedad:", 6, 0)
        
        tk.Label(f_grid, text="Título:", bg=mgc.COLORS['bg_card'], font=('Segoe UI', 9)).grid(row=6, column=2, sticky='w', padx=3, pady=3)
        self.cb_titulo = ttk.Combobox(f_grid, values=["Sin Título", "Título Secundario", "Título Técnico", "Título Universitario"], 
                                     state="readonly", width=20, font=('Segoe UI', 9))
        self.cb_titulo.current(0)
        self.cb_titulo.grid(row=6, column=3, sticky='w', padx=3, pady=3)
        
        # Separador
        sep3 = tk.Frame(f_grid, bg=mgc.COLORS['green'], height=1)
        sep3.grid(row=7, column=0, columnspan=4, sticky='ew', pady=(6, 5))
        
        # Sección: Presentismo
        lbl_pres = tk.Label(f_grid, text="✓ Presentismo", bg=mgc.COLORS['bg_card'], 
                           font=('Segoe UI', 8, 'bold'), fg=mgc.COLORS['green'])
        lbl_pres.grid(row=8, column=0, columnspan=4, sticky='w', padx=3, pady=(3, 2))
        
        self.var_pres_uecara = tk.StringVar(value="Con Presentismo")
        tk.Checkbutton(f_grid, text="Aplicar Presentismo", variable=self.var_pres_uecara, 
                      onvalue="Con Presentismo", offvalue="Sin Presentismo", 
                      bg=mgc.COLORS['bg_card'], font=('Segoe UI', 9), 
                      activebackground=mgc.COLORS['bg_card']).grid(row=9, column=0, columnspan=2, sticky='w', pady=3, padx=3)
        
        # Separador
        sep4 = tk.Frame(f_grid, bg=mgc.COLORS['orange'], height=1)
        sep4.grid(row=10, column=0, columnspan=4, sticky='ew', pady=(6, 5))
        
        # Sección: Bono
        lbl_bono = tk.Label(f_grid, text="💰 Bono", bg=mgc.COLORS['bg_card'], 
                           font=('Segoe UI', 8, 'bold'), fg=mgc.COLORS['orange'])
        lbl_bono.grid(row=11, column=0, columnspan=4, sticky='w', padx=3, pady=(3, 2))
        
        tk.Label(f_grid, text="Monto Bono:", bg=mgc.COLORS['bg_card'], font=('Segoe UI', 9)).grid(row=12, column=0, sticky='w', padx=3, pady=3)
        self.entries_uecara['bono'] = tk.Entry(f_grid, width=10, font=('Segoe UI', 9), relief=tk.GROOVE, bd=1)
        self.entries_uecara['bono'].insert(0, "0")
        self.entries_uecara['bono'].grid(row=12, column=1, sticky='w', padx=3, pady=3)
        
        self.cb_bono_tipo_uecara = ttk.Combobox(f_grid, values=["Aporta Obra Social", "Sin Retenciones", "Con Retenciones"], 
                                               state="readonly", width=20, font=('Segoe UI', 9))
        self.cb_bono_tipo_uecara.current(0)
        self.cb_bono_tipo_uecara.grid(row=12, column=2, columnspan=2, sticky='w', padx=3, pady=3)
        
        # Botón CALCULAR
        btn_frame = tk.Frame(self.tab_uecara, bg=mgc.COLORS['bg_primary'])
        btn_frame.pack(fill=tk.X, pady=(3, 0), padx=5)
        
        btn_calc = mgc.create_large_button(btn_frame, "CALCULAR", self.calculate_uecara, color='purple', icon_image=self.icon_check)
        btn_calc.pack()
        
        # Cargar categorías
        cats = self.excel_data["uecara_data"]["categorias"]
        self.cb_cat_uecara['values'] = cats
        if cats: self.cb_cat_uecara.current(0)
    
    def create_compact_entry(self, parent, label_text, row, col, default="0"):
        """Crea un campo de entrada compacto."""
        tk.Label(parent, text=label_text, bg=mgc.COLORS['bg_card'], font=('Segoe UI', 9)).grid(
            row=row, column=col, sticky='w', padx=3, pady=2)
        entry = tk.Entry(parent, width=10, font=('Segoe UI', 9), relief=tk.GROOVE, bd=1)
        entry.insert(0, default)
        entry.grid(row=row, column=col+1, sticky='w', padx=3, pady=2)
        return entry
    
    def update_categories_uocra(self, event=None):
        convenio = self.cb_convenio.get()
        if convenio in self.excel_data["hourly_rates"]:
            cats = list(self.excel_data["hourly_rates"][convenio].keys())
            self.cb_categoria['values'] = cats
            if cats: self.cb_categoria.current(0)
        
        # Autocompletar Adicional 3% para NASA
        if convenio == "NASA":
            self.entries_uocra['porc_adicional'].delete(0, tk.END)
            self.entries_uocra['porc_adicional'].insert(0, "3")
        else:
            self.entries_uocra['porc_adicional'].delete(0, tk.END)
            self.entries_uocra['porc_adicional'].insert(0, "0")
    
    def calculate_uocra(self):
        data = {
            'convenio': self.cb_convenio.get(),
            'categoria': self.cb_categoria.get(),
            'bono_tipo': self.cb_bono_tipo.get(),
            'presentismo': self.var_pres_uocra.get()
        }
        for key, entry in self.entries_uocra.items():
            data[key] = entry.get()
        
        result = calculators.calculate_payroll_uocra_quilmes_nasa(data)
        calc_type = f"{self.cb_convenio.get()} - {self.cb_categoria.get()}"
        ResultsWindow(self.root, "Resultados del Cálculo", result, calc_type)
        self.status_var.set(f"✓ Cálculo completado para {calc_type}")
    
    def calculate_uecara(self):
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

# ==============================================================================
# MAIN ENTRY POINT
# ==============================================================================

def main():
    # Si hay argumentos, usar CLI
    if len(sys.argv) > 1:
        print("Modo CLI detectado pero no implementado completamente en esta versión híbrida.")
        print("Por favor ejecute sin argumentos para la Interfaz Gráfica.")
        return

    # Lanzar GUI unificada
    root = tk.Tk()
    app = UnifiedCalculatorApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()