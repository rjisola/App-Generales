# -*- coding: utf-8 -*-
"""
🚀 LAUNCHER MODERNO - SISTEMA DE SUELDOS
Versión nativa en Python que replica el diseño Ultra Premium del Launcher Web.
"""

import os
import sys

# Asegurar que el directorio 03_OTROS esté en el path para los componentes compartidos
script_dir = os.path.dirname(os.path.abspath(__file__))
others_dir = os.path.join(script_dir, "03_OTROS")
if others_dir not in sys.path:
    sys.path.insert(0, others_dir)

import subprocess
import tkinter as tk
import customtkinter as ctk
from PIL import Image
import modern_gui_components as mgc
import tkinter.messagebox as messagebox

# =============================================================================
# ROUTER DE EJECUCIÓN (Para modo Portable / OneFile)
# =============================================================================

def run_sub_app(app_id):
    """ Busca y ejecuta la sub-aplicación dentro del contexto del bundle """
    from modern_gui_components import get_resource_path
    
    all_apps = MAIN_APPS + EXTRA_APPS
    app = next((a for a in all_apps if a['id'] == app_id), None)
    
    if not app:
        print(f"Error: Aplicación {app_id} no encontrada.")
        return

    # Obtener ruta absoluta (considerando PyInstaller _MEIPASS)
    file_path = get_resource_path(app['file'])
    
    if not os.path.exists(file_path):
        # Fallback: intentar en _internal si estamos en modo onedir
        file_path = get_resource_path(os.path.join("_internal", app['file']))

    if os.path.exists(file_path):
        import runpy
        # Asegurar que el entorno sea correcto para la sub-app
        os.environ['CHILD_APP_MODE'] = '1'
        
        # Inyectar rutas para que la sub-app encuentre modern_gui_components
        sub_app_dir = os.path.dirname(file_path)
        if sub_app_dir not in sys.path:
            sys.path.insert(0, sub_app_dir)
        if others_dir not in sys.path:
            sys.path.insert(0, others_dir)
        
        # Ejecutar el script como __main__
        runpy.run_path(file_path, run_name="__main__")
    else:
        print(f"Error crítico: No se encontró el archivo {file_path}")

# ... (resto de las definiciones de APPS igual, pero para ahorrar contexto solo cambio start_launcher en la siguiente llamada si es necesario)
# En realidad, el prompt pide un cambio quirúrgico. Voy a reemplazar todo el bloque start_launcher y las importaciones.

# =============================================================================
# ESTRUCTURA DE DATOS DE LAS APLICACIONES
# =============================================================================

MAIN_APPS = [
    {
        'id': 'preparar',
        'file': '01_APLICACIONES/actualizar_quincena_gui.pyw',
        'icon': 'dates.png',
        'emoji': '📅',
        'name': 'Preparar Quincena',
        'short': 'Realiza la limpieza integral del archivo Excel, actualiza los días y resetea las fórmulas.',
        'description': 'Limpia el archivo de depósito y actualiza el calendario y formatos para el nuevo período.',
        'details': ['Borrado de horas y resultados anteriores', 'Configuración de calendario', 'Pintado de fines de semana/feriados', 'Limpieza de reportes'],
        'color': 'rose',
        'tags': ['Inicio', 'Nuevo'],
        'badge': 'Fundacional'
    },
    {
        'id': 'deposito',
        'file': '01_APLICACIONES/15-PASAR_HORAS_DEPOSITO.pyw',
        'icon': 'warehouse.png',
        'emoji': '🏪',
        'name': 'Horas Depósito',
        'short': 'Procesa las novedades de horas del personal de logística y depósito.',
        'description': 'Carga y transfiere las horas del personal de depósito al sistema.',
        'details': ['Carga horas normales y extras', 'Sincroniza con nómina principal', 'Copias de seguridad automáticas', 'Módulo esencial post-reset'],
        'color': 'teal',
        'tags': ['Logística', 'Carga Inicial']
    },
    {
        'id': 'calculadora',
        'file': '01_APLICACIONES/10-CALCULAR_HORAS.pyw',
        'icon': 'calculator.png',
        'emoji': '🧮',
        'name': 'Calculadora de Horas',
        'short': 'Calcula automáticamente el sueldo bruto según horas trabajadas, categoría y sindicato.',
        'description': 'Calcula jornales UOCRA, NASA y UECARA con promedios, feriados y extras.',
        'details': ['Carga de horas por empleado', 'Cálculo de promedios y extras', 'Aplicación de feriados y adicionales', 'Exporta resultados a Excel'],
        'color': 'blue',
        'tags': ['UOCRA', 'NASA', 'UECARA']
    },
    {
        'id': 'procesar',
        'file': '01_APLICACIONES/B-PROCESARSUELDOS.pyw',
        'icon': 'payroll.png',
        'emoji': '💼',
        'name': 'Procesar Sueldos',
        'short': 'Motor central: toma las horas, aplica descuentos y adicionales, y cierra la liquidación.',
        'description': 'Motor central de procesamiento: cierra la nómina y genera el Excel operativo.',
        'details': ['Procesa retenciones de ley', 'Aplica retención judicial', 'Calcula SAC y vacaciones', 'Genera Excel final liquidación'],
        'color': 'purple',
        'tags': ['Nómina', 'Excel']
    },
    {
        'id': 'recibos',
        'file': '01_APLICACIONES/A-GENERAR_RECIBOS_CONTROL.pyw',
        'icon': 'receipts.png',
        'emoji': '📄',
        'name': 'Generar Recibos',
        'short': 'Genera un recibo completo en PDF para cada empleado desde la nómina procesada.',
        'description': 'Crea recibos de sueldo individuales en PDF, listos para imprimir o enviar.',
        'details': ['Generación masiva en un clic', 'Incluye todos los conceptos', 'Formato legal con firma', 'Organización automática en carpetas'],
        'color': 'green',
        'tags': ['PDF', 'Masivo']
    },
    {
        'id': 'sobres',
        'file': '01_APLICACIONES/imprimir_sobres_gui.pyw',
        'icon': 'bonus_white.png',
        'emoji': '✉️',
        'name': 'Imprimir Sobres',
        'short': 'Configura e imprime en lote los sobres C5 con los datos de cada empleado.',
        'description': 'Asistente para configurar e imprimir sobres C5 con datos de cada empleado.',
        'details': ['Toma datos del listado activo', 'Configuración de márgenes', 'Impresión masiva por lotes', 'Vista previa de impresión'],
        'color': 'cyan',
        'tags': ['Impresión', 'C5']
    },
    {
        'id': 'email',
        'file': '01_APLICACIONES/Enviar_Documentacion_Email.pyw',
        'icon': 'bonus_white.png',
        'emoji': '✈️',
        'name': 'Enviar Documentación',
        'short': 'Envía automáticamente los recibos y archivos por correo electrónico (Gmail).',
        'description': 'Envía los archivos ZIP de sueldos y recibos por correo electrónico via Gmail.',
        'details': ['Adjunta recibos o paquetes ZIP', 'Envío masivo en un clic', 'Asuntos y cuerpos personalizables', 'Registro de confirmación'],
        'color': 'rose',
        'tags': ['Gmail', 'ZIP']
    }
]

EXTRA_APPS = [
    {
        'id': 'aguinaldo',
        'file': '01_APLICACIONES/Asistente_Aguinaldo_UNIFICADO.pyw',
        'icon': 'bonus_black.png',
        'emoji': '🎯',
        'name': 'Aguinaldo Unificado',
        'short': 'Asistente integral para el cálculo del SAC semestral de todo el personal.',
        'description': 'Calcula el aguinaldo (SAC) de todos los tipos de empleados.',
        'details': ['Búsqueda automática de quincenas', 'Soporta Blanco, Negro y Efectivos', 'Genera planilla de mejores sueldos', 'Cruce con índice maestro'],
        'color': 'indigo',
        'tags': ['SAC', 'Liquidación']
    },
    {
        'id': 'fechas_pdf',
        'file': '01_APLICACIONES/18-EXTRAER_FECHAS_INGRESO.pyw',
        'icon': 'dates.png',
        'emoji': '📅',
        'name': 'Extraer Fechas (PDF)',
        'short': 'Analiza recibos PDF y extrae automáticamente los datos de ingreso de los empleados.',
        'description': 'Extrae Legajo, Nombre y Fecha de Ingreso directamente desde PDFs de recibos.',
        'details': ['Extracción directa de archivos PDF', 'Identifica Legajo, Nombre y Fecha de Ingreso', 'Calcula antigüedad de forma automática', 'Soporta ordenamiento por índice externo'],
        'color': 'blue',
        'tags': ['Nuevo', 'PDF']
    },
    {
        'id': 'buscar_recibos',
        'file': '01_APLICACIONES/BuscarRecibosPDF.pyw',
        'icon': 'search.png',
        'emoji': '🔍',
        'name': 'Buscar Recibos PDF',
        'short': 'Busca y abre rápidamente cualquier recibo de sueldo en PDF filtrando por nombre o legajo.',
        'description': 'Localiza y abre recibos de sueldo individuales en PDF por nombre o legajo.',
        'details': ['Búsqueda instantánea por múltiples criterios', 'Abre el PDF directamente desde el resultado', 'Escanea subcarpetas automáticamente', 'Interfaz optimizada para acceso rápido'],
        'color': 'orange',
        'tags': ['Búsqueda']
    },
    {
        'id': 'acomodar',
        'file': '01_APLICACIONES/gui_acomodar_pdf.pyw',
        'icon': 'pdf.png',
        'emoji': '📂',
        'name': 'Acomodar PDF',
        'short': 'Organiza, separa o une archivos PDF según criterios de nombre o legajo.',
        'description': 'Organiza, separa o une archivos PDF según criterios de nombre o legajo.',
        'details': ['Separa PDFs masivos por empleado', 'Une múltiples PDFs en un solo archivo', 'Renombra archivos según contenido', 'Organiza la documentación eficientemente'],
        'color': 'rose',
        'tags': ['Gestión']
    },
    {
        'id': 'conceptos',
        'file': '01_APLICACIONES/gui_buscador_conceptos.pyw',
        'icon': 'concepts.png',
        'emoji': '🔎',
        'name': 'Buscador de Conceptos',
        'short': 'Extrae y consolida los importes de conceptos específicos desde múltiples recibos PDF.',
        'description': 'Extrae importes de conceptos específicos desde recibos PDF.',
        'details': ['Busca conceptos por nombre o código', 'Consolida resultados de múltiples archivos', 'Útil para auditoría de retenciones', 'Exporta tabla comparativa a Excel'],
        'color': 'purple',
        'tags': ['Auditoría']
    },
    {
        'id': 'promedio',
        'file': '01_APLICACIONES/Promedio_Sueldos.pyw',
        'icon': 'chart.png',
        'emoji': '📊',
        'name': 'Promedios de Sueldos',
        'short': 'Genera un informe con los promedios mensualizados de sueldos discriminados por tipo.',
        'description': 'Calcula promedios mensualizados de sueldos Blanco, Negro y Total.',
        'details': ['Calcula promedios por período y categoría', 'Discrimina entre sueldo blanco y negro', 'Exporta resultados a Excel para informes', 'Comparación histórica entre períodos'],
        'color': 'blue',
        'tags': ['Análisis']
    },
    {
        'id': 'pago',
        'file': '01_APLICACIONES/GUIA_PAGO_BANCARIO.pyw',
        'icon': 'bank.png',
        'emoji': '🏦',
        'name': 'Guía Pago Bancario',
        'short': 'Organiza y filtra los datos de pago por banco para transferencias masivas.',
        'description': 'Filtra empleados por banco y actualiza la guía de pago bancario.',
        'details': ['Filtra empleados por entidad bancaria', 'Genera la guía de pago en formato oficial', 'Actualiza automáticamente los importes', 'Soporta múltiples entidades bancarias'],
        'color': 'green',
        'tags': ['Finanzas']
    },
    {
        'id': 'epp',
        'file': '01_APLICACIONES/generar_epp_excel.pyw',
        'icon': 'receipts.png',
        'emoji': '🛡️',
        'name': 'Formulario EPP',
        'short': 'Genera las planillas oficiales de entrega de Elementos de Protección Personal.',
        'description': 'Genera planillas de entrega de Elementos de Protección Personal.',
        'details': ['Selección de empleados desde listado activo', 'Checklist de EPP entregados por fecha', 'Genera Excel con formato de planilla legal', 'Registro histórico por cada operario'],
        'color': 'orange',
        'tags': ['Seguridad']
    },
    {
        'id': 'planilla',
        'file': '01_APLICACIONES/gui_planilla.pyw',
        'icon': 'bonus_assist.png',
        'emoji': '📑',
        'name': 'Planilla x Índice',
        'short': 'Herramienta para acomodar y ordenar planillas Excel basándose en el índice maestro.',
        'description': 'Herramienta para acomodar y ordenar planillas Excel basándose en el índice.',
        'details': ['Ordena planillas Excel automáticamente', 'Sincroniza con el índice maestro de personal', 'Ajusta filas y columnas por legajo', 'Cruce de datos inteligente entre archivos'],
        'color': 'teal',
        'tags': ['Excel']
    },
    {
        'id': 'firmador',
        'file': '01_APLICACIONES/Firmador_Masivo_PDF.pyw',
        'icon': 'pdf.png',
        'emoji': '✍️',
        'name': 'Firmador Masivo PDF',
        'short': 'Estampa una firma PNG en todas las hojas de un PDF con posición y escala personalizable.',
        'description': 'Permite estampar una firma PNG en todas las hojas de un PDF masivamente.',
        'details': ['Selección de PDF y archivo de firma (PNG)', 'Posiciones predefinidas y ajuste de escala', 'Genera un nuevo PDF sin alterar el original', 'Optimizado para firmas masivas de conformidad'],
        'color': 'teal',
        'tags': ['PDF', 'Firmas']
    }
]

# =============================================================================
# COMPONENTES VISUALES OPTIMIZADOS
# =============================================================================

class WorkflowBar(ctk.CTkFrame):
    def __init__(self, parent, **kwargs):
        fg_color = kwargs.pop('fg_color', mgc.COLORS['bg_card'])
        super().__init__(parent, fg_color=fg_color, border_color=mgc.COLORS['border'], border_width=1, corner_radius=16, **kwargs)
        
        inner = ctk.CTkFrame(self, fg_color="transparent")
        inner.pack(fill=tk.BOTH, expand=True, padx=30, pady=15)
        
        ctk.CTkLabel(inner, text="FLUJO DE TRABAJO RECOMENDADO", font=('Segoe UI', 12, 'bold'), 
                     text_color=mgc.COLORS['accent_blue']).pack(anchor='w', pady=(0, 20))
        
        self.steps_frame = ctk.CTkFrame(inner, fg_color="transparent")
        self.steps_frame.pack(fill=tk.X)
        
        steps = [
            ("1", "Preparar Quincena", "Resetear archivo Excel", mgc.COLORS['rose']),
            ("2", "Horas Depósito", "Carga inicial de logística", mgc.COLORS['teal']),
            ("3", "Calcular Horas", "Procesar jornales UOCRA", mgc.COLORS['blue']),
            ("4", "Procesar Sueldos", "Cierre final de nómina", mgc.COLORS['purple']),
            ("5", "Generar Recibos", "Crear PDFs individuales", mgc.COLORS['green']),
            ("6", "Enviar Documentación", "Distribuir por Email", mgc.COLORS['cyan'])
        ]
        
        for i, (num, name, subtitle, color) in enumerate(steps):
            step = ctk.CTkFrame(self.steps_frame, fg_color="transparent")
            step.pack(side=tk.LEFT, expand=True)
            
            content = ctk.CTkFrame(step, fg_color="transparent")
            content.pack()
            
            ctk.CTkLabel(content, text=num, width=28, height=28, corner_radius=14, 
                         fg_color=color, text_color="white", font=('Segoe UI', 12, 'bold')).pack(side=tk.LEFT, padx=(0, 12))
            
            info = ctk.CTkFrame(content, fg_color="transparent")
            info.pack(side=tk.LEFT)
            
            ctk.CTkLabel(info, text=name, font=('Segoe UI', 14, 'bold'), 
                         text_color=mgc.COLORS['text_primary']).pack(anchor='w')
            ctk.CTkLabel(info, text=subtitle, font=('Segoe UI', 11), 
                         text_color=mgc.COLORS['text_secondary']).pack(anchor='w')
            
            if i < len(steps) - 1:
                ctk.CTkLabel(self.steps_frame, text="→", font=('Segoe UI', 16), 
                             text_color=mgc.COLORS['border']).pack(side=tk.LEFT, padx=8)

# =============================================================================
# APLICACIÓN PRINCIPAL (Optimización de Carga y Redibujado)
# =============================================================================

class ModernLauncher(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.withdraw() # Ocultar mientras se construye para evitar parpadeo
        
        self.title("🚀 Sistema de Sueldos — Launcher PRO")
        mgc.center_window(self, 1200, 950)
        
        # Header
        self.header = mgc.create_header(self, "Sueldos Carjor", "Ecosistema inteligente para la liquidación masiva y gestión de personal.")
        self.header.pack(fill=tk.X, padx=150, pady=(20, 0))
        
        # Main Scrollable Area
        self.scroll = ctk.CTkScrollableFrame(self, fg_color="transparent", corner_radius=0,
                                             scrollbar_button_color=mgc.COLORS['border'],
                                             scrollbar_button_hover_color=mgc.COLORS['blue'])
        self.scroll.pack(fill=tk.BOTH, expand=True, padx=150, pady=10)
        
        # Status Bar
        self.status_frame, self.status_var = mgc.create_status_bar(self, "✦ Cargando ecosistema inteligente...")

        # Carga diferida de iconos y renderizado para arranque instantáneo
        self.icons = {}
        self.after(10, self.async_init)

    def async_init(self):
        """Carga asíncrona simulada de componentes pesados"""
        self.workflow = WorkflowBar(self.scroll, fg_color=mgc.COLORS['bg_card']) 
        self.workflow.pack(fill=tk.X, pady=(10, 20))
        
        self.apps_container = ctk.CTkFrame(self.scroll, fg_color="transparent")
        self.apps_container.pack(fill=tk.BOTH, expand=True)

        self.load_icons_batch() # Carga optimizada
        self.render_all_sections()
        
        self.deiconify() # Mostrar ventana ya lista
        self.status_var.set("✦ Pasá el mouse sobre una aplicación para ver detalles — hacé clic para abrirla")

    def load_icons_batch(self):
        """Carga de iconos usando CTkImage para mejor rendimiento y escalado DPI"""
        icons_dir = mgc.get_resource_path(os.path.join("02_CARPETAS", "launcher_icons"))
        all_apps = MAIN_APPS + EXTRA_APPS
        
        # Evitar recargas duplicadas
        unique_icons = {app.get('icon') for app in all_apps if app.get('icon')}
        
        for icon_file in unique_icons:
            path = os.path.join(icons_dir, icon_file)
            if os.path.exists(path):
                try:
                    img = Image.open(path)
                    # CTkImage es mucho más eficiente que PhotoImage en CustomTkinter
                    self.icons[icon_file] = ctk.CTkImage(light_image=img, dark_image=img, size=(32, 32))
                except Exception:
                    self.icons[icon_file] = None

    def render_all_sections(self):
        self.create_section_label(self.apps_container, "⚙️ Aplicaciones principales")
        self.main_grid = ctk.CTkFrame(self.apps_container, fg_color="transparent")
        self.main_grid.pack(fill=tk.X, pady=(0, 20))
        self.main_grid.columnconfigure((0, 1, 2), weight=1, uniform="col")
        self.render_apps(self.main_grid, MAIN_APPS)

        self.create_section_label(self.apps_container, "🔧 Herramientas adicionales")
        self.extra_grid = ctk.CTkFrame(self.apps_container, fg_color="transparent")
        self.extra_grid.pack(fill=tk.X, pady=(0, 40))
        self.extra_grid.columnconfigure((0, 1, 2), weight=1, uniform="col")
        self.render_apps(self.extra_grid, EXTRA_APPS)
        
        footer = ctk.CTkLabel(self.apps_container, text="Sistema de Sueldos · Gestión Interna de Nómina · v2.2 (Optimized)", 
                             font=('Segoe UI', 11), text_color=mgc.COLORS['text_secondary'])
        footer.pack(pady=(20, 40))

    def create_section_label(self, parent, text):
        frame = ctk.CTkFrame(parent, fg_color="transparent")
        frame.pack(fill=tk.X, pady=(20, 15))
        ctk.CTkLabel(frame, text=text, font=('Segoe UI', 14, 'bold'), 
                     text_color=mgc.COLORS['text_secondary']).pack(side=tk.LEFT)
        ctk.CTkFrame(frame, height=1, fg_color=mgc.COLORS['border']).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(15, 0))

    def render_apps(self, container, apps):
        row, col = 0, 0
        for app in apps:
            icon_img = self.icons.get(app.get('icon'))
            description = app.get('short', app.get('description'))
            card_outer, _ = mgc.create_web_card(
                container, title=app['name'], description=description, icon_image=icon_img,
                color=app.get('color', 'blue'), tags=app.get('tags', []), command=lambda a=app: self.launch_app(a),
                details=app.get('details', []), emoji=app.get('emoji', '🚀'), badge_text=app.get('badge', 'Disponible'),
                on_hover_enter=lambda e, a=app: self.status_var.set(f"▸ {a['name'].upper()} — {a.get('description')}"),
                on_hover_leave=lambda e: self.status_var.set("✦ Pasá el mouse sobre una aplicación para ver detalles — hacé clic para abrirla")
            )
            card_outer.grid(row=row, column=col, sticky='nsew', padx=18, pady=12)
            col += 1
            if col == 3: col = 0; row += 1

    def launch_app(self, app):
        """ Lanza la aplicación secundaria con manejo de entorno optimizado """
        try:
            self.status_var.set(f"🚀 Iniciando {app['name']}...")
            app_path = mgc.get_resource_path(app['file'])
            
            if os.path.exists(app_path):
                new_env = os.environ.copy()
                new_env['CHILD_APP_MODE'] = '1'
                # Usar creationflags para evitar ventanas de consola si se desea (en Windows)
                subprocess.Popen([sys.executable, app_path], env=new_env)
                self.after(2000, lambda: self.status_var.set(f"✓ {app['name']} ejecutándose"))
            else:
                messagebox.showerror("Error", f"No se encontró el archivo:\n{app_path}")
        except Exception as e:
            messagebox.showerror("Error de ejecución", f"No se pudo iniciar la aplicación:\n{str(e)}")

def start_launcher():
    app = ModernLauncher()
    app.mainloop()

# =============================================================================
# INICIO
# =============================================================================
# INICIO
# =============================================================================

if __name__ == "__main__":
    if len(sys.argv) > 2 and sys.argv[1] == "--run":
        run_sub_app(sys.argv[2])
    else:
        start_launcher()
