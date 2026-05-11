# -*- coding: utf-8 -*-
"""
Módulo de Componentes GUI Modernos usando CustomTkinter
Proporciona estilos, colores y componentes reutilizables "Ultra Premium"
"""

import tkinter as tk
import customtkinter as ctk
import os
import tempfile

# Monkey-patch para retrocompatibilidad con apps que hacían btn.config(bg=...)
def _ctk_config_patch(self, **kwargs):
    pass
ctk.CTkButton.config = _ctk_config_patch

# Configuración Base de CustomTkinter
ctk.set_appearance_mode("Light")  # Match the current bright/clean style
ctk.set_default_color_theme("blue")

# =============================================================================
# PALETA DE COLORES MODERNA
# =============================================================================
COLORS = {
    'bg_primary': '#f5f7fa',      # Fondo principal
    'bg_card': '#ffffff',          # Fondo de cards
    'purple': '#8b5cf6',           # Morado vibrante
    'blue': '#3b82f6',             # Azul
    'green': '#10b981',            # Verde éxito
    'orange': '#f59e0b',           # Naranja advertencia
    'red': '#ef4444',              # Rojo error
    'gray': '#6b7280',             # Gris secundario
    'text_primary': '#1f2937',     # Texto principal
    'text_secondary': '#6b7280',   # Texto secundario
    'border': '#e5e7eb',           # Bordes
    'shadow': '#000000'            # Sombra base
}

HOVER_COLORS = {
    'purple': '#7c3aed',
    'blue': '#2563eb',
    'green': '#059669',
    'orange': '#d97706',
    'red': '#dc2626',
    'gray': '#4b5563'
}

# =============================================================================
# CONFIGURACIÓN DE FUENTES
# =============================================================================
FONTS = {
    'title': ('Segoe UI', 32, 'bold'),
    'subtitle': ('Segoe UI', 16, 'bold'),
    'heading': ('Segoe UI', 14, 'bold'),
    'normal': ('Segoe UI', 13),
    'button': ('Segoe UI', 14, 'bold'),
    'button_large': ('Segoe UI', 16, 'bold'),
    'small': ('Segoe UI', 11),
    'tiny': ('Segoe UI', 10)
}

# =============================================================================
# CLASE BASE PARA VENTANAS MODERNAS
# =============================================================================
class ModernWindow:
    """Clase base para ventanas con estilo moderno (Usa CTk)"""
    def __init__(self, root, title, width=800, height=600, resizable=True):
        self.root = root  # Root nativo de tk o ctk (si es tk puro, soportaremos CTk interno)
        self.root.title(title)
        self.root.geometry(f"{width}x{height}")
        
        try:
            self.root.configure(bg=COLORS['bg_primary'])
            if hasattr(self.root, 'config'):
                self.root.config(bg=COLORS['bg_primary'])
        except:
            pass

        if not resizable:
            self.root.resizable(False, False)


# =============================================================================
# FUNCIONES HELPER PARA CREAR COMPONENTES
# =============================================================================

def create_header(parent, title, subtitle=None, icon="", icon_image=None):
    """Crea un header moderno con título y subtítulo opcional usando CTkLabel"""
    header_frame = ctk.CTkFrame(parent, fg_color="transparent")
    header_frame.pack(fill=tk.X, pady=(0, 20))
    
    if icon_image:
        # Use ctk.CTkImage internally if possible, but tk.PhotoImage is supported by CTkLabel
        icon_label = ctk.CTkLabel(header_frame, text="", image=icon_image)
        icon_label.pack(pady=(0, 10))
        icon_label.image = icon_image
        
        title_label = ctk.CTkLabel(header_frame, text=title, font=FONTS['title'], text_color=COLORS['text_primary'])
        title_label.pack(pady=(0, 5))
    else:
        title_text = f"{icon} {title}" if icon else title
        title_label = ctk.CTkLabel(header_frame, text=title_text, font=FONTS['title'], text_color=COLORS['text_primary'])
        title_label.pack(pady=(0, 5))
    
    if subtitle:
        subtitle_label = ctk.CTkLabel(header_frame, text=subtitle, font=FONTS['normal'], text_color=COLORS['text_secondary'])
        subtitle_label.pack()
    
    return header_frame


def create_card(parent, title=None, padding=20):
    """Crea una tarjeta UI con bordes redondeados nativos de CustomTkinter"""
    card_outer = ctk.CTkFrame(parent, fg_color="transparent")
    
    if title:
        title_lbl = ctk.CTkLabel(card_outer, text=title, font=FONTS['heading'], text_color=COLORS['text_primary'])
        title_lbl.pack(anchor='w', padx=8, pady=(0, 5))
    
    card = ctk.CTkFrame(card_outer, fg_color=COLORS['bg_card'], 
                        border_color=COLORS['border'], border_width=1, corner_radius=12)
    card.pack(fill=tk.BOTH, expand=True)
    
    inner = ctk.CTkFrame(card, fg_color="transparent")
    inner.pack(fill=tk.BOTH, expand=True, padx=padding, pady=padding)
    
    return card_outer, inner


def create_button(parent, text, command, color='blue', icon="", icon_image=None, text_color='white', **kwargs):
    """Crea un botón redondeado interactivo animado CTkButton"""
    bg_color = COLORS.get(color, COLORS['blue'])
    h_color = HOVER_COLORS.get(color, bg_color)
    f_color = COLORS.get(text_color, text_color)
    
    btn_text = f"{icon} {text}" if icon and not icon_image else text
    
    button = ctk.CTkButton(parent, text=btn_text, image=icon_image,
                           command=command,
                           font=FONTS['button'],
                           fg_color=bg_color,
                           hover_color=h_color,
                           text_color=f_color,
                           corner_radius=8,
                           cursor="hand2")
                           
    if icon_image:
        button.image = icon_image
    
    return button


def create_large_button(parent, text, command, color='green', icon="", icon_image=None, text_color='white', **kwargs):
    """Botón grande para acciones destacadas"""
    bg_color = COLORS.get(color, COLORS['green'])
    h_color = HOVER_COLORS.get(color, bg_color)
    f_color = COLORS.get(text_color, text_color)
    
    btn_text = f"{icon} {text}" if icon and not icon_image else text
    
    button = ctk.CTkButton(parent, text=btn_text, image=icon_image,
                           command=command,
                           font=FONTS['button_large'],
                           fg_color=bg_color,
                           hover_color=h_color,
                           text_color=f_color,
                           corner_radius=12,
                           height=45,
                           cursor="hand2")
                           
    if icon_image:
        button.image = icon_image
        
    return button


def create_status_bar(parent, initial_text="Listo"):
    """Barra de estado inferior limpia"""
    status_var = tk.StringVar(value=initial_text)
    
    status_frame = ctk.CTkFrame(parent, fg_color=COLORS['bg_card'], corner_radius=0, border_width=1, border_color=COLORS['border'], height=35)
    status_frame.pack(side=tk.BOTTOM, fill=tk.X)
    
    status_label = ctk.CTkLabel(status_frame, textvariable=status_var, font=FONTS['small'], text_color=COLORS['text_secondary'])
    status_label.pack(side=tk.LEFT, padx=15, pady=5)
    
    return status_frame, status_var


def create_file_selector(parent, label_text, string_var, button_command, icon="📁", icon_image=None):
    """Selector de archivo estilizado y redondeado"""
    container = ctk.CTkFrame(parent, fg_color="transparent")
    
    label = ctk.CTkLabel(container, text=label_text, font=FONTS['normal'], text_color=COLORS['text_primary'], width=200, anchor="w")
    label.grid(row=0, column=0, sticky='w', padx=5, pady=5)
    
    # Campo de entrada redondeado
    entry = ctk.CTkEntry(container, textvariable=string_var, font=FONTS['normal'], 
                         fg_color='#f8fafc', border_color=COLORS['border'], border_width=1, corner_radius=6, state="readonly")
    entry.grid(row=0, column=1, sticky='we', padx=8, pady=5)
    
    button = create_button(container, "Seleccionar", button_command, color='purple', icon=icon, icon_image=icon_image)
    button.grid(row=0, column=2, padx=5, pady=5)
    
    container.columnconfigure(0, weight=0)
    container.columnconfigure(1, weight=1)
    container.columnconfigure(2, weight=0)
    
    return container


def create_progress_section(parent):
    """Barra de progreso redondeada"""
    container = ctk.CTkFrame(parent, fg_color="transparent")
    
    progress_var = tk.StringVar(value="Esperando...")
    progress_label = ctk.CTkLabel(container, textvariable=progress_var, font=FONTS['normal'], text_color=COLORS['text_secondary'])
    progress_label.pack(pady=(0, 8))
    
    progress_bar = ctk.CTkProgressBar(container, width=300, height=12, corner_radius=6, 
                                      fg_color=COLORS['border'], progress_color=COLORS['blue'])
    progress_bar.set(0) # Inicializar vacía
    progress_bar.pack(fill=tk.X, pady=(0, 5))
    
    # Monkey-patch para que asimile la API de tk.ttk.Progressbar
    class ProgressBarWrapper:
        def __init__(self, ctk_pb):
            self.pb = ctk_pb
        def __setitem__(self, key, value):
            if key == 'value':
                self.pb.set(value / 100.0) # CTkProgressbar usa escala 0-1
    
    wrapper = ProgressBarWrapper(progress_bar)
    return container, wrapper, progress_var


def create_icon_label(parent, icon, text, color='text_primary'):
    """Mini tarjeta de menú con super icono emoji"""
    container = ctk.CTkFrame(parent, fg_color="transparent")
    
    icon_label = ctk.CTkLabel(container, text=icon, font=('Segoe UI', 46), text_color=COLORS['text_primary'])
    icon_label.pack(pady=(0, 8))
    
    text_label = ctk.CTkLabel(container, text=text, font=FONTS['subtitle'], text_color=COLORS.get(color, COLORS['text_primary']))
    text_label.pack()
    
    return container


# =============================================================================
# UTILIDADES GLOBALES (COMPATIBILIDAD)
# =============================================================================

def _get_cascade_offset():
    try:
        temp_dir = tempfile.gettempdir()
        counter_file = os.path.join(temp_dir, "mgc_cascade_counter.txt")
        count = 0
        if os.path.exists(counter_file):
            try:
                with open(counter_file, "r") as f:
                    count = int(f.read().strip())
            except:
                pass
        next_count = (count + 1) % 8
        with open(counter_file, "w") as f:
            f.write(str(next_count))
        return count
    except:
        return 0

def center_window(root, width, height):
    """
    Centra la ventana en pantalla sin que quede cortada.
    - Standalone (abierta sola): centrada perfectamente.
    - Hija del launcher (CHILD_APP_MODE=1): centrada en área disponible
      con efecto cascada a medida que se abren más apps.
    """
    import os
    screen_width  = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()

    # Márgenes para barra de tareas y barra de título del SO
    taskbar_margin  = 70
    titlebar_margin = 30

    # Tamaño efectivo: nunca superar el área visible
    eff_width  = min(width,  screen_width  - 20)
    eff_height = min(height, screen_height - taskbar_margin - titlebar_margin)

    # Detectar si fue lanzada desde el launcher mediante variable de entorno
    es_app_hija = os.environ.get('CHILD_APP_MODE') == '1'

    if es_app_hija:
        # Área disponible: lado izquierdo (el launcher compacto ocupa ~460px a la derecha)
        available_width = screen_width - 460
        
        # FORZAR DIMENSIONES ESTÁNDAR (900x700) SEGÚN REQUERIMIENTO
        eff_width = 900
        eff_height = 700
        
        count = _get_cascade_offset()

        # Primera app centrada en el área disponible; las siguientes en la misma posición (sin cascada)
        x = max(10, (available_width - eff_width) // 2)
        y = max(titlebar_margin, (screen_height - eff_height) // 2)
        # x += count * 25  # Comentado para mantener posición fija
        # y += count * 25  # Comentado para mantener posición fija

        # Evitar que la cascada saque la ventana de pantalla
        x = min(x, available_width - eff_width - 10)
        y = min(y, screen_height - eff_height - taskbar_margin)
    else:
        # Standalone o el propio launcher: perfectamente centrada
        x = (screen_width  - eff_width)  // 2
        y = max(titlebar_margin, (screen_height - eff_height) // 2)

    x = max(10, int(x))
    y = max(titlebar_margin, int(y))

    root.update_idletasks()
    root.geometry(f"{eff_width}x{eff_height}+{x}+{y}")
    # Evitar que al redimensionar manualmente supere el área visible
    root.maxsize(screen_width - 10, screen_height - taskbar_margin)
    root.minsize(900, 700) # Mantener mínimo estándar

def create_main_container(parent, padding=15):
    """
    Crea un contenedor principal con scrollbar automático.
    Usa CTkScrollableFrame para asegurar que todo el contenido sea accesible.
    """
    container = ctk.CTkScrollableFrame(parent, fg_color="transparent", 
                                       corner_radius=0)
    container.pack(fill=tk.BOTH, expand=True, padx=padding, pady=padding)
    return container

def disable_button(button, disabled_color='gray'):
    button.configure(state='disabled', fg_color=COLORS[disabled_color])

def enable_button(button, color='blue'):
    button.configure(state='normal', fg_color=COLORS[color])
