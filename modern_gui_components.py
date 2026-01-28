# -*- coding: utf-8 -*-
"""
Módulo de Componentes GUI Modernos
Proporciona estilos, colores y componentes reutilizables para aplicaciones con Tkinter
"""

import tkinter as tk
from tkinter import ttk

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
    'shadow': '#00000015'          # Sombra
}

# =============================================================================
# CONFIGURACIÓN DE FUENTES
# =============================================================================
FONTS = {
    'title': ('Segoe UI', 28, 'bold'),
    'subtitle': ('Segoe UI', 14, 'bold'),
    'heading': ('Segoe UI', 11, 'bold'),
    'normal': ('Segoe UI', 10),
    'button': ('Segoe UI', 11, 'bold'),
    'button_large': ('Segoe UI', 12, 'bold'),
    'small': ('Segoe UI', 9),
    'tiny': ('Segoe UI', 8)
}

# =============================================================================
# CLASE BASE PARA VENTANAS MODERNAS
# =============================================================================
class ModernWindow:
    """Clase base para ventanas con estilo moderno"""
    
    def __init__(self, root, title, width=800, height=600, resizable=True):
        self.root = root
        self.root.title(title)
        self.root.geometry(f"{width}x{height}")
        self.root.configure(bg=COLORS['bg_primary'])
        
        if not resizable:
            self.root.resizable(False, False)
        
        # Configurar estilos ttk
        self._configure_styles()
    
    def _configure_styles(self):
        """Configura los estilos ttk para la ventana"""
        style = ttk.Style(self.root)
        
        # Estilo para botones
        style.configure('Modern.TButton',
                       font=FONTS['button'],
                       padding=10)
        
        # Estilo para labels
        style.configure('Modern.TLabel',
                       font=FONTS['normal'],
                       background=COLORS['bg_primary'])
        
        style.configure('Title.TLabel',
                       font=FONTS['title'],
                       background=COLORS['bg_primary'],
                       foreground=COLORS['text_primary'])
        
        style.configure('Subtitle.TLabel',
                       font=FONTS['subtitle'],
                       background=COLORS['bg_primary'],
                       foreground=COLORS['text_secondary'])
        
        style.configure('Heading.TLabel',
                       font=FONTS['heading'],
                       background=COLORS['bg_card'],
                       foreground=COLORS['text_primary'])
        
        # Estilo para frames
        style.configure('Card.TFrame',
                       background=COLORS['bg_card'],
                       relief=tk.RAISED,
                       borderwidth=1)
        
        style.configure('Modern.TFrame',
                       background=COLORS['bg_primary'])
        
        # Estilo para LabelFrame
        style.configure('Modern.TLabelframe',
                       background=COLORS['bg_card'],
                       borderwidth=2,
                       relief=tk.SOLID)
        
        style.configure('Modern.TLabelframe.Label',
                       font=FONTS['heading'],
                       background=COLORS['bg_card'],
                       foreground=COLORS['text_primary'])
        
        # Estilo para Entry
        style.configure('Modern.TEntry',
                       font=FONTS['normal'])
        
        # Estilo para barra de progreso
        style.configure('Modern.Horizontal.TProgressbar',
                       troughcolor=COLORS['border'],
                       background=COLORS['blue'],
                       bordercolor=COLORS['border'],
                       lightcolor=COLORS['blue'],
                       darkcolor=COLORS['blue'],
                       thickness=10)

# =============================================================================
# FUNCIONES HELPER PARA CREAR COMPONENTES
# =============================================================================

def create_header(parent, title, subtitle=None, icon="", icon_image=None):
    """
    Crea un header moderno con título y subtítulo opcional
    
    Args:
        parent: Widget padre
        title: Texto del título
        subtitle: Texto del subtítulo (opcional)
        icon: Emoji o icono (opcional, fallback si no hay icon_image)
        icon_image: PhotoImage del icono PNG (opcional, preferido sobre icon)
    
    Returns:
        Frame contenedor del header
    """
    header_frame = tk.Frame(parent, bg=COLORS['bg_primary'])
    header_frame.pack(fill=tk.X, pady=(0, 20))
    
    # Si hay icono PNG, mostrarlo
    if icon_image:
        icon_label = tk.Label(header_frame,
                             image=icon_image,
                             bg=COLORS['bg_primary'])
        icon_label.pack(pady=(0, 10))
        icon_label.image = icon_image  # Mantener referencia
        
        # Título sin icono de texto
        title_label = tk.Label(header_frame, 
                              text=title,
                              font=FONTS['title'],
                              bg=COLORS['bg_primary'],
                              fg=COLORS['text_primary'])
        title_label.pack(pady=(0, 5))
    else:
        # Título con icono emoji (comportamiento original)
        title_text = f"{icon} {title}" if icon else title
        title_label = tk.Label(header_frame, 
                              text=title_text,
                              font=FONTS['title'],
                              bg=COLORS['bg_primary'],
                              fg=COLORS['text_primary'])
        title_label.pack(pady=(0, 5))
    
    # Subtítulo
    if subtitle:
        subtitle_label = tk.Label(header_frame,
                                 text=subtitle,
                                 font=FONTS['normal'],
                                 bg=COLORS['bg_primary'],
                                 fg=COLORS['text_secondary'])
        subtitle_label.pack()
    
    return header_frame


def create_card(parent, title=None, padding=20):
    """
    Crea una card (tarjeta) con estilo moderno
    
    Args:
        parent: Widget padre
        title: Título de la card (opcional)
        padding: Padding interno
    
    Returns:
        Frame interno de la card donde agregar contenido
    """
    # Frame externo (card)
    card_outer = tk.Frame(parent, bg=COLORS['bg_primary'])
    
    if title:
        # LabelFrame si tiene título
        card = ttk.LabelFrame(card_outer, 
                             text=title,
                             style='Modern.TLabelframe',
                             padding=padding)
    else:
        # Frame simple con borde
        card = tk.Frame(card_outer, 
                       bg=COLORS['bg_card'],
                       relief=tk.RAISED,
                       bd=1)
    
    card.pack(fill=tk.BOTH, expand=True, padx=2, pady=2)
    
    # Frame interno para contenido
    inner = tk.Frame(card, bg=COLORS['bg_card'])
    inner.pack(fill=tk.BOTH, expand=True, padx=padding, pady=padding)
    
    return card_outer, inner


def create_button(parent, text, command, color='blue', icon="", icon_image=None, text_color='white', **kwargs):
    """
    Crea un botón moderno con estilo
    
    Args:
        parent: Widget padre
        text: Texto del botón
        command: Función a ejecutar
        color: Color del botón ('blue', 'green', 'purple', 'red', 'gray', 'orange')
        icon: Emoji o icono (opcional, fallback si no hay icon_image)
        icon_image: PhotoImage del icono PNG (opcional)
        text_color: Color del texto (default: 'white')
        **kwargs: Argumentos adicionales para el botón
    
    Returns:
        Button widget
    """
    bg_color = COLORS.get(color, COLORS['blue'])
    fg_color = COLORS.get(text_color, text_color)
    
    # Si hay icono PNG, usarlo
    if icon_image:
        default_kwargs = {
            'text': text,
            'image': icon_image,
            'compound': tk.LEFT,  # Icono a la izquierda del texto
            'command': command,
            'font': FONTS['button'],
            'bg': bg_color,
            'fg': fg_color,
            'relief': tk.FLAT,
            'padx': 20,
            'pady': 10,
            'cursor': 'hand2'
        }
    else:
        # Usar emoji (comportamiento original)
        button_text = f"{icon} {text}" if icon else text
        default_kwargs = {
            'text': button_text,
            'command': command,
            'font': FONTS['button'],
            'bg': bg_color,
            'fg': fg_color,
            'relief': tk.FLAT,
            'padx': 20,
            'pady': 10,
            'cursor': 'hand2'
        }
    
    # Combinar kwargs
    default_kwargs.update(kwargs)
    
    button = tk.Button(parent, **default_kwargs)
    
    # Mantener referencia al icono si existe
    if icon_image:
        button.image = icon_image
    
    return button


def create_large_button(parent, text, command, color='green', icon="", icon_image=None, text_color='white', **kwargs):
    """
    Crea un botón grande para acciones principales
    
    Args:
        parent: Widget padre
        text: Texto del botón
        command: Función a ejecutar
        color: Color del botón
        icon: Emoji o icono (opcional, fallback si no hay icon_image)
        icon_image: PhotoImage del icono PNG (opcional)
        text_color: Color del texto (default: 'white')
        **kwargs: Argumentos adicionales
    
    Returns:
        Button widget
    """
    bg_color = COLORS.get(color, COLORS['green'])
    fg_color = COLORS.get(text_color, text_color)
    
    # Si hay icono PNG, usarlo
    if icon_image:
        default_kwargs = {
            'text': text,
            'image': icon_image,
            'compound': tk.LEFT,
            'command': command,
            'font': FONTS['button_large'],
            'bg': bg_color,
            'fg': fg_color,
            'relief': tk.FLAT,
            'padx': 40,
            'pady': 15,
            'cursor': 'hand2'
        }
    else:
        # Usar emoji (comportamiento original)
        button_text = f"{icon} {text}" if icon else text
        default_kwargs = {
            'text': button_text,
            'command': command,
            'font': FONTS['button_large'],
            'bg': bg_color,
            'fg': fg_color,
            'relief': tk.FLAT,
            'padx': 40,
            'pady': 15,
            'cursor': 'hand2'
        }
    
    default_kwargs.update(kwargs)
    
    button = tk.Button(parent, **default_kwargs)
    
    # Mantener referencia al icono si existe
    if icon_image:
        button.image = icon_image
    
    return button


def create_status_bar(parent, initial_text="Listo"):
    """
    Crea una barra de estado en la parte inferior
    
    Args:
        parent: Widget padre (generalmente root)
        initial_text: Texto inicial
    
    Returns:
        Tuple (frame, StringVar) - Frame de la barra y variable de texto
    """
    status_var = tk.StringVar(value=initial_text)
    
    status_frame = tk.Frame(parent, bg=COLORS['border'], relief=tk.SUNKEN, bd=1)
    status_frame.pack(side=tk.BOTTOM, fill=tk.X)
    
    status_label = tk.Label(status_frame,
                           textvariable=status_var,
                           font=FONTS['small'],
                           bg=COLORS['bg_card'],
                           fg=COLORS['text_secondary'],
                           anchor='w',
                           padx=10,
                           pady=5)
    status_label.pack(side=tk.LEFT, fill=tk.X, expand=True)
    
    return status_frame, status_var


def create_file_selector(parent, label_text, string_var, button_command, icon="📁", icon_image=None):
    """
    Crea un selector de archivo con label, entry y botón
    
    Args:
        parent: Widget padre
        label_text: Texto del label
        string_var: StringVar para el path del archivo
        button_command: Función para el botón
        icon: Icono del botón (emoji, fallback)
        icon_image: PhotoImage del icono PNG (opcional)
    
    Returns:
        Frame contenedor
    """
    container = tk.Frame(parent, bg=COLORS['bg_card'])
    
    # Label
    label = tk.Label(container,
                    text=label_text,
                    font=FONTS['normal'],
                    bg=COLORS['bg_card'],
                    fg=COLORS['text_primary'])
    label.grid(row=0, column=0, sticky='w', pady=5, padx=5)
    
    # Entry (readonly)
    entry = tk.Entry(container,
                    textvariable=string_var,
                    font=FONTS['normal'],
                    state='readonly',
                    bg='white',
                    relief=tk.GROOVE,
                    bd=1)
    entry.grid(row=0, column=1, sticky='we', padx=5)
    
    # Botón
    button = create_button(container,
                          text="Seleccionar",
                          command=button_command,
                          color='purple',
                          icon=icon,
                          icon_image=icon_image,
                          padx=15,
                          pady=8)
    button.grid(row=0, column=2, padx=5)
    
    # Configurar expansión
    container.columnconfigure(1, weight=1)
    
    return container


def create_progress_section(parent):
    """
    Crea una sección de progreso con barra y label
    
    Args:
        parent: Widget padre
    
    Returns:
        Tuple (progressbar, label_var) - Barra de progreso y variable de texto
    """
    container = tk.Frame(parent, bg=COLORS['bg_card'])
    
    # Label de progreso
    progress_var = tk.StringVar(value="Esperando...")
    progress_label = tk.Label(container,
                             textvariable=progress_var,
                             font=FONTS['normal'],
                             bg=COLORS['bg_card'],
                             fg=COLORS['text_secondary'])
    progress_label.pack(pady=(0, 10))
    
    # Barra de progreso
    progress_bar = ttk.Progressbar(container,
                                  length=300,
                                  mode='determinate',
                                  style='Modern.Horizontal.TProgressbar')
    progress_bar.pack(fill=tk.X, pady=(0, 5))
    
    return container, progress_bar, progress_var


def create_icon_label(parent, icon, text, color='text_primary'):
    """
    Crea un label con icono grande (para cards de menú)
    
    Args:
        parent: Widget padre
        icon: Emoji grande
        text: Texto descriptivo
        color: Color del texto
    
    Returns:
        Frame contenedor
    """
    container = tk.Frame(parent, bg=COLORS['bg_card'])
    
    # Icono grande
    icon_label = tk.Label(container,
                         text=icon,
                         font=('Segoe UI', 48),
                         bg=COLORS['bg_card'])
    icon_label.pack(pady=(0, 10))
    
    # Texto
    text_label = tk.Label(container,
                         text=text,
                         font=FONTS['subtitle'],
                         bg=COLORS['bg_card'],
                         fg=COLORS[color])
    text_label.pack()
    
    return container


# =============================================================================
# UTILIDADES
# =============================================================================

def center_window(root, width, height):
    """Centra una ventana en la pantalla"""
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    
    x = (screen_width - width) // 2
    y = (screen_height - height) // 2
    
    root.geometry(f"{width}x{height}+{x}+{y}")


def disable_button(button, disabled_color='gray'):
    """Deshabilita un botón y cambia su color"""
    button.config(state='disabled', bg=COLORS[disabled_color], disabledforeground='white')


def enable_button(button, color='blue'):
    """Habilita un botón y restaura su color"""
    button.config(state='normal', bg=COLORS[color])
