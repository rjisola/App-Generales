# -*- coding: utf-8 -*-
"""
Módulo de Componentes GUI Modernos usando CustomTkinter
Proporciona estilos, colores y componentes reutilizables "Ultra Premium"
"""

import tkinter as tk
import customtkinter as ctk
import os
import tempfile
from PIL import Image, ImageDraw, ImageTk, ImageFilter, ImageOps

# Monkey-patch para retrocompatibilidad profunda con apps que hacían btn.config(bg=...) etc.
def _ctk_config_patch(self, **kwargs):
    if 'bg' in kwargs: kwargs['fg_color'] = kwargs.pop('bg')
    if 'fg' in kwargs: kwargs['text_color'] = kwargs.pop('fg')
    if 'highlightbackground' in kwargs: kwargs['border_color'] = kwargs.pop('highlightbackground')
    if 'highlightthickness' in kwargs: kwargs['border_width'] = kwargs.pop('highlightthickness')
    try:
        self.configure(**kwargs)
    except Exception as e:
        pass

ctk.windows.widgets.core_widget_classes.ctk_base_class.CTkBaseClass.config = _ctk_config_patch

# Configuración Base de CustomTkinter
ctk.set_appearance_mode("Dark")   # Paleta oscura premium — igual al Launcher Web
ctk.set_default_color_theme("blue")

# =============================================================================
# PALETA DE COLORES OSCURA PREMIUM (igual al Launcher_Web.html)
# =============================================================================
COLORS = {
    'bg_primary': '#0d1117',      # Fondo principal — oscuro
    'bg_card': '#161b22',          # Fondo de cards — ligeramente más claro
    'bg_input': '#0d1117',         # Fondo de campos de entrada
    'purple': '#bc8cff',           # Morado acento
    'blue': '#58a6ff',             # Azul acento
    'green': '#3fb950',            # Verde éxito
    'orange': '#f0883e',           # Naranja advertencia
    'red': '#f47067',              # Rojo error
    'gray': '#8b949e',             # Gris secundario
    'rose': '#f47067',
    'teal': '#39c5bb',
    'cyan': '#53d0ef',
    'indigo': '#6366f1',
    'pink': '#db61a2',
    'text_primary': '#c9d1d9',     # Texto principal
    'text_secondary': '#8b949e',   # Texto secundario
    'border': '#30363d',           # Bordes
    'shadow': '#000000',           # Sombra base
    'success': '#3fb950',          # Alias de green
    'accent_blue': '#58a6ff',      # Azul para acentos
    'bg_glass': 'rgba(22, 27, 34, 0.7)', # Color vidrio para simulación
}

# Colores para Blobs (Gradientes)
BLOB_COLORS = {
    'blue_glow': (59, 130, 246, 30),     # Azul suave semi-transparente
    'purple_glow': (139, 92, 246, 25),    # Púrpura suave semi-transparente
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
    'title': ('Segoe UI', 42, 'bold'),
    'subtitle': ('Segoe UI', 18),
    'heading': ('Segoe UI', 18, 'bold'),
    'normal': ('Segoe UI', 14),
    'button': ('Segoe UI', 14, 'bold'),
    'button_large': ('Segoe UI', 16, 'bold'),
    'small': ('Segoe UI', 12),
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
    # Intentar obtener el color del padre, si es "transparent" o no existe, usar el del tema
    bg = COLORS['bg_primary']
    try:
        parent_bg = parent.cget("fg_color")
        if parent_bg and parent_bg != "transparent":
            bg = parent_bg
    except:
        try:
            parent_bg = parent.cget("bg")
            if parent_bg and parent_bg != "transparent":
                bg = parent_bg
        except: pass
        
    header_frame = tk.Frame(parent, bg=bg)
    header_frame.pack(fill=tk.X, pady=(0, 20))
    
    if icon_image:
        icon_label = ctk.CTkLabel(header_frame, text="", image=icon_image)
        icon_label.pack(pady=(0, 10))
        icon_label.image = icon_image
        
        title_label = ctk.CTkLabel(header_frame, text=title, font=FONTS['title'],
                                   text_color=COLORS['accent_blue'])
        title_label.pack(pady=(0, 5))
    else:
        title_text = f"{icon} {title}" if icon else title
        title_label = ctk.CTkLabel(header_frame, text=title_text, font=FONTS['title'],
                                   text_color='#ffffff')
        title_label.pack(pady=(0, 5))
    
    if subtitle:
        # Ajustado a FONTS['subtitle'] (18) según requerimiento para igualar detalles
        subtitle_label = ctk.CTkLabel(header_frame, text=subtitle, font=FONTS['subtitle'],
                                      text_color=COLORS['text_secondary'])
        subtitle_label.pack()
    
    # Línea separadora sutil
    sep = tk.Frame(header_frame, height=1, bg=COLORS['border'])
    sep.pack(fill=tk.X, pady=(10, 0))
    
    return header_frame


def create_card(parent, title=None, padding=20):
    """Crea una tarjeta UI con bordes redondeados — tema oscuro premium"""
    card_outer = ctk.CTkFrame(parent, fg_color="transparent")
    
    if title:
        title_lbl = ctk.CTkLabel(card_outer, text=title, font=FONTS['heading'],
                                  text_color=COLORS['accent_blue'])
        title_lbl.pack(anchor='w', padx=8, pady=(0, 5))
    
    card = ctk.CTkFrame(card_outer, fg_color=COLORS['bg_card'],
                        border_color=COLORS['border'], border_width=1, corner_radius=14)
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
    """Barra de estado inferior — tema oscuro"""
    status_var = tk.StringVar(value=initial_text)
    
    status_frame = ctk.CTkFrame(parent, fg_color='#0d1526', corner_radius=0,
                                 border_width=1, border_color=COLORS['border'], height=35)
    status_frame.pack(side=tk.BOTTOM, fill=tk.X)
    
    dot = ctk.CTkLabel(status_frame, text="●", font=FONTS['small'], text_color=COLORS['green'])
    dot.pack(side=tk.LEFT, padx=(12, 4), pady=5)
    
    status_label = ctk.CTkLabel(status_frame, textvariable=status_var,
                                 font=FONTS['small'], text_color=COLORS['text_secondary'])
    status_label.pack(side=tk.LEFT, pady=5)
    
    return status_frame, status_var


def create_file_selector(parent, label_text, string_var, button_command, icon="📁", icon_image=None):
    """Selector de archivo estilizado — tema oscuro"""
    container = ctk.CTkFrame(parent, fg_color="transparent")
    
    label = ctk.CTkLabel(container, text=label_text, font=FONTS['normal'],
                          text_color=COLORS['text_primary'], width=200, anchor="w")
    label.grid(row=0, column=0, sticky='w', padx=5, pady=5)
    
    # Campo de entrada oscuro
    entry = ctk.CTkEntry(container, textvariable=string_var, font=FONTS['normal'],
                         fg_color=COLORS['bg_input'], border_color=COLORS['border'],
                         border_width=1, corner_radius=6, state="readonly",
                         text_color=COLORS['text_primary'])
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
# COMPONENTES PREMIUM ESTILO WEB (OPTIMIZADOS)
# =============================================================================

_BLOB_TEXTURE_CACHE = None

def _get_precomputed_blob_texture(w=160, h=120):
    """Generat la textura en resolución MINÚSCULA para carga instantánea."""
    global _BLOB_TEXTURE_CACHE
    if _BLOB_TEXTURE_CACHE is not None:
        return _BLOB_TEXTURE_CACHE
    
    # Crear imagen base oscura
    bg_img = Image.new('RGBA', (w, h), (10, 14, 26, 255)) 
    
    def draw_blob(img, center, radius, color):
        blob_layer = Image.new('RGBA', (w, h), (0, 0, 0, 0))
        b_draw = ImageDraw.Draw(blob_layer)
        left, top = center[0] - radius, center[1] - radius
        right, bottom = center[0] + radius, center[1] + radius
        b_draw.ellipse([left, top, right, bottom], fill=color)
        blurred = blob_layer.filter(ImageFilter.GaussianBlur(radius * 0.7))
        return Image.alpha_composite(img, blurred)
    
    # Dibujar Blobs una sola vez con alta calidad
    res_img = draw_blob(bg_img, (0, 0), int(min(w, h) * 0.8), BLOB_COLORS['blue_glow'])
    res_img = draw_blob(res_img, (w, h), int(min(w, h) * 0.7), BLOB_COLORS['purple_glow'])
    
    _BLOB_TEXTURE_CACHE = res_img
    return _BLOB_TEXTURE_CACHE

def create_blob_background(parent):
    """
    Crea un fondo de color sólido para máximo rendimiento (Instantáneo).
    Mantiene la elegancia del Dark Mode.
    """
    canvas = tk.Canvas(parent, bg=COLORS['bg_primary'], highlightthickness=0, borderwidth=0)
    canvas.place(x=0, y=0, relwidth=1, relheight=1)
    return canvas

def create_web_card(parent, title, description, icon_image, color='blue', tags=[], command=None, details=[], emoji='🚀', badge_text="Disponible"):
    """
    Versión REPLICA ORIGINAL de la Web Card.
    Usa CTkFrame para esquinas redondeadas y layout proporcional.
    """
    accent_color = COLORS.get(color, COLORS['blue'])
    
    # --- CONTENEDOR EXTERNO ---
    card_outer = ctk.CTkFrame(parent, fg_color="transparent")
    
    # --- CARD PRINCIPAL (Rounded) ---
    card = ctk.CTkFrame(card_outer, fg_color=COLORS['bg_card'], 
                        corner_radius=12, border_width=1, border_color=COLORS['border'])
    card.pack(fill=tk.BOTH, expand=True, padx=0, pady=0)
    
    # Franja de Acento (Izquierda, redondeada con la card)
    accent_strip = ctk.CTkFrame(card, fg_color=accent_color, width=4, corner_radius=0)
    accent_strip.place(relx=0, rely=0, relheight=1, x=0)
    
    # Margen interno - Usamos tk.Frame para mayor rendimiento
    inner = tk.Frame(card, bg=COLORS['bg_card'])
    inner.pack(fill=tk.BOTH, expand=True, padx=24, pady=24)

    # --- HEADER SECTION (Icon + Title) ---
    header = tk.Frame(inner, bg=COLORS['bg_card'])
    header.pack(fill=tk.X, pady=(0, 15))
    
    # Icon Container (Rounded)
    icon_container = ctk.CTkFrame(header, fg_color='#22272e', 
                                  border_width=0,
                                  width=48, height=48, corner_radius=12)
    icon_container.pack_propagate(False)
    icon_container.pack(side=tk.LEFT, padx=(0, 15))
    
    if icon_image:
        icon_lbl = ctk.CTkLabel(icon_container, text="", image=icon_image)
        icon_lbl.pack(expand=True)
    else:
        icon_lbl = tk.Label(icon_container, bg='#22272e')
        icon_lbl.configure(text=emoji, fg='#ffffff', font=('Segoe UI Emoji', 22))
        icon_lbl.pack(expand=True)
    
    # Título y Tags
    title_meta = tk.Frame(header, bg=COLORS['bg_card'])
    title_meta.pack(side=tk.LEFT, fill=tk.Y)
    
    title_lbl = ctk.CTkLabel(title_meta, text=title, font=('Segoe UI', 20, 'bold'), 
                             text_color=COLORS['text_primary'], anchor='w')
    title_lbl.pack(anchor='w', pady=(2, 0))
    
    # Renderizar Tags
    tag_container = tk.Frame(title_meta, bg=COLORS['bg_card'])
    tag_container.pack(anchor='w', pady=(4, 0))
    if tags:
        for t in tags:
            is_new = t.upper() == "NUEVO"
            tag_bg = '#12243a' if is_new else '#22272e'
            tag_fg = COLORS['accent_blue'] if is_new else COLORS['text_secondary']
            
            t_lbl = ctk.CTkLabel(tag_container, text=t.upper(), font=('Segoe UI', 11, 'bold'), 
                                 fg_color=tag_bg, text_color=tag_fg, corner_radius=10,
                                 padx=10, pady=3)
            t_lbl.pack(side=tk.LEFT, padx=(0, 8))
    
    # --- DESCRIPTION (Responsive Wrap) ---
    # Usamos tk.Label para el wraplength dinámico
    desc_lbl = tk.Label(inner, text=description, font=('Segoe UI', 15), 
                         bg=COLORS['bg_card'], fg=COLORS['text_secondary'], 
                         justify='left', anchor='w')
    desc_lbl.pack(fill=tk.X, pady=(0, 20))
    
    # --- DIVIDER ---
    divider = tk.Frame(inner, height=1, bg=COLORS['border'])
    divider.pack(fill=tk.X, pady=(0, 20))
    
    # --- DETAILS LIST ---
    d_text_list = []
    if details:
        details_frame = tk.Frame(inner, bg=COLORS['bg_card'])
        details_frame.pack(fill=tk.X, pady=(0, 2))
        for d in details:
            d_line = tk.Frame(details_frame, bg=COLORS['bg_card'])
            d_line.pack(fill=tk.X, pady=2)
            
            tk.Label(d_line, text="•", font=('Segoe UI', 20, 'bold'), 
                     fg=COLORS['accent_blue'], bg=COLORS['bg_card']).pack(side=tk.LEFT)
            
            d_text = tk.Label(d_line, text=d, font=('Segoe UI', 14), 
                             fg=COLORS['text_secondary'], bg=COLORS['bg_card'],
                             justify='left', anchor='w')
            d_text.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=12)
            d_text_list.append(d_text)
            
    # Auto-wrap binding
    def update_wrap(e, lbl=desc_lbl, d_lbls=d_text_list):
        new_wrap = e.width - 20
        if new_wrap > 0:
            lbl.config(wraplength=new_wrap)
            for dl in d_lbls: 
                try: dl.config(wraplength=new_wrap - 30)
                except: pass
    
    inner.bind('<Configure>', update_wrap)
    
    # --- FOOTER ---
    # Spacer flexible para empujar el footer hacia abajo
    spacer = tk.Frame(inner, bg=COLORS['bg_card'])
    spacer.pack(fill=tk.BOTH, expand=True)
    
    footer = tk.Frame(inner, bg=COLORS['bg_card'])
    footer.pack(fill=tk.X, side=tk.BOTTOM, pady=(15, 0))
    
    badge = ctk.CTkLabel(footer, text=f" {badge_text} ", font=('Segoe UI', 11, 'bold'), 
                         fg_color='#238636', text_color='white', corner_radius=6,
                         padx=12, pady=4)
    badge.pack(side=tk.LEFT)
                 
    arrow_lbl = ctk.CTkLabel(footer, text="→", font=('Segoe UI', 18, 'bold'), 
                             text_color=COLORS['text_secondary'])
    arrow_lbl.pack(side=tk.RIGHT)
    
    # --- INTERACTION ---
    def on_enter(e):
        card.configure(border_color='#444c56', fg_color='#1c2128')
        inner.configure(bg='#1c2128')
        header.configure(bg='#1c2128')
        title_meta.configure(bg='#1c2128')
        tag_container.configure(bg='#1c2128')
        icon_lbl.configure(bg='#22272e')
        desc_lbl.configure(bg='#1c2128')
        spacer.configure(bg='#1c2128')
        footer.configure(bg='#1c2128')
        arrow_lbl.configure(text_color=COLORS['text_primary'])
        if details:
            details_frame.configure(bg='#1c2128')
            for child in details_frame.winfo_children():
                child.configure(bg='#1c2128')
                for sub in child.winfo_children():
                    sub.configure(bg='#1c2128')
        
    def on_leave(e):
        card.configure(border_color=COLORS['border'], fg_color=COLORS['bg_card'])
        inner.configure(bg=COLORS['bg_card'])
        header.configure(bg=COLORS['bg_card'])
        title_meta.configure(bg=COLORS['bg_card'])
        tag_container.configure(bg=COLORS['bg_card'])
        desc_lbl.configure(bg=COLORS['bg_card'])
        spacer.configure(bg=COLORS['bg_card'])
        footer.configure(bg=COLORS['bg_card'])
        arrow_lbl.configure(text_color=COLORS['text_secondary'])
        if details:
            details_frame.configure(bg=COLORS['bg_card'])
            for child in details_frame.winfo_children():
                child.configure(bg=COLORS['bg_card'])
                for sub in child.winfo_children():
                    sub.configure(bg=COLORS['bg_card'])
        
    for w in [card, inner]: # Solo bindear a los frames principales para evitar flicker
        w.bind('<Enter>', on_enter)
        w.bind('<Leave>', on_leave)
        
    if command:
        # Función para bindear recursivamente el comando a todos los hijos
        def bind_click_recursive(widget):
            def on_click(e):
                print(f"DEBUG: Clic detectado en widget {widget}")
                command()
                
            widget.bind('<Button-1>', on_click)
            try: widget.configure(cursor='hand2')
            except: pass
            for child in widget.winfo_children():
                bind_click_recursive(child)
        
        # Empezar desde el contenedor más externo para no dejar zonas muertas
        bind_click_recursive(card_outer)
            
    return card_outer, card
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
    
    # Inyectar chequeo de cerrado seguro desde Launcher
    _setup_graceful_shutdown(root)

def _setup_graceful_shutdown(root):
    """Verifica si el Launcher ordenó un apagado seguro y evita corrupciones deteniéndose."""
    def check_signal():
        try:
            temp_dir = tempfile.gettempdir()
            signal_file = os.path.join(temp_dir, "launcher_kill_signal.txt")
            if os.path.exists(signal_file):
                # Permite que el proceso cierre de manera ordenada
                root.destroy()
                import sys
                sys.exit(0)
        except:
            pass
        # Revisión cada segundo
        try: root.after(1000, check_signal)
        except: pass
    
    try: root.after(1000, check_signal)
    except: pass

def create_main_container(parent, padding=15):
    """
    Crea un contenedor principal con scrollbar automático — tema oscuro.
    """
    container = ctk.CTkScrollableFrame(parent, fg_color="transparent",
                                       corner_radius=0,
                                       scrollbar_button_color=COLORS['border'],
                                       scrollbar_button_hover_color=COLORS['blue'])
    container.pack(fill=tk.BOTH, expand=True, padx=padding, pady=padding)
    return container

def disable_button(button, disabled_color='gray'):
    button.configure(state='disabled', fg_color=COLORS[disabled_color])

def enable_button(button, color='blue'):
    button.configure(state='normal', fg_color=COLORS[color])
