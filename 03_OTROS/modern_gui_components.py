# -*- coding: utf-8 -*-
"""
Módulo de Componentes GUI Modernos usando CustomTkinter
Proporciona estilos, colores y componentes reutilizables "Ultra Premium"
"""

import os
import tempfile
import json

# --- CARGA DE TEMA DINÁMICA (Solo lógica de datos) ---
_APPEARANCE_MODE = "Dark"
try:
    _config_path = os.path.join(os.path.dirname(__file__), 'config.json')
    if os.path.exists(_config_path):
        with open(_config_path, 'r', encoding='utf-8') as _f:
            _config = json.load(_f)
            _APPEARANCE_MODE = _config.get("appearance_mode", "Dark")
except Exception:
    pass

# =============================================================================
# PALETA DE COLORES PREMIUM (Estática)
# =============================================================================
# Definir ambos esquemas para evitar recálculos lentos
LIGHT_COLORS = {
    'bg_primary': '#f3f4f6', 'bg_card': '#ffffff', 'bg_input': '#f9fafb',
    'purple': '#7c3aed', 'blue': '#2563eb', 'green': '#059669',
    'orange': '#d97706', 'red': '#dc2626', 'gray': '#4b5563',
    'rose': '#e11d48', 'teal': '#0d9488', 'cyan': '#0891b2',
    'indigo': '#4f46e5', 'pink': '#db2777',
    'text_primary': '#000000', 'text_secondary': '#374151',
    'border': '#d1d5db', 'shadow': '#000000', 'success': '#059669',
    'accent_blue': '#1d4ed8', 'bg_glass': 'rgba(255, 255, 255, 0.8)',
    'icon_bg': '#e5e7eb', 'tag_bg': '#f3f4f6', 'tag_bg_new': '#dbeafe',
    'status_bar_bg': '#ffffff',
}

DARK_COLORS = {
    'bg_primary': '#0d1117', 'bg_card': '#161b22', 'bg_input': '#0d1117',
    'purple': '#bc8cff', 'blue': '#58a6ff', 'green': '#3fb950',
    'orange': '#f0883e', 'red': '#f47067', 'gray': '#8b949e',
    'rose': '#f47067', 'teal': '#39c5bb', 'cyan': '#53d0ef',
    'indigo': '#6366f1', 'pink': '#db61a2',
    'text_primary': '#ffffff', 'text_secondary': '#d1d5db',
    'border': '#30363d', 'shadow': '#000000', 'success': '#3fb950',
    'accent_blue': '#79c0ff', 'bg_glass': 'rgba(22, 27, 34, 0.7)',
    'icon_bg': '#22272e', 'tag_bg': '#22272e', 'tag_bg_new': '#12243a',
    'status_bar_bg': '#0d1526',
}

COLORS = DARK_COLORS if _APPEARANCE_MODE == "Dark" else LIGHT_COLORS

# Colores para Blobs y Hovers
BLOB_COLORS = {'blue_glow': (59, 130, 246, 30), 'purple_glow': (139, 92, 246, 25)}
HOVER_COLORS = {'purple': '#7c3aed', 'blue': '#2563eb', 'green': '#059669', 'orange': '#d97706', 'red': '#dc2626', 'gray': '#4b5563'}

# Configuración de Fuentes
FONTS = {
    'title': ('Segoe UI', 42, 'bold'), 'subtitle': ('Segoe UI', 18),
    'heading': ('Segoe UI', 18, 'bold'), 'normal': ('Segoe UI', 14),
    'button': ('Segoe UI', 14, 'bold'), 'button_large': ('Segoe UI', 16, 'bold'),
    'small': ('Segoe UI', 12), 'tiny': ('Segoe UI', 10)
}

# --- INICIALIZACIÓN DIFERIDA DE CTK ---
_GUI_INITIALIZED = False
tk = None
ctk = None

def init_gui_engine():
    """Inicializa CustomTkinter y parches de GUI bajo demanda"""
    global _GUI_INITIALIZED, tk, ctk
    if _GUI_INITIALIZED: return
    
    import tkinter as _tk
    import customtkinter as _ctk
    tk = _tk
    ctk = _ctk
    
    # Optimización de Parche
    _original_config = ctk.windows.widgets.core_widget_classes.ctk_base_class.CTkBaseClass.configure
    def _ctk_config_patch(self, **kwargs):
        if not kwargs: return _original_config(self)
        
        # Traducir nombres clásicos a CTk
        if 'bg' in kwargs: kwargs['fg_color'] = kwargs.pop('bg')
        if 'fg' in kwargs: kwargs['text_color'] = kwargs.pop('fg')
        if 'highlightbackground' in kwargs: kwargs['border_color'] = kwargs.pop('highlightbackground')
        if 'highlightthickness' in kwargs: kwargs['border_width'] = kwargs.pop('highlightthickness')
        
        # Intentar aplicar todos los cambios de una vez para mantener atomicidad y evitar recursión
        try:
            return _original_config(self, **kwargs)
        except (ValueError, TypeError):
            # Si falla (ej. border_width en un Label), intentar aplicar solo lo básico o ignorar fallos individuales
            # Para máxima compatibilidad en apps antiguas, si falla el bloque entero, lo intentamos individualmente 
            # de forma segura solo si es necesario, pero aquí preferimos el fail-safe atómico.
            for k, v in list(kwargs.items()):
                try: _original_config(self, **{k: v})
                except: pass
            return

    ctk.windows.widgets.core_widget_classes.ctk_base_class.CTkBaseClass.config = _ctk_config_patch
    
    ctk.set_appearance_mode(_APPEARANCE_MODE)
    ctk.set_default_color_theme("blue")
    _GUI_INITIALIZED = True

# =============================================================================
# CLASE BASE PARA VENTANAS MODERNAS
# =============================================================================
class ModernWindow:
    """Clase base para ventanas con estilo moderno"""
    def __init__(self, root, title, width=800, height=600, resizable=True):
        init_gui_engine()
        self.root = root
        self.root.title(title)
        self.root.geometry(f"{width}x{height}")
        
        try:
            self.root.configure(bg=COLORS['bg_primary'])
        except: pass

        if not resizable:
            self.root.resizable(False, False)

# =============================================================================
# FUNCIONES HELPER PARA CREAR COMPONENTES
# =============================================================================

def create_header(parent, title, subtitle=None, icon="", icon_image=None):
    init_gui_engine()
    import tkinter as tk
    import customtkinter as ctk
    
    bg = COLORS['bg_primary']
    header_frame = tk.Frame(parent, bg=bg)
    header_frame.pack(fill=tk.X, pady=(0, 20))
    
    if icon_image:
        icon_label = ctk.CTkLabel(header_frame, text="", image=icon_image)
        icon_label.pack(pady=(0, 10))
        icon_label.image = icon_image
        title_label = ctk.CTkLabel(header_frame, text=title, font=FONTS['title'], text_color=COLORS['accent_blue'])
        title_label.pack(pady=(0, 5))
    else:
        title_text = f"{icon} {title}" if icon else title
        title_label = ctk.CTkLabel(header_frame, text=title_text, font=FONTS['title'], text_color=COLORS['text_primary'])
        title_label.pack(pady=(0, 5))
    
    if subtitle:
        subtitle_label = ctk.CTkLabel(header_frame, text=subtitle, font=FONTS['subtitle'], text_color=COLORS['text_secondary'])
        subtitle_label.pack()
    
    sep = tk.Frame(header_frame, height=1, bg=COLORS['border'])
    sep.pack(fill=tk.X, pady=(10, 0))
    return header_frame

def create_card(parent, title=None, padding=20):
    init_gui_engine()
    import tkinter as tk
    import customtkinter as ctk
    
    card_outer = ctk.CTkFrame(parent, fg_color="transparent")
    if title:
        title_lbl = ctk.CTkLabel(card_outer, text=title, font=FONTS['heading'], text_color=COLORS['accent_blue'])
        title_lbl.pack(anchor='w', padx=8, pady=(0, 5))
    
    card = ctk.CTkFrame(card_outer, fg_color=COLORS['bg_card'], border_color=COLORS['border'], border_width=1, corner_radius=14)
    card.pack(fill=tk.BOTH, expand=True)
    inner = ctk.CTkFrame(card, fg_color="transparent")
    inner.pack(fill=tk.BOTH, expand=True, padx=padding, pady=padding)
    return card_outer, inner

def create_button(parent, text, command, color='blue', icon="", icon_image=None, text_color='white', **kwargs):
    init_gui_engine()
    import customtkinter as ctk
    bg_color = COLORS.get(color, COLORS['blue'])
    h_color = HOVER_COLORS.get(color, bg_color)
    f_color = COLORS.get(text_color, text_color)
    btn_text = f"{icon} {text}" if icon and not icon_image else text
    button = ctk.CTkButton(parent, text=btn_text, image=icon_image, command=command, font=FONTS['button'],
                           fg_color=bg_color, hover_color=h_color, text_color=f_color, corner_radius=8, cursor="hand2")
    if icon_image: button.image = icon_image
    return button

def create_large_button(parent, text, command, color='green', icon="", icon_image=None, text_color='white', **kwargs):
    init_gui_engine()
    import customtkinter as ctk
    bg_color = COLORS.get(color, COLORS['green'])
    h_color = HOVER_COLORS.get(color, bg_color)
    f_color = COLORS.get(text_color, text_color)
    btn_text = f"{icon} {text}" if icon and not icon_image else text
    button = ctk.CTkButton(parent, text=btn_text, image=icon_image, command=command, font=FONTS['button_large'],
                           fg_color=bg_color, hover_color=h_color, text_color=f_color, corner_radius=12, height=45, cursor="hand2")
    if icon_image: button.image = icon_image
    return button

def create_status_bar(parent, initial_text="Listo"):
    init_gui_engine()
    import tkinter as tk
    import customtkinter as ctk
    status_var = tk.StringVar(value=initial_text)
    status_frame = ctk.CTkFrame(parent, fg_color=COLORS['status_bar_bg'], corner_radius=0, border_width=1, border_color=COLORS['border'], height=35)
    status_frame.pack(side=tk.BOTTOM, fill=tk.X)
    dot = ctk.CTkLabel(status_frame, text="●", font=FONTS['small'], text_color=COLORS['green'])
    dot.pack(side=tk.LEFT, padx=(12, 4), pady=5)
    status_label = ctk.CTkLabel(status_frame, textvariable=status_var, font=FONTS['small'], text_color=COLORS['text_secondary'])
    status_label.pack(side=tk.LEFT, pady=5)
    return status_frame, status_var

def create_file_selector(parent, label_text, string_var, button_command, icon="📁", icon_image=None):
    init_gui_engine()
    import customtkinter as ctk
    container = ctk.CTkFrame(parent, fg_color="transparent")
    # Reducimos el ancho del label de 200 a un valor más razonable (110) o 0 si está vacío para ganar espacio para el entry
    lbl_width = 110 if label_text else 0
    label = ctk.CTkLabel(container, text=label_text, font=FONTS['normal'], text_color=COLORS['text_primary'], width=lbl_width, anchor="w")
    label.grid(row=0, column=0, sticky='w', padx=5, pady=5)
    entry = ctk.CTkEntry(container, textvariable=string_var, font=FONTS['normal'], fg_color=COLORS['bg_input'], border_color=COLORS['border'],
                         border_width=1, corner_radius=6, state="readonly", text_color=COLORS['text_primary'])
    entry.grid(row=0, column=1, sticky='we', padx=8, pady=5)
    button = create_button(container, "Seleccionar", button_command, color='purple', icon=icon, icon_image=icon_image)
    button.grid(row=0, column=2, padx=5, pady=5)
    container.columnconfigure(0, weight=0); container.columnconfigure(1, weight=1); container.columnconfigure(2, weight=0)
    return container

def create_progress_section(parent):
    init_gui_engine()
    import tkinter as tk
    import customtkinter as ctk
    container = ctk.CTkFrame(parent, fg_color="transparent")
    progress_var = tk.StringVar(value="Esperando...")
    progress_label = ctk.CTkLabel(container, textvariable=progress_var, font=FONTS['normal'], text_color=COLORS['text_secondary'])
    progress_label.pack(pady=(0, 8))
    progress_bar = ctk.CTkProgressBar(container, width=300, height=12, corner_radius=6, fg_color=COLORS['border'], progress_color=COLORS['blue'])
    progress_bar.set(0); progress_bar.pack(fill=tk.X, pady=(0, 5))
    class ProgressBarWrapper:
        def __init__(self, ctk_pb): self.pb = ctk_pb
        def __setitem__(self, key, value):
            if key == 'value': self.pb.set(value / 100.0)
    return container, ProgressBarWrapper(progress_bar), progress_var

def create_icon_label(parent, icon, text, color='text_primary'):
    init_gui_engine()
    import customtkinter as ctk
    container = ctk.CTkFrame(parent, fg_color="transparent")
    icon_label = ctk.CTkLabel(container, text=icon, font=('Segoe UI', 46), text_color=COLORS['text_primary'])
    icon_label.pack(pady=(0, 8))
    text_label = ctk.CTkLabel(container, text=text, font=FONTS['subtitle'], text_color=COLORS.get(color, COLORS['text_primary']))
    text_label.pack()
    return container

# =============================================================================
# COMPONENTES PREMIUM ESTILO WEB (OPTIMIZADOS)
# =============================================================================

def create_blob_background(parent):
    import tkinter as tk
    canvas = tk.Canvas(parent, bg=COLORS['bg_primary'], highlightthickness=0, borderwidth=0)
    canvas.place(x=0, y=0, relwidth=1, relheight=1)
    return canvas

def create_web_card(parent, title, description, icon_image, color='blue', tags=[], command=None, details=[], emoji='🚀', badge_text="Disponible", on_hover_enter=None, on_hover_leave=None):
    """
    Crea una card estilo web ultra-optimizada usando componentes nativos de CTk.
    """
    init_gui_engine()
    import customtkinter as ctk
    
    accent_color = COLORS.get(color, COLORS['blue'])
    hover_bg = '#e5e7eb' if _APPEARANCE_MODE == 'Light' else '#1c2128'
    
    # Contenedor Principal (Usamos CTkFrame para aceleración)
    card_outer = ctk.CTkFrame(parent, fg_color="transparent")
    
    # La Card real
    card = ctk.CTkFrame(card_outer, fg_color=COLORS['bg_card'], corner_radius=12, border_width=1, border_color=COLORS['border'])
    card.pack(fill="both", expand=True)
    
    # Línea de acento lateral (Optimizado: un simple frame pequeño)
    accent_strip = ctk.CTkFrame(card, fg_color=accent_color, width=4, corner_radius=0)
    accent_strip.place(relx=0, rely=0, relheight=1, x=0)
    
    # Contenedor Interno (Reducimos padding de 24 a 16 para compactar)
    inner = ctk.CTkFrame(card, fg_color="transparent")
    inner.pack(fill="both", expand=True, padx=20, pady=16)
    
    # Header: Icono + Título (Reducimos pady de 15 a 10)
    header = ctk.CTkFrame(inner, fg_color="transparent")
    header.pack(fill="x", pady=(0, 10))
    
    icon_container = ctk.CTkFrame(header, fg_color=COLORS['icon_bg'], width=42, height=42, corner_radius=10)
    icon_container.pack_propagate(False)
    icon_container.pack(side="left", padx=(0, 12))
    
    if icon_image:
        ctk.CTkLabel(icon_container, text="", image=icon_image).pack(expand=True)
    else:
        ctk.CTkLabel(icon_container, text=emoji, font=('Segoe UI Emoji', 20)).pack(expand=True)
    
    title_meta = ctk.CTkFrame(header, fg_color="transparent")
    title_meta.pack(side="left", fill="y")
    
    title_lbl = ctk.CTkLabel(title_meta, text=title, font=('Segoe UI', 15, 'bold'), text_color=COLORS['text_primary'], anchor='w')
    title_lbl.pack(anchor='w', pady=(0, 0))
    
    tag_container = ctk.CTkFrame(title_meta, fg_color="transparent")
    tag_container.pack(anchor='w', pady=(2, 0))
    
    for t in tags:
        is_new = t.upper() == "NUEVO"
        t_bg = COLORS['tag_bg_new'] if is_new else COLORS['tag_bg']
        t_fg = COLORS['accent_blue'] if is_new else COLORS['text_secondary']
        ctk.CTkLabel(tag_container, text=t.upper(), font=('Segoe UI', 8, 'bold'), 
                     fg_color=t_bg, text_color=t_fg, corner_radius=4, padx=6).pack(side="left", padx=(0, 4))
    
    # Descripción (Reducimos pady de 20 a 10)
    desc_lbl = ctk.CTkLabel(inner, text=description, font=('Segoe UI', 13), text_color=COLORS['text_secondary'], 
                           justify='left', anchor='w', wraplength=230)
    desc_lbl.pack(fill="x", pady=(0, 10))
    
    # Separador (Reducimos pady de 20 a 10)
    ctk.CTkFrame(inner, height=1, fg_color=COLORS['border']).pack(fill="x", pady=(0, 10))
    
    # Detalles
    if details:
        details_frame = ctk.CTkFrame(inner, fg_color="transparent")
        details_frame.pack(fill="x")
        for d in details:
            d_line = ctk.CTkFrame(details_frame, fg_color="transparent")
            d_line.pack(fill="x", pady=1)
            ctk.CTkLabel(d_line, text="•", font=('Segoe UI', 16, 'bold'), text_color=COLORS['accent_blue']).pack(side="left")
            ctk.CTkLabel(d_line, text=d, font=('Segoe UI', 12), text_color=COLORS['text_secondary'], 
                         justify='left', anchor='w', wraplength=210).pack(side="left", fill="x", expand=True, padx=8)
            
    # Footer (Reducimos pady de 15 a 0 ya que el contenedor tiene pady inferior)
    footer = ctk.CTkFrame(inner, fg_color="transparent")
    footer.pack(fill="x", side="bottom", pady=(5, 0))
    
    ctk.CTkLabel(footer, text=f" {badge_text} ", font=('Segoe UI', 9, 'bold'), fg_color='#238636', text_color='white', corner_radius=4).pack(side="left")
    arrow_lbl = ctk.CTkLabel(footer, text="→", font=('Segoe UI', 16, 'bold'), text_color=COLORS['text_secondary'])
    arrow_lbl.pack(side="right")
    
    # Lógica de Hover Ultra-Rápida
    def on_enter(e):
        card.configure(border_color=COLORS['blue'], fg_color=hover_bg)
        arrow_lbl.configure(text_color=COLORS['text_primary'])
        if on_hover_enter: on_hover_enter(e)
        
    def on_leave(e):
        card.configure(border_color=COLORS['border'], fg_color=COLORS['bg_card'])
        arrow_lbl.configure(text_color=COLORS['text_secondary'])
        if on_hover_leave: on_hover_leave(e)
    
    # Vincular eventos a los componentes clave
    for w in [card, inner, desc_lbl, footer, title_lbl, arrow_lbl]:
        w.bind('<Enter>', on_enter)
        w.bind('<Leave>', on_leave)
        if command:
            w.bind('<Button-1>', lambda e: command())
            w.configure(cursor='hand2')
            
    return card_outer, card

# =============================================================================
# UTILIDADES GLOBALES
# =============================================================================

def center_window(root, width, height):
    init_gui_engine()
    screen_width  = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    taskbar_margin = 70; titlebar_margin = 30
    eff_width = min(width, screen_width - 20)
    eff_height = min(height, screen_height - taskbar_margin - titlebar_margin)
    es_app_hija = os.environ.get('CHILD_APP_MODE') == '1'
    if es_app_hija:
        available_width = screen_width - 460
        eff_width = 900; eff_height = 700
        x = max(10, (available_width - eff_width) // 2)
        y = max(titlebar_margin, (screen_height - eff_height) // 2)
    else:
        x = (screen_width - eff_width) // 2
        y = max(titlebar_margin, (screen_height - eff_height) // 2)
    root.geometry(f"{eff_width}x{eff_height}+{int(x)}+{int(y)}")
    _setup_graceful_shutdown(root)

def _setup_graceful_shutdown(root):
    def check_signal():
        try:
            signal_file = os.path.join(tempfile.gettempdir(), "launcher_kill_signal.txt")
            if os.path.exists(signal_file):
                root.destroy(); import sys; sys.exit(0)
        except: pass
        try: root.after(1000, check_signal)
        except: pass
    try: root.after(1000, check_signal)
    except: pass

def create_main_container(parent, padding=15):
    init_gui_engine()
    import customtkinter as ctk
    import tkinter as tk
    container = ctk.CTkScrollableFrame(parent, fg_color="transparent", corner_radius=0, scrollbar_button_color=COLORS['border'], scrollbar_button_hover_color=COLORS['blue'])
    container.pack(fill=tk.BOTH, expand=True, padx=padding, pady=padding)
    return container

def disable_button(button, disabled_color='gray'):
    button.configure(state='disabled', fg_color=COLORS[disabled_color])

def enable_button(button, color='blue'):
    button.configure(state='normal', fg_color=COLORS[color])

def get_resource_path(relative_path):
    init_gui_engine()
    try:
        import sys
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)
