# -*- coding: utf-8 -*-
"""
LAUNCHER PRINCIPAL
Página principal para ejecutar todos los programas del sistema de sueldos
"""

import tkinter as tk
from tkinter import messagebox
from PIL import Image, ImageTk
import subprocess
import os
import sys

# Asegurar que el directorio del script esté en sys.path para imports locales
script_dir = os.path.dirname(os.path.abspath(__file__))
if script_dir not in sys.path:
    sys.path.insert(0, script_dir)

# Importar componentes modernos
import modern_gui_components as mgc

# ==============================================================================
# CONFIGURACIÓN DE PROGRAMAS
# ==============================================================================

PROGRAMS = [
    {
        'file': '10-CALCULAR_HORAS.pyw',
        'icon': 'calculator.png',
        'name': 'Calculadora de Horas',
        'description': 'Calcular sueldos UOCRA/NASA/UECARA',
        'color': 'blue'
    },
    {
        'file': 'A-GENERAR_RECIBOS_CONTROL.pyw',
        'icon': 'receipts.png',
        'name': 'Generar Recibos',
        'description': 'Crear recibos de sueldo en PDF',
        'color': 'green'
    },
    {
        'file': 'B-PROCESARSUELDOS.pyw',
        'icon': 'payroll.png',
        'name': 'Procesar Sueldos',
        'description': 'Procesar archivo de sueldos',
        'color': 'purple'
    },
    {
        'file': 'BuscarRecibosPDF.pyw',
        'icon': 'search.png',
        'name': 'Buscar Recibos',
        'description': 'Buscar recibos PDF por empleado',
        'color': 'orange'
    },
    {
        'file': '15-PASAR_HORAS_DEPOSITO.pyw',
        'icon': 'warehouse.png',
        'name': 'Horas Depósito',
        'description': 'Pasar horas a depósito',
        'color': 'blue'
    },

    {
        'file': 'imprimir_sobres_gui.pyw',
        'icon': 'receipts.png',
        'name': 'Imprimir Sobres',
        'description': 'Impresión masiva de sobres C5',
        'color': 'blue'
    },
    {
        'file': 'Asistente_Aguinaldo_UNIFICADO.pyw',
        'icon': 'bonus_black.png',
        'name': 'Aguinaldo Efectivo',
        'description': 'Asistente para aguinaldo sueldo y efectivo',
        'color': 'green'
    }
]

# ==============================================================================
# CLASE PRINCIPAL
# ==============================================================================

class LauncherApp:
    """Aplicación launcher principal con cards para ejecutar programas."""
    
    def __init__(self, root):
        self.root = root
        self.root.title("🚀 Launcher - Sistema de Sueldos")
        self.root.geometry("900x750")
        self.root.resizable(True, True)
        self.root.configure(bg=mgc.COLORS['bg_primary'])
        
        mgc.center_window(self.root, 900, 750)
        
        # Obtener directorio base
        self.base_dir = os.path.dirname(os.path.abspath(__file__))
        
        # Cargar iconos
        self.icons = {}
        self.load_icons()
        
        # Contenedor principal
        main_frame = tk.Frame(self.root, bg=mgc.COLORS['bg_primary'])
        main_frame.pack(fill=tk.BOTH, expand=True, padx=15, pady=10)
        
        # Header compacto
        header_frame = tk.Frame(main_frame, bg=mgc.COLORS['bg_primary'])
        header_frame.pack(fill=tk.X, pady=(0, 10))
        
        title_label = tk.Label(header_frame, text="🚀 Sistema de Sueldos", 
                              font=('Segoe UI', 16, 'bold'), bg=mgc.COLORS['bg_primary'], 
                              fg=mgc.COLORS['text_primary'])
        title_label.pack()
        
        subtitle_label = tk.Label(header_frame, text="Seleccione una herramienta para comenzar", 
                                 font=('Segoe UI', 9), bg=mgc.COLORS['bg_primary'], 
                                 fg=mgc.COLORS['text_secondary'])
        subtitle_label.pack()
        
        # Contenedor de cards con grid
        cards_container = tk.Frame(main_frame, bg=mgc.COLORS['bg_primary'])
        cards_container.pack(fill=tk.BOTH, expand=True, pady=(0, 5))
        
        # Configurar grid para 3 columnas
        for i in range(3):
            cards_container.columnconfigure(i, weight=1, uniform='col')
        
        # Crear cards en grid
        self.create_program_cards(cards_container)
        
        # Barra de estado
        self.status_frame, self.status_var = mgc.create_status_bar(self.root, "✓ Listo - Haga clic en una card para ejecutar")
    
    def load_icons(self):
        """Carga los iconos PNG desde el directorio launcher_icons."""
        icons_dir = os.path.join(self.base_dir, "launcher_icons")
        
        if not os.path.exists(icons_dir):
            print(f"⚠ Directorio de iconos no encontrado: {icons_dir}")
            return
        
        for program in PROGRAMS:
            icon_file = program['icon']
            icon_path = os.path.join(icons_dir, icon_file)
            
            if os.path.exists(icon_path):
                try:
                    # Cargar imagen y redimensionar
                    img = Image.open(icon_path)
                    img = img.resize((80, 80), Image.Resampling.LANCZOS)
                    photo = ImageTk.PhotoImage(img)
                    
                    # Guardar referencia (importante para evitar que se borre)
                    self.icons[icon_file] = photo
                except Exception as e:
                    print(f"⚠ Error al cargar icono {icon_file}: {e}")
            else:
                print(f"⚠ Icono no encontrado: {icon_path}")
    
    def create_program_cards(self, parent):
        """Crea las cards de programas en un grid."""
        row = 0
        col = 0
        
        for program in PROGRAMS:
            # Verificar si el archivo existe
            file_path = os.path.join(self.base_dir, program['file'])
            exists = os.path.exists(file_path)
            
            # Crear card
            card = self.create_card(parent, program, exists)
            card.grid(row=row, column=col, padx=5, pady=5, sticky='nsew')
            
            # Avanzar posición en grid
            col += 1
            if col >= 3:
                col = 0
                row += 1
        
        # Configurar filas para que se expandan uniformemente
        for i in range(row + 1):
            parent.rowconfigure(i, weight=1, uniform='row')
    
    def create_card(self, parent, program, exists):
        """Crea una card individual para un programa."""
        # Frame de la card con borde
        card_frame = tk.Frame(parent, bg=mgc.COLORS['bg_card'], 
                             relief=tk.RAISED, bd=2, cursor='hand2' if exists else 'arrow')
        
        # Contenedor interno
        inner = tk.Frame(card_frame, bg=mgc.COLORS['bg_card'])
        inner.pack(fill=tk.BOTH, expand=True, padx=8, pady=5)
        
        # Icono grande
        icon_file = program['icon']
        if icon_file in self.icons:
            # Usar icono PNG
            icon_label = tk.Label(inner, image=self.icons[icon_file], 
                                 bg=mgc.COLORS['bg_card'], bd=0)
        else:
            # Fallback a emoji si no se cargó el icono
            icon_label = tk.Label(inner, text="📱", 
                                 font=('Segoe UI', 40), bg=mgc.COLORS['bg_card'])
        icon_label.pack(pady=(2, 2))
        
        # Nombre del programa
        name_label = tk.Label(inner, text=program['name'], 
                             font=('Segoe UI', 11, 'bold'), 
                             bg=mgc.COLORS['bg_card'], 
                             fg=mgc.COLORS[program['color']])
        name_label.pack(pady=(0, 2))
        
        # Descripción
        desc_label = tk.Label(inner, text=program['description'], 
                             font=('Segoe UI', 8), 
                             bg=mgc.COLORS['bg_card'], 
                             fg=mgc.COLORS['text_secondary'],
                             wraplength=250)
        desc_label.pack(pady=(0, 2))
        
        # Indicador de estado
        if exists:
            status_label = tk.Label(inner, text="✓ Disponible", 
                                   font=('Segoe UI', 7), 
                                   bg=mgc.COLORS['bg_card'], 
                                   fg=mgc.COLORS['green'])
        else:
            status_label = tk.Label(inner, text="✗ No encontrado", 
                                   font=('Segoe UI', 7), 
                                   bg=mgc.COLORS['bg_card'], 
                                   fg=mgc.COLORS['red'])
        status_label.pack()
        
        # Bind eventos si el archivo existe
        if exists:
            # Un solo clic para ejecutar
            card_frame.bind('<Button-1>', lambda e, p=program: self.launch_program(p))
            inner.bind('<Button-1>', lambda e, p=program: self.launch_program(p))
            icon_label.bind('<Button-1>', lambda e, p=program: self.launch_program(p))
            name_label.bind('<Button-1>', lambda e, p=program: self.launch_program(p))
            desc_label.bind('<Button-1>', lambda e, p=program: self.launch_program(p))
            status_label.bind('<Button-1>', lambda e, p=program: self.launch_program(p))
            
            # Efectos hover
            def on_enter(e):
                card_frame.config(relief=tk.SUNKEN, bd=3)
                card_frame.config(bg=mgc.COLORS['bg_primary'])
                inner.config(bg=mgc.COLORS['bg_primary'])
                icon_label.config(bg=mgc.COLORS['bg_primary'])
                name_label.config(bg=mgc.COLORS['bg_primary'])
                desc_label.config(bg=mgc.COLORS['bg_primary'])
                status_label.config(bg=mgc.COLORS['bg_primary'])
            
            def on_leave(e):
                card_frame.config(relief=tk.RAISED, bd=2)
                card_frame.config(bg=mgc.COLORS['bg_card'])
                inner.config(bg=mgc.COLORS['bg_card'])
                icon_label.config(bg=mgc.COLORS['bg_card'])
                name_label.config(bg=mgc.COLORS['bg_card'])
                desc_label.config(bg=mgc.COLORS['bg_card'])
                status_label.config(bg=mgc.COLORS['bg_card'])
            
            card_frame.bind('<Enter>', on_enter)
            card_frame.bind('<Leave>', on_leave)
            inner.bind('<Enter>', on_enter)
            inner.bind('<Leave>', on_leave)
            icon_label.bind('<Enter>', on_enter)
            icon_label.bind('<Leave>', on_leave)
            name_label.bind('<Enter>', on_enter)
            name_label.bind('<Leave>', on_leave)
            desc_label.bind('<Enter>', on_enter)
            desc_label.bind('<Leave>', on_leave)
            status_label.bind('<Enter>', on_enter)
            status_label.bind('<Leave>', on_leave)
        
        return card_frame
    
    def launch_program(self, program):
        """Ejecuta un programa .pyw."""
        file_path = os.path.join(self.base_dir, program['file'])
        
        try:
            # Actualizar estado
            self.status_var.set(f"🚀 Ejecutando {program['name']}...")
            self.root.update()
            
            # Ejecutar el programa
            if sys.platform == 'win32':
                # En Windows, usar pythonw.exe para archivos .pyw
                subprocess.Popen(['pythonw', file_path], 
                               cwd=self.base_dir,
                               creationflags=subprocess.CREATE_NO_WINDOW)
            else:
                subprocess.Popen(['python', file_path], cwd=self.base_dir)
            
            # Actualizar estado
            self.status_var.set(f"✓ {program['name']} ejecutado correctamente")
            
        except Exception as e:
            messagebox.showerror("Error", 
                               f"No se pudo ejecutar {program['name']}:\n{str(e)}")
            self.status_var.set(f"✗ Error al ejecutar {program['name']}")

# ==============================================================================
# MAIN ENTRY POINT
# ==============================================================================

def main():
    try:
        root = tk.Tk()
        app = LauncherApp(root)
        root.mainloop()
    except Exception as e:
        # Intentar mostrar error en ventana emergente (para .pyw)
        try:
            import ctypes
            ctypes.windll.user32.MessageBoxW(0, f"Error fatal en Launcher:\n{str(e)}", "Error de Inicio", 0x10)
        except:
            pass
        # También imprimir a stderr por si se ejecuta desde consola
        import sys
        sys.stderr.write(f"Error fatal: {e}\n")

if __name__ == "__main__":
    main()
