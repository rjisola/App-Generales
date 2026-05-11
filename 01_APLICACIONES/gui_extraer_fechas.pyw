# -*- coding: utf-8 -*-
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkcalendar import DateEntry
import traceback
import os
import sys
import datetime

# Asegurar que el directorio del script esté en sys.path
script_dir = os.path.dirname(os.path.abspath(__file__))
if script_dir not in sys.path:
    sys.path.insert(0, script_dir)

# También agregar 03_OTROS
others_dir = os.path.abspath(os.path.join(script_dir, "..", "03_OTROS"))
if others_dir not in sys.path:
    sys.path.insert(0, others_dir)

try:
    import modern_gui_components as mgc
except ImportError:
    import tkinter.messagebox as messagebox
    messagebox.showerror("Error", "Falta modern_gui_components.py")
    sys.exit(1)

class FechasApp:
    def __init__(self, root):
        self.root = root
        self.root.title("📅 Fechas de Ingreso")
        self.root.geometry("900x700")
        self.root.configure(bg=mgc.COLORS['bg_primary'])
        mgc.center_window(self.root, 900, 700)

        # Barra de estado inferior
        self.status_frame, self.status_var = mgc.create_status_bar(
            self.root, "✓ Listo — Seleccione archivo Excel para comenzar"
        )

        main_frame = mgc.create_main_container(self.root)

        # Header Premium
        mgc.create_header(
            main_frame, 
            "Fechas de Ingreso y Vacaciones", 
            "Cálculo de antigüedad y períodos vacacionales correspondientes",
            icon="📅"
        )

        # Variables
        self.input_path = tk.StringVar()

        # Input Card
        card_outer, card_inner = mgc.create_card(main_frame, "📂 ARCHIVO DE ENTRADA")
        card_outer.pack(fill=tk.X, pady=(10, 10))

        mgc.create_file_selector(
            card_inner, "Seleccione Planilla Excel:",
            self.input_path, lambda: self.browse(self.input_path), "📂"
        ).pack(fill=tk.X, pady=4)

        tk.Label(card_inner, text="Fecha de Corte para el Cálculo:",
                 font=mgc.FONTS['normal'], bg=mgc.COLORS['bg_card'],
                 fg=mgc.COLORS['text_primary']).pack(anchor='w', pady=(15, 0))

        self.date_entry = DateEntry(
            card_inner, width=20,
            background=mgc.COLORS['blue'],
            foreground='white', borderwidth=2,
            date_pattern='yyyy-mm-dd',
            font=mgc.FONTS['normal']
        )
        self.date_entry.pack(anchor='w', pady=(5, 0))
        self.date_entry.set_date(datetime.date.today())

        # Action Card
        card_act_outer, card_act_inner = mgc.create_card(main_frame, "⚡ ACCIÓN")
        card_act_outer.pack(fill=tk.X, pady=(10, 20))

        btn = mgc.create_large_button(
            card_act_inner, "GENERAR INFORME DE ANTIGÜEDAD",
            self.run_process, color='blue', icon="⚙️"
        )
        btn.pack(pady=10)

    def browse(self, var):
        f = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx;*.xlsm")])
        if f:
            var.set(f)
            self.status_var.set(f"✓ Seleccionado: {os.path.basename(f)}")

    def run_process(self):
        in_f = self.input_path.get()
        calc_d = self.date_entry.get_date().strftime('%Y-%m-%d')

        if not in_f:
            messagebox.showerror("Error", "Debe seleccionar un archivo.")
            return

        try:
            self.status_var.set("⏳ Procesando...")
            self.root.update()

            output_file = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                initialfile="Antiguedad.xlsx"
            )
            if not output_file:
                return

            import subprocess
            script_a_ejecutar = os.path.join(others_dir, "extraer_fechas.py")
            cmd = [sys.executable, script_a_ejecutar, "--input", in_f, "--output", output_file, "--calc-date", calc_d]
            
            creationflags = 0
            if sys.platform == 'win32':
                creationflags = 0x08000000 # CREATE_NO_WINDOW
                
            subprocess.run(cmd, creationflags=creationflags)
            
            self.status_var.set(f"✓ Terminando.")
            messagebox.showinfo("Éxito", f"Archivo generado:\n{output_file}")
            os.startfile(output_file)

        except Exception as e:
            traceback.print_exc()
            messagebox.showerror("Error", str(e))

if __name__ == "__main__":
    try:
        root = tk.Tk()
        app = FechasApp(root)
        root.mainloop()
    except Exception as e:
        # Crear un log de error específico para esta aplicación
        log_dir = os.path.dirname(os.path.abspath(__file__))
        log_file = os.path.join(log_dir, "gui_extraer_fechas_CRASH.log")
        with open(log_file, "w", encoding="utf-8") as f:
            f.write(f"La aplicación falló al iniciar el {datetime.datetime.now()}.\n")
            f.write("="*80 + "\n")
            f.write(traceback.format_exc())
        # También intentar mostrar un messagebox, aunque podría fallar si Tkinter está roto
        messagebox.showerror("Error Crítico - FechasApp", f"La aplicación no pudo iniciarse.\n\nError: {e}\n\nSe ha creado un archivo 'gui_extraer_fechas_CRASH.log' con los detalles.")
