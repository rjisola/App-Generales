# -*- coding: utf-8 -*-
import tkinter as tk
import customtkinter as ctk
from tkinter import filedialog, messagebox, ttk
import traceback
import os
import sys

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

class PlanillaApp:
    def __init__(self, root):
        self.root = root
        self.root.title("📊 Planilla por Índice")
        self.root.geometry("900x700")
        self.root.resizable(True, True)
        self.root.configure(bg=mgc.COLORS['bg_primary'])

        mgc.center_window(self.root, 900, 700)

        if _has_icon_loader:
            set_window_icon(self.root, 'bonus_assist')

        # Barra de estado inferior
        self.status_frame, self.status_var = mgc.create_status_bar(
            self.root, "✓ Listo — Seleccione los 3 archivos requeridos para comenzar"
        )

        # Contenedor principal con scrollbar
        main_frame = mgc.create_main_container(self.root)

        # Header Premium
        mgc.create_header(
            main_frame, 
            "Planilla por Índice", 
            "Acomodar planilla Excel usando referencias de índice y horas adicionales",
            icon="📊"
        )

        # Variables
        self.main_path = tk.StringVar()
        self.ref_path = tk.StringVar()
        self.horas_path = tk.StringVar()
        self.bono_var = tk.BooleanVar()

        # Card 1: Archivos de Entrada
        card_outer, card_inner = mgc.create_card(main_frame, "📂 ARCHIVOS DE ENTRADA")
        card_outer.pack(fill=tk.X, pady=(10, 10))

        mgc.create_file_selector(
            card_inner, "1. Archivo Principal (HORAS CONTADOR):",
            self.main_path, lambda: self.browse(self.main_path), "📂"
        ).pack(fill=tk.X, pady=6)

        mgc.create_file_selector(
            card_inner, "2. Archivo de Referencia (VALOR_HORAS_SUELDOS):",
            self.ref_path, lambda: self.browse(self.ref_path), "📂"
        ).pack(fill=tk.X, pady=6)

        mgc.create_file_selector(
            card_inner, "3. Archivo de Años Antigüedad (ANTIGÜEDAD):",
            self.horas_path, lambda: self.browse(self.horas_path), "📂"
        ).pack(fill=tk.X, pady=6)

        # Card 2: Configuración Adicional
        card_cfg_outer, card_cfg_inner = mgc.create_card(main_frame, "⚙️ OPCIONES")
        card_cfg_outer.pack(fill=tk.X, pady=(5, 15))

        tk.Checkbutton(
            card_cfg_inner, text="Aplicar Bono Paritarias (Asignación Extraordinaria)",
            variable=self.bono_var,
            bg=mgc.COLORS['bg_card'],
            fg=mgc.COLORS['text_primary'],
            activebackground=mgc.COLORS['bg_card'],
            selectcolor=mgc.COLORS['bg_primary'],
            font=mgc.FONTS['normal']
        ).pack(anchor='w', pady=5)

        # Card 3: Ejecución
        card_act_outer, card_act_inner = mgc.create_card(main_frame, "▶️ EJECUCIÓN")
        card_act_outer.pack(fill=tk.X, pady=(5, 20))

        btn = mgc.create_large_button(
            card_act_inner, "PROCESAR Y GENERAR PLANILLA",
            self.run_process, color='purple', icon="▶️"
        )
        btn.pack(pady=10)

    def _add_hover(self, btn, normal_color_key, hover_hex):
        normal_bg = mgc.COLORS.get(normal_color_key, normal_color_key)
        btn.bind("<Enter>", lambda e: btn.configure(bg=hover_hex))
        btn.bind("<Leave>", lambda e: btn.configure(bg=normal_bg))

    def browse(self, var):
        f = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx;*.xlsm")])
        if f:
            var.set(f)
            self.status_var.set(f"✓ Seleccionado: {os.path.basename(f)}")

    def run_process(self):
        main_f = self.main_path.get()
        ref_f = self.ref_path.get()
        horas_f = self.horas_path.get()

        if not all([main_f, ref_f, horas_f]):
            messagebox.showerror("Error", "Debe seleccionar los 3 archivos requeridos.")
            return

        try:
            self.status_var.set("⏳ Procesando...")
            self.root.update()

            output_file = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                initialfile="Planilla_Procesada.xlsx",
                title="Guardar Resultado Como"
            )
            if not output_file:
                self.status_var.set("✗ Cancelado.")
                return

            planilla.procesar_planilla(main_f, ref_f, horas_f, output_file, self.bono_var.get())

            self.status_var.set("✓ Proceso terminado correctamente.")
            messagebox.showinfo("Éxito", f"Archivo generado correctamente:\n{output_file}")
            os.startfile(output_file)

        except Exception as e:
            traceback.print_exc()
            self.status_var.set("✗ Error durante el proceso.")
            messagebox.showerror("Error", str(e))


if __name__ == "__main__":
    try:
        root = ctk.CTk()
        app = PlanillaApp(root)
        root.mainloop()
    except Exception as e:
        try:
            import ctypes
            ctypes.windll.user32.MessageBoxW(0, f"Error fatal en Planilla por Índice:\n{str(e)}", "Error de Inicio", 0x10)
        except:
            pass
        sys.stderr.write(f"Error fatal: {e}\n")
