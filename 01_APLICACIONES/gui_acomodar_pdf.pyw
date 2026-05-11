# -*- coding: utf-8 -*-
import tkinter as tk
import customtkinter as ctk
from tkinter import filedialog, messagebox, ttk
import traceback
import os
import sys
import subprocess

# Asegurar que el directorio del script esté en sys.path
# Asegurar que el directorio del script esté en sys.path para imports locales
script_dir = os.path.dirname(os.path.abspath(__file__))
if script_dir not in sys.path:
    sys.path.insert(0, script_dir)

# Agregar 03_OTROS
root_dir = os.path.dirname(script_dir)
others_dir = os.path.join(root_dir, "03_OTROS")
if others_dir not in sys.path:
    sys.path.insert(0, others_dir)

class AcomodadorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("📑 Acomodar PDF")
        self.root.geometry("900x700")
        self.root.configure(bg=mgc.COLORS['bg_primary'])
        mgc.center_window(self.root, 900, 700)

        self.status_frame, self.status_var = mgc.create_status_bar(
            self.root, "✓ Listo — Seleccione PDF e Índice para comenzar"
        )

        main_frame = mgc.create_main_container(self.root)

        # Header Premium
        mgc.create_header(
            main_frame, 
            "Acomodar PDF por Índice", 
            "Genera un único PDF ordenado según las referencias de un archivo Excel",
            icon="📑"
        )

        self.pdf_path = tk.StringVar()
        self.index_path = tk.StringVar()

        # Input Card
        card_outer, card_inner = mgc.create_card(main_frame, "📂 ARCHIVOS DE ENTRADA")
        card_outer.pack(fill=tk.X, pady=(10, 10))

        mgc.create_file_selector(card_inner, "1. Archivo PDF (Recibos Sueltos):", 
                               self.pdf_path, lambda: self.browse_file(self.pdf_path, "*.pdf"), "📄").pack(fill=tk.X, pady=8)
        
        mgc.create_file_selector(card_inner, "2. Índice Excel (Orden deseado):", 
                               self.index_path, lambda: self.browse_file(self.index_path, "*.xlsx"), "📊").pack(fill=tk.X, pady=8)

        # Action Card
        card_act_outer, card_act_inner = mgc.create_card(main_frame, "⚡ ACCIÓN")
        card_act_outer.pack(fill=tk.X, pady=(10, 20))

        btn = mgc.create_large_button(card_act_inner, "ORDENAR Y GENERAR PDF FINAL", 
                                    self.run_process, color='orange', icon="⚙️")
        btn.pack(pady=10)

    def browse_file(self, var, ext):
        f = filedialog.askopenfilename(filetypes=[("Archivo", ext)])
        if f:
            var.set(f)
            self.status_var.set(f"✓ Seleccionado: {os.path.basename(f)}")

    def run_process(self):
        pdf = self.pdf_path.get()
        index = self.index_path.get()
        
        if not pdf or not index:
            messagebox.showerror("Error", "Seleccione ambos archivos para procesar.")
            return

        try:
            self.status_var.set("⏳ Reordenando páginas...")
            self.root.update()

            output_file = filedialog.asksaveasfilename(
                defaultextension=".zip",
                initialfile="Recibos_Ordenados.zip",
                title="Guardar ZIP Final (PDF + CSV)"
            )
            if not output_file:
                return

            cmd = [sys.executable, os.path.join(others_dir, "acomodar_pdf_unificado.py"), "--pdf", pdf, "--index", index, "--output", output_file]
            
            creationflags = 0x08000000 if sys.platform == 'win32' else 0 # CREATE_NO_WINDOW
            result = subprocess.run(cmd, capture_output=True, text=True, creationflags=creationflags)

            if result.returncode == 0:
                self.status_var.set("✓ PDF Generado con éxito.")
                messagebox.showinfo("Proceso Completo", f"PDF generado correctamente:\n{output_file}")
                os.startfile(output_file)
            else:
                self.status_var.set("✗ Error en el proceso.")
                messagebox.showerror("Error", f"Falla en el ordenamiento:\n{result.stdout}\n{result.stderr}")

        except Exception as e:
            traceback.print_exc()
            messagebox.showerror("Error", str(e))

if __name__ == "__main__":
    root = ctk.CTk()
    app = AcomodadorApp(root)
    root.mainloop()
