import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox
import customtkinter as ctk
import fitz  # PyMuPDF
from PIL import Image, ImageChops, ImageStat
import threading
import numpy as np

# Asegurar que el directorio del script esté en sys.path para imports locales
script_dir = os.path.dirname(os.path.abspath(__file__))
if script_dir not in sys.path:
    sys.path.insert(0, script_dir)

# También agregar 03_OTROS (donde se encuentran los submódulos del sistema)
others_dir = os.path.abspath(os.path.join(script_dir, "..", "03_OTROS"))
if others_dir not in sys.path:
    sys.path.insert(0, others_dir)
# Configurar rutas para encontrar módulos en 03_OTROS
parent_dir = os.path.dirname(script_dir)
otros_dir = os.path.join(parent_dir, '03_OTROS')
if otros_dir not in sys.path:
    sys.path.insert(0, otros_dir)

import modern_gui_components as mgc
from icon_loader import set_window_icon, load_icon

class ExtractorFirmasApp(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("🖋️ EXTRACTOR DE FIRMAS PDF ULTRA")
        self.geometry("900x700")
        mgc.center_window(self, 900, 700)
        self.resizable(True, True)
        
        # Tema y Colores
        ctk.set_appearance_mode("dark")
        self.colors = {
            'bg': mgc.COLORS['bg_primary'],
            'accent': mgc.COLORS['blue'],
            'success': mgc.COLORS['green'],
            'danger': mgc.COLORS['red']
        }
        self.configure(fg_color=self.colors['bg'])
        
        set_window_icon(self, 'printer')

        # Variables
        self.pdf_list = []
        self.output_dir = tk.StringVar(value=r"C:\Users\rjiso\OneDrive\Escritorio\PROCESAR ARCHIVO SUELDOS\02_CARPETAS\test_firmas")
        self.page_mode = tk.StringVar(value="Auto") # "Auto", "Manual", "Ultima"
        self.specific_page = tk.IntVar(value=3)
        self.is_processing = False
        self.blue_only = tk.BooleanVar(value=True) 
        self.anchor_text = tk.StringVar(value="Firma del Empleado, Firma del Trabajador")

        self.setup_ui()

    def setup_ui(self):
        # Contenedor principal con scroll
        self.scroll_container = mgc.create_main_container(self, padding=0)

        # Header
        self.header = mgc.create_header(self.scroll_container, "Extractor de Firmas ULTRA", 
                                       "Detección automática de página y anclaje inteligente",
                                       icon_image=load_icon('printer', (64, 64)))
        self.header.pack(fill='x', padx=30, pady=(30, 10))

        # Main Layout
        self.main_container = ctk.CTkFrame(self.scroll_container, fg_color="transparent")
        self.main_container.pack(fill='both', expand=True, padx=30, pady=10)

        # Left Column: List
        self.left_col = ctk.CTkFrame(self.main_container, fg_color="#111827", corner_radius=15, border_width=1, border_color="#1e2a3a")
        self.left_col.pack(side='left', fill='both', expand=True, padx=(0, 10))

        ctk.CTkLabel(self.left_col, text="Lista de PDFs (Batch)", font=("Segoe UI", 16, "bold")).pack(pady=15)
        
        list_frame = ctk.CTkFrame(self.left_col, fg_color="transparent")
        list_frame.pack(fill='both', expand=True, padx=20, pady=5)
        
        self.file_listbox = tk.Listbox(list_frame, font=("Segoe UI", 10), selectmode='extended', relief='flat', borderwidth=0, highlightthickness=0, bg="#0a0e1a", fg="#f1f5f9", selectbackground="#3b82f6", selectforeground="white")
        self.file_listbox.pack(side='left', fill='both', expand=True)
        
        scrollbar = ctk.CTkScrollbar(list_frame, command=self.file_listbox.yview)
        scrollbar.pack(side='right', fill='y')
        self.file_listbox.configure(yscrollcommand=scrollbar.set)

        btn_frame = ctk.CTkFrame(self.left_col, fg_color="transparent")
        btn_frame.pack(fill='x', padx=20, pady=15)
        
        ctk.CTkButton(btn_frame, text="+ AGREGAR", command=self.add_files, fg_color=self.colors['accent']).pack(side='left', fill='x', expand=True, padx=(0, 5))
        ctk.CTkButton(btn_frame, text="- QUITAR", command=self.remove_files, fg_color=self.colors['danger']).pack(side='left', fill='x', expand=True, padx=(5, 0))

        # Right Column: Config
        self.right_col = ctk.CTkFrame(self.main_container, fg_color="#111827", corner_radius=15, border_width=1, border_color="#1e2a3a", width=420)
        self.right_col.pack(side='right', fill='both', padx=(10, 0))
        self.right_col.pack_propagate(False)

        ctk.CTkLabel(self.right_col, text="Configuración Avanzada", font=("Segoe UI", 16, "bold")).pack(pady=15)

        config_frame = ctk.CTkScrollableFrame(self.right_col, fg_color="transparent")
        config_frame.pack(fill='both', expand=True, padx=15)

        # Mode Selection
        ctk.CTkLabel(config_frame, text="Modo de Búsqueda:", font=("Segoe UI", 12, "bold")).pack(anchor='w', pady=(5, 5))
        ctk.CTkRadioButton(config_frame, text="Automático (Escanea Págs 1-10)", variable=self.page_mode, value="Auto").pack(anchor='w', pady=5)
        ctk.CTkRadioButton(config_frame, text="Manual (Página específica)", variable=self.page_mode, value="Manual").pack(anchor='w', pady=5)
        ctk.CTkRadioButton(config_frame, text="Última Página", variable=self.page_mode, value="Ultima").pack(anchor='w', pady=5)

        self.entry_pg = ctk.CTkEntry(config_frame, textvariable=self.specific_page, width=70)
        self.entry_pg.pack(anchor='e', pady=(0, 15))

        # Anchor List
        ctk.CTkLabel(config_frame, text="Anclas de texto (separadas por coma):", font=("Segoe UI", 12, "bold")).pack(anchor='w', pady=(5, 5))
        self.entry_anchor = ctk.CTkEntry(config_frame, textvariable=self.anchor_text, font=("Segoe UI", 10))
        self.entry_anchor.pack(fill='x', pady=(0, 15))

        # Visual Filters
        ctk.CTkLabel(config_frame, text="Filtros Visuales:", font=("Segoe UI", 12, "bold")).pack(anchor='w', pady=(5, 5))
        ctk.CTkSwitch(config_frame, text="Priorizar Tinta AZUL/COLOR", variable=self.blue_only).pack(anchor='w', pady=5)
        
        # Destination
        ctk.CTkLabel(config_frame, text="Destino de Imágenes:", font=("Segoe UI", 12, "bold")).pack(anchor='w', pady=(15, 5))
        self.entry_out = ctk.CTkEntry(config_frame, textvariable=self.output_dir, font=("Segoe UI", 10))
        self.entry_out.pack(fill='x', pady=(0, 10))
        ctk.CTkButton(config_frame, text="Examinar", command=self.select_output, fg_color="#607D8B", height=30).pack(fill='x', pady=(0, 10))

        # Progress
        self.progress = ctk.CTkProgressBar(self.right_col)
        self.progress.pack(fill='x', padx=20, pady=10)
        self.progress.set(0)
        
        self.status_label = ctk.CTkLabel(self.right_col, text="Listo.", font=("Segoe UI", 11), text_color="gray")
        self.status_label.pack(pady=5)

        self.btn_process = ctk.CTkButton(self.right_col, text="🚀 PROCESAR LOTE", command=self.start_processing, 
                                        fg_color=self.colors['success'], height=60, font=("Segoe UI", 14, "bold"))
        self.btn_process.pack(fill='x', padx=20, pady=(0, 25), side='bottom')

    def add_files(self):
        files = filedialog.askopenfilenames(title="Seleccionar PDFs", filetypes=[("PDF", "*.pdf")])
        if files:
            for f in files:
                if f not in self.pdf_list:
                    self.pdf_list.append(f)
                    self.file_listbox.insert(tk.END, os.path.basename(f))

    def remove_files(self):
        for i in reversed(self.file_listbox.curselection()):
            self.pdf_list.pop(i)
            self.file_listbox.delete(i)

    def select_output(self):
        path = filedialog.askdirectory()
        if path: self.output_dir.set(path)

    def start_processing(self):
        if not self.pdf_list:
            messagebox.showwarning("Aviso", "Añade archivos PDF.")
            return
        self.is_processing = True
        self.btn_process.configure(state='disabled')
        threading.Thread(target=self.run_logic, daemon=True).start()

    def run_logic(self):
        try:
            out = self.output_dir.get()
            os.makedirs(out, exist_ok=True)
            total = len(self.pdf_list)
            
            for i, pdf in enumerate(self.pdf_list):
                self.status_label.configure(text=f"Analizando: {os.path.basename(pdf)}")
                self.progress.set((i + 0.1) / total)
                self.extract_core(pdf, out)
                self.progress.set((i + 1) / total)

            self.after(0, self.finish_processing)
        except Exception as e:
            self.after(0, lambda: messagebox.showerror("Error", str(e)))
            self.after(0, self.finish_processing)

    def extract_core(self, pdf_path, output_folder):
        doc = fitz.open(pdf_path)
        mode = self.page_mode.get()
        anchors = [a.strip() for a in self.anchor_text.get().split(",")]
        
        target_page = -1
        target_rect = None
        
        if mode == "Auto":
            # Escanear páginas 1 a 10 (o total)
            scan_limit = min(10, len(doc))
            for p_idx in range(scan_limit):
                page = doc.load_page(p_idx)
                for anchor in anchors:
                    res = page.search_for(anchor)
                    if res:
                        target_page = p_idx
                        inst = res[0]
                        target_rect = fitz.Rect(inst.x0 - 50, inst.y0 - 150, inst.x1 + 50, inst.y0)
                        break
                if target_page != -1: break
        elif mode == "Manual":
            target_page = min(self.specific_page.get() - 1, len(doc) - 1)
        elif mode == "Ultima":
            target_page = len(doc) - 1

        if target_page == -1: target_page = 0 # Fallback

        page = doc.load_page(target_page)
        
        # Si no hay rect de anclaje (porque era manual o ultima), usar tercio inferior
        if not target_rect:
            h_pdf = page.rect.height
            target_rect = fitz.Rect(0, h_pdf * 0.5, page.rect.width, h_pdf)

        # Render
        zoom = 4
        pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), clip=target_rect, alpha=True)
        img = Image.frombytes("RGBA", [pix.width, pix.height], pix.samples)
        
        # Filters
        data = np.array(img)
        r, g, b, a = data[:,:,0], data[:,:,1], data[:,:,2], data[:,:,3]
        
        # Transparency for white
        white_mask = (r > 240) & (g > 240) & (b > 240)
        a[white_mask] = 0
        
        # Blue filter
        if self.blue_only.get():
            blue_mask = (b > r + 25) & (b > g + 25)
            a[~blue_mask] = 0
            
        data[:,:,3] = a
        filtered_img = Image.fromarray(data, 'RGBA')
        
        bbox = filtered_img.getbbox()
        if bbox:
            final_img = filtered_img.crop(bbox)
            base_name = os.path.splitext(os.path.basename(pdf_path))[0]
            final_img.save(os.path.join(output_folder, f"Firma_{base_name}.png"))
            
        doc.close()

    def finish_processing(self):
        self.is_processing = False
        self.btn_process.configure(state='normal')
        self.status_label.configure(text="Sincronizado.")
        if messagebox.askyesno("Éxito", "¿Abrir carpeta de firmas?"):
            os.startfile(self.output_dir.get())

if __name__ == "__main__":
    app = ExtractorFirmasApp()
    app.mainloop()
