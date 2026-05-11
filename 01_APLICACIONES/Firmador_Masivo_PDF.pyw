# -*- coding: utf-8 -*-
"""
✍️ HERRAMIENTA PARA FIRMAR PDFS EN MASA
Permite elegir una firma PNG y estamparla en todas las hojas de un PDF.
Diseño Premium integrado con el Sistema de Sueldos.
"""

import sys
import tkinter as tk
from tkinter import filedialog, messagebox
import os
import fitz  # PyMuPDF

# Asegurar que el directorio del script esté en sys.path para imports locales
script_dir = os.path.dirname(os.path.abspath(__file__))
if script_dir not in sys.path:
    sys.path.insert(0, script_dir)

# También agregar 03_OTROS (donde se encuentran los submódulos del sistema)
others_dir = os.path.abspath(os.path.join(script_dir, "03_OTROS"))
if others_dir not in sys.path:
    sys.path.insert(0, others_dir)

# Importar componentes modernos
import customtkinter as ctk
import modern_gui_components as mgc
from icon_loader import set_window_icon, load_icon

# Configuración de apariencia
ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

class FirmadorPDFApp:
    def __init__(self, root):
        self.root = root
        self.root.title("✍️ Firmador Masivo de PDF")
        self.root.geometry("900x700")
        self.root.resizable(False, False)
        self.root.configure(fg_color=mgc.COLORS['bg_primary'])
        
        mgc.center_window(self.root, 900, 700)
        
        # Establecer icono de ventana
        set_window_icon(self.root, 'printer')
        
        # Cargar iconos PNG
        self.icon_app = load_icon('printer', (64, 64))
        self.icon_check = load_icon('check', (24, 24))
        
        # Variables
        self.pdf_path = tk.StringVar()
        self.img_path = tk.StringVar()
        self.posicion = tk.StringVar(value="Abajo Derecha")
        self.escala = tk.DoubleVar(value=1.0)
        
        self.setup_ui()

    def setup_ui(self):
        # Contenedor principal con scroll
        main_frame = mgc.create_main_container(self.root)
        
        # Header moderno
        mgc.create_header(main_frame, "Firmador Masivo PDF", 
                         "Estampa firmas en todas las hojas de tus documentos PDF de forma automática", 
                         icon_image=self.icon_app)
        
        # Separador inicial
        tk.Frame(main_frame, height=1, bg=mgc.COLORS['border']).pack(fill=tk.X, pady=(0, 20))
        
        # Card de Selección de Archivos
        files_card, files_inner = mgc.create_card(main_frame, "📂 Selección de Archivos", padding=20)
        files_card.pack(fill=tk.X, pady=(0, 20))
        
        # Fila PDF
        ctk.CTkLabel(files_inner, text="1. Documento PDF a firmar:", font=mgc.FONTS['small'], text_color=mgc.COLORS['text_primary']).pack(anchor="w")
        f_pdf = ctk.CTkFrame(files_inner, fg_color="transparent")
        f_pdf.pack(fill=tk.X, pady=(5, 15))
        ctk.CTkEntry(f_pdf, textvariable=self.pdf_path, placeholder_text="Seleccione el archivo PDF...", 
                    height=35, font=mgc.FONTS['small'], fg_color=mgc.COLORS['bg_input'], border_color=mgc.COLORS['border']).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))
        mgc.create_button(f_pdf, "Explorar", self.buscar_pdf, color='blue', width=100).pack(side=tk.RIGHT)
        
        # Fila Firma
        ctk.CTkLabel(files_inner, text="2. Imagen de la firma (PNG transparente recomendada):", font=mgc.FONTS['small'], text_color=mgc.COLORS['text_primary']).pack(anchor="w")
        f_img = ctk.CTkFrame(files_inner, fg_color="transparent")
        f_img.pack(fill=tk.X, pady=(5, 0))
        ctk.CTkEntry(f_img, textvariable=self.img_path, placeholder_text="Seleccione la imagen de la firma...", 
                    height=35, font=mgc.FONTS['small'], fg_color=mgc.COLORS['bg_input'], border_color=mgc.COLORS['border']).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))
        mgc.create_button(f_img, "Explorar", self.buscar_firma, color='blue', width=100).pack(side=tk.RIGHT)
        
        # Card de Configuración de Estampado
        config_card, config_inner = mgc.create_card(main_frame, "⚙️ Configuración de Estampado", padding=20)
        config_card.pack(fill=tk.X, pady=(0, 20))
        
        # Layout de dos columnas para configuración
        col_frame = ctk.CTkFrame(config_inner, fg_color="transparent")
        col_frame.pack(fill=tk.X)
        
        # Columna Izquierda: Posición
        left_col = ctk.CTkFrame(col_frame, fg_color="transparent")
        left_col.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 20))
        
        ctk.CTkLabel(left_col, text="3. Ubicación en la hoja:", font=mgc.FONTS['small'], text_color=mgc.COLORS['text_primary']).pack(anchor="w")
        opciones_pos = ["Arriba Izquierda", "Arriba Derecha", "Abajo Izquierda", "Abajo Derecha"]
        self.combo_pos = ctk.CTkComboBox(left_col, values=opciones_pos, variable=self.posicion, 
                                        width=250, height=35, font=mgc.FONTS['small'], 
                                        fg_color=mgc.COLORS['bg_input'], border_color=mgc.COLORS['border'],
                                        button_color=mgc.COLORS['blue'], dropdown_fg_color=mgc.COLORS['bg_card'])
        self.combo_pos.pack(anchor="w", pady=5)
        
        # Columna Derecha: Escala
        right_col = ctk.CTkFrame(col_frame, fg_color="transparent")
        right_col.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        ctk.CTkLabel(right_col, text="4. Tamaño de la firma (Escala):", font=mgc.FONTS['small'], text_color=mgc.COLORS['text_primary']).pack(anchor="w")
        self.slider = ctk.CTkSlider(right_col, from_=0.1, to=2.0, number_of_steps=19, variable=self.escala,
                                   button_color=mgc.COLORS['blue'], button_hover_color=mgc.COLORS['accent_blue'],
                                   progress_color=mgc.COLORS['blue'])
        self.slider.pack(fill=tk.X, pady=12)
        
        # Botón Procesar
        self.btn_procesar = mgc.create_large_button(main_frame, "FIRMAR TODAS LAS HOJAS", 
                                                  self.procesar, color='green', icon_image=self.icon_check)
        self.btn_procesar.pack(pady=20)
        
        # Barra de estado
        mgc.create_status_bar(self.root, "Listo para estampar firmas")

    def buscar_pdf(self):
        filename = filedialog.askopenfilename(title="Seleccionar PDF", filetypes=[("PDF files", "*.pdf")])
        if filename:
            self.pdf_path.set(filename)

    def buscar_firma(self):
        filename = filedialog.askopenfilename(title="Seleccionar Firma PNG", filetypes=[("Image files", "*.png")])
        if filename:
            self.img_path.set(filename)

    def procesar(self):
        pdf = self.pdf_path.get()
        img = self.img_path.get()
        
        if not pdf or not img:
            messagebox.showerror("Error", "Debe seleccionar el PDF y la firma.")
            return
            
        try:
            # Abrir PDF
            doc = fitz.open(pdf)
            
            # Abrir Imagen para obtener dimensiones
            img_doc = fitz.open(img)
            img_rect = img_doc[0].rect
            img_w = img_rect.width
            img_h = img_rect.height
            img_doc.close()
            
            # Aplicar escala
            factor = self.escala.get()
            w = img_w * factor
            h = img_h * factor
            
            # Margen
            margin = 30
            
            pos = self.posicion.get()
            
            for page in doc:
                p_rect = page.rect
                pw = p_rect.width
                ph = p_rect.height
                
                if pos == "Arriba Izquierda":
                    rect = fitz.Rect(margin, margin, margin + w, margin + h)
                elif pos == "Arriba Derecha":
                    rect = fitz.Rect(pw - margin - w, margin, pw - margin, margin + h)
                elif pos == "Abajo Izquierda":
                    rect = fitz.Rect(margin, ph - margin - h, margin + w, ph - margin)
                else: # Abajo Derecha
                    rect = fitz.Rect(pw - margin - w, ph - margin - h, pw - margin, ph - margin)
                
                page.insert_image(rect, filename=img)
            
            # Guardar
            output_path = pdf.replace(".pdf", "_firmado.pdf")
            doc.save(output_path)
            doc.close()
            
            messagebox.showinfo("Éxito", f"PDF firmado generado correctamente:\n{os.path.basename(output_path)}")
            os.startfile(os.path.dirname(output_path))
            
        except Exception as e:
            messagebox.showerror("Error Crítico", f"Ocurrió un error: {str(e)}")

if __name__ == "__main__":
    root = ctk.CTk()
    app = FirmadorPDFApp(root)
    root.mainloop()
