import os
import PyPDF2
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import tkinter.scrolledtext as scrolledtext
import sys

# Asegurar que el directorio del script esté en sys.path para imports locales
script_dir = os.path.dirname(os.path.abspath(__file__))
if script_dir not in sys.path:
    sys.path.insert(0, script_dir)

# Importar componentes modernos
import modern_gui_components as mgc
from icon_loader import set_window_icon, load_icon

class SeparadorRecibosApp:
    def __init__(self, root):
        self.root = root
        self.root.title("🔍 Gestor de Recibos PDF")
        self.root.geometry("900x650")
        self.root.resizable(True, True)
        self.root.configure(bg=mgc.COLORS['bg_primary'])
        
        mgc.center_window(self.root, 900, 650)
        
        # Establecer icono de ventana
        set_window_icon(self.root, 'search')
        
        # Cargar iconos PNG para la interfaz
        self.icon_search = load_icon('search', (64, 64))
        self.icon_pdf = load_icon('pdf', (24, 24))
        self.icon_folder = load_icon('folder', (24, 24))
        self.icon_warning = load_icon('warning', (24, 24))
        self.icon_check = load_icon('check', (24, 24))

        # Barra de estado inferior (Crear primero)
        self.status_frame, self.status_var = mgc.create_status_bar(self.root, "Listo para comenzar")

        self.archivos_pdf_seleccionados = []
        self.recibos_encontrados_pages = []
        self.open_pdf_readers = []

        # Frame Principal
        main_frame = tk.Frame(self.root, bg=mgc.COLORS['bg_primary'])
        main_frame.pack(fill=tk.BOTH, expand=True, padx=30, pady=20)
        
        # Header
        mgc.create_header(main_frame, "Buscador de Recibos PDF", 
                         "Busca y unifica recibos de una persona específica", 
                         icon_image=self.icon_search)

        # Grid para Cards Superiores
        top_grid = tk.Frame(main_frame, bg=mgc.COLORS['bg_primary'])
        top_grid.columnconfigure(0, weight=1)
        top_grid.columnconfigure(1, weight=1)
        top_grid.rowconfigure(0, weight=1)

        # --- Card 1: Selección de PDFs ---
        card1_outer, card1_inner = mgc.create_card(top_grid, "1. Archivos PDF a procesar", padding=15)
        card1_outer.grid(row=0, column=0, sticky="nsew", padx=(0, 5))

        list_container = tk.Frame(card1_inner, bg=mgc.COLORS['bg_card'])
        list_container.pack(fill=tk.BOTH, expand=True)

        self.listbox_archivos = tk.Listbox(list_container, selectmode=tk.MULTIPLE, height=5, 
                                          font=('Consolas', 9), bg='white', relief=tk.GROOVE, bd=1)
        self.listbox_archivos.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        scrollbar = ttk.Scrollbar(list_container, orient="vertical", command=self.listbox_archivos.yview)
        scrollbar.pack(side=tk.LEFT, fill=tk.Y)
        self.listbox_archivos.config(yscrollcommand=scrollbar.set)

        btn_files_frame = tk.Frame(card1_inner, bg=mgc.COLORS['bg_card'])
        btn_files_frame.pack(fill=tk.X, pady=(10, 0))
        
        mgc.create_button(btn_files_frame, "Añadir PDFs", self.seleccionar_archivos_pdf_gui, color='blue', icon_image=self.icon_folder, padx=10, pady=5).pack(side=tk.LEFT, padx=2)
        mgc.create_button(btn_files_frame, "Eliminar", self.eliminar_archivos_seleccionados, color='red', icon_image=self.icon_warning, padx=10, pady=5).pack(side=tk.LEFT, padx=2)

        # --- Card 2: Búsqueda y Acción ---
        card2_outer, card2_inner = mgc.create_card(top_grid, "2. Parámetros de Búsqueda", padding=15)
        card2_outer.grid(row=0, column=1, sticky="nsew", padx=(5, 0))

        tk.Label(card2_inner, text="Nombre y Apellido a buscar:", bg=mgc.COLORS['bg_card'], font=mgc.FONTS['normal']).pack(anchor='w', pady=(0, 5))
        self.entry_nombre_apellido = tk.Entry(card2_inner, font=mgc.FONTS['normal'], relief=tk.GROOVE, bd=1)
        self.entry_nombre_apellido.pack(fill=tk.X, pady=(0, 15))

        self.btn_procesar = mgc.create_large_button(card2_inner, "BUSCAR Y UNIFICAR", self.procesar_recibos, color='green', icon_image=self.icon_check, padx=40, pady=15)
        self.btn_procesar.pack(pady=5)

        # --- Card 3: Estado (Área de Log) ---
        card3_outer, card3_inner = mgc.create_card(main_frame, "Estado del Proceso", padding=10)
        card3_outer.pack(side=tk.BOTTOM, fill=tk.X, pady=(5, 0))

        self.status_text = scrolledtext.ScrolledText(card3_inner, height=6, font=('Consolas', 9), bg='#f8f9fa', state='disabled', relief=tk.FLAT)
        self.status_text.pack(fill=tk.BOTH, expand=True)

        # Empaquetar Grid Superior al final para que ocupe el espacio restante
        top_grid.pack(side=tk.TOP, fill=tk.BOTH, expand=True, pady=(0, 10))

        self.update_file_listbox()
        self.update_status("Bienvenido. Seleccione los archivos PDF para comenzar.")

    def seleccionar_archivos_pdf_gui(self):
        nuevos_archivos = filedialog.askopenfilenames(
            title="Añadir uno o más archivos PDF de recibos",
            filetypes=[("Archivos PDF", "*.pdf")]
        )
        if nuevos_archivos:
            for archivo in nuevos_archivos:
                if archivo not in self.archivos_pdf_seleccionados:
                    self.archivos_pdf_seleccionados.append(archivo)
            self.update_file_listbox()
            self.update_status(f"Archivos añadidos. Total: {len(self.archivos_pdf_seleccionados)}.")
            self.status_var.set(f"✓ {len(self.archivos_pdf_seleccionados)} archivos listos")
        else:
            self.update_status("Selección de archivos cancelada.")

    def eliminar_archivos_seleccionados(self):
        selected_indices = self.listbox_archivos.curselection()
        if not selected_indices:
            messagebox.showwarning("Advertencia", "Por favor, selecciona al menos un archivo para eliminar.")
            return

        for index in selected_indices[::-1]:
            del self.archivos_pdf_seleccionados[index]
        
        self.update_file_listbox()
        self.update_status(f"Archivos eliminados. Total: {len(self.archivos_pdf_seleccionados)}.")
        self.status_var.set(f"Archivos restantes: {len(self.archivos_pdf_seleccionados)}")

    def update_file_listbox(self):
        self.listbox_archivos.delete(0, tk.END)
        if not self.archivos_pdf_seleccionados:
            self.listbox_archivos.insert(tk.END, "Ningún PDF seleccionado...")
            self.listbox_archivos.itemconfig(0, fg=mgc.COLORS['gray'])
        else:
            for archivo in self.archivos_pdf_seleccionados:
                self.listbox_archivos.insert(tk.END, f"📄 {os.path.basename(archivo)}")

    def buscar_y_separar_recibos(self, nombre_apellido):
        self.recibos_encontrados_pages = []
        nombre_apellido_lower = nombre_apellido.lower()
        recibos_encontrados_count = 0
        self.open_pdf_readers = []

        if not self.archivos_pdf_seleccionados:
            messagebox.showwarning("Advertencia", "Por favor, selecciona al menos un archivo PDF.")
            return []

        try:
            for archivo_pdf in self.archivos_pdf_seleccionados:
                self.update_status(f"Buscando en: {os.path.basename(archivo_pdf)}...")
                self.root.update_idletasks()

                try:
                    reader = PyPDF2.PdfReader(archivo_pdf)
                    self.open_pdf_readers.append(reader)

                    for page_num in range(len(reader.pages)):
                        page = reader.pages[page_num]
                        text = page.extract_text()
                        if text and nombre_apellido_lower in text.lower():
                            self.recibos_encontrados_pages.append(page)
                            recibos_encontrados_count += 1
                            self.update_status(f"✓ Encontrado en página {page_num + 1} de '{os.path.basename(archivo_pdf)}'")
                            self.root.update_idletasks()
                except Exception as e:
                    self.update_status(f"❌ ERROR en {os.path.basename(archivo_pdf)}: {e}")
        finally:
            pass
        return self.recibos_encontrados_pages

    def unificar_recibos(self, recibos_encontrados):
        if not recibos_encontrados:
            messagebox.showinfo("Sin Recibos", "No se encontraron recibos para unificar.")
            self.update_status("Proceso finalizado: No se encontraron recibos.")
            return

        writer = PyPDF2.PdfWriter()
        for page in recibos_encontrados:
            writer.add_page(page)

        try:
            output_path = filedialog.asksaveasfilename(
                defaultextension=".pdf",
                filetypes=[("Archivos PDF", "*.pdf")],
                initialfile=f"Recibos_{self.entry_nombre_apellido.get().replace(' ', '_')}.pdf",
                title="Guardar archivo unificado como..."
            )
            if output_path:
                with open(output_path, 'wb') as output_pdf:
                    writer.write(output_pdf)
                messagebox.showinfo("Éxito", f"Los recibos se han unificado en '{output_path}'.")
                self.update_status(f"🎯 ÉXITO: Archivo guardado en '{output_path}'.")
            else:
                self.update_status("⚠ Guardado cancelado por el usuario.")
        except Exception as e:
            messagebox.showerror("Error al guardar PDF", f"No se pudo guardar el archivo de salida: {e}")
            self.update_status(f"❌ ERROR al guardar: {e}")
        finally:
            for reader in self.open_pdf_readers:
                try:
                    if hasattr(reader.stream, 'close'):
                        reader.stream.close()
                except: pass
            self.open_pdf_readers = []

    def procesar_recibos(self):
        nombre_apellido = self.entry_nombre_apellido.get().strip()

        if not self.archivos_pdf_seleccionados:
            return messagebox.showwarning("Advertencia", "Por favor, selecciona los archivos PDF primero.")

        if not nombre_apellido:
            return messagebox.showwarning("Advertencia", "Por favor, ingresa el nombre y apellido a buscar.")

        self.update_status("🚀 Iniciando búsqueda de recibos...")
        mgc.disable_button(self.btn_procesar)
        self.status_var.set("⏳ Procesando PDFs...")
        self.root.update_idletasks()

        recibos = self.buscar_y_separar_recibos(nombre_apellido)
        if recibos is not None:
            self.unificar_recibos(recibos)

        mgc.enable_button(self.btn_procesar, 'green')
        self.status_var.set("Listo")

    def update_status(self, message):
        self.status_text.config(state="normal")
        self.status_text.insert(tk.END, message + "\n")
        self.status_text.see(tk.END)
        self.status_text.config(state="disabled")

if __name__ == "__main__":
    root = tk.Tk()
    app = SeparadorRecibosApp(root)
    root.mainloop()
