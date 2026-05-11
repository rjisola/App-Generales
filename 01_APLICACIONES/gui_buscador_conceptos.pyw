# -*- coding: utf-8 -*-
import tkinter as tk
import customtkinter as ctk
from tkinter import filedialog, messagebox, ttk
import traceback
import os
import sys
import pdfplumber  # <-- Asegurado para la nueva lógica
import re          # <-- Asegurado para la nueva lógica

def format_nombre_propio(nombre):
    if not nombre or nombre == "No encontrado": return "No encontrado"
    nombre = str(nombre).replace(',', ' ')
    return ' '.join(nombre.split()).title()

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

class BuscadorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("🔍 Extracción de Datos a Lección")
        self.root.geometry("900x700")
        self.root.resizable(True, True)
        self.root.configure(bg=mgc.COLORS['bg_primary'])

        mgc.center_window(self.root, 900, 700)

        if _has_icon_loader:
            set_window_icon(self.root, 'concepts')

        # Barra de estado inferior
        self.status_frame, self.status_var = mgc.create_status_bar(
            self.root, "✓ Listo — Seleccione archivo(s) y opciones"
        )

        # Contenedor principal con scrollbar
        main_frame = mgc.create_main_container(self.root)

        # Header Premium
        mgc.create_header(
            main_frame, 
            "Extracción de Datos a Lección", 
            "Extraer importes o pivotear datos de conceptos (ej: Vacaciones) por empleado",
            icon="🔍"
        )

        # Variables
        self.pdf_paths = []
        self.concept_var = tk.StringVar(value="VACACIONES")
        
        # Variables Modo Avanzado
        self.var_modo = tk.StringVar(value="regular")
        self.var_fecha_tope = tk.StringVar(value="31/12/2026")
        self.var_indice = tk.StringVar()

        # Card: Configuración
        card_outer, card_inner = mgc.create_card(main_frame, "📄  Archivos y Filtros", padding=12)
        card_outer.pack(fill=tk.X, pady=(0, 10))

        # Fila Split: Izquierda (PDFs), Derecha (Concepto y Modo)
        split_frame = tk.Frame(card_inner, bg=mgc.COLORS['bg_card'])
        split_frame.pack(fill=tk.X)
        
        left_f = tk.Frame(split_frame, bg=mgc.COLORS['bg_card'])
        left_f.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 10))
        
        right_f = tk.Frame(split_frame, bg=mgc.COLORS['bg_card'])
        right_f.pack(side=tk.LEFT, fill=tk.BOTH, expand=False)

        # --- LEFT (PDFs) ---
        tk.Label(left_f, text="1. Archivos PDF a analizar (Quincenas):",
                 font=mgc.FONTS['normal'], bg=mgc.COLORS['bg_card'],
                 fg=mgc.COLORS['text_primary']).pack(anchor='w', pady=(0, 4))
                 
        pdf_frame = tk.Frame(left_f, bg=mgc.COLORS['bg_card'])
        pdf_frame.pack(fill=tk.X)
        
        list_frame = tk.Frame(pdf_frame, bg=mgc.COLORS['bg_card'])
        list_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        self.pdf_listbox = tk.Listbox(list_frame, height=4, font=('Segoe UI', 9), relief=tk.GROOVE, bd=1, selectmode=tk.EXTENDED)
        self.pdf_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        ll_scroll = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.pdf_listbox.yview)
        ll_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.pdf_listbox.configure(yscrollcommand=ll_scroll.set)
        
        btn_frame = tk.Frame(pdf_frame, bg=mgc.COLORS['bg_card'])
        btn_frame.pack(side=tk.LEFT, fill=tk.Y, padx=(5, 0))
        
        btn_add = tk.Button(btn_frame, text="➕", font=mgc.FONTS['normal'], 
                            bg=mgc.COLORS.get('primary', '#4f46e5'), fg='white', relief=tk.FLAT,
                            command=self.add_pdfs, width=3)
        btn_add.pack(pady=(0, 2))
        self._add_hover(btn_add, 'primary', '#4338ca')
        
        btn_rem = tk.Button(btn_frame, text="🗑", font=mgc.FONTS['normal'],
                            bg=mgc.COLORS.get('danger', '#ef4444'), fg='white', relief=tk.FLAT,
                            command=self.remove_pdfs, width=3)
        btn_rem.pack(pady=(2, 0))
        self._add_hover(btn_rem, 'danger', '#dc2626')

        # --- RIGHT (Concepto) ---
        tk.Label(right_f, text="2. Concepto a Buscar:",
                 font=mgc.FONTS['normal'], bg=mgc.COLORS['bg_card'],
                 fg=mgc.COLORS['text_primary']).pack(anchor='w', pady=(0, 4))
        tk.Entry(right_f, textvariable=self.concept_var,
                 font=mgc.FONTS['normal'], width=25,
                 relief=tk.GROOVE, bd=1).pack(anchor='w', pady=(0, 15))
                 
        # Modo de Extracción
        tk.Label(right_f, text="Modo de Análisis:", font=mgc.FONTS['normal'], 
                 bg=mgc.COLORS['bg_card'], fg=mgc.COLORS['text_primary']).pack(anchor='w', pady=(0, 2))
                 
        tk.Radiobutton(right_f, text="Modo Regular: Extrae el Importe ($) por renglón", 
                       variable=self.var_modo, value="regular", command=self.toggle_mode,
                       bg=mgc.COLORS['bg_card'], fg=mgc.COLORS['text_primary'],
                       font=mgc.FONTS['small'], activebackground=mgc.COLORS['bg_card'], selectcolor=mgc.COLORS['bg_primary']).pack(anchor='w')
                       
        tk.Radiobutton(right_f, text="Modo Vacaciones: Extrae Días, Pivotea y Resta", 
                       variable=self.var_modo, value="vacaciones", command=self.toggle_mode,
                       bg=mgc.COLORS['bg_card'], fg=mgc.COLORS['text_primary'],
                       font=mgc.FONTS['small'], activebackground=mgc.COLORS['bg_card'], selectcolor=mgc.COLORS['bg_primary']).pack(anchor='w', pady=(0,10))

        # Filtro de Firma
        self.var_filtro_firma = tk.StringVar(value="empleador")
        tk.Label(right_f, text="Filtrar por Copia (para evitar duplicados):", font=mgc.FONTS['normal'], 
                 bg=mgc.COLORS['bg_card'], fg=mgc.COLORS['text_primary']).pack(anchor='w', pady=(0, 2))
                 
        f_firmas = tk.Frame(right_f, bg=mgc.COLORS['bg_card'])
        f_firmas.pack(anchor='w', pady=(0,5))
        
        tk.Radiobutton(f_firmas, text="Ambas", variable=self.var_filtro_firma, value="ambas",
                       bg=mgc.COLORS['bg_card'], fg=mgc.COLORS['text_primary'], font=mgc.FONTS['small'],
                       activebackground=mgc.COLORS['bg_card'], selectcolor=mgc.COLORS['bg_primary']).pack(side=tk.LEFT)
        tk.Radiobutton(f_firmas, text="Solo Empleador", variable=self.var_filtro_firma, value="empleador",
                       bg=mgc.COLORS['bg_card'], fg=mgc.COLORS['text_primary'], font=mgc.FONTS['small'],
                       activebackground=mgc.COLORS['bg_card'], selectcolor=mgc.COLORS['bg_primary']).pack(side=tk.LEFT, padx=(5,0))
        tk.Radiobutton(f_firmas, text="Solo Empleado", variable=self.var_filtro_firma, value="empleado",
                       bg=mgc.COLORS['bg_card'], fg=mgc.COLORS['text_primary'], font=mgc.FONTS['small'],
                       activebackground=mgc.COLORS['bg_card'], selectcolor=mgc.COLORS['bg_primary']).pack(side=tk.LEFT, padx=(5,0))

        # Card: Modo Avanzado (Ocultable o Mostrable)
        self.card_opt_outer, self.card_opt_inner = mgc.create_card(main_frame, "⚙️  Opciones Avanzadas (Control de Vacaciones)", padding=10)
        self.card_opt_outer.pack(fill=tk.X, pady=(0, 10))
        
        # Opciones Avanzadas
        f_idx = tk.Frame(self.card_opt_inner, bg=mgc.COLORS['bg_card'])
        f_idx.pack(fill=tk.X, pady=(0, 5))
        
        mgc.create_file_selector(
            f_idx, "Índice de Personal Excel (Opcional pero Recomendado):",
            self.var_indice, self.browse_indice, "📊"
        ).pack(fill=tk.X)
        
        f_date = tk.Frame(self.card_opt_inner, bg=mgc.COLORS['bg_card'])
        f_date.pack(fill=tk.X, pady=(10, 0))
        tk.Label(f_date, text="Fecha Tope para Calcular Antigüedad (DD/MM/YYYY):",
                 font=mgc.FONTS['normal'], bg=mgc.COLORS['bg_card'],
                 fg=mgc.COLORS['text_primary']).pack(side=tk.LEFT)
        tk.Entry(f_date, textvariable=self.var_fecha_tope,
                 font=mgc.FONTS['normal'], width=15,
                 relief=tk.GROOVE, bd=1).pack(side=tk.LEFT, padx=(10, 0))

        # Card: Ejecutar
        self.card_act_outer, card_act_inner = mgc.create_card(main_frame, "✅  Ejecutar Búsqueda", padding=12)
        self.card_act_outer.pack(fill=tk.X, pady=(0, 10))

        btn = mgc.create_large_button(
            card_act_inner, "Realizar Búsqueda / Armar Reporte",
            self.run_process, color='green', icon="🔍"
        )
        btn.pack(pady=5)
        self._add_hover(btn, 'green', '#059669')

        # Card: Resultados
        card_res_outer, self.card_res_inner = mgc.create_card(main_frame, "📊  Resultados", padding=12)
        card_res_outer.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        # Frame for Treeview and Scrollbar
        self.tree_frame = tk.Frame(self.card_res_inner, bg=mgc.COLORS['bg_card'])
        self.tree_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        # Inicializar Treeview (se reescribirá en runtime)
        self.tree = ttk.Treeview(self.tree_frame, show='headings', height=8)
        self.build_tree_columns([("Archivo", 150), ("Legajo", 80), ("Nombre y Apellido", 200), ("Concepto", 120), ("Importe", 80)])
        
        # Export Button
        self.btn_export = mgc.create_large_button(
            self.card_res_inner, "Exportar Reporte a Excel",
            self.export_excel, color='blue', icon="📊"
        )
        self.btn_export.pack(pady=5)
        self._add_hover(self.btn_export, 'blue', '#1d4ed8')
        
        self.toggle_mode()

    # --- NUEVA FUNCIÓN INTERNA DE EXTRACCIÓN ROBUSTA ---
    def extraer_datos_pdf_regular(self, pdf_paths, concepto_buscado, callback=None, filtro_firma="ambas"):
        resultados = []
        for ruta_pdf in pdf_paths:
            nombre_archivo = os.path.basename(ruta_pdf)
            if callback:
                callback(f"Analizando: {nombre_archivo}...")
                
            try:
                with pdfplumber.open(ruta_pdf) as pdf:
                    for i, pagina in enumerate(pdf.pages):
                        texto = pagina.extract_text(layout=True) 
                        if not texto: continue
                        
                        text_clean = texto.lower().replace(" ", "")
                        if filtro_firma != "ambas":
                            is_empleador = 'firmadelempleador' in text_clean
                            is_empleado = 'firmadelempleado' in text_clean and not is_empleador
                            
                            if filtro_firma == "empleador" and not is_empleador:
                                continue
                            elif filtro_firma == "empleado" and not is_empleado:
                                continue

                        lineas = texto.split('\n')
                        legajo = "No encontrado"
                        nombre = "No encontrado"

                        # 1. Buscar Legajo y Nombre (leyendo el encabezado)
                        for idx, linea in enumerate(lineas):
                            if "LEGAJO" in linea.upper() and "NOMBRE" in linea.upper():
                                if idx + 1 < len(lineas):
                                    linea_datos = lineas[idx + 1]
                                    match_empleado = re.search(r'\b(\d{1,4})\s+([A-Za-zÑñÁÉÍÓÚáéíóú\s,]{5,}[A-Za-zÑñÁÉÍÓÚáéíóú])', linea_datos)
                                    
                                    if match_empleado:
                                        legajo = match_empleado.group(1)
                                        nombre = format_nombre_propio(match_empleado.group(2))
                                    else:
                                        if idx + 2 < len(lineas):
                                            linea_datos_2 = lineas[idx + 2]
                                            match_empleado_2 = re.search(r'\b(\d{1,4})\s+([A-Za-zÑñÁÉÍÓÚáéíóú\s,]{5,}[A-Za-zÑñÁÉÍÓÚáéíóú])', linea_datos_2)
                                            if match_empleado_2:
                                                legajo = match_empleado_2.group(1)
                                                nombre = format_nombre_propio(match_empleado_2.group(2))
                                break

                        # 2. Buscar el concepto y su importe
                        for linea_idx, linea in enumerate(lineas):
                            if concepto_buscado.lower() in linea.lower():
                                importe = None
                                montos = re.findall(r'\d{1,3}(?:\.\d{3})*(?:,\d{2})', linea)
                                if montos:
                                    importe = montos[-1]
                                else:
                                    # Buscar en las siguientes 3 líneas si no está en la misma
                                    for offset in range(1, 4):
                                        if linea_idx + offset < len(lineas):
                                            linea_sig = lineas[linea_idx + offset]
                                            montos_sig = re.findall(r'\d{1,3}(?:\.\d{3})*(?:,\d{2})', linea_sig)
                                            if montos_sig:
                                                importe = montos_sig[-1]
                                                break
                                
                                if importe:
                                    resultados.append({
                                        "Archivo": nombre_archivo,
                                        "Legajo": legajo,
                                        "Nombre y Apellido": nombre,
                                        "Concepto": concepto_buscado,
                                        "Importe": importe
                                    })
                                break
            except Exception as e:
                print(f"Error procesando {nombre_archivo}: {e}")
                
        resultados.sort(key=lambda x: x.get("Nombre y Apellido", ""))
        return resultados
    # ----------------------------------------------------

    def build_tree_columns(self, cols_with_widths):
        # Destruir viejo scrollbar si hay
        if hasattr(self, 'tree_scrollbarX'): self.tree_scrollbarX.destroy()
        if hasattr(self, 'tree_scrollbarY'): self.tree_scrollbarY.destroy()
        self.tree.destroy()
        
        # Recrear
        self.tree = ttk.Treeview(self.tree_frame, show='headings', height=8)
        
        cols = [c[0] for c in cols_with_widths]
        self.tree["columns"] = cols
        for col_name, col_width in cols_with_widths:
            self.tree.heading(col_name, text=col_name)
            self.tree.column(col_name, width=col_width, anchor='center' if col_width < 100 else 'w')
            
        # Scrollbars
        self.tree_scrollbarY = ttk.Scrollbar(self.tree_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree_scrollbarY.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.tree_scrollbarX = ttk.Scrollbar(self.tree_frame, orient=tk.HORIZONTAL, command=self.tree.xview)
        self.tree_scrollbarX.pack(side=tk.BOTTOM, fill=tk.X)
        
        self.tree.configure(yscrollcommand=self.tree_scrollbarY.set, xscrollcommand=self.tree_scrollbarX.set)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

    def _add_hover(self, btn, normal_color_key, hover_hex):
        normal_bg = mgc.COLORS.get(normal_color_key, normal_color_key)
        btn.bind("<Enter>", lambda e: btn.configure(bg=hover_hex))
        btn.bind("<Leave>", lambda e: btn.configure(bg=normal_bg))

    def toggle_mode(self):
        if self.var_modo.get() == "vacaciones":
            self.card_opt_outer.pack(fill=tk.X, pady=(0, 10), before=self.card_act_outer)
        else:
            self.card_opt_outer.pack_forget()

    def update_status_count(self):
        count = self.pdf_listbox.size()
        self.status_var.set(f"✓ {count} archivo(s) en lista.")
        
    def add_pdfs(self):
        files = filedialog.askopenfilenames(title="Seleccionar PDFs", filetypes=[("Archivos PDF", "*.pdf")])
        if files:
            for f in files:
                if f not in self.pdf_paths:
                    self.pdf_paths.append(f)
                    self.pdf_listbox.insert(tk.END, os.path.basename(f))
            self.update_status_count()
            
    def remove_pdfs(self):
        sel = self.pdf_listbox.curselection()
        if not sel: return
        for i in reversed(sel):
            self.pdf_paths.pop(i)
            self.pdf_listbox.delete(i)
        self.update_status_count()
        
    def browse_indice(self):
        f = filedialog.askopenfilename(title="Seleccionar Índice", filetypes=[("Excel", "*.xlsx;*.xls")])
        if f: self.var_indice.set(f)

    def run_process(self):
        concept = self.concept_var.get().strip()

        if not self.pdf_paths or not concept:
            messagebox.showerror("Error", "Seleccione al menos un PDF e ingrese el concepto a buscar.")
            return

        try:
            self.status_var.set("⏳ Buscando en PDF(s)...")
            self.root.update()

            def callback(msg):
                self.status_var.set(msg)
                self.root.update()

            if self.var_modo.get() == "vacaciones":
                indice = self.var_indice.get()
                tope = self.var_fecha_tope.get()
                
                self.last_results_raw = buscador_conceptos.search_in_pdfs_pivot(
                    self.pdf_paths, concept, 
                    indice_path=indice, 
                    fecha_tope=tope,
                    extract_units=True, 
                    callback=callback,
                    filtro_firma=self.var_filtro_firma.get()
                )
                
                if not self.last_results_raw:
                    messagebox.showinfo("Sin resultados", "No se encontró el concepto en los archivos.")
                    self.status_var.set("✓ Búsqueda terminada sin resultados.")
                    return
                cols_def = [("Legajo", 70), ("Nombre y Apellido", 180), ("Fecha Ingreso", 90)]
                pdfs_cols = self.last_results_raw[0][1]
                for p in pdfs_cols:
                    cols_def.append((p, 80))
                cols_def.append(("Días Tomados (Total)", 120))
                cols_def.append(("Días que Corresponden", 130))
                cols_def.append(("Saldo (Resto)", 90))
                
                self.build_tree_columns(cols_def)
                for item in self.tree.get_children(): self.tree.delete(item)
                
                for row_dict, _ in self.last_results_raw:
                    vals = [row_dict.get(c[0], "") for c in cols_def]
                    self.tree.insert('', tk.END, values=vals)
                self.status_var.set(f"✓ Proceso terminado: {len(self.last_results_raw)} legajos analizados.")
            else:
                # ---> AQUÍ SE USA LA NUEVA LÓGICA LOCAL EN LUGAR DE BUSCADOR_CONCEPTOS <---
                self.last_results_flat = self.extraer_datos_pdf_regular(self.pdf_paths, concept, callback=callback, filtro_firma=self.var_filtro_firma.get())
                
                if not self.last_results_flat:
                    messagebox.showinfo("Sin resultados", "No se encontró el concepto.")
                    self.status_var.set("✓ Búsqueda terminada sin resultados.")
                    return
                cols_def = [("Archivo", 150), ("Legajo", 80), ("Nombre y Apellido", 200), ("Concepto", 120), ("Importe", 80)]
                self.build_tree_columns(cols_def)
                for item in self.tree.get_children(): self.tree.delete(item)
                
                for r in self.last_results_flat:
                     self.tree.insert('', tk.END, values=(r.get("Archivo",""), r.get("Legajo",""), r.get("Nombre y Apellido",""), r.get("Concepto",""), r.get("Importe","")))
                self.status_var.set(f"✓ Proceso terminado: {len(self.last_results_flat)} coincidencias.")
                
        except Exception as e:
            traceback.print_exc()
            self.status_var.set("✗ Error durante el proceso.")
            messagebox.showerror("Error", str(e))

    def export_excel(self):
        concept = self.concept_var.get().strip()
        modo_vac = (self.var_modo.get() == "vacaciones")
        
        if modo_vac:
            if getattr(self, "last_results_raw", None) is None: return messagebox.showerror("Error", "No hay resultados de vacaciones.")
            flat_data = [d[0] for d in self.last_results_raw]
        else:
            if getattr(self, "last_results_flat", None) is None: return messagebox.showerror("Error", "No hay resultados regulares.")
            flat_data = self.last_results_flat
            
        output_file = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel XLSX", "*.xlsx"), ("Archivo CSV", "*.csv")],
            initialfile=f"Reporte_{concept}.xlsx"
        )
        if not output_file:
            return
            
        try:
            import pandas as pd
            df = pd.DataFrame(flat_data)
            
            if output_file.endswith('.csv'):
                df.to_csv(output_file, index=False, sep=';', encoding='utf-8-sig')
            else:
                df.to_excel(output_file, index=False)
                
            self.status_var.set("✓ Reporte exportado correctamente.")
            messagebox.showinfo("Éxito", f"Reporte finalizado.\n{output_file}")
            os.startfile(output_file)
        except ImportError:
            messagebox.showerror("Falta Librería", "La librería 'pandas' o 'openpyxl' no está instalada.\nNo se puede generar Excel.")
        except Exception as e:
            traceback.print_exc()
            self.status_var.set("✗ Error al exportar.")
            messagebox.showerror("Error", f"Error al crear archivo Excel:\n{str(e)}")


if __name__ == "__main__":
    try:
        root = ctk.CTk()
        app = BuscadorApp(root)
        root.mainloop()
    except Exception as e:
        try:
            import ctypes
            ctypes.windll.user32.MessageBoxW(0, f"Error fatal :\n{str(e)}", "Error", 0x10)
        except:
            pass
        sys.stderr.write(f"Error fatal: {e}\n")