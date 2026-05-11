import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
import sys
import os

# Asegurar que el directorio del script esté en sys.path para imports locales
script_dir = os.path.dirname(os.path.abspath(__file__))
if script_dir not in sys.path:
    sys.path.insert(0, script_dir)

# También agregar 03_OTROS (donde se encuentran los submódulos del sistema)
others_dir = os.path.abspath(os.path.join(script_dir, "..", "03_OTROS"))
if others_dir not in sys.path:
    sys.path.insert(0, others_dir)

# Importar componentes modernos
try:
    import modern_gui_components as mgc
    from icon_loader import set_window_icon, load_icon
    HAS_MGC = True
except ImportError:
    HAS_MGC = False

# --- CONFIGURACIÓN TÉCNICA TOTAL (Cortes, Pesos y Descuentos) ---
PRECIOS_REF = {"Mosquitero_m2": 8500}

DATA = {
    "A30 New (Alta Gama)": {
        "Ventana Corrediza": {
            "formulas": lambda w, h: [
                ("Marco Horiz. (6206)", 2, w, "45°", 1.34), ("Marco Vert. (6206)", 2, h, "45°", 1.34),
                ("Hoja Horiz. (6211)", 4, (w+34)/2, "45°", 1.25), ("Hoja Vert. (6211)", 4, h-75, "45°", 1.25)
            ],
            "vidrio": lambda w, h: ((w+34)/2 - 138, (h-75) - 138),
            "panos": 2, "cota_a": 60, "cota_v": 24, "tag": "corrediza",
            "mosq": lambda w, h: [("Perfil Mosq (6213) H", 2, (w+34)/2, "45°", 0.45), ("Perfil Mosq (6213) V", 2, h-75, "45°", 0.45)],
            "accesorios": lambda w, h: [
                ("Cierre Multipunto Seg." if h > 2000 else "Cierre Lat. Embutido", 2, "un"), 
                ("Ruedas H64 Aguja", 4, "un"), ("Felpa Fin-Seal", round((w*2+h*4)/1000*1.1, 1), "mts"), 
                ("Escuadra Tirón", 8, "un"), ("Kit Antielevación", 2, "un")
            ]
        },
        "Puerta de Abrir": {
            "formulas": lambda w, h: [
                ("Marco (6203) Vertical", 2, h, "45°", 1.10), ("Marco (6203) Horiz.", 1, w, "45°", 1.10),
                ("Hoja H (6214)", 2, w-86, "45°", 1.65), ("Hoja V (6214)", 2, h-55, "45°", 1.65)
            ],
            "vidrio": lambda w, h: (w-220, h-189),
            "panos": 1, "cota_a": 75, "cota_v": 24, "tag": "puerta",
            "accesorios": lambda w, h: [("Cerradura Seg.", 1, "un"), ("Bisagra Ref.", 3, "un"), ("Picaporte", 1, "juego")]
        },
        "Paño Fijo": {
            "formulas": lambda w, h: [
                ("Marco (6201) H", 2, w, "45°", 0.95), ("Marco (6201) V", 2, h, "45°", 0.95),
                ("Contravidrio (6217) H", 2, w-60, "90°", 0.25), ("Contravidrio (6217) V", 2, h-60, "90°", 0.25)
            ],
            "vidrio": lambda w, h: (w-60, h-60),
            "panos": 1, "cota_a": 60, "cota_v": 24, "tag": "fijo",
            "accesorios": lambda w, h: [("Burlete en U", round((w*2+h*2)/1000, 1), "mts")]
        }
    },
    "Modena (Estándar)": {
        "Ventana Corrediza": {
            "formulas": lambda w, h: [
                ("Marco H (6144)", 2, w, "90°", 0.82), ("Marco V (6144)", 2, h, "90°", 0.82),
                ("Hoja H (6146)", 4, (w+22)/2, "45°", 0.75), ("Hoja V (6146)", 4, h-60, "45°", 0.75)
            ],
            "vidrio": lambda w, h: ((w+22)/2 - 100, (h-60) - 100),
            "panos": 2, "cota_a": 45, "cota_v": 18, "tag": "corrediza",
            "mosq": lambda w, h: [("Perfil Mosq (6150) H", 2, (w+22)/2, "45°", 0.35), ("Perfil Mosq (6150) V", 2, h-60, "45°", 0.35)],
            "accesorios": lambda w, h: [("Cierre Modena", 2, "un"), ("Rueda Modena", 4, "un"), ("Felpa 7x7", round((w*2+h*4)/1000*1.1, 1), "mts")]
        },
        "Puerta de Abrir": {
            "formulas": lambda w, h: [
                ("Marco (6142) Vert.", 2, h, "45°", 0.75), ("Marco (6142) Horiz.", 1, w, "45°", 0.75),
                ("Hoja H (6147)", 2, w-80, "45°", 0.92), ("Hoja V (6147)", 2, h-50, "45°", 0.92)
            ],
            "vidrio": lambda w, h: (w-180, h-150),
            "panos": 1, "cota_a": 45, "cota_v": 18, "tag": "puerta",
            "accesorios": lambda w, h: [("Cerradura Modena", 1, "un"), ("Bisagra Modena", 3, "un")]
        },
        "Paño Fijo": {
            "formulas": lambda w, h: [
                ("Marco (6141) H", 2, w, "45°", 0.65), ("Marco (6141) V", 2, h, "45°", 0.65),
                ("Contravidrio (6218) H", 2, w-55, "90°", 0.18), ("Contravidrio (6218) V", 2, h-55, "90°", 0.18)
            ],
            "vidrio": lambda w, h: (w-55, h-55),
            "panos": 1, "cota_a": 45, "cota_v": 18, "tag": "fijo",
            "accesorios": lambda w, h: [("Burlete Cuña", round((w*2+h*2)/1000, 1), "mts")]
        }
    },
    "Herrero (Económica)": {
        "Ventana Corrediza": {
            "formulas": lambda w, h: [
                ("Marco H (7001)", 2, w, "90°", 0.45), ("Marco V (7001)", 2, h, "90°", 0.45),
                ("Hoja H (7003)", 4, (w+12)/2, "45°", 0.48), ("Hoja V (7003)", 4, h-45, "45°", 0.48)
            ],
            "vidrio": lambda w, h: ((w+12)/2 - 70, (h-45) - 70),
            "panos": 2, "cota_a": 35, "cota_v": 10, "tag": "corrediza",
            "accesorios": lambda w, h: [("Aldaba Herrero", 1, "un"), ("Ruedas Herrero", 4, "un")]
        },
        "Paño Fijo": {
            "formulas": lambda w, h: [
                ("Marco (7001) H", 2, w, "90°", 0.45), ("Marco (7001) V", 2, h, "90°", 0.45),
                ("Contravidrio (7008) H", 2, w-40, "90°", 0.15), ("Contravidrio (7008) V", 2, h-40, "90°", 0.15)
            ],
            "vidrio": lambda w, h: (w-40, h-40),
            "panos": 1, "cota_a": 35, "cota_v": 10, "tag": "fijo",
            "accesorios": lambda w, h: [("Burlete en U", round((w*2+h*2)/1000, 1), "mts")]
        }
    }
}

class SuiteCarjorMaster:
    def __init__(self, root):
        self.root = root
        self.root.title("Suite Carjor v12.1 | Absolute Master Edition")
        self.root.geometry("900x700")
        self.root.resizable(True, True)
        self.root.configure(bg="#F0F2F5")
        
        if HAS_MGC:
            mgc.center_window(self.root, 900, 700)
            set_window_icon(self.root, 'calculator')
            self.icon_main = load_icon('calculator', (64, 64))
        
        self.setup_styles()
        
        # Contenedor principal con scroll (Helper de mgc)
        self.scroll_container = mgc.create_main_container(self.root, padding=0)
        
        # Header modernizado
        self.header = mgc.create_header(
            self.scroll_container, 
            "Cálculo de Aluminio", 
            "Optimización de cortes y presupuesto para aberturas", 
            icon_image=self.icon_main
        )

        # Frame de contenido (dentro del scroll)
        content_frame = ctk.CTkFrame(self.scroll_container, fg_color="transparent")
        content_frame.pack(fill='both', expand=True, padx=30, pady=10)

        # SECCIÓN SUPERIOR: CONFIGURACIÓN Y VISUALIZADOR
        row_top = ctk.CTkFrame(content_frame, fg_color="transparent")
        row_top.pack(fill='x', pady=(0, 20))
        
        # Card 1: Datos Técnicos
        card_cfg_outer, card_cfg_inner = mgc.create_card(row_top, "⚙️ PARÁMETROS DE CARPINTERÍA", padding=15)
        card_cfg_outer.pack(side='left', fill='both', expand=True, padx=(0, 10))
        
        def create_label(parent, text):
            return ctk.CTkLabel(parent, text=text, font=mgc.FONTS['small'], text_color=mgc.COLORS['text_secondary'])

        row1 = ctk.CTkFrame(card_cfg_inner, fg_color="transparent")
        row1.pack(fill='x', pady=5)
        
        col1 = ctk.CTkFrame(row1, fg_color="transparent")
        col1.pack(side='left', fill='x', expand=True, padx=(0, 5))
        create_label(col1, "Línea de Perfil:").pack(anchor='w')
        self.cb_linea = ttk.Combobox(col1, values=list(DATA.keys()), state="readonly")
        self.cb_linea.pack(fill="x", pady=5); self.cb_linea.bind("<<ComboboxSelected>>", self.on_linea_change)

        col2 = ctk.CTkFrame(row1, fg_color="transparent")
        col2.pack(side='left', fill='x', expand=True, padx=(5, 0))
        create_label(col2, "Tipología:").pack(anchor='w')
        self.cb_tipo = ttk.Combobox(col2, state="readonly")
        self.cb_tipo.pack(fill="x", pady=5); self.cb_tipo.bind("<<ComboboxSelected>>", lambda e: self.update_preview())

        row2 = ctk.CTkFrame(card_cfg_inner, fg_color="transparent")
        row2.pack(fill='x', pady=10)
        
        col_w = ctk.CTkFrame(row2, fg_color="transparent")
        col_w.pack(side='left', fill='x', expand=True, padx=(0, 5))
        create_label(col_w, "Ancho (mm):").pack(anchor='w')
        self.ent_w = ctk.CTkEntry(col_w, placeholder_text="0", font=mgc.FONTS['title'])
        self.ent_w.pack(fill='x', pady=5)

        col_h = ctk.CTkFrame(row2, fg_color="transparent")
        col_h.pack(side='left', fill='x', expand=True, padx=(5, 0))
        create_label(col_h, "Alto (mm):").pack(anchor='w')
        self.ent_h = ctk.CTkEntry(col_h, placeholder_text="0", font=mgc.FONTS['title'])
        self.ent_h.pack(fill='x', pady=5)

        # Card 2: Visualizador y Precios
        card_vis_outer, card_vis_inner = mgc.create_card(row_top, "🖼️ VISTA PREVIA Y COSTOS", padding=15)
        card_vis_outer.pack(side='left', fill='both', expand=True, padx=(10, 0))
        
        self.canvas = tk.Canvas(card_vis_inner, width=200, height=120, bg="#1a202c", highlightthickness=0)
        self.canvas.pack(pady=5); self.canvas.bind("<Button-1>", self.ver_zoom)

        row_prices = ctk.CTkFrame(card_vis_inner, fg_color="transparent")
        row_prices.pack(fill='x', pady=5)
        
        col_alu = ctk.CTkFrame(row_prices, fg_color="transparent")
        col_alu.pack(side='left', fill='x', expand=True, padx=(0, 5))
        create_label(col_alu, "P. Alum ($/kg):").pack(anchor='w')
        self.ent_p_alu = ctk.CTkEntry(col_alu, height=28); self.ent_p_alu.insert(0, "9500"); self.ent_p_alu.pack(fill='x')

        col_vid = ctk.CTkFrame(row_prices, fg_color="transparent")
        col_vid.pack(side='left', fill='x', expand=True, padx=(5, 0))
        create_label(col_vid, "P. Vidrio ($/m2):").pack(anchor='w')
        self.ent_p_vid = ctk.CTkEntry(col_vid, height=28); self.ent_p_vid.insert(0, "65000"); self.ent_p_vid.pack(fill='x')

        self.var_mosq = ctk.CTkCheckBox(card_vis_inner, text="Incluir Mosquitero", font=mgc.FONTS['small'], command=self.update_preview)
        self.var_mosq.pack(pady=5)

        # SECCIÓN INFERIOR: RESULTADOS
        card_res_outer, card_res_inner = mgc.create_card(content_frame, "📋 DESGLOSE DE MATERIALES Y CORTES", padding=15)
        card_res_outer.pack(fill='both', expand=True)

        self.tree = ttk.Treeview(card_res_inner, columns=("cat", "desc", "can", "med", "cost"), show="headings", height=10)
        for c, h in [("cat","TIPO"), ("desc","DESCRIPCIÓN"), ("can","CANT"), ("med","MEDIDA"), ("cost","SUBTOTAL")]:
            self.tree.heading(c, text=h); self.tree.column(c, width=100, anchor="center")
        self.tree.column("desc", width=350, anchor="w")
        
        sb = ttk.Scrollbar(card_res_inner, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        self.tree.pack(side='left', fill='both', expand=True)
        sb.pack(side='right', fill='y')

        # Botones de Acción Fijos Inferiores
        footer = ctk.CTkFrame(self.root, fg_color=mgc.COLORS['bg_card'], height=100, corner_radius=0)
        footer.pack(fill="x", side="bottom")
        
        btn_frame = ctk.CTkFrame(footer, fg_color="transparent")
        btn_frame.pack(pady=10)
        
        self.btn_calc = mgc.create_large_button(btn_frame, "CALCULAR PROYECTO", self.calcular, color='blue', icon="⚡")
        self.btn_calc.pack(side='left', padx=10)
        
        mgc.create_button(btn_frame, "Exportar Excel", self.exportar, color='green', icon="📄").pack(side='left', padx=10)

        self.lbl_total = ctk.CTkLabel(footer, text="PRESUPUESTO TOTAL: $ 0.00", font=mgc.FONTS['title'], text_color=mgc.COLORS['green'])
        self.lbl_total.pack(pady=(0, 10))

    def on_linea_change(self, e):
        l = self.cb_linea.get()
        self.cb_tipo['values'] = list(DATA[l].keys()); self.cb_tipo.current(0)
        self.update_preview()

    def update_preview(self):
        self.canvas.delete("all")
        l, t = self.cb_linea.get(), self.cb_tipo.get()
        if not l or not t: return
        tag = DATA[l][t]["tag"]
        if tag == "corrediza":
            self.canvas.create_rectangle(60, 15, 160, 115, outline="#3182CE", width=2)
            self.canvas.create_line(110, 15, 110, 115, fill="#CBD5E0")
            if self.var_mosq.get(): self.canvas.create_rectangle(55, 10, 110, 120, outline="#38A169", dash=(4,2))
        elif tag == "puerta":
            self.canvas.create_rectangle(80, 15, 140, 120, outline="#3182CE", width=2)
            self.canvas.create_oval(125, 65, 130, 70, fill="#718096")
        elif tag == "fijo":
            self.canvas.create_rectangle(70, 15, 150, 115, outline="#3182CE", width=3)
        self.canvas.create_text(110, 130, text=f"{l} - {t}".upper(), fill="#718096", font=("Segoe UI", 7, "bold"))

    def ver_zoom(self, e):
        l, t = self.cb_linea.get(), self.cb_tipo.get()
        if not l: return
        top = tk.Toplevel(self.root); top.title("Catálogo Técnico"); top.geometry("450x450"); top.configure(bg="#FFFFFF")
        top.transient(self.root)
        top.grab_set()
        cv = tk.Canvas(top, width=450, height=450, bg="#FFFFFF", highlightthickness=0); cv.pack()
        anc, can = DATA[l][t]["cota_a"], DATA[l][t]["cota_v"]
        
        # DIBUJO TÉCNICO AVANZADO (Simulación de extrusión de catálogo)
        if "Corrediza" in t:
            # Dibujo de Umbral 2 guías
            cv.create_rectangle(80, 250, 370, 290, fill="#E2E8F0", outline="#2D3748")
            cv.create_rectangle(130, 150, 160, 250, fill="#F7FAFC", outline="#3182CE") # Riel 1
            cv.create_rectangle(290, 150, 320, 250, fill="#F7FAFC", outline="#3182CE") # Riel 2
            cv.create_text(225, 330, text=f"UMBRAL 2 GUÍAS: {anc}mm", font=("Segoe UI", 10, "bold"))
        else:
            # Perfil tubular (Hoja de puerta o Paño Fijo)
            cv.create_rectangle(125, 125, 325, 275, fill="#E2E8F0", outline="#2D3748", width=2)
            cv.create_rectangle(145, 145, 305, 255, fill="white", outline="#A0AEC0") # Cámara de aire
            cv.create_text(225, 330, text=f"PERFIL TUBULAR: {anc}mm", font=("Segoe UI", 10, "bold"))
            
        cv.create_line(100, 370, 350, 370, fill="#3182CE", arrow=tk.BOTH, width=2)
        cv.create_text(225, 80, text=f"ALOJAMIENTO VIDRIO: {can}mm", fill="#3182CE", font=("Segoe UI", 10, "bold"))

    def calcular(self):
        try:
            l, t = self.cb_linea.get(), self.cb_tipo.get()
            w, h = float(self.ent_w.get()), float(self.ent_h.get())
            p_alu, p_vid = float(self.ent_p_alu.get()), float(self.ent_p_vid.get())
            for i in self.tree.get_children(): self.tree.delete(i)
            tot = 0
            # Aluminio (Corregido: ahora solo procesa números)
            for n, c, v, ang, pm in DATA[l][t]["formulas"](w, h):
                cost = (v/1000) * pm * c * p_alu; tot += cost
                self.tree.insert("", "end", values=("ALUMINIO", n, c, f"{v:.1f}mm", f"$ {cost:,.0f}"))
            # Mosquitero
            if self.var_mosq.get() and "mosq" in DATA[l][t]:
                for n, c, v, ang, pm in DATA[l][t]["mosq"](w, h):
                    cost = (v/1000) * pm * c * p_alu; tot += cost
                    self.tree.insert("", "end", values=("MOSQUITERO", n, c, f"{v:.1f}mm", f"$ {cost:,.0f}"))
                mw, mh = (w+34)/2, h-75
                ctela = (mw * mh / 1_000_000) * PRECIOS_REF["Mosquitero_m2"]; tot += ctela
                self.tree.insert("", "end", values=("VIDRIO", "TELA MOSQ.", "1", f"{mw:.0f}x{mh:.0f}mm", f"$ {ctela:,.0f}"))
            # Vidrio
            vw, vh = DATA[l][t]["vidrio"](w, h); paños = DATA[l][t]["panos"]
            m2 = (vw * vh / 1_000_000) * paños; cvid = m2 * p_vid; tot += cvid
            self.tree.insert("", "end", values=("VIDRIO", "CRISTAL", paños, f"{vw:.0f}x{vh:.0f}mm", f"$ {cvid:,.0f}"))
            # Accesorios
            for n, c, u in DATA[l][t]["accesorios"](w, h):
                self.tree.insert("", "end", values=("ACCESORIOS", n.upper(), c, u, "---"))
            self.lbl_total.config(text=f"PRESUPUESTO TOTAL: $ {tot:,.2f}")
        except Exception as e:
            messagebox.showerror("Error", "Check dimensions and values.")

    def exportar(self):
        desktop = "C:\\Users\\rjiso\\OneDrive\\Escritorio\\"
        p = filedialog.asksaveasfilename(initialdir=desktop, defaultextension=".xlsx")
        if not p: return
        wb = openpyxl.Workbook(); ws = wb.active; ws.append(["TIPO", "ITEM", "CANT", "MEDIDA", "SUB"])
        for r in self.tree.get_children(): ws.append(self.tree.item(r)['values'])
        wb.save(p); messagebox.showinfo("Suite Carjor", "Excel Guardado en Escritorio.")

if __name__ == "__main__":
    root = tk.Tk(); app = SuiteCarjorMaster(root); root.mainloop()