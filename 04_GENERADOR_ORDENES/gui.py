import customtkinter as ctk
from database import Database
from pdf_generator import GeneradorOrdenPDF
import datetime
from tkinter import messagebox
import os

# --- CONFIGURACIÓN DE COLORES PROFESIONALES ---
COLORS = {
    "primary": "#1B3C73",      # Azul Profundo
    "secondary": "#40679E",    # Azul Medio
    "accent": "#2ECC71",       # Verde Esmeralda (Éxito)
    "danger": "#E74C3C",       # Rojo (Borrar)
    "warning": "#F39C12",      # Naranja (Editar)
    "bg_light": "#F5F6FA",     # Fondo muy claro
    "text_dark": "#2C3E50",    # Texto principal
    "white": "#FFFFFF",
    "border": "#DCDDE1"
}

class AppOrdenes(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.db = Database("ordenes.db")
        self.title("CARJOR - Sistema de Gestión de Órdenes")
        
        # Centrar ventana y ajustar tamaño para que no se corte
        w = 1100
        h = 850
        ws = self.winfo_screenwidth()
        hs = self.winfo_screenheight()
        x = (ws/2) - (w/2)
        y = (hs/2) - (h/2) - 30 # Un poco más arriba para compensar la barra de tareas
        self.geometry('%dx%d+%d+%d' % (w, h, x, y))
        
        # Aplicar tema base
        ctk.set_appearance_mode("light")
        
        # Variables de datos
        self.prov_var = ctk.StringVar()
        self.obra_var = ctk.StringVar()
        self.autorizado_var = ctk.StringVar()
        self.pago_var = ctk.StringVar()
        self.fecha_ent_var = ctk.StringVar()
        self.retira_var = ctk.StringVar()
        self.destino_var = ctk.StringVar()
        self.fecha_var = ctk.StringVar(value=datetime.date.today().strftime("%d/%m/%Y"))
        self.orden_num_var = ctk.IntVar(value=self.db.get_ultima_orden_num())

        self.iva_perc_var = ctk.StringVar(value="21")
        self.iibb_perc_var = ctk.StringVar(value="0")
        self.ley23966_perc_var = ctk.StringVar(value="0")
        self.ley27430_perc_var = ctk.StringVar(value="0")
        
        self.porc_iva = float(self.db.get_config("porcentaje_iva", 21))
        
        # Configuración de Grid Principal
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1) # El cuerpo principal se expande
        self.configure(fg_color=COLORS["bg_light"])

        self.setup_ui()
        self.load_proveedores()

    def create_card(self, parent, title=None):
        """Crea un marco con estilo de tarjeta profesional"""
        card = ctk.CTkFrame(parent, fg_color=COLORS["white"], corner_radius=10, border_width=1, border_color=COLORS["border"])
        if title:
            lbl = ctk.CTkLabel(card, text=title.upper(), font=("Helvetica", 13, "bold"), text_color=COLORS["primary"])
            lbl.pack(pady=(10, 5), padx=15, anchor="w")
        return card

    def setup_ui(self):
        # === TOP BAR (HEADER) ===
        top_bar = ctk.CTkFrame(self, height=70, fg_color=COLORS["primary"], corner_radius=0)
        top_bar.grid(row=0, column=0, sticky="ew")
        top_bar.grid_columnconfigure(1, weight=1)

        header_title = ctk.CTkLabel(top_bar, text="GENERACIÓN DE ÓRDENES DE COMPRA", 
                                  font=("Helvetica", 20, "bold"), text_color=COLORS["white"])
        header_title.grid(row=0, column=0, padx=25, pady=20)

        # Info de orden a la derecha
        ord_info_frame = ctk.CTkFrame(top_bar, fg_color="transparent")
        ord_info_frame.grid(row=0, column=1, sticky="e", padx=25)
        
        ctk.CTkLabel(ord_info_frame, text="ORDEN N°:", text_color=COLORS["white"], font=("Helvetica", 12, "bold")).pack(side="left", padx=5)
        self.entry_orden = ctk.CTkEntry(ord_info_frame, textvariable=self.orden_num_var, width=80, 
                                      fg_color=COLORS["white"], text_color=COLORS["primary"], font=("Helvetica", 14, "bold"))
        self.entry_orden.pack(side="left", padx=5)
        
        ctk.CTkButton(ord_info_frame, text="⚙", width=35, height=35, command=self.configurar_numero_inicio,
                     fg_color="transparent", hover_color="#2c52a0", border_width=1, border_color=COLORS["white"]).pack(side="left", padx=5)

        # === CONTENIDO PRINCIPAL (SCROLLABLE) ===
        main_scroll = ctk.CTkScrollableFrame(self, fg_color="transparent")
        main_scroll.grid(row=1, column=0, sticky="nsew", padx=10, pady=10)
        main_scroll.grid_columnconfigure(0, weight=1)

        # --- SECCIÓN 1: DATOS GENERALES ---
        info_card = self.create_card(main_scroll, "Información de la Orden")
        info_card.pack(fill="x", padx=15, pady=10)
        
        inner_info = ctk.CTkFrame(info_card, fg_color="transparent")
        inner_info.pack(fill="x", padx=15, pady=10)
        inner_info.grid_columnconfigure((1, 3, 5), weight=1)

        # Campos en una fila
        fields = [
            ("FECHA", self.fecha_var, 120),
            ("OBRA", self.obra_var, 250),
            ("AUTORIZA", self.autorizado_var, 250)
        ]
        
        for i, (label, var, w) in enumerate(fields):
            ctk.CTkLabel(inner_info, text=label+":", font=("Helvetica", 11, "bold")).grid(row=0, column=i*2, padx=5, pady=5, sticky="e")
            ctk.CTkEntry(inner_info, textvariable=var, width=w).grid(row=0, column=i*2+1, padx=10, pady=5, sticky="w")

        # --- SECCIÓN 2: PROVEEDOR ---
        prov_card = self.create_card(main_scroll, "Proveedor")
        prov_card.pack(fill="x", padx=15, pady=10)

        prov_grid = ctk.CTkFrame(prov_card, fg_color="transparent")
        prov_grid.pack(fill="x", padx=15, pady=10)
        prov_grid.grid_columnconfigure(0, weight=1)

        # Selector de proveedor mejorado
        sel_frame = ctk.CTkFrame(prov_grid, fg_color="transparent")
        sel_frame.grid(row=0, column=0, sticky="ew")

        self.combo_prov = ctk.CTkComboBox(sel_frame, values=[], variable=self.prov_var, width=500, 
                                        height=35, command=self.on_prov_select,
                                        font=("Helvetica", 13), dropdown_font=("Helvetica", 12))
        self.combo_prov.pack(side="left", padx=(0, 10))
        self.combo_prov._entry.bind("<KeyRelease>", self.filter_providers)
        self.combo_prov._entry.bind("<FocusIn>", self.on_prov_focus_in)

        self.btn_edit_prov = ctk.CTkButton(sel_frame, text="EDITAR DATOS", width=120, height=35, 
                                         command=self.abrir_edicion_proveedor, fg_color=COLORS["warning"], 
                                         hover_color="#D68910", text_color="black", font=("Helvetica", 11, "bold"))
        self.btn_edit_prov.pack(side="left", padx=5)

        self.btn_del_prov = ctk.CTkButton(sel_frame, text="BORRAR", width=80, height=35, 
                                        command=self.borrar_proveedor, fg_color=COLORS["danger"], 
                                        hover_color="#C0392B", font=("Helvetica", 11, "bold"))
        self.btn_del_prov.pack(side="left", padx=5)

        self.lbl_prov_info = ctk.CTkLabel(prov_grid, text="Seleccione un proveedor para ver sus detalles...", 
                                        font=("Helvetica", 11, "italic"), text_color="gray")
        self.lbl_prov_info.grid(row=1, column=0, sticky="w", pady=(10, 0))

        # --- SECCIÓN 3: ITEMS ---
        items_card = self.create_card(main_scroll, "Detalle de Ítems")
        items_card.pack(fill="x", padx=15, pady=10)

        # Encabezados de tabla
        table_header = ctk.CTkFrame(items_card, fg_color=COLORS["secondary"], height=30, corner_radius=5)
        table_header.pack(fill="x", padx=10, pady=(10, 5))
        
        ctk.CTkLabel(table_header, text="DESCRIPCIÓN", text_color="white", font=("Helvetica", 11, "bold")).place(relx=0.05, rely=0.5, anchor="w")
        ctk.CTkLabel(table_header, text="CANT.", text_color="white", font=("Helvetica", 11, "bold")).place(relx=0.72, rely=0.5, anchor="center")
        ctk.CTkLabel(table_header, text="PRECIO UNIT.", text_color="white", font=("Helvetica", 11, "bold")).place(relx=0.85, rely=0.5, anchor="center")
        ctk.CTkLabel(table_header, text="TOTAL", text_color="white", font=("Helvetica", 11, "bold")).place(relx=0.95, rely=0.5, anchor="center")

        self.item_rows = []
        rows_container = ctk.CTkFrame(items_card, fg_color="transparent")
        rows_container.pack(fill="x", padx=10, pady=5)

        for i in range(20):
            row_f = ctk.CTkFrame(rows_container, fg_color="transparent", height=32)
            row_f.pack(fill="x", pady=1)

            desc_e = ctk.CTkEntry(row_f, height=28, placeholder_text=f"Ítem {i+1}...")
            desc_e.place(relx=0, rely=0, relwidth=0.65)

            cant_e = ctk.CTkEntry(row_f, height=28, width=70, justify="center")
            cant_e.place(relx=0.66, rely=0)
            cant_e.bind("<KeyRelease>", lambda e: self.update_totals())

            prec_e = ctk.CTkEntry(row_f, height=28, width=100, justify="right")
            prec_e.place(relx=0.75, rely=0)
            prec_e.bind("<KeyRelease>", lambda e: self.update_totals())

            total_ent = ctk.CTkEntry(row_f, height=28, width=100, justify="right")
            total_ent.insert(0, "0,00")
            total_ent.configure(state="readonly")
            total_ent.place(relx=0.87, rely=0)

            self.item_rows.append({"desc": desc_e, "cant": cant_e, "prec": prec_e, "total_ent": total_ent})

        # --- SECCIÓN 4: CONDICIONES Y TOTALES ---
        footer_container = ctk.CTkFrame(main_scroll, fg_color="transparent")
        footer_container.pack(fill="x", padx=15, pady=10)
        footer_container.grid_columnconfigure(0, weight=2)
        footer_container.grid_columnconfigure(1, weight=1)

        # Card Opciones (Izquierda)
        opt_card = self.create_card(footer_container, "Condiciones e Impuestos")
        opt_card.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
        
        opt_inner = ctk.CTkFrame(opt_card, fg_color="transparent")
        opt_inner.pack(fill="both", padx=15, pady=10)

        # Columna 1: Condiciones
        cond_frame = ctk.CTkFrame(opt_inner, fg_color="transparent")
        cond_frame.grid(row=0, column=0, padx=(0, 20), sticky="nw")

        opt_fields = [
            ("FORMA DE PAGO", self.pago_var),
            ("FECHA ENT.", self.fecha_ent_var),
            ("RETIRA", self.retira_var),
            ("DESTINO", self.destino_var)
        ]

        for i, (label, var) in enumerate(opt_fields):
            ctk.CTkLabel(cond_frame, text=label+":", font=("Helvetica", 10, "bold")).grid(row=i, column=0, padx=5, pady=2, sticky="e")
            ctk.CTkEntry(cond_frame, textvariable=var, width=200, height=24).grid(row=i, column=1, padx=5, pady=2, sticky="w")

        # Columna 2: Porcentajes
        perc_frame = ctk.CTkFrame(opt_inner, fg_color="transparent")
        perc_frame.grid(row=0, column=1, sticky="nw")

        ctk.CTkLabel(perc_frame, text="I.V.A. (%):", font=("Helvetica", 10, "bold")).grid(row=0, column=0, padx=5, pady=2, sticky="e")
        self.combo_iva = ctk.CTkComboBox(perc_frame, values=["21", "10.5", "0"], variable=self.iva_perc_var, width=80, height=24, command=lambda x: self.update_totals())
        self.combo_iva.grid(row=0, column=1, padx=5, pady=2)

        ctk.CTkLabel(perc_frame, text="IIBB (%):", font=("Helvetica", 10, "bold")).grid(row=1, column=0, padx=5, pady=2, sticky="e")
        ctk.CTkEntry(perc_frame, textvariable=self.iibb_perc_var, width=80, height=24).grid(row=1, column=1, padx=5, pady=2)
        self.iibb_perc_var.trace_add("write", lambda *args: self.update_totals())

        ctk.CTkLabel(perc_frame, text="L.23966 (%):", font=("Helvetica", 10, "bold")).grid(row=2, column=0, padx=5, pady=2, sticky="e")
        ctk.CTkEntry(perc_frame, textvariable=self.ley23966_perc_var, width=80, height=24).grid(row=2, column=1, padx=5, pady=2)
        self.ley23966_perc_var.trace_add("write", lambda *args: self.update_totals())

        ctk.CTkLabel(perc_frame, text="L.27430 (%):", font=("Helvetica", 10, "bold")).grid(row=3, column=0, padx=5, pady=2, sticky="e")
        ctk.CTkEntry(perc_frame, textvariable=self.ley27430_perc_var, width=80, height=24).grid(row=3, column=1, padx=5, pady=2)
        self.ley27430_perc_var.trace_add("write", lambda *args: self.update_totals())

        # Card Totales (Derecha)
        tot_card = self.create_card(footer_container, "Resumen de Totales")
        tot_card.grid(row=0, column=1, sticky="nsew")
        
        tot_inner = ctk.CTkFrame(tot_card, fg_color="transparent")
        tot_inner.pack(fill="both", expand=True, padx=15, pady=10)

        def add_summary_row(label_text):
            f = ctk.CTkFrame(tot_inner, fg_color="transparent")
            f.pack(fill="x", pady=1)
            ctk.CTkLabel(f, text=label_text, font=("Helvetica", 11)).pack(side="left")
            ent = ctk.CTkEntry(f, width=130, height=24, justify="right")
            ent.insert(0, "$ 0,00")
            ent.configure(state="readonly")
            ent.pack(side="right")
            return ent

        self.ent_subtotal = add_summary_row("Subtotal:")
        self.ent_iibb = add_summary_row("IIBB:")
        self.ent_l23 = add_summary_row("L.23966:")
        self.ent_l27 = add_summary_row("L.27430:")
        self.ent_iva = add_summary_row("IVA:")

        ctk.CTkFrame(tot_inner, height=2, fg_color=COLORS["border"]).pack(fill="x", pady=5)

        f_total = ctk.CTkFrame(tot_inner, fg_color="transparent")
        f_total.pack(fill="x", pady=5)
        ctk.CTkLabel(f_total, text="TOTAL:", font=("Helvetica", 14, "bold"), text_color=COLORS["primary"]).pack(side="left")
        self.ent_total = ctk.CTkEntry(f_total, width=180, height=35, justify="right", font=("Helvetica", 18, "bold"), text_color=COLORS["primary"])
        self.ent_total.insert(0, "$ 0,00")
        self.ent_total.configure(state="readonly")
        self.ent_total.pack(side="right")

        # === BOTTOM BAR (BUTTONS) ===
        actions_bar = ctk.CTkFrame(self, height=80, fg_color=COLORS["white"], border_width=1, border_color=COLORS["border"])
        actions_bar.grid(row=2, column=0, sticky="ew")
        
        btn_container = ctk.CTkFrame(actions_bar, fg_color="transparent")
        btn_container.pack(expand=True)

        self.btn_gen = ctk.CTkButton(btn_container, text="GENERAR PDF Y GUARDAR", height=45, width=300,
                                   command=self.generar_orden, font=("Helvetica", 14, "bold"),
                                   fg_color=COLORS["accent"], hover_color="#27AE60")
        self.btn_gen.pack(side="left", padx=20, pady=15)

        self.btn_clear = ctk.CTkButton(btn_container, text="LIMPIAR FORMULARIO", height=45, width=200,
                                     command=self.limpiar, font=("Helvetica", 13),
                                     fg_color="#BDC3C7", hover_color="#95A5A6", text_color="black")
        self.btn_clear.pack(side="left", padx=20, pady=15)

    def on_prov_focus_in(self, event):
        if self.prov_var.get() == "<- SELECCIONE AQUÍ":
            self.prov_var.set("")
        try:
            self.combo_prov._open_dropdown_menu()
        except:
            pass

    def filter_providers(self, event):
        if event.keysym in ("Up", "Down", "Return", "Escape", "Tab"):
            return
        typed = self.combo_prov.get().upper()
        if typed == "<- SELECCIONE AQUÍ": return
        if not typed:
            data = list(self.proveedores_data.keys())
        else:
            starts = [n for n in self.proveedores_data.keys() if n.upper().startswith(typed)]
            contains = [n for n in self.proveedores_data.keys() if typed in n.upper() and not n.upper().startswith(typed)]
            data = starts + contains
        self.combo_prov.configure(values=data)
        if data:
            try: self.combo_prov._open_dropdown_menu()
            except: pass

    def configurar_numero_inicio(self):
        new_val = ctk.CTkInputDialog(text="Ingrese el número de la PRÓXIMA orden:", title="Configurar Inicio").get_input()
        if new_val:
            try:
                num = int(new_val)
                self.db.set_config("proxima_orden", num)
                self.orden_num_var.set(num)
                messagebox.showinfo("Éxito", f"El sistema ahora continuará desde la orden: {num}")
            except ValueError:
                messagebox.showerror("Error", "Por favor ingrese un número válido.")

    def borrar_proveedor(self):
        nombre = self.prov_var.get().strip().upper()
        if not nombre or nombre == "<- SELECCIONE AQUÍ":
            messagebox.showwarning("Atención", "Seleccione un proveedor para borrar.")
            return
        if nombre not in self.proveedores_data:
            messagebox.showwarning("Atención", "El proveedor no existe en la base de datos.")
            return
        if messagebox.askyesno("Confirmar", f"¿Está seguro de que desea borrar permanentemente al proveedor:\n{nombre}?"):
            self.db.delete_proveedor(nombre)
            self.load_proveedores()
            self.limpiar_proveedor()
            messagebox.showinfo("Éxito", "Proveedor eliminado correctamente.")

    def limpiar_proveedor(self):
        self.prov_var.set("")
        self.lbl_prov_info.configure(text="Seleccione un proveedor para ver sus detalles...")

    def abrir_edicion_proveedor(self):
        nombre = self.prov_var.get().strip().upper()
        if not nombre:
            messagebox.showwarning("Atención", "Seleccione o escriba un nombre de proveedor primero.")
            return
        p = self.proveedores_data.get(nombre, {"nombre": nombre, "domicilio": "", "domicilio2": "", "categoria_iva": "INSCRIPTO", "cuit": ""})
        dialog = ctk.CTkToplevel(self)
        dialog.title(f"Datos de: {nombre}")
        dialog.geometry("450x500")
        dialog.grab_set()
        dialog.configure(fg_color=COLORS["bg_light"])
        ctk.CTkLabel(dialog, text=f"DETALLES DEL PROVEEDOR", font=("Helvetica", 16, "bold"), text_color=COLORS["primary"]).pack(pady=20)
        fields_entries = {}
        container = ctk.CTkFrame(dialog, fg_color="white", corner_radius=10, border_width=1, border_color=COLORS["border"])
        container.pack(padx=20, pady=10, fill="both", expand=True)
        def create_field(label, key, value):
            f = ctk.CTkFrame(container, fg_color="transparent")
            f.pack(fill="x", padx=15, pady=10)
            ctk.CTkLabel(f, text=label, font=("Helvetica", 11, "bold")).pack(anchor="w")
            entry = ctk.CTkEntry(f, width=350)
            entry.insert(0, value or "")
            entry.pack(pady=(5, 0), fill="x")
            fields_entries[key] = entry
        create_field("DOMICILIO 1:", "domicilio", p.get('domicilio'))
        create_field("DOMICILIO 2 (Localidad/Prov):", "domicilio2", p.get('domicilio2'))
        create_field("CATEGORÍA IVA:", "categoria_iva", p.get('categoria_iva'))
        create_field("CUIT:", "cuit", p.get('cuit'))
        def guardar():
            p_new = {
                "nombre": nombre,
                "domicilio": fields_entries['domicilio'].get().upper(),
                "domicilio2": fields_entries['domicilio2'].get().upper(),
                "categoria_iva": fields_entries['categoria_iva'].get().upper(),
                "cuit": fields_entries['cuit'].get()
            }
            self.db.upsert_proveedor(p_new['nombre'], p_new['domicilio'], p_new['domicilio2'], p_new['categoria_iva'], p_new['cuit'])
            self.load_proveedores()
            self.on_prov_select(nombre)
            dialog.destroy()
            messagebox.showinfo("Éxito", "Datos del proveedor guardados.")
        ctk.CTkButton(dialog, text="GUARDAR CAMBIOS", command=guardar, fg_color=COLORS["accent"], hover_color="#27AE60", height=40, font=("Helvetica", 12, "bold")).pack(pady=25)

    def load_proveedores(self):
        provs = self.db.get_proveedores()
        self.proveedores_data = {p['nombre']: dict(p) for p in provs if p['nombre'] and p['nombre'].strip()}
        self.combo_prov.configure(values=list(self.proveedores_data.keys()))

    def on_prov_select(self, choice):
        p = self.proveedores_data.get(choice)
        if p:
            info = f"CUIT: {p['cuit']}  |  DOMICILIO: {p['domicilio']} {p['domicilio2']}  |  IVA: {p['categoria_iva']}"
            self.lbl_prov_info.configure(text=info, text_color=COLORS["primary"], font=("Helvetica", 11, "bold"))

    def format_n(self, value):
        try:
            val = float(value)
            formatted = f"{val:,.2f}"
            return formatted.replace(",", "X").replace(".", ",").replace("X", ".")
        except (ValueError, TypeError):
            return "0,00"

    def update_totals(self):
        subtotal = 0.0
        for row in self.item_rows:
            try:
                c_val = row['cant'].get().replace(",", ".")
                p_val = row['prec'].get().replace(",", ".")
                cant = float(c_val or 0)
                prec = float(p_val or 0)
                line_total = cant * prec
                subtotal += line_total
                row['total_ent'].configure(state="normal")
                row['total_ent'].delete(0, "end")
                row['total_ent'].insert(0, self.format_n(line_total))
                row['total_ent'].configure(state="readonly")
            except ValueError:
                row['total_ent'].configure(state="normal")
                row['total_ent'].delete(0, "end")
                row['total_ent'].insert(0, "0,00")
                row['total_ent'].configure(state="readonly")

        def get_p(var):
            try: return float(var.get().replace(",", "."))
            except: return 0.0
        p_iva = get_p(self.iva_perc_var)
        p_iibb = get_p(self.iibb_perc_var)
        p_l23 = get_p(self.ley23966_perc_var)
        p_l27 = get_p(self.ley27430_perc_var)
        v_iibb = subtotal * (p_iibb / 100.0)
        v_l23 = subtotal * (p_l23 / 100.0)
        v_l27 = subtotal * (p_l27 / 100.0)
        v_iva = subtotal * (p_iva / 100.0)
        total = subtotal + v_iibb + v_l23 + v_l27 + v_iva

        def set_sum_val(ent, val):
            ent.configure(state="normal")
            ent.delete(0, "end")
            ent.insert(0, f"$ {self.format_n(val)}")
            ent.configure(state="readonly")

        set_sum_val(self.ent_subtotal, subtotal)
        set_sum_val(self.ent_iibb, v_iibb)
        set_sum_val(self.ent_l23, v_l23)
        set_sum_val(self.ent_l27, v_l27)
        set_sum_val(self.ent_iva, v_iva)
        set_sum_val(self.ent_total, total)

        self.current_calc = {
            "subtotal": subtotal, "iibb": v_iibb, "p_iibb": p_iibb,
            "ley23966": v_l23, "p_l23": p_l23, "ley27430": v_l27, "p_l27": p_l27,
            "iva": v_iva, "p_iva": p_iva, "total": total
        }

    def generar_orden(self):
        prov_nombre = self.prov_var.get().strip().upper()
        if not prov_nombre:
            messagebox.showerror("Error", "Debe seleccionar o escribir un proveedor.")
            return
        self.update_totals()
        items_to_save = []
        for row in self.item_rows:
            desc = row['desc'].get().strip()
            cant_str = row['cant'].get().strip().replace(",", ".")
            prec_str = row['prec'].get().strip().replace(",", ".")
            if cant_str:
                try:
                    cant = float(cant_str)
                    prec = float(prec_str or 0)
                    if cant > 0:
                        items_to_save.append({"descripcion": desc.upper(), "cantidad": cant, "precio_unitario": prec, "total_item": cant * prec})
                except ValueError: continue
        if not items_to_save:
            messagebox.showwarning("Atención", "No hay ítems válidos en la orden.")
            return
        p_info = self.proveedores_data.get(prov_nombre)
        if not p_info:
            p_info = {"nombre": prov_nombre, "domicilio": "", "domicilio2": "", "categoria_iva": "INSCRIPTO", "cuit": ""}
            self.db.save_proveedor(p_info)
            self.load_proveedores()
            p_info = self.proveedores_data.get(prov_nombre)
        full_dir = f"{p_info.get('domicilio', '')} {p_info.get('domicilio2', '')}".strip()
        orden_data = {
            "numero_orden": self.orden_num_var.get(), "fecha": self.fecha_var.get(), "proveedor_nombre": prov_nombre,
            "obra": self.obra_var.get().upper(), "autorizado": self.autorizado_var.get().upper(), "forma_pago": self.pago_var.get().upper(),
            "fecha_entrega": self.fecha_ent_var.get().upper(), "retira": self.retira_var.get().upper(), "destino": self.destino_var.get().upper(),
            "subtotal": self.current_calc['subtotal'], "iibb": self.current_calc['iibb'], "p_iibb": self.current_calc['p_iibb'],
            "ley23966": self.current_calc['ley23966'], "p_l23": self.current_calc['p_l23'], "ley27430": self.current_calc['ley27430'],
            "p_l27": self.current_calc['p_l27'], "iva": self.current_calc['iva'], "p_iva": self.current_calc['p_iva'],
            "total": self.current_calc['total'], "domicilio": full_dir, "cuit": p_info.get('cuit', ''), "categoria_iva": p_info.get('categoria_iva', '')
        }
        # Determinar ruta de salida (Escritorio o Carpeta de la App)
        out_dir = ""
        # Intento 1: Escritorio estándar o OneDrive
        possibles_desktops = [
            os.path.join(os.environ.get("USERPROFILE", ""), "Desktop"),
            os.path.join(os.environ.get("USERPROFILE", ""), "OneDrive", "Escritorio"),
            os.path.join(os.environ.get("USERPROFILE", ""), "OneDrive", "Desktop")
        ]
        
        for d in possibles_desktops:
            if os.path.exists(d):
                out_dir = os.path.join(d, "ORDENES")
                break
        
        # Intento 2: Si no se encontró escritorio o no se puede usar, usar carpeta de la App
        if not out_dir:
            out_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "ORDENES")

        if not os.path.exists(out_dir):
            try:
                os.makedirs(out_dir)
            except:
                # Fallback final: Directorio actual
                out_dir = "ORDENES"
                if not os.path.exists(out_dir): os.makedirs(out_dir)

        clean_prov = "".join([c for c in prov_nombre if c.isalnum() or c in (' ', '-', '_')]).strip()
        filename = f"{clean_prov}_{orden_data['numero_orden']}_{orden_data['obra'].replace(' ', '_')}.pdf"
        filepath = os.path.join(out_dir, filename)
        try:
            self.db.save_orden(orden_data, items_to_save)
            gen = GeneradorOrdenPDF(filepath)
            gen.generar(orden_data, items_to_save)
            messagebox.showinfo("Éxito", f"Orden generada:\n{filepath}")
            if os.path.exists(filepath): os.startfile(filepath)
            self.limpiar()
            self.orden_num_var.set(self.db.get_ultima_orden_num())
            self.load_proveedores()
        except Exception as e:
            msg = str(e)
            if "UNIQUE constraint failed" in msg: msg = f"El número de orden {orden_data['numero_orden']} ya existe."
            messagebox.showerror("Error", f"Fallo: {msg}")

    def limpiar(self):
        self.prov_var.set("")
        self.obra_var.set("")
        self.autorizado_var.set("")
        self.pago_var.set("")
        self.fecha_ent_var.set("")
        self.retira_var.set("")
        self.destino_var.set("")
        self.lbl_prov_info.configure(text="Seleccione un proveedor para ver sus detalles...", text_color="gray", font=("Helvetica", 11, "italic"))
        for row in self.item_rows:
            row['desc'].delete(0, 'end')
            row['cant'].delete(0, 'end')
            row['prec'].delete(0, 'end')
            row['total_ent'].configure(state="normal")
            row['total_ent'].delete(0, "end")
            row['total_ent'].insert(0, "0,00")
            row['total_ent'].configure(state="readonly")
        self.update_totals()

if __name__ == "__main__":
    app = AppOrdenes()
    app.mainloop()
