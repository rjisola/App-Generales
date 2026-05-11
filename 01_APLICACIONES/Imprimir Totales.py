import tkinter as tk
from tkinter import ttk

class FormularioRecibo(tk.Frame):
    def __init__(self, parent):
        # Fondo negro para que actúe como las líneas de la grilla (bordes)
        super().__init__(parent, bg="black")
        self.parent = parent
        
        # Colores de fondo (Fieles a las capturas)
        self.color_bg = {
            "AMARILLO": "#D4C100",
            "CELESTE": "#5B9BD5",
            "AZUL": "#4472C4",
            "GRIS": "#999999",
            "BLANCO": "#FFFFFF"
        }
        
        # Variables de control
        self.datos_vars = {
            "color": tk.StringVar(value="AMARILLO"),
            "legajo": tk.StringVar(),
            "nombre": tk.StringVar(),
            "periodo": tk.StringVar(),
            "categoria": tk.StringVar(),
            "hs_50_cant": tk.StringVar(value="0,0"),
            "hs_50_monto": tk.StringVar(value="$0,00"),
            "hs_100_cant": tk.StringVar(value="0,0"),
            "hs_100_monto": tk.StringVar(value="$0,00"),
            "reintegro": tk.StringVar(value="$0,00"),
            "total_extras": tk.StringVar(value="$0,00"),
            "presentismo": tk.StringVar(value="SI"),
            "ajuste_alquiler": tk.StringVar(value="$0,00"),
            "sueldo_sobre": tk.StringVar(value="$0,00"),
            "total_quincena": tk.StringVar(value="$0,00"),
            "adelanto": tk.StringVar(value="$0,00"),
            "efectivo": tk.StringVar(value="$0,00"),
            "sueldo_acordado": tk.StringVar(value="$0,00"),
            "banco": tk.StringVar(value="$0,00"),
            "caja_ahorro_2": tk.StringVar(value="$0,00"),
        }

        self.construir_interfaz()

    def crear_celda(self, texto_fijo=None, variable=None, row=0, col=0, colspan=1, bg="white", font=("Arial", 11), anchor="w"):
        """
        Crea un Label. El truco del borde es el padx/pady en el grid sobre el fondo negro del Frame padre.
        """
        if variable:
            lbl = tk.Label(self, textvariable=variable, bg=bg, font=font, anchor=anchor, padx=8, pady=4)
        else:
            lbl = tk.Label(self, text=texto_fijo, bg=bg, font=font, anchor=anchor, padx=8, pady=4)
            
        # El padding de 1px en el grid sobre el fondo negro crea la línea de la cuadrícula
        lbl.grid(row=row, column=col, columnspan=colspan, sticky="nsew", padx=(1, 0), pady=(0, 1))
        return lbl

    def construir_interfaz(self):
        for widget in self.winfo_children():
            widget.destroy()
            
        color_tipo = self.datos_vars["color"].get().upper()
        bg_actual = self.color_bg.get(color_tipo, "#FFFFFF")
        
        # Pesos de columnas: A=1, B=0.5, C=0.5 (Aproximado a Excel)
        self.columnconfigure(0, weight=2, minsize=180)
        self.columnconfigure(1, weight=1, minsize=110)
        self.columnconfigure(2, weight=1, minsize=110)

        # --- FILA 0: Legajo y Nombre ---
        leg_val = self.datos_vars['legajo'].get()
        self.crear_celda(f"Leg N° {leg_val}", row=0, col=0, bg=bg_actual, font=("Arial", 11, "bold"))
        self.crear_celda(variable=self.datos_vars["nombre"], row=0, col=1, colspan=2, bg=bg_actual, font=("Arial", 11, "bold"))
        
        # --- FILA 1: QUINCENA ---
        self.crear_celda("QUINCENA", row=1, col=0, bg=bg_actual, font=("Arial", 11, "bold"))
        self.crear_celda(variable=self.datos_vars["periodo"], row=1, col=1, colspan=2, bg=bg_actual, font=("Arial", 11, "bold"))
        
        # --- FILA 2: Categoría ---
        self.crear_celda("Categoría", row=2, col=0, bg=bg_actual)
        self.crear_celda(variable=self.datos_vars["categoria"], row=2, col=1, colspan=2, bg=bg_actual, font=("Arial", 11, "bold"))
        
        # --- FILA 3: Cabecera Horas ---
        self.crear_celda("", row=3, col=0, bg=bg_actual)
        self.crear_celda("HORAS", row=3, col=1, bg=bg_actual, font=("Arial", 11, "bold"), anchor="center")
        self.crear_celda("($)", row=3, col=2, bg=bg_actual, font=("Arial", 11, "bold"), anchor="center")

        if color_tipo == "AMARILLO":
            # Filas con divisoria central (A, B, C)
            self.crear_celda("HS.50%", row=4, col=0, bg=bg_actual, font=("Arial", 11, "bold"))
            self.crear_celda(variable=self.datos_vars["hs_50_cant"], row=4, col=1, bg=bg_actual, anchor="center")
            self.crear_celda(variable=self.datos_vars["hs_50_monto"], row=4, col=2, bg=bg_actual, font=("Arial", 11, "bold"))
            
            self.crear_celda("HS.100%", row=5, col=0, bg=bg_actual, font=("Arial", 11, "bold"))
            self.crear_celda(variable=self.datos_vars["hs_100_cant"], row=5, col=1, bg=bg_actual, anchor="center")
            self.crear_celda(variable=self.datos_vars["hs_100_monto"], row=5, col=2, bg=bg_actual, font=("Arial", 11, "bold"))
            
            # Filas COMBINADAS (A, B+C) - Aquí se elimina la línea vertical entre B y C
            conceptos = [
                ("REINTEGRO", "reintegro"), ("TOTAL EXTRAS", "total_extras"),
                ("PRESENTISMO", "presentismo"), ("AJUSTE-ALQUILER", "ajuste_alquiler"),
                ("SUELDO SOBRE", "sueldo_sobre")
            ]
            for i, (lab, var) in enumerate(conceptos, 6):
                self.crear_celda(lab, row=i, col=0, bg=bg_actual, font=("Arial", 11, "bold"))
                self.crear_celda(variable=self.datos_vars[var], row=i, col=1, colspan=2, bg=bg_actual, font=("Arial", 11, "bold"), anchor="w")
            
            # TOTAL QUINCENA - Centrado y grande
            self.crear_celda("TOTAL QUINCENA", row=11, col=0, bg=bg_actual, font=("Arial", 12, "bold"), anchor="center")
            self.crear_celda(variable=self.datos_vars["total_quincena"], row=11, col=1, colspan=2, bg=bg_actual, font=("Arial", 22, "bold"), anchor="center")
            
            # ADELANTO
            self.crear_celda("ADELANTO", row=12, col=0, bg=bg_actual, font=("Arial", 11, "bold"))
            self.crear_celda(variable=self.datos_vars["adelanto"], row=12, col=1, colspan=2, bg=bg_actual, anchor="w")
            
            # Espacios vacíos combinados
            self.crear_celda("", row=13, col=0, bg=bg_actual); self.crear_celda("", row=13, col=1, colspan=2, bg=bg_actual)
            self.crear_celda("", row=14, col=0, bg=bg_actual); self.crear_celda("", row=14, col=1, colspan=2, bg=bg_actual)
            
            # EFECTIVO
            self.crear_celda("EFECTIVO", row=15, col=0, bg=bg_actual, font=("Arial", 11, "bold"))
            self.crear_celda(variable=self.datos_vars["efectivo"], row=15, col=1, colspan=2, bg=bg_actual, font=("Arial", 11, "bold"), anchor="w")

        elif color_tipo == "CELESTE" or color_tipo == "GRIS" or color_tipo == "AZUL":
            # Lógica similar para otros colores respetando sus campos específicos
            # ... (implementación simplificada para brevedad, siguiendo el mismo patrón de merged cells)
            pass

    def cargar_datos(self, diccionario_datos):
        for clave, valor in diccionario_datos.items():
            if clave in self.datos_vars:
                self.datos_vars[clave].set(valor)
        self.construir_interfaz()

if __name__ == "__main__":
    root = tk.Tk()
    root.title("Visor de Recibos - Replicación Excel")
    root.geometry("500x750")
    root.configure(bg="#f0f0f0", padx=20, pady=20)
    
    formulario = FormularioRecibo(root)
    formulario.pack(fill="both", expand=True)

    # Datos reales de la imagen 104738
    datos_originales = {
        "color": "AMARILLO",
        "legajo": "9008",
        "nombre": "Acland Frantl Hector",
        "periodo": "1ERA ABRIL 2026",
        "categoria": "ESPECIALIZADO",
        "hs_50_cant": "0,0",
        "hs_50_monto": "$0,00",
        "hs_100_cant": "0,0",
        "hs_100_monto": "$0,00",
        "reintegro": "$196.900,00",
        "total_extras": "$0,00",
        "presentismo": "SI",
        "ajuste_alquiler": "$17.413,00",
        "sueldo_sobre": "$462.595,00",
        "total_quincena": "$676.908,00",
        "adelanto": "$0,00",
        "efectivo": "$676.908,00"
    }

    formulario.cargar_datos(datos_originales)
    root.mainloop()
