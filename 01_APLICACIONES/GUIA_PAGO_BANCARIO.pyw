# -*- coding: utf-8 -*-
"""
GUIA_PAGO_BANCARIO.pyw
Procesa un archivo XLSM (hoja "RECUENTO TOTAL (2)"), filtra por banco en columna O,
y copia columnas A-N al archivo de Guía de Pago correspondiente.
"""

import sys
import os
import tkinter as tk
import customtkinter as ctk
from tkinter import ttk, messagebox, filedialog
import threading

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

def _safe_save(wb, path):
    try:
        wb.save(path)
    except PermissionError:
        import tempfile, shutil, uuid
        temp_dir = tempfile.gettempdir()
        temp_path = os.path.join(temp_dir, f"temp_{uuid.uuid4().hex}.xlsx")
        wb.save(temp_path)
        shutil.copy2(temp_path, path)
        os.remove(temp_path)

# =============================================================================
# CONSTANTES
# =============================================================================
HOJA_ORIGEN   = "RECUENTO TOTAL (2)"
COLUMNA_BANCO = "O"   # Columna donde está el nombre del banco (índice 15, base-1)
COL_BANCO_IDX = 15    # Número de columna O (1-based)
COL_INICIO    = 1     # Columna A
COL_FIN       = 14    # Columna N

ARCHIVOS_DESTINO = {
    "SANTANDER":     r"C:\Users\rjiso\OneDrive\Escritorio\Guia Pago Simple.xlsx",
    "OTROS BANCOS":  r"C:\Users\rjiso\OneDrive\Escritorio\Guia Pago SimpleOTROS.xlsx",
}

CELDA_CONTADOR = "F3"
FILA_INICIO_DATOS  = 8    # A partir de A8 se borran y se pegan datos
FILA_FIN_BORRADO   = 100  # Hasta N100

COLORES = mgc.COLORS if HAS_MGC else {
    'bg_primary': '#0a0e1a', 'bg_card': '#111827',
    'purple': '#8b5cf6', 'blue': '#3b82f6', 'green': '#10b981',
    'orange': '#f59e0b', 'red': '#ef4444', 'gray': '#6b7280',
    'text_primary': '#f1f5f9', 'text_secondary': '#94a3b8',
    'border': '#1e2a3a',
}

FUENTES = mgc.FONTS if HAS_MGC else {
    'title': ('Segoe UI', 20, 'bold'),
    'subtitle': ('Segoe UI', 13, 'bold'),
    'heading': ('Segoe UI', 11, 'bold'),
    'normal': ('Segoe UI', 10),
    'button': ('Segoe UI', 11, 'bold'),
    'button_large': ('Segoe UI', 12, 'bold'),
    'small': ('Segoe UI', 9),
}

# =============================================================================
# LÓGICA DE PROCESAMIENTO
# =============================================================================

def procesar_guia_pago(archivo_origen: str, banco_filtro: str,
                       log_callback=None, done_callback=None):
    """
    Procesa el archivo XLSM de origen y copia las filas correspondientes al
    banco indicado en el archivo de Guía de Pago destino.

    :param archivo_origen:  Ruta completa al .xlsm
    :param banco_filtro:    'SANTANDER' o 'OTROS BANCOS'
    :param log_callback:    Función(str) para mostrar mensajes en la GUI
    :param done_callback:   Función(bool, str) llamada al finalizar
    """
    def log(msg):
        if log_callback:
            log_callback(msg)

    def finalizar(ok, msg):
        if done_callback:
            done_callback(ok, msg)

    archivo_destino = ARCHIVOS_DESTINO.get(banco_filtro)
    if not archivo_destino:
        finalizar(False, f"No se encontró archivo destino para '{banco_filtro}'.")
        return

    # --- Verificación de archivos ---
    if not os.path.isfile(archivo_origen):
        finalizar(False, f"No se encontró el archivo origen:\n{archivo_origen}")
        return

    if not os.path.isfile(archivo_destino):
        finalizar(False, f"No se encontró el archivo destino:\n{archivo_destino}")
        return

    # --- Abrir origen (keep_vba=True para no corromper el XLSM) ---
    log(f"📂 Abriendo archivo origen de forma segura...")
    try:
        wb_origen = safe_openpyxl_load(archivo_origen, data_only=True, keep_vba=True)
    except Exception as e:
        finalizar(False, f"Error al abrir el archivo origen:\n{e}")
        return

    if HOJA_ORIGEN not in wb_origen.sheetnames:
        wb_origen.close()
        finalizar(False, f"No se encontró la hoja '{HOJA_ORIGEN}' en el archivo origen.\n"
                         f"Hojas disponibles: {wb_origen.sheetnames}")
        return

    ws_origen = wb_origen[HOJA_ORIGEN]
    log(f"✅ Hoja '{HOJA_ORIGEN}' encontrada.")

    # --- Recolectar filas que coinciden con el banco ---
    filas_a_copiar = []
    filtro_upper = banco_filtro.upper()

    for row in ws_origen.iter_rows(min_row=2, values_only=True):
        # row es 0-indexed; columna O es índice 14 (0-based)
        if len(row) < COL_BANCO_IDX:
            continue
        valor_banco = row[COL_BANCO_IDX - 1]  # índice 14 → columna O
        if valor_banco is None:
            continue

        valor_str = str(valor_banco).strip().upper()

        # Lógica de filtrado: comparación EXACTA con la leyenda de la columna O
        # La columna O contiene exactamente "SANTANDER" u "OTROS BANCOS"
        coincide = (valor_str == filtro_upper)

        if coincide:
            # Copiar columnas A..N (índices 0..13)
            filas_a_copiar.append(row[COL_INICIO - 1: COL_FIN])

    wb_origen.close()
    log(f"🔍 Filas encontradas para '{banco_filtro}': {len(filas_a_copiar)}")

    if not filas_a_copiar:
        finalizar(False, f"No se encontraron registros para '{banco_filtro}' "
                         f"en la columna {COLUMNA_BANCO}.")
        return

    # --- Abrir destino ---
    log(f"📂 Abriendo archivo destino de forma segura...")
    try:
        wb_destino = safe_openpyxl_load(archivo_destino)
    except Exception as e:
        finalizar(False, f"Error al abrir el archivo destino:\n{e}")
        return

    ws_destino = wb_destino.active
    log(f"✅ Hoja activa destino: '{ws_destino.title}'")

    # --- Leer y aumentar contador en F3 ---
    valor_f3 = ws_destino[CELDA_CONTADOR].value
    try:
        nuevo_f3 = int(valor_f3) + 1 if valor_f3 is not None else 1
    except (TypeError, ValueError):
        nuevo_f3 = 1
    log(f"🔢 Contador {CELDA_CONTADOR}: {valor_f3} → {nuevo_f3}")
    ws_destino[CELDA_CONTADOR] = nuevo_f3

    # --- Borrar rango A8:N100 ---
    log(f"🗑️ Borrando rango A{FILA_INICIO_DATOS}:N{FILA_FIN_BORRADO}...")
    for fila in range(FILA_INICIO_DATOS, FILA_FIN_BORRADO + 1):
        for col in range(COL_INICIO, COL_FIN + 1):
            ws_destino.cell(row=fila, column=col).value = None

    # --- Pegar datos ---
    log(f"📋 Pegando {len(filas_a_copiar)} filas desde la fila {FILA_INICIO_DATOS}...")
    for i, fila_data in enumerate(filas_a_copiar):
        fila_destino = FILA_INICIO_DATOS + i
        for j, valor in enumerate(fila_data):
            ws_destino.cell(row=fila_destino, column=COL_INICIO + j).value = valor

    # --- Guardar y cerrar ---
    log(f"💾 Guardando archivo destino...")
    try:
        _safe_save(wb_destino, archivo_destino)
        wb_destino.close()
    except Exception as e:
        finalizar(False, f"Error al guardar el archivo destino:\n{e}")
        return

    log(f"✅ Proceso completado exitosamente.")
    finalizar(True, f"Se procesaron {len(filas_a_copiar)} filas correctamente.\n"
                    f"Contador {CELDA_CONTADOR} actualizado a: {nuevo_f3}\n"
                    f"Archivo destino: {archivo_destino}")


# =============================================================================
# APLICACIÓN GUI
# =============================================================================

class GuiaPagoBancarioApp:
    """Aplicación principal para procesar Guía de Pago Bancario."""

    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("💳 Guía de Pago Bancario")
        self.root.resizable(False, False)
        self.root.configure(bg=COLORES['bg_primary'])

        # Centrar ventana
        self.root.geometry("900x700")
        if HAS_MGC:
            mgc.center_window(self.root, 900, 700)
        else:
            ancho, alto = 900, 700
            sw = root.winfo_screenwidth()
            sh = root.winfo_screenheight()
            x = (sw - ancho) // 2
            y = (sh - alto) // 2
            self.root.geometry(f"{ancho}x{alto}+{x}+{y}")
        
        self.root.resizable(True, True)

        # Icono de ventana
        if HAS_ICON:
            try:
                set_window_icon(self.root, 'payment')
            except Exception:
                pass

        # Variable para el archivo origen
        self.archivo_origen_var = tk.StringVar(value="")

        # Variable para la opción de banco
        self.banco_var = tk.StringVar(value="SANTANDER")

        self._build_ui()

        # Verificar openpyxl disponible
        if not HAS_OPENPYXL:
            messagebox.showerror(
                "Dependencia faltante",
                "No se encontró la librería 'openpyxl'.\n"
                "Instálela con: pip install openpyxl"
            )

    # -------------------------------------------------------------------------
    def _build_ui(self):
        main_container = mgc.create_main_container(self.root)
        main = tk.Frame(main_container, bg=COLORES['bg_primary'])
        main.pack(fill=tk.BOTH, expand=True, padx=20, pady=15)

        # ── HEADER ────────────────────────────────────────────────────────────
        # Header moderno (estilo launcher)
        mgc.create_header(
            main, 
            "Guía Pago Bancario", 
            "Filtración de haberes por banco y actualización automática de guías de pago", 
            icon_image=load_icon('payment', (64, 64)) if HAS_ICON else None
        )

        # ── CARDS ─────────────────────────────────────────────────────────────
        self._card_banco(main)
        self._card_archivo(main)
        self._card_destino(main)

        # ── BOTÓN PROCESAR ────────────────────────────────────────────────────
        btn_frame = tk.Frame(main, bg=COLORES['bg_primary'])
        btn_frame.pack(fill=tk.X, pady=(15, 0))

        self.btn_procesar = tk.Button(
            btn_frame,
            text="▶  PROCESAR GUÍA DE PAGO",
            command=self._iniciar_proceso,
            font=FUENTES['button_large'],
            bg=COLORES['green'],
            fg='white',
            relief=tk.FLAT,
            padx=40,
            pady=14,
            cursor='hand2',
            activebackground='#059669',
            activeforeground='white',
        )
        self.btn_procesar.pack()
        self.btn_procesar.bind("<Enter>", lambda e: self.btn_procesar.configure(bg='#059669'))
        self.btn_procesar.bind("<Leave>", lambda e: self.btn_procesar.configure(bg=COLORES['green']))

        # ── LOG / PROGRESO ────────────────────────────────────────────────────
        self._card_log(main)

        # ── BARRA DE ESTADO ───────────────────────────────────────────────────
        status_bar = tk.Frame(self.root, bg=COLORES['border'], relief=tk.SUNKEN, bd=1)
        status_bar.pack(side=tk.BOTTOM, fill=tk.X)

        self.status_var = tk.StringVar(value="Listo")
        tk.Label(status_bar,
                 textvariable=self.status_var,
                 font=FUENTES['small'],
                 bg=COLORES['bg_card'],
                 fg=COLORES['text_secondary'],
                 anchor='w', padx=10, pady=4
                 ).pack(side=tk.LEFT, fill=tk.X, expand=True)

    # -------------------------------------------------------------------------
    def _card_banco(self, parent):
        """Card de selección de banco."""
        card_outer, card_inner = mgc.create_card(parent, "🏦  Seleccionar Banco", padding=15)
        card_outer.pack(fill=tk.X, pady=(0, 10))

        inner = tk.Frame(card_inner, bg=COLORES['bg_card'])
        inner.pack(fill=tk.X)

        # Botón SANTANDER
        self.btn_snt = tk.Radiobutton(
            inner,
            text="🏦  SANTANDER",
            variable=self.banco_var,
            value="SANTANDER",
            command=self._on_banco_change,
            font=FUENTES['subtitle'],
            bg=COLORES['bg_card'],
            fg='#cc0000',
            activebackground=COLORES['bg_card'],
            activeforeground='#cc0000',
            selectcolor=COLORES['bg_card'],
            indicatoron=0,
            relief=tk.RAISED,
            padx=20, pady=10,
            cursor='hand2',
            width=18,
        )
        self.btn_snt.pack(side=tk.LEFT, padx=(0, 15))

        # Botón OTROS BANCOS
        self.btn_otros = tk.Radiobutton(
            inner,
            text="🏛️  OTROS BANCOS",
            variable=self.banco_var,
            value="OTROS BANCOS",
            command=self._on_banco_change,
            font=FUENTES['subtitle'],
            bg=COLORES['bg_card'],
            fg=COLORES['blue'],
            activebackground=COLORES['bg_card'],
            activeforeground=COLORES['blue'],
            selectcolor=COLORES['bg_card'],
            indicatoron=0,
            relief=tk.RAISED,
            padx=20, pady=10,
            cursor='hand2',
            width=18,
        )
        self.btn_otros.pack(side=tk.LEFT)

        self._on_banco_change()  # Aplicar estilo inicial

    # -------------------------------------------------------------------------
    def _on_banco_change(self):
        """Actualiza el aspecto visual de los botones de banco y el label del destino."""
        banco = self.banco_var.get()

        if banco == "SANTANDER":
            self.btn_snt.configure(relief=tk.SUNKEN, bg='#3b0a0a')
            self.btn_otros.configure(relief=tk.RAISED, bg=COLORES['bg_card'])
        else:
            self.btn_otros.configure(relief=tk.SUNKEN, bg='#0a1a3b')
            self.btn_snt.configure(relief=tk.RAISED, bg=COLORES['bg_card'])

        # Actualizar label destino
        ruta = ARCHIVOS_DESTINO.get(banco, "")
        if hasattr(self, 'lbl_destino_val'):
            self.lbl_destino_val.configure(text=ruta)

    # -------------------------------------------------------------------------
    def _card_archivo(self, parent):
        """Card para seleccionar el archivo origen XLSM."""
        card_outer, card_inner = mgc.create_card(parent, "📂  Archivo Origen (XLSM)", padding=15)
        card_outer.pack(fill=tk.X, pady=(0, 10))

        inner = tk.Frame(card_inner, bg=COLORES['bg_card'])
        inner.pack(fill=tk.X)
        inner.columnconfigure(0, weight=1)
        inner.columnconfigure(1, weight=0)

        self.entry_origen = tk.Entry(
            inner,
            textvariable=self.archivo_origen_var,
            font=FUENTES['small'],
            state='readonly',
            bg='#1a2235',
            fg='#f1f5f9',
            insertbackground='white',
            relief=tk.GROOVE,
            bd=1,
        )
        self.entry_origen.grid(row=0, column=0, sticky='we', padx=(0, 8), pady=2)

        btn_sel = tk.Button(
            inner,
            text="📁  Seleccionar",
            command=self._seleccionar_archivo,
            font=FUENTES['button'],
            bg=COLORES['purple'],
            fg='white',
            relief=tk.FLAT,
            padx=15, pady=6,
            cursor='hand2',
            activebackground='#7c3aed',
            activeforeground='white',
        )
        btn_sel.grid(row=0, column=1, sticky='e')

        tk.Label(
            card_outer,
            text=f'Hoja usada: "{HOJA_ORIGEN}"  ·  Columna de banco: {COLUMNA_BANCO}  ·  Se copian columnas A – N',
            font=FUENTES['small'],
            bg=COLORES['bg_card'],
            fg=COLORES['text_secondary'],
        ).pack(anchor='w', pady=(5, 0))

    # -------------------------------------------------------------------------
    def _card_destino(self, parent):
        """Card informativa con el archivo destino seleccionado."""
        card_outer, card_inner = mgc.create_card(parent, "📄  Archivo Destino (Automático)", padding=15)
        card_outer.pack(fill=tk.X, pady=(0, 0))

        self.lbl_destino_val = tk.Label(
            card_inner,
            text=ARCHIVOS_DESTINO.get(self.banco_var.get(), ""),
            font=FUENTES['small'],
            bg=COLORES['bg_card'],
            fg=COLORES['blue'],
            anchor='w',
            wraplength=560,
            justify=tk.LEFT,
        )
        self.lbl_destino_val.pack(fill=tk.X)

        tk.Label(
            card_outer,
            text=f'Se borrará A{FILA_INICIO_DATOS}:N{FILA_FIN_BORRADO}  ·  Se incrementará {CELDA_CONTADOR} en +1',
            font=FUENTES['small'],
            bg=COLORES['bg_card'],
            fg=COLORES['text_secondary'],
        ).pack(anchor='w', pady=(4, 0))

    # -------------------------------------------------------------------------
    def _card_log(self, parent):
        """Área de texto para mostrar el progreso."""
        card_outer, card_inner = mgc.create_card(parent, "📋  Registro de Actividad", padding=10)
        card_outer.pack(fill=tk.BOTH, expand=True, pady=(12, 0))

        frame_txt = tk.Frame(card_inner, bg=COLORES['bg_card'])
        frame_txt.pack(fill=tk.BOTH, expand=True)

        scroll = tk.Scrollbar(frame_txt)
        scroll.pack(side=tk.RIGHT, fill=tk.Y)

        self.txt_log = tk.Text(
            frame_txt,
            font=FUENTES['small'],
            bg='#0d1526',
            fg=COLORES['text_primary'],
            relief=tk.FLAT,
            bd=0,
            height=6,
            yscrollcommand=scroll.set,
            state='disabled',
            wrap=tk.WORD,
        )
        self.txt_log.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scroll.configure(command=self.txt_log.yview)

        # Tags de color
        self.txt_log.tag_config('ok',    foreground=COLORES['green'])
        self.txt_log.tag_config('error', foreground=COLORES['red'])
        self.txt_log.tag_config('info',  foreground=COLORES['blue'])

    # -------------------------------------------------------------------------
    def _seleccionar_archivo(self):
        """Abre el diálogo para seleccionar el archivo XLSM origen."""
        ruta = filedialog.askopenfilename(
            title="Seleccionar archivo XLSM de origen",
            filetypes=[
                ("Archivos Excel con macros", "*.xlsm"),
                ("Archivos Excel", "*.xlsx *.xls"),
                ("Todos los archivos", "*.*"),
            ]
        )
        if ruta:
            self.archivo_origen_var.set(ruta)
            self._log(f"📂 Archivo seleccionado: {os.path.basename(ruta)}", 'info')
            self.status_var.set(f"Archivo: {os.path.basename(ruta)}")

    # -------------------------------------------------------------------------
    def _log(self, mensaje: str, tag: str = ''):
        """Escribe un mensaje en el área de log."""
        def _write():
            self.txt_log.configure(state='normal')
            if tag:
                self.txt_log.insert(tk.END, mensaje + "\n", tag)
            else:
                self.txt_log.insert(tk.END, mensaje + "\n")
            self.txt_log.see(tk.END)
            self.txt_log.configure(state='disabled')
        self.root.after(0, _write)

    def _log_clear(self):
        """Limpia el área de log."""
        self.txt_log.configure(state='normal')
        self.txt_log.delete('1.0', tk.END)
        self.txt_log.configure(state='disabled')

    # -------------------------------------------------------------------------
    def _iniciar_proceso(self):
        """Valida los datos y lanza el procesamiento en un hilo separado."""
        if not HAS_OPENPYXL:
            messagebox.showerror("Error", "No está disponible la librería 'openpyxl'.\nInstálela con: pip install openpyxl")
            return

        archivo_origen = self.archivo_origen_var.get().strip()
        if not archivo_origen:
            messagebox.showwarning("Archivo no seleccionado",
                                   "Debe seleccionar el archivo XLSM de origen.")
            return

        banco = self.banco_var.get()

        # Confirmación
        archivo_destino = ARCHIVOS_DESTINO.get(banco, "")
        resp = messagebox.askyesno(
            "Confirmar proceso",
            f"Se procesará:\n\n"
            f"  Banco:   {banco}\n"
            f"  Origen:  {os.path.basename(archivo_origen)}\n"
            f"  Destino: {os.path.basename(archivo_destino)}\n\n"
            f"Se borrará el rango A{FILA_INICIO_DATOS}:N{FILA_FIN_BORRADO} "
            f"y se incrementará {CELDA_CONTADOR}.\n\n"
            f"¿Desea continuar?"
        )
        if not resp:
            return

        # Limpiar log y deshabilitar botón
        self._log_clear()
        self.btn_procesar.configure(state='disabled',
                                  bg=COLORES['gray'],
                                  text="⏳  Procesando...")
        self.status_var.set("Procesando...")

        # Lanzar en hilo separado para no bloquear la GUI
        hilo = threading.Thread(
            target=procesar_guia_pago,
            args=(archivo_origen, banco, self._log, self._on_done),
            daemon=True,
        )
        hilo.start()

    # -------------------------------------------------------------------------
    def _on_done(self, ok: bool, mensaje: str):
        """Callback llamado al finalizar el procesamiento."""
        def _actualizar():
            self.btn_procesar.configure(
                state='normal',
                bg=COLORES['green'],
                text="▶  PROCESAR GUÍA DE PAGO",
            )
            if ok:
                self.status_var.set("✅ Proceso completado con éxito")
                messagebox.showinfo("Proceso finalizado", mensaje)
            else:
                self.status_var.set("❌ Error durante el proceso")
                messagebox.showerror("Error", mensaje)
        self.root.after(0, _actualizar)


# =============================================================================
# MAIN
# =============================================================================

def main():
    root = ctk.CTk()
    app = GuiaPagoBancarioApp(root)
    root.mainloop()


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        try:
            import ctypes
            ctypes.windll.user32.MessageBoxW(0, f"Error fatal en Guía de Pago Bancario:\n{str(e)}", "Error de Inicio", 0x10)
        except:
            pass
        import sys
        sys.stderr.write(f"Error fatal: {e}\n")
