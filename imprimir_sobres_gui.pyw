# -*- coding: utf-8 -*-
"""
Generador de Sobres C5 desde Excel
Replica la lógica VBA para imprimir sobres masivamente seleccionando columnas.
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import os
import sys
import pythoncom
import win32com.client
import traceback
import threading

# Asegurar que el directorio del script esté en sys.path para imports locales
script_dir = os.path.dirname(os.path.abspath(__file__))
if script_dir not in sys.path:
    sys.path.insert(0, script_dir)

# Importar componentes modernos
try:
    import modern_gui_components as mgc
    from icon_loader import set_window_icon, load_icon
except ImportError:
    # Fallback si no existen los módulos
    import tkinter.messagebox as messagebox
    messagebox.showerror("Error", "Faltan módulos necesarios (modern_gui_components.py)")
    sys.exit(1)

class EnvelopePrinterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("🖨️ Impresión Masiva de Sobres")
        self.root.geometry("800x650")
        self.root.resizable(False, False)
        self.root.configure(bg=mgc.COLORS['bg_primary'])
        
        mgc.center_window(self.root, 800, 650)
        set_window_icon(self.root, 'printer') 

        # Cargar iconos
        self.icon_excel = load_icon('excel', (24, 24))
        self.icon_printer = load_icon('printer', (24, 24))
        self.icon_preview = load_icon('search', (24, 24)) # Usar lupa para vista previa

        # Variables
        self.file_path = tk.StringVar()
        self.sheet_name = tk.StringVar(value="RECUENTO TOTAL (2)")
        self.status_var = tk.StringVar(value="Listo. Seleccione un archivo.")
        
        self.create_widgets()

    def create_widgets(self):
        # Frame principal
        main_frame = tk.Frame(self.root, bg=mgc.COLORS['bg_primary'])
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # Header
        mgc.create_header(main_frame, 
                         "Impresión de Sobres", 
                         "Genera y formatea sobres C5 desde Excel", 
                         "✉️")

        # --- CARD 1: Selección de Archivo ---
        card1_outer, card1_inner = mgc.create_card(main_frame, "1. Archivo de Datos", padding=15)
        card1_outer.pack(fill=tk.X, pady=(0, 15))
        
        selector = mgc.create_file_selector(
            card1_inner, "Archivo Excel:", self.file_path, self.select_file, "📂"
            , icon_image=self.icon_excel
        )
        selector.pack(fill=tk.X)

        # --- CARD 2: Configuración de Datos ---
        card2_outer, card2_inner = mgc.create_card(main_frame, "2. Selección de Datos", padding=15)
        card2_outer.pack(fill=tk.X, pady=(0, 15))
        
        # Grid para configuración
        grid_frame = tk.Frame(card2_inner, bg=mgc.COLORS['bg_card'])
        grid_frame.pack(fill=tk.X)
        
        # Fila 1: Hoja y Fila de Inicio
        tk.Label(grid_frame, text="Nombre de Hoja:", font=mgc.FONTS['normal'], bg=mgc.COLORS['bg_card']).grid(row=0, column=0, sticky='w', padx=5, pady=5)
        tk.Entry(grid_frame, textvariable=self.sheet_name, width=25, font=mgc.FONTS['normal']).grid(row=0, column=1, sticky='w', padx=5, pady=5)

        tk.Label(grid_frame, text="ℹ️ Al generar, se abrirá Excel para que selecciones las filas y columnas con el mouse.", 
                 font=mgc.FONTS['small'], bg=mgc.COLORS['bg_card'], fg=mgc.COLORS['text_secondary']).grid(row=1, column=0, columnspan=2, sticky='w', padx=5, pady=5)
        
        # --- CARD 3: Acciones ---
        card3_outer, card3_inner = mgc.create_card(main_frame, "3. Generar", padding=15)
        card3_outer.pack(fill=tk.X, pady=(0, 15))
        
        btn_frame = tk.Frame(card3_inner, bg=mgc.COLORS['bg_card'])
        btn_frame.pack()
        
        self.btn_preview = mgc.create_large_button(
            btn_frame, "GENERAR Y VISTA PREVIA", 
            lambda: self.run_process(preview=True), 
            color='blue', icon="👁️", text_color='white', icon_image=self.icon_preview
        )
        self.btn_preview.pack(side=tk.LEFT, padx=10)
        
        self.btn_print = mgc.create_large_button(
            btn_frame, "IMPRIMIR DIRECTO", 
            lambda: self.run_process(preview=False), 
            color='green', icon="🖨️", text_color='white', icon_image=self.icon_printer
        )
        self.btn_print.pack(side=tk.LEFT, padx=10)

        # Barra de estado y progreso
        self.status_frame, self.status_var = mgc.create_status_bar(self.root)
        
        # Barra de progreso (inicialmente oculta o en 0)
        self.progress = ttk.Progressbar(self.status_frame, orient=tk.HORIZONTAL, length=200, mode='determinate')
        self.progress.pack(side=tk.RIGHT, padx=10)

    def _get_col_letters(self):
        return [chr(i) for i in range(65, 91)] + ['AA', 'AB', 'AC', 'AD', 'AE']

    def select_file(self):
        path = filedialog.askopenfilename(
            title="Seleccionar archivo Excel",
            filetypes=[("Excel Files", "*.xlsx *.xlsm *.xls")]
        )
        if path:
            self.file_path.set(path)
            self.status_var.set(f"Archivo seleccionado: {os.path.basename(path)}")
            # Resetear progreso
            self.progress['value'] = 0

        if not path or not os.path.exists(path):
            messagebox.showwarning("Falta archivo", "Seleccione un archivo válido.")
            return

    def run_process(self, preview):
        path = self.file_path.get()
        if not path or not os.path.exists(path):
             messagebox.showwarning("Falta archivo", "Seleccione un archivo válido.")
             return

        # Deshabilitar botones
        mgc.disable_button(self.btn_preview)
        mgc.disable_button(self.btn_print)
        self.status_var.set("Iniciando Excel...")
        self.progress['value'] = 0 # Reset
        
        # Leer valores de UI en el hilo principal
        sheet_name_val = self.sheet_name.get()
        self.root.update()

        try:
            # Ejecutar en hilo secundario para evitar congelamiento
            threading.Thread(target=self._run_thread, args=(path, sheet_name_val, preview), daemon=True).start()
        except Exception as e:
            self._on_error(e)

    def _run_thread(self, path, sheet_name, preview):
        try:
            self._process_excel(path, sheet_name, preview)
        except Exception as e:
            # Invocar callback de error en main thread (a través de after quizás, o manejo directo si safe)
            # Tkinter no es thread-safe, mejor manejarlo así:
            self.root.after(0, lambda: self._on_error(e))

    def _on_error(self, e):
        tb = traceback.format_exc()
        messagebox.showerror("Error Crítico", f"Ocurrió un error:\n{e}\n\nDetalles:\n{tb}")
        self._reset_ui()

    def _reset_ui(self):
        mgc.enable_button(self.btn_preview, 'blue')
        mgc.enable_button(self.btn_print, 'green')
        self.status_var.set("Proceso finalizado.")
        self.progress['value'] = 0

    def _update_progress(self, value, max_value):
        # Actualiza la barra de progreso desde el thread secundario de forma segura
        self.root.after(0, lambda: self._set_progress(value, max_value))

    def _set_progress(self, value, max_value):
        if self.progress['maximum'] != max_value:
             self.progress['maximum'] = max_value
        self.progress['value'] = value

    # --- Métodos Thread-Safe para Diálogos ---
    def _safe_ask_ok_cancel(self, title, message):
        result = [False]
        event = threading.Event()
        def _ask():
            result[0] = messagebox.askokcancel(title, message, parent=self.root)
            event.set()
        self.root.after(0, _ask)
        event.wait()
        return result[0]

    def _safe_show_error(self, title, message):
        event = threading.Event()
        def _show():
            messagebox.showerror(title, message, parent=self.root)
            event.set()
        self.root.after(0, _show)
        event.wait()

    def _safe_ask_yes_no(self, title, message):
        result = [False]
        event = threading.Event()
        def _ask():
            result[0] = messagebox.askyesno(title, message, parent=self.root)
            event.set()
        self.root.after(0, _ask)
        event.wait()
        return result[0]

    def _process_excel(self, path, sheet_name, preview):
        excel = None
        wb = None
        try:
            # Configurar barra de progreso total
            self._update_progress(0, 100)
            
            # 1. Iniciar Excel
            pythoncom.CoInitialize()
            try:
                 excel = win32com.client.GetActiveObject("Excel.Application")
            except:
                 excel = win32com.client.Dispatch("Excel.Application")
            
            excel.Visible = True 
            excel.DisplayAlerts = False
            self._update_progress(10, 100)

            # 2. Abrir Libro
            self.root.after(0, lambda: self.status_var.set("Abriendo libro..."))
            
            # Verificar si ya está abierto
            wb = None
            abs_path = os.path.abspath(path)
            for w in excel.Workbooks:
                if w.FullName == abs_path:
                    wb = w
                    break
            
            if not wb:
                wb = excel.Workbooks.Open(abs_path)
            self._update_progress(20, 100)
            
            try:
                ws_data = wb.Sheets(sheet_name)
                ws_data.Activate()
            except:
                self._safe_show_error("Error", f"No se encontró la hoja '{sheet_name}'")
                self.root.after(0, self._reset_ui)
                return
            self._update_progress(30, 100)

            # 3. Selección Interactiva (Refactorizado)
            # En lugar de InputBox (que falla devolviendo valores), pedimos al usuario que seleccione y confirme.
            
            excel.Visible = True
            
            # Mensaje instructivo (Modal)
            msg_response = self._safe_ask_ok_cancel(
                "Seleccionar Datos",
                "Por favor, realice los siguientes pasos:\n\n"
                "1. Vaya a Excel y seleccione con el mouse el rango de datos (Legajos y Nombres).\n"
                "2. Vuelva aquí y presione ACEPTAR.\n"
            )
            
            if not msg_response:
                # Usuario canceló
                self.root.after(0, self._reset_ui)
                return

            try:
                # Capturar la selección actual de Excel
                user_range = excel.Selection
                
                # Validar que sea un Rango verificando si tiene propiedad Address
                try:
                    _ = user_range.Address
                except:
                     raise Exception("La selección actual no corresponde a celdas válidas.")
                     
            except Exception as e:
                self._safe_show_error("Error de Selección", f"No se pudo obtener la selección de Excel.\nAsegúrese de seleccionar celdas.\n\nError: {e}")
                self.root.after(0, self._reset_ui)
                return
            self._update_progress(40, 100)

            # 4. Configurar Hoja de Impresión
            PRINT_SHEET_NAME = "ImpresionSobres_Temp"
            
            try:
                wb.Sheets(PRINT_SHEET_NAME).Delete()
            except:
                pass
            
            ws_print = wb.Sheets.Add(After=wb.Sheets(wb.Sheets.Count))
            ws_print.Name = PRINT_SHEET_NAME

            # OPTIMIZACIÓN: Desactivar actualizaciones visuales y de impresora ANTES de configurar página
            excel.ScreenUpdating = False
            excel.Calculation = -4135 # xlCalculationManual
            try:
                excel.PrintCommunication = False
                ws_print.DisplayPageBreaks = False
            except:
                pass

            # 5. Configurar Página (Sobre C5)
            self.root.after(0, lambda: self.status_var.set("Configurando página..."))
            try:
                ws_print.PageSetup.PaperSize = 28 # xlPaperEnvelopeC5
                ws_print.PageSetup.Orientation = 1 
                ws_print.PageSetup.TopMargin = 3 * 28.35
                ws_print.PageSetup.LeftMargin = 2 * 28.35
                ws_print.PageSetup.RightMargin = 2 * 28.35
                ws_print.PageSetup.BottomMargin = 2 * 28.35
                ws_print.PageSetup.CenterHorizontally = True
            except Exception as e:
                print(f"Advertencia setup: {e}")

            self._update_progress(50, 100)

            # 6. Procesamiento en Lote (Batch)
            self.root.after(0, lambda: self.status_var.set("Generando etiquetas (Optimizado)..."))
            
            # Leer datos a memoria
            raw_values = user_range.Value
            # Normalizar a lista de tuplas
            if not isinstance(raw_values, tuple): raw_values = ((raw_values,),) # Un solo valor
            
            output_data = []
            
            for row in raw_values:
                # row es una tupla
                legajo = row[0] if len(row) > 0 else ""
                nombre = row[1] if len(row) > 1 else ""
                
                if legajo or nombre:
                    l_str = str(int(legajo)) if isinstance(legajo, (int, float)) else str(legajo)
                    n_str = str(nombre) if nombre else ""
                    text = f"({l_str}) {n_str}"
                    output_data.append([text])
            
            num_sobres = len(output_data)
            
            # Configurar maximo de barra de progreso para la parte de sobres (de 50% a 100%)
            progress_start = 50
            progress_range = 50  # de 50 a 100
            
            if num_sobres > 0:
                # Escribir DE GOLPE
                dest_range = ws_print.Range(ws_print.Cells(1, 1), ws_print.Cells(num_sobres, 1))
                dest_range.Value = output_data
                
                # Formatear DE GOLPE
                dest_range.Font.Name = "Arial"
                dest_range.Font.Size = 14
                dest_range.Font.Bold = True
                dest_range.HorizontalAlignment = -4108 # xlCenter
                dest_range.VerticalAlignment = -4160 # xlTop
                dest_range.WrapText = True
                dest_range.RowHeight = 30
                
                # Ancho de columna
                ws_print.Columns("A:A").ColumnWidth = 50
                
                # Insertar saltos solo si son necesarios
                for i in range(2, num_sobres + 1):
                    ws_print.HPageBreaks.Add(Before=ws_print.Cells(i, 1))
                    
                    # Actualizar progreso cada 5 sobres o al final para no sobrecargar
                    if i % 5 == 0 or i == num_sobres:
                         current_progress = progress_start + int((i / num_sobres) * progress_range)
                         self._update_progress(current_progress, 100)

                # Progreso al 100%
                self._update_progress(100, 100)
                
                try:
                    excel.PrintCommunication = True
                except:
                    pass
            
            # Restaurar cálculo auto
            excel.Calculation = -4105 # xlCalculationAutomatic

            self.root.after(0, lambda: self.status_var.set(f"Generados {num_sobres} sobres."))
            
            excel.ScreenUpdating = True
            excel.Visible = True
            ws_print.Activate()

            if preview:
                # Usar invoke en main thread para dialogos modales que podrian bloquear COM?
                # excel.Dialogs(8).Show() es bloqueante.
                # Si lo llamamos aqui en thread, Excel se muestra.
                try:
                    excel.Dialogs(8).Show()
                except:
                    pass
            else:
                # Confirmación antes de imprimir
                if self._safe_ask_yes_no("Confirmar Impresión", "¿Está seguro que desea enviar los sobres a la impresora?"):
                    ws_print.PrintOut()
                else:
                    self.root.after(0, lambda: self.status_var.set("Impresión cancelada por el usuario."))

            self.root.after(0, self._reset_ui)

        except Exception as e:
            if excel: excel.ScreenUpdating = True
            raise e
        finally:
             if excel: excel.DisplayAlerts = True

if __name__ == "__main__":
    root = tk.Tk()
    app = EnvelopePrinterApp(root)
    root.mainloop()