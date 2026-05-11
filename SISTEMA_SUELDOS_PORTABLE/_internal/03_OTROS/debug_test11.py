import sys
import shutil

src = r"C:\Users\rjiso\OneDrive\Escritorio\2DA FEBRERO.xlsx"
src_test = r"C:\Users\rjiso\OneDrive\Escritorio\PROCESAR ARCHIVO SUELDOS\TEST_2DA_FEBRERO_VADO.xlsx"
shutil.copy(src, src_test)

dst = r"C:\Users\rjiso\OneDrive\Escritorio\2DA FEBRERO 2026\PROGRAMA DEPOSITO 2DA FEBRERO2026.xlsm"
dst_test = r"C:\Users\rjiso\OneDrive\Escritorio\PROCESAR ARCHIVO SUELDOS\TEST_PROGRAMA_DEPOSITO_VADO.xlsm"
shutil.copy(dst, dst_test)

indice = r"c:\Users\rjiso\OneDrive\Escritorio\PROCESAR ARCHIVO SUELDOS\indice.xlsx"

import importlib
pasar_horas = importlib.import_module('15-PASAR_HORAS_DEPOSITO')

class MockRoot:
    def after(self, ms, func, *args):
        try:
            func(*args)
        except Exception as e:
            print("After Error:", e)

class MockVar:
    def __init__(self, val=None):
        self.val = val
    def get(self): return self.val
    def set(self, val): self.val = val

class MockApp:
    def __init__(self):
        self.root = MockRoot()
        self.validar_legajos = MockVar(True)
        self.status_var = MockVar("")
        
    def _update_status(self, msg):
        pass
        
    def processing_finished(self, success, msg):
        print(f"FINISHED: Success={success}\nMsg:\n{msg}")
        
    def validar_y_corregir_legajos(self, ws, start_row, file_path, indice_path_selected):
        return pasar_horas.App.validar_y_corregir_legajos(self, ws, start_row, file_path, indice_path_selected)

app = MockApp()
pasar_horas.App.run_process_background(app, src_test, dst_test, indice)
