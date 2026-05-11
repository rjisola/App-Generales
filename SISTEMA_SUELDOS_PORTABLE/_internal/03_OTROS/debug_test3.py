import sys
import importlib

sys.path.insert(0, r'c:\Users\rjiso\OneDrive\Escritorio\PROCESAR ARCHIVO SUELDOS')
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
        print(f"FINISHED: Success={success}")
        
    def validar_y_corregir_legajos(self, ws, start_row, file_path, indice_path_selected):
        return pasar_horas.App.validar_y_corregir_legajos(self, ws, start_row, file_path, indice_path_selected)

src = r"C:\Users\rjiso\OneDrive\Escritorio\2DA FEBRERO.xlsx"
dst = r"C:\Users\rjiso\OneDrive\Escritorio\PROCESAR ARCHIVO SUELDOS\TEST_PROGRAMA_DEPOSITO.xlsm"
indice = r"c:\Users\rjiso\OneDrive\Escritorio\PROCESAR ARCHIVO SUELDOS\indice.xlsx"

import shutil
shutil.copy(r"C:\Users\rjiso\OneDrive\Escritorio\2DA FEBRERO 2026\PROGRAMA DEPOSITO 2DA FEBRERO2026.xlsm", dst)

# Monkey-patch to capture val mapping
original_run = pasar_horas.App.run_process_background
def patched_run(self, src_path, dst_path, indice_path):
    print("Running patched process...")
    # Call the original method but let's intercept just the legajo loop
    pass 

# Since we don't want to copy a lot of code, let's just create a modified version of the module in memory
import time
time.sleep(1)

with open(r'c:\Users\rjiso\OneDrive\Escritorio\PROCESAR ARCHIVO SUELDOS\15-PASAR_HORAS_DEPOSITO.pyw', 'r', encoding='utf-8') as f:
    code = f.read()

code = code.replace(
    'if valor_actual is None or valor_actual == 0:',
    'if "Cardoso" in str(ws_dst.cell(row=dict_dst_row_map[legajo], column=1).value):\n                                print(f"Cardoso: col_src={col_src}, col_dst={col_dst}, val={val}")\n                            if valor_actual is None or valor_actual == 0:'
)

with open('debug_15_temp.py', 'w', encoding='utf-8') as f:
    f.write(code)
