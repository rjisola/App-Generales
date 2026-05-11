import sys
import shutil
import importlib

sys.path.insert(0, r'c:\Users\rjiso\OneDrive\Escritorio\PROCESAR ARCHIVO SUELDOS')

# We'll patch the real module before running
pasar_horas = importlib.import_module('15-PASAR_HORAS_DEPOSITO')

original_run = pasar_horas.App.run_process_background

def patched_run(self, src_path, dst_path, indice_path_selected):
    # This is a bit complex, let's just use the file text approach again but cleanly
    pass

import time
time.sleep(1)

with open(r'c:\Users\rjiso\OneDrive\Escritorio\PROCESAR ARCHIVO SUELDOS\15-PASAR_HORAS_DEPOSITO.pyw', 'r', encoding='utf-8') as f:
    code = f.read()

code = code.replace(
    'if valor_actual is None or valor_actual == 0:',
    '''                            # INSTRUMENTATION
                            nom_destino = str(ws_dst.cell(row=dict_dst_row_map[legajo], column=1).value).lower()
                            if 'cardoso' in nom_destino or 'perez' in nom_destino or 'pérez' in nom_destino:
                                with open('debug_out5.txt', 'a', encoding='utf-8') as dbg_f:
                                    dbg_f.write(f"{nom_destino.upper()}: col_src={col_src}, col_dst={col_dst}, val={val}\\n")
                            if valor_actual is None or valor_actual == 0:'''
)

with open('debug_15_temp.py', 'w', encoding='utf-8') as f:
    f.write(code)

import debug_15_temp

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
        return debug_15_temp.App.validar_y_corregir_legajos(self, ws, start_row, file_path, indice_path_selected)

src = r"C:\Users\rjiso\OneDrive\Escritorio\2DA FEBRERO.xlsx"
src_test = r"C:\Users\rjiso\OneDrive\Escritorio\PROCESAR ARCHIVO SUELDOS\TEST_2DA_FEBRERO.xlsx"
shutil.copy(src, src_test)

dst = r"C:\Users\rjiso\OneDrive\Escritorio\2DA FEBRERO 2026\PROGRAMA DEPOSITO 2DA FEBRERO2026.xlsm"
dst_test = r"C:\Users\rjiso\OneDrive\Escritorio\PROCESAR ARCHIVO SUELDOS\TEST_PROGRAMA_DEPOSITO.xlsm"
shutil.copy(dst, dst_test)

indice = r"c:\Users\rjiso\OneDrive\Escritorio\PROCESAR ARCHIVO SUELDOS\indice.xlsx"

open('debug_out5.txt', 'w').close() # clear
app = MockApp()
debug_15_temp.App.run_process_background(app, src_test, dst_test, indice)
