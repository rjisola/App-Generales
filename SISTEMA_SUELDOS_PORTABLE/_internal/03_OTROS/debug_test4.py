import sys
import shutil

src = r"C:\Users\rjiso\OneDrive\Escritorio\2DA FEBRERO.xlsx"
src_test = r"C:\Users\rjiso\OneDrive\Escritorio\PROCESAR ARCHIVO SUELDOS\TEST_2DA_FEBRERO.xlsx"
shutil.copy(src, src_test)

dst = r"C:\Users\rjiso\OneDrive\Escritorio\2DA FEBRERO 2026\PROGRAMA DEPOSITO 2DA FEBRERO2026.xlsm"
dst_test = r"C:\Users\rjiso\OneDrive\Escritorio\PROCESAR ARCHIVO SUELDOS\TEST_PROGRAMA_DEPOSITO.xlsm"
shutil.copy(dst, dst_test)

indice = r"c:\Users\rjiso\OneDrive\Escritorio\PROCESAR ARCHIVO SUELDOS\indice.xlsx"

import debug_15_temp
app = debug_15_temp.App(None)

class MockRoot:
    def after(self, ms, func, *args):
        try:
            func(*args)
        except Exception as e:
            print("After Error:", e)

app.root = MockRoot()
app.validar_legajos.set(True)

debug_15_temp.App.run_process_background(app, src_test, dst_test, indice)
