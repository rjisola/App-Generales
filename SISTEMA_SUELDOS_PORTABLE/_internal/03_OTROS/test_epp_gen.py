import sys
sys.path.append('.')
from generar_epp_excel import EPPGeneratorApp
class DummyApp(EPPGeneratorApp):
    def __init__(self):
        class DummyVar:
            def __init__(self, val):
                self.val = val
            def get(self):
                return self.val
        self.vars = {
            'nombre': DummyVar('AGUILAR IVAN ARTURO'),
            'dni': DummyVar('20277430534'),
            'proyecto': DummyVar('CARJOR DEPOSITO'),
            'cargo': DummyVar('OPERARIO'),
            'jefe': DummyVar('TORCHIANA AGOSTINA'),
            'fecha_entrega': DummyVar('05-01-2026')
        }
app = DummyApp()
datos = {k: v.get() for k, v in app.vars.items()}
app.generar_excel_logic('test_epp.xlsx', datos)
print("Generacion exitosa")
