import time

with open(r'c:\Users\rjiso\OneDrive\Escritorio\PROCESAR ARCHIVO SUELDOS\15-PASAR_HORAS_DEPOSITO.pyw', 'r', encoding='utf-8') as f:
    code = f.read()

replacement = """                            nom_destino = str(ws_dst.cell(row=dict_dst_row_map[legajo], column=1).value).lower()
                            if 'cardoso' in nom_destino or 'perez' in nom_destino or 'pérez' in nom_destino:
                                with open('debug_out7.txt', 'a', encoding='utf-8') as dbg_f:
                                    dbg_f.write(f"DESTINO: {nom_destino.upper()} -> LEGAJO: {legajo}, col_src={col_src}, col_dst={col_dst}, val={val}\\n")
                            if valor_actual is None or valor_actual == 0:"""

code = code.replace('                            if valor_actual is None or valor_actual == 0:', replacement)

with open('debug_15_temp2.py', 'w', encoding='utf-8') as f:
    f.write(code)
