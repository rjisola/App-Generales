import sys
import os
import argparse
import csv
import re
from PyPDF2 import PdfReader
import pandas as pd
import unicodedata

def normalize_text(text):
    if not text: return ""
    nfkd_form = unicodedata.normalize('NFKD', text)
    return "".join([c for c in nfkd_form if not unicodedata.combining(c)]).lower().strip()

def clean_number(val_str):
    if not val_str: return 0.0
    val_str = str(val_str)
    if ',' in val_str and '.' in val_str:
        if val_str.rfind(',') > val_str.rfind('.'):
             val_str = val_str.replace('.', '').replace(',', '.')
        else:
             val_str = val_str.replace(',', '')
    else:
        val_str = val_str.replace(',', '.')
    try:
        return float(val_str)
    except:
        return 0.0

def extract_amount_from_line(line, extract_units=False):
    matches = re.findall(r'-?[\d\.,]+', line)
    if not matches: return ""
    
    if extract_units:
        # Extraer Días/Unidades de forma robusta.
        # Días típicos: 1 a 32.
        for m in matches:
             try:
                 valor = clean_number(m)
                 if 1.0 <= valor <= 32.0 and ('.' in m or ',' in m):
                      return m
             except:
                 pass
                 
        for m in matches:
             try:
                 valor = clean_number(m)
                 if 1.0 <= valor <= 32.0:
                      return m
             except:
                 pass

        if len(matches) >= 3:
             return matches[1] 
        elif len(matches) >= 2:
             return matches[0]
        return matches[-1]
    else:
        return matches[-1]

def extract_employee_data(text):
    legajo = ""
    nombre = ""
    fecha = ""
    lines = text.split('\n')
    
    # 1. Tratar de sacar todo de una linea (Legajo + Nombre + Fecha)
    for line in lines:
        match = re.search(r'(?:^\s*|(?:DO[\'´]?|LEGAJO)\s*)(\d+)\s+(.*?)\s+(?:DU|D\.U\.|CUIL|C\.U\.I\.L\.|D\.N\.I\.|DNI).*?(\d{2}/\d{2}/\d{2,4})', line, re.IGNORECASE)
        if match:
             return match.group(1), match.group(2).strip(), match.group(3)
        
        # 2. Solo Legajo y Nombre
        if not legajo:
            match_simple = re.search(r'(?:^\s*|(?:DO[\'´]?|LEGAJO)\s*)(\d+)\s+(.*?)\s+(?:DU|D\.U\.|CUIL|C\.U\.I\.L\.|D\.N\.I\.|DNI)', line, re.IGNORECASE)
            if match_simple:
                legajo = match_simple.group(1)
                nombre = match_simple.group(2).strip()

    # 3. Buscar fecha independiente
    if not fecha:
        for line in lines:
            if "FECHA DE INGRESO" in line.upper():
                dmatches = re.findall(r'\d{2}/\d{2}/\d{2,4}', line)
                if dmatches: 
                    fecha = dmatches[0]
                    break

    return legajo, nombre, fecha

def extract_quincena(text, default_name):
    # Buscar patrón de fecha "23/12/25 12/25" del layout PDF
    m_doble = re.search(r'\d{2}/\d{2}/\d{2,4}\s+((?:0[1-9]|1[0-2])/\d{2})\b', text[:1500])
    if m_doble:
        return m_doble.group(1)
        
    lines = text.split('\n')
    for i, line in enumerate(lines[:30]):
        if "MES" in line.upper() or "LIQUIDACION" in line.upper():
            entorno = "\n".join(lines[max(0, i-5):min(len(lines), i+5)])
            m = re.findall(r'\b(?:0[1-9]|1[0-2])/\d{2}\b', entorno)
            if m:
                return m[-1]
    return default_name

def calcular_dias_vacaciones(str_ingreso, str_tope):
    try:
        if not str_ingreso or not str_tope: return 0
        
        if len(str_ingreso.split('/')[-1]) == 2:
            parts = str_ingreso.split('/')
            year = int(parts[2])
            # Todas las fechas son desde 2004. Asumimos siempre +2000.
            year = year + 2000
            str_ingreso = f"{parts[0]}/{parts[1]}/{year}"
            
        ingreso = pd.to_datetime(str_ingreso, format="%d/%m/%Y", errors='coerce')
            
        tope = pd.to_datetime(str_tope, format="%d/%m/%Y", errors='coerce')
        
        if pd.isna(ingreso) or pd.isna(tope):
            return 0
            
        dias_antiguedad = (tope - ingreso).days
        
        if dias_antiguedad < 180:
            return max(0, dias_antiguedad // 20)
        elif dias_antiguedad < 5 * 365:
            return 14
        elif dias_antiguedad < 10 * 365:
            return 21
        elif dias_antiguedad < 20 * 365:
            return 28
        else:
            return 35
    except Exception:
        return 0



def search_in_pdfs_pivot(pdf_paths, concept, indice_path=None, fecha_tope="", extract_units=False, callback=None):
    normalized_concept = normalize_text(concept)
    total_files = len(pdf_paths)
    
    maestro = {}
    todas_quincenas = set()
            
    # 1. Cargar el Maestro
    if indice_path and os.path.exists(indice_path):
        if callback: callback("Cargando Índice Excel...")
        try:
            df_ind = pd.read_excel(indice_path, header=None)
            for idx, row in df_ind.iterrows():
                legajo = str(row.iloc[0]).strip()
                nombre = str(row.iloc[1]).strip() if len(row) > 1 else ""
                
                maestro[legajo] = {
                    "Legajo": legajo,
                    "Nombre y Apellido": nombre,
                    "Fecha Ingreso": "",
                    "pdfs": {}
                }
        except Exception as e:
            if callback: callback(f"Error cargando índice: {e}")
            
    # 2. Procesar PDFs
    for index, pdf_path in enumerate(pdf_paths):
        file_name = os.path.basename(pdf_path)
        if callback: callback(f"Procesando {file_name} ({index+1}/{total_files})...")
            
        try:
            reader = PdfReader(pdf_path)
            for i, page in enumerate(reader.pages):
                text = page.extract_text()
                if not text: continue
                # Procesar sólo el original o el duplicado (según lo que traiga) "FIRMA DEL EMPLEADOR"
                if "FIRMA DEL EMPLEADOR" not in text.upper():
                    continue
                    
                leg, nombre_extracted, f_ingreso = extract_employee_data(text)
                
                if not leg: continue
                
                if leg not in maestro:
                    maestro[leg] = {
                        "Legajo": leg,
                        "Nombre y Apellido": nombre_extracted,
                        "Fecha Ingreso": f_ingreso,
                        "pdfs": {}
                    }
                
                if f_ingreso and not maestro[leg].get("Fecha Ingreso", ""):
                    maestro[leg]["Fecha Ingreso"] = f_ingreso
                    
                amount_found = ""
                lines = text.split('\n')
                for line in lines:
                    normalized_line = normalize_text(line)
                    if normalized_concept in normalized_line or concept.lower() in line.lower():
                        if extract_units:
                            extract = extract_amount_from_line(line, True)
                        else:
                            extract = extract_amount_from_line(line, False)
                        
                        if extract:
                            amount_found = extract
                            break
                            
                if amount_found:
                    quincena_hdr = extract_quincena(text, file_name)
                    todas_quincenas.add(quincena_hdr)
                    
                    actual_val = maestro[leg]["pdfs"].get(quincena_hdr, 0.0)
                    if extract_units:
                        # Acumular sumando los días (ej: 9 + 4 = 13 en el mismo PDF y misma quincena_hdr)
                        sumado = clean_number(actual_val) + clean_number(amount_found)
                        maestro[leg]["pdfs"][quincena_hdr] = int(sumado) if int(sumado) == sumado else sumado
                    else:
                        sumado = clean_number(actual_val) + clean_number(amount_found)
                        maestro[leg]["pdfs"][quincena_hdr] = sumado
                    
        except Exception as e:
            if callback: callback(f"Error procesando {file_name}: {e}")

    # 3. Formatear y calcular totales/saldos para cada legajo
    final_results = []
    
    pdf_column_names = sorted(list(todas_quincenas))
    
    ordered_keys = list(maestro.keys())
    if indice_path and os.path.exists(indice_path):
        try:
            df_ind = pd.read_excel(indice_path, header=None)
            idx_legs = []
            for _, r in df_ind.iterrows():
                try:
                    val = str(r.iloc[0]).replace('.0', '').strip()
                    idx_legs.append(val)
                except:
                    pass
            
            ordered_keys_tmp = []
            for lg in idx_legs:
                if lg in maestro:
                    ordered_keys_tmp.append(lg)
                    
            ordered_keys = ordered_keys_tmp
        except:
            ordered_keys = sorted(list(maestro.keys()), key=lambda k: str(maestro[k].get("Nombre y Apellido", "")).upper())
    else:
        ordered_keys = sorted(list(maestro.keys()), key=lambda k: str(maestro[k].get("Nombre y Apellido", "")).upper())
            
    for leg in ordered_keys:
        data = maestro[leg]
        # Inicializar fila de salida
        row = {
            "Legajo": data["Legajo"],
            "Nombre y Apellido": data["Nombre y Apellido"],
            "Fecha Ingreso": data.get("Fecha Ingreso", "")
        }
        
        total_unidades = 0.0
        
        for p_col in pdf_column_names:
            val = data["pdfs"].get(p_col, "")
            
            if val != "":
                num_val = clean_number(val)
                row[p_col] = int(num_val) if int(num_val) == num_val else num_val
                if extract_units:
                    total_unidades += num_val
            else:
                row[p_col] = ""
                
        if extract_units:
            row["Días Tomados (Total)"] = int(total_unidades) if int(total_unidades) == total_unidades else total_unidades
            
            corresponden = 0
            saldo = 0
            if fecha_tope and row["Fecha Ingreso"]:
                corresponden = calcular_dias_vacaciones(row["Fecha Ingreso"], fecha_tope)
                saldo = corresponden - total_unidades
                
            row["Días que Corresponden"] = corresponden
            row["Saldo (Resto)"] = saldo
            
        final_results.append((row, pdf_column_names))
        
    return final_results

# Fallback compatibility
def search_in_pdfs(pdf_paths, concept, callback=None):
    results = search_in_pdfs_pivot(pdf_paths, concept, callback=callback)
    if not results: return []
    # results is list of (dict, p_cols). Convert to old flat list for UI compatibility if needed.
    flat_res = []
    for r, p_cols in results:
        for p in p_cols:
            if r.get(p):
                 flat_res.append({
                     "Archivo": p,
                     "Legajo": r["Legajo"],
                     "Nombre y Apellido": r["Nombre y Apellido"],
                     "Concepto": concept,
                     "Importe": r[p]
                 })
    return flat_res

def search_in_pdf(pdf_path, concept, output_csv):
    results = search_in_pdfs([pdf_path], concept)
    with open(output_csv, 'w', newline='', encoding='utf-8') as f:
        writer = csv.writer(f, delimiter=';', quoting=csv.QUOTE_NONE, escapechar='\\') 
        writer.writerow(["Leg", "Nombre y Apellido", "Importe"])
        for r in results:
            writer.writerow([r["Legajo"], r["Nombre y Apellido"], r["Importe"]])
    print(f"Found {len(results)} matches. Saved to {output_csv}")

def main():
    pass

if __name__ == "__main__":
    main()
