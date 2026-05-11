
import os
import PyPDF2
import pandas as pd
import unicodedata
import zipfile
import io
import re

def normalize_text(text):
    if not isinstance(text, str):
        text = str(text)
    text = ''.join(
        c for c in unicodedata.normalize('NFD', text)
        if unicodedata.category(c) != 'Mn'
    )
    text = re.sub(r'[^\w\s]', ' ', text)
    return ' '.join(text.lower().split())

def check_signature(text_norm, filter_type='ambas'):
    if filter_type == 'ambas':
        return True
    text_clean = text_norm.replace(" ", "")
    if filter_type == 'empleado':
        has_empleado = 'firmadelempleado' in text_clean or ('firma' in text_clean and 'empleado' in text_clean)
        is_empleador = 'empleador' in text_clean
        return has_empleado and not is_empleador
    if filter_type == 'empleador':
        return 'firmadelempleador' in text_clean or ('firma' in text_clean and 'empleador' in text_clean)
    return True

def get_search_variations(full_name_norm):
    parts = full_name_norm.split()
    variations = [full_name_norm]
    if len(parts) >= 2:
        variations.append(f"{parts[0]} {parts[1]}")
        variations.append(f"{parts[1]} {parts[0]}")
        if len(parts) >= 3:
            variations.append(f"{' '.join(parts[1:])} {parts[0]}")
    return list(set(variations))

def main():
    pdfs = [
        r"C:\Users\rjiso\OneDrive\Escritorio\algo\escritorio\Reparaciones\Nueva carpeta\2025\1ERA DICIEMBRE 2025\recibos-ordenadosVACACIONES.pdf",
        r"C:\Users\rjiso\OneDrive\Escritorio\algo\escritorio\Reparaciones\Nueva carpeta\2025\1ERA DICIEMBRE 2025\SAC recibos-ordenados.pdf",
        r"C:\Users\rjiso\OneDrive\Escritorio\algo\escritorio\Reparaciones\Nueva carpeta\2025\1ERA DICIEMBRE 2025\RECIBOS QUINCENA_ORDENADA.pdf"
    ]
    excel_path = r"C:\Users\rjiso\OneDrive\Escritorio\carjorExcelJS\API\public\ACOMODAR_PDF\index.xlsx"
    output_zip = r"C:\Users\rjiso\OneDrive\Escritorio\PROCESAR ARCHIVO SUELDOS\Resultado_Prueba_3Pasadas.zip"
    
    print(f"Leyendo Excel: {excel_path}")
    df = pd.read_excel(excel_path, header=None)
    # Columna B es index 1
    nombres_raw = df.iloc[:, 1].dropna().astype(str).tolist()
    nombres = [n.strip() for n in nombres_raw if n.strip() and n.lower() != 'nombre']
    print(f"Encontrados {len(nombres)} nombres.")

    print("Caché de PDFs...")
    pdf_data_cache = []
    for path in pdfs:
        print(f"  -> {os.path.basename(path)}")
        reader = PyPDF2.PdfReader(path)
        for page in reader.pages:
            pdf_data_cache.append((page, normalize_text(page.extract_text() or "")))
    
    zip_buffer = io.BytesIO()
    archivos_generados = 0
    filtro = 'ambas'

    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
        for i, nombre in enumerate(nombres):
            nombre_norm = normalize_text(nombre)
            nombre_agresivo = nombre_norm.replace(" ", "")
            parts = nombre_norm.split()
            variations = get_search_variations(nombre_norm)
            
            writer = PyPDF2.PdfWriter()
            paginas_agregadas = 0
            pasadas_usadas = set()

            for page_obj, page_text in pdf_data_cache:
                if not check_signature(page_text, filtro): continue
                
                encontrado_en_pag = 0
                # 1. Pasada Exacta
                for v in variations:
                    if v in page_text:
                        encontrado_en_pag = 1
                        break
                
                # 2. Pasada por partes
                if not encontrado_en_pag and len(parts) >= 2:
                    if parts[0] in page_text and parts[1] in page_text:
                        encontrado_en_pag = 2
                
                # 3. Pasada Agresiva
                if not encontrado_en_pag:
                    texto_sin_espacios = page_text.replace(" ", "")
                    found_agresivo = False
                    if nombre_agresivo in texto_sin_espacios:
                        found_agresivo = True
                    else:
                        for v in variations:
                            if v.replace(" ", "") in texto_sin_espacios:
                                found_agresivo = True
                                break
                    if found_agresivo:
                        encontrado_en_pag = 3
                
                if encontrado_en_pag:
                    writer.add_page(page_obj)
                    paginas_agregadas += 1
                    pasadas_usadas.add(encontrado_en_pag)
            
            if paginas_agregadas > 0:
                out_pdf = io.BytesIO()
                writer.write(out_pdf)
                fname = f"{nombre.replace(' ', '_')}.pdf"
                zip_file.writestr(fname, out_pdf.getvalue())
                archivos_generados += 1
                pasadas_str = ",".join(map(str, sorted(list(pasadas_usadas))))
                print(f"✔ [{i+1}/{len(nombres)}] Generado (Pasadas {pasadas_str}): {fname} ({paginas_agregadas} págs)")

    if archivos_generados > 0:
        with open(output_zip, 'wb') as f:
            f.write(zip_buffer.getvalue())
        print(f"\n¡ÉXITO! Se generaron {archivos_generados} archivos en: {output_zip}")
    else:
        print("\nNo se encontro ninguna coincidencia.")

if __name__ == "__main__":
    main()
