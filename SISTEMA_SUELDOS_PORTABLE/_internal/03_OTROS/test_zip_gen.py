
import os
import io
import zipfile
import pandas as pd
from reportlab.pdfgen import canvas
import PyPDF2
import unicodedata

# 1. Setup Dummy Files
def create_dummy_pdf(filename, text_content):
    c = canvas.Canvas(filename)
    c.drawString(100, 750, text_content)
    c.save()
    print(f"Created {filename}")

def create_dummy_excel(filename, names):
    # Create DataFrame with name in Column B (index 1)
    df = pd.DataFrame({'ColA': range(len(names)), 'ColB': names})
    df.to_excel(filename, header=False, index=False)
    print(f"Created {filename}")

def normalize_text(text):
    if not isinstance(text, str):
        text = str(text)
    return ''.join(
        c for c in unicodedata.normalize('NFD', text)
        if unicodedata.category(c) != 'Mn'
    ).lower().strip()

def run_test():
    print("--- Starting Test ---")
    
    # Clean up old files
    for f in ['test_receipt_1.pdf', 'test_receipt_2.pdf', 'test_index.xlsx', 'test_output.zip']:
        if os.path.exists(f): 
            try: os.remove(f)
            except: pass

    # Create inputs
    create_dummy_pdf('test_receipt_1.pdf', "Recibo de Haberes - Empleado: Juan Perez - Legajo 123")
    create_dummy_pdf('test_receipt_2.pdf', "Recibo de Haberes - Empleado: Ana Gomez - Legajo 456")
    
    # Excel with names to find (Column B)
    # Adding a name that exists (Juan Perez) and one that doesn't (Carlos)
    create_dummy_excel('test_index.xlsx', ["Juan Perez", "Carlos Desconocido", "Ana Gomez"])
    
    # 2. Simulate Batch Logic (Copied from BuscarRecibosPDF.pyw)
    
    archivo_indice_path = 'test_index.xlsx'
    archivos_pdf_seleccionados = ['test_receipt_1.pdf', 'test_receipt_2.pdf']
    
    print("\n[Logic] Reading Excel...")
    df = pd.read_excel(archivo_indice_path, header=None)
    nombres = df.iloc[:, 1].dropna().astype(str).tolist()
    nombres = [n.strip() for n in nombres if n.strip()]
    print(f"[Logic] Names found: {nombres}")
    
    print("[Logic] Preparing Readers...")
    readers = []
    files_handles = []
    
    for path in archivos_pdf_seleccionados:
        f = open(path, 'rb')
        files_handles.append(f)
        readers.append((os.path.basename(path), PyPDF2.PdfReader(f)))
        
    zip_buffer = io.BytesIO()
    archivos_generados = 0
    
    print("[Logic] Processing...")
    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
        for nombre in nombres:
            nombre_norm = normalize_text(nombre)
            writer = PyPDF2.PdfWriter()
            paginas_agregadas = 0
            
            for pdf_name, reader in readers:
                for page in reader.pages:
                    txt = normalize_text(page.extract_text() or "")
                    if nombre_norm in txt:
                        writer.add_page(page)
                        paginas_agregadas += 1
            
            if paginas_agregadas > 0:
                pdf_bytes = io.BytesIO()
                writer.write(pdf_bytes)
                pdf_filename = f"{nombre.replace(' ', '_')}.pdf"
                zip_file.writestr(pdf_filename, pdf_bytes.getvalue())
                archivos_generados += 1
                print(f"  + Generated: {pdf_filename}")
            else:
                print(f"  - Not found: {nombre}")

    # Close handles
    for f in files_handles: f.close()
    
    # Save ZIP
    with open('test_output.zip', 'wb') as f:
        f.write(zip_buffer.getvalue())
        
    print(f"\n[Result] Created test_output.zip with size: {os.path.getsize('test_output.zip')} bytes")
    print(f"[Result] Total files in ZIP: {archivos_generados}")

    # Verify ZIP content
    with zipfile.ZipFile('test_output.zip', 'r') as z:
        print(f"[Verify] ZIP Content: {z.namelist()}")
        if 'Juan_Perez.pdf' in z.namelist() and 'Ana_Gomez.pdf' in z.namelist():
             print("SUCCESS: Expected files found in ZIP.")
        else:
             print("FAILURE: Expected files NOT found.")

if __name__ == "__main__":
    run_test()
