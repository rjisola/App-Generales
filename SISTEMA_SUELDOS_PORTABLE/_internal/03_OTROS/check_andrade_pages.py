import fitz
from PIL import Image
import os

def render_first_pages(pdf_path, output_folder, num_pages=5):
    os.makedirs(output_folder, exist_ok=True)
    doc = fitz.open(pdf_path)
    
    for i in range(min(num_pages, len(doc))):
        page = doc.load_page(i)
        pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        out_file = os.path.join(output_folder, f"page_{i+1}.png")
        img.save(out_file)
        print(f"Rendered Page {i+1} to {out_file}")
    
    doc.close()

if __name__ == "__main__":
    pdf = r"G:\Mi unidad\PERSONAL Y VEHICULOS\ARCHIVOS DEL PERSONAL\EXAMENES PREOCUPACIONALES\ANDRADE DIEGO PREOCUP.pdf"
    out = r"c:\Users\rjiso\OneDrive\Escritorio\PROCESAR ARCHIVO SUELDOS\02_CARPETAS\test_firmas\andrade_pages"
    render_first_pages(pdf, out)
