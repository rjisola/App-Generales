import fitz
from PIL import Image
import os

def analyze_pdf_signatures(pdf_path, output_path):
    doc = fitz.open(pdf_path)
    page_count = len(doc)
    print(f"Total pages: {page_count}")
    
    # Usually signatures are on the last page or second to last
    page = doc.load_page(page_count - 1)
    zoom = 3
    mat = fitz.Matrix(zoom, zoom)
    pix = page.get_pixmap(matrix=mat, alpha=True)
    img = Image.frombytes("RGBA", [pix.width, pix.height], pix.samples)
    
    # Save the full page for diagnosis
    img.save(output_path)
    print(f"Full page rendered to: {output_path}")
    doc.close()

if __name__ == "__main__":
    pdf_file = r"G:\Mi unidad\PERSONAL Y VEHICULOS\ARCHIVOS DEL PERSONAL\EXAMENES PREOCUPACIONALES\AGUILAR IVAN PREOCUP.pdf"
    out_file = r"c:\Users\rjiso\OneDrive\Escritorio\PROCESAR ARCHIVO SUELDOS\02_CARPETAS\test_firmas\full_page_aguilar.png"
    os.makedirs(os.path.dirname(out_file), exist_ok=True)
    analyze_pdf_signatures(pdf_file, out_file)
