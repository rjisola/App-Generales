import fitz
from PIL import Image
import numpy as np
import os

def extract_signature(pdf_path, output_path, page_idx):
    doc = fitz.open(pdf_path)
    if page_idx >= len(doc):
        print(f"Error: Page {page_idx} not found in {pdf_path}")
        return
    
    page = doc.load_page(page_idx)
    # Render at 400 DPI (zoom 4) for high quality
    pix = page.get_pixmap(matrix=fitz.Matrix(4, 4), alpha=True)
    img = Image.frombytes("RGBA", [pix.width, pix.height], pix.samples)
    
    data = np.array(img)
    r, g, b, a = data[:,:,0], data[:,:,1], data[:,:,2], data[:,:,3]
    
    # Blue ink filter
    blue_mask = (b > r + 30) & (b > g + 30)
    a[~blue_mask] = 0
    data[:,:,3] = a
    
    filtered_img = Image.fromarray(data, 'RGBA')
    
    # Autocrop
    bbox = filtered_img.getbbox()
    if bbox:
        # Add 10px margin
        left, top, right, bottom = bbox
        left = max(0, left - 10)
        top = max(0, top - 10)
        right = min(img.width, right + 10)
        bottom = min(img.height, bottom + 10)
        filtered_img.crop((left, top, right, bottom)).save(output_path)
        print(f"Extracted: {output_path}")
    else:
        print(f"No signature found on page {page_idx+1} of {pdf_path}")
    doc.close()

if __name__ == "__main__":
    base_out = r"C:\Users\rjiso\.gemini\antigravity\brain\8864ff70-01cc-41bd-9f86-d741ece4b766"
    
    # Aguilar
    pdf1 = r"G:\Mi unidad\PERSONAL Y VEHICULOS\ARCHIVOS DEL PERSONAL\EXAMENES PREOCUPACIONALES\AGUILAR IVAN PREOCUP.pdf"
    extract_signature(pdf1, os.path.join(base_out, "resultado_aguilar_page3.png"), 2)
    
    # Andrade
    pdf2 = r"G:\Mi unidad\PERSONAL Y VEHICULOS\ARCHIVOS DEL PERSONAL\EXAMENES PREOCUPACIONALES\ANDRADE DIEGO PREOCUP.pdf"
    extract_signature(pdf2, os.path.join(base_out, "resultado_andrade_page3.png"), 2)
