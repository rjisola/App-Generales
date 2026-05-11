import fitz
from PIL import Image
import numpy as np
import os

def test_blue_extraction(pdf_path, output_path):
    doc = fitz.open(pdf_path)
    page = doc.load_page(len(doc) - 1) # Last page
    
    pix = page.get_pixmap(matrix=fitz.Matrix(3, 3), alpha=True)
    img = Image.frombytes("RGBA", [pix.width, pix.height], pix.samples)
    
    data = np.array(img)
    r, g, b, a = data[:,:,0], data[:,:,1], data[:,:,2], data[:,:,3]
    
    # Filter for BLUE ink (Signature in the image is clearly blue)
    # Blue pixels: B > R and B > G
    blue_mask = (b > r + 30) & (b > g + 30)
    
    # Make everything else transparent
    a[~blue_mask] = 0
    data[:,:,3] = a
    
    filtered_img = Image.fromarray(data, 'RGBA')
    
    # Focus on bottom area where signatures live
    h = filtered_img.height
    bottom_half = filtered_img.crop((0, int(h*0.5), filtered_img.width, h))
    
    bbox = bottom_half.getbbox()
    if bbox:
        final = bottom_half.crop(bbox)
        final.save(output_path)
        print(f"Success! Signature extracted to {output_path}")
    else:
        print("No blue signature found in bottom half.")

if __name__ == "__main__":
    pdf = r"G:\Mi unidad\PERSONAL Y VEHICULOS\ARCHIVOS DEL PERSONAL\EXAMENES PREOCUPACIONALES\AGUILAR IVAN PREOCUP.pdf"
    out = r"c:\Users\rjiso\OneDrive\Escritorio\PROCESAR ARCHIVO SUELDOS\02_CARPETAS\test_firmas\final_signature_aguilar.png"
    test_blue_extraction(pdf, out)
