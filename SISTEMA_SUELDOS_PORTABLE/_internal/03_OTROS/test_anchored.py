import fitz
from PIL import Image
import numpy as np
import os

def anchored_extraction(pdf_path, output_path, page_idx, anchor_text):
    doc = fitz.open(pdf_path)
    page = doc.load_page(page_idx)
    
    # Buscar el ancla
    text_instances = page.search_for(anchor_text)
    if not text_instances:
        print(f"Anchor '{anchor_text}' not found in {pdf_path}")
        doc.close()
        return False
    
    # Primera instancia
    inst = text_instances[0]
    # Rect de búsqueda (150 unidades arriba del texto)
    search_rect = fitz.Rect(inst.x0 - 50, inst.y0 - 150, inst.x1 + 50, inst.y0)
    
    # Renderizar alta calidad
    pix = page.get_pixmap(matrix=fitz.Matrix(4, 4), clip=search_rect, alpha=True)
    img = Image.frombytes("RGBA", [pix.width, pix.height], pix.samples)
    
    # Filtrado azul
    data = np.array(img)
    r, g, b, a = data[:,:,0], data[:,:,1], data[:,:,2], data[:,:,3]
    blue_mask = (b > r + 30) & (b > g + 30)
    a[~blue_mask] = 0
    data[:,:,3] = a
    
    filtered_img = Image.fromarray(data, 'RGBA')
    
    # Recorte milimétrico
    bbox = filtered_img.getbbox()
    if bbox:
        # Recortar al tamaño exacto de la firma
        final = filtered_img.crop(bbox)
        final.save(output_path)
        print(f"Success! Exact signature saved to {output_path} (Size: {final.size})")
        doc.close()
        return True
    else:
        print(f"No ink found near anchor in {pdf_path}")
        doc.close()
        return False

if __name__ == "__main__":
    out_dir = r"c:\Users\rjiso\OneDrive\Escritorio\PROCESAR ARCHIVO SUELDOS\02_CARPETAS\test_firmas\anchored"
    os.makedirs(out_dir, exist_ok=True)
    
    # Aguilar
    pdf1 = r"G:\Mi unidad\PERSONAL Y VEHICULOS\ARCHIVOS DEL PERSONAL\EXAMENES PREOCUPACIONALES\AGUILAR IVAN PREOCUP.pdf"
    anchored_extraction(pdf1, os.path.join(out_dir, "firma_aguilar_anchored.png"), 2, "Firma del Trabajador")
    
    # Andrade
    pdf2 = r"G:\Mi unidad\PERSONAL Y VEHICULOS\ARCHIVOS DEL PERSONAL\EXAMENES PREOCUPACIONALES\ANDRADE DIEGO PREOCUP.pdf"
    anchored_extraction(pdf2, os.path.join(out_dir, "firma_andrade_anchored.png"), 2, "Firma del Trabajador")
