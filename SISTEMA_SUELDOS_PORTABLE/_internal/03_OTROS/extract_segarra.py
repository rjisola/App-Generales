import fitz
from PIL import Image
import numpy as np
import os

def extract_signature_segarra(pdf_path, output_path, page_idx):
    doc = fitz.open(pdf_path)
    page = doc.load_page(page_idx)
    
    # Buscar ancla
    search_res = page.search_for("Firma del Trabajador") or page.search_for("Firma del Empleado")
    if not search_res:
        print("No se encontró el ancla en la página.")
        doc.close()
        return
        
    inst = search_res[0]
    clip = fitz.Rect(inst.x0 - 50, inst.y0 - 150, inst.x1 + 50, inst.y0)
    
    # Render
    pix = page.get_pixmap(matrix=fitz.Matrix(4, 4), clip=clip, alpha=True)
    img = Image.frombytes("RGBA", [pix.width, pix.height], pix.samples)
    
    data = np.array(img)
    r, g, b, a = data[:,:,0], data[:,:,1], data[:,:,2], data[:,:,3]
    
    # Filtro azul
    blue_mask = (b > r + 30) & (b > g + 30)
    a[~blue_mask] = 0
    data[:,:,3] = a
    
    filtered_img = Image.fromarray(data, 'RGBA')
    bbox = filtered_img.getbbox()
    if bbox:
        filtered_img.crop(bbox).save(output_path)
        print(f"Extracción exitosa: {output_path}")
    else:
        print("No se detectó tinta azul en la zona del ancla.")
    doc.close()

if __name__ == "__main__":
    pdf = r"G:\Mi unidad\PERSONAL Y VEHICULOS\ARCHIVOS DEL PERSONAL\EXAMENES PREOCUPACIONALES\SEGARRA RUBEN PREOCUP.pdf"
    out = r"C:\Users\rjiso\.gemini\antigravity\brain\8864ff70-01cc-41bd-9f86-d741ece4b766\resultado_segarra_final.png"
    extract_signature_segarra(pdf, out, 3) # Hoja 4 es índice 3
