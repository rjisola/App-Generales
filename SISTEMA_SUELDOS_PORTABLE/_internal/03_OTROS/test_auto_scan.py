import fitz
from PIL import Image
import numpy as np
import os

def auto_extract(pdf_path, output_folder, anchors):
    doc = fitz.open(pdf_path)
    target_page = -1
    target_rect = None
    
    # Scan first 10 pages
    for p_idx in range(min(10, len(doc))):
        page = doc.load_page(p_idx)
        for anchor in anchors:
            res = page.search_for(anchor)
            if res:
                target_page = p_idx
                inst = res[0]
                # Area above the text
                target_rect = fitz.Rect(inst.x0 - 50, inst.y0 - 150, inst.x1 + 50, inst.y0)
                break
        if target_page != -1: break

    if target_page == -1:
        print(f"No anchor found in {pdf_path}")
        doc.close()
        return False

    page = doc.load_page(target_page)
    pix = page.get_pixmap(matrix=fitz.Matrix(4, 4), clip=target_rect, alpha=True)
    img = Image.frombytes("RGBA", [pix.width, pix.height], pix.samples)
    
    data = np.array(img)
    r, g, b, a = data[:,:,0], data[:,:,1], data[:,:,2], data[:,:,3]
    white_mask = (r > 240) & (g > 240) & (b > 240)
    a[white_mask] = 0
    blue_mask = (b > r + 30) & (b > g + 30)
    a[~blue_mask] = 0
    data[:,:,3] = a
    
    f = Image.fromarray(data, 'RGBA')
    bbox = f.getbbox()
    if bbox:
        base = os.path.splitext(os.path.basename(pdf_path))[0]
        f.crop(bbox).save(os.path.join(output_folder, f"AutoFirma_{base}.png"))
        print(f"Success! {base} (Page {target_page+1})")
        doc.close()
        return True
    else:
        print(f"No ink found near anchor in {pdf_path}")
        doc.close()
        return False

if __name__ == "__main__":
    out_dir = r"c:\Users\rjiso\OneDrive\Escritorio\PROCESAR ARCHIVO SUELDOS\02_CARPETAS\test_firmas\auto_scan"
    os.makedirs(out_dir, exist_ok=True)
    anchors = ["Firma del Trabajador", "Firma del Empleado"]
    
    pdfs = [
        r"G:\Mi unidad\PERSONAL Y VEHICULOS\ARCHIVOS DEL PERSONAL\EXAMENES PREOCUPACIONALES\ARAUJO WALTER PREOCUP.pdf",
        r"G:\Mi unidad\PERSONAL Y VEHICULOS\ARCHIVOS DEL PERSONAL\EXAMENES PREOCUPACIONALES\BARRIOS DUGLAS RAFAEL PREOCUP.pdf"
    ]
    
    for p in pdfs:
        auto_extract(p, out_dir, anchors)
