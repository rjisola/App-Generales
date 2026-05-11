import fitz
from PIL import Image
import os

def render_page_to_transparent_png(pdf_path, page_num, output_path):
    doc = fitz.open(pdf_path)
    page = doc.load_page(page_num)
    
    # Increase resolution for better quality
    zoom = 3  # 3x zoom
    mat = fitz.Matrix(zoom, zoom)
    pix = page.get_pixmap(matrix=mat, alpha=True)
    
    # Convert to Pillow image
    img = Image.frombytes("RGBA", [pix.width, pix.height], pix.samples)
    
    # Simple background removal for white pixels
    datas = img.getdata()
    new_data = []
    for item in datas:
        # If it's pure white or very close to white, make it transparent
        if item[0] > 240 and item[1] > 240 and item[2] > 240:
            new_data.append((255, 255, 255, 0)) # Fully transparent
        else:
            new_data.append(item)
            
    img.putdata(new_data)
    img.save(output_path)
    doc.close()

if __name__ == "__main__":
    pdf_file = r"c:\Users\rjiso\OneDrive\Escritorio\PROCESAR ARCHIVO SUELDOS\02_CARPETAS\borrar\Recibos.pdf"
    out_file = r"c:\Users\rjiso\OneDrive\Escritorio\PROCESAR ARCHIVO SUELDOS\02_CARPETAS\test_firmas\page_render_transparent.png"
    
    os.makedirs(os.path.dirname(out_file), exist_ok=True)
    render_page_to_transparent_png(pdf_file, 0, out_file)
    print(f"Rendered: {out_file}")
