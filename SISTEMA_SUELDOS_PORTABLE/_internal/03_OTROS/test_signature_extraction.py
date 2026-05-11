import fitz # PyMuPDF
from PIL import Image
import io
import os

def extract_signatures(pdf_path, output_folder):
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)
    
    doc = fitz.open(pdf_path)
    img_count = 0
    
    for i in range(len(doc)):
        images = doc.get_page_images(i)
        for img_index, img in enumerate(images):
            xref = img[0]
            base_image = doc.extract_image(xref)
            image_bytes = base_image["image"]
            image_ext = base_image["ext"]
            
            # Load into Pillow
            image = Image.open(io.BytesIO(image_bytes))
            
            # Basic background removal (convert white to transparent)
            if image.mode != 'RGBA':
                image = image.convert('RGBA')
            
            datas = image.getdata()
            new_data = []
            for item in datas:
                # If pixel is white or near white, make it transparent
                if item[0] > 220 and item[1] > 220 and item[2] > 220:
                    new_data.append((255, 255, 255, 0))
                else:
                    new_data.append(item)
            
            image.putdata(new_data)
            
            # Save
            output_name = f"signature_{i}_{img_index}.png"
            image.save(os.path.join(output_folder, output_name))
            print(f"Saved: {output_name}")
            img_count += 1
            
    doc.close()
    return img_count

if __name__ == "__main__":
    pdf_file = r"c:\Users\rjiso\OneDrive\Escritorio\PROCESAR ARCHIVO SUELDOS\02_CARPETAS\borrar\Recibos.pdf"
    out_dir = r"c:\Users\rjiso\OneDrive\Escritorio\PROCESAR ARCHIVO SUELDOS\02_CARPETAS\test_firmas"
    count = extract_signatures(pdf_file, out_dir)
    print(f"Total images extracted: {count}")
