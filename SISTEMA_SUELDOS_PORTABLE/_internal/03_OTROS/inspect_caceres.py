import fitz
import os

pdf_path = r"C:\Users\rjiso\OneDrive\Escritorio\2DA MARZO 2026\recibos-ordenados.pdf"

if os.path.exists(pdf_path):
    doc = fitz.open(pdf_path)
    full_text = ""
    # Analizar más páginas para encontrar a Caceres (403)
    for i in range(len(doc)):
        txt = doc[i].get_text()
        if "403" in txt and "CACERES" in txt.upper():
            print(f"\n=== ENCONTRADO EN PAGINA {i} ===")
            blocks = txt.split("FIRMA DEL EMPLEADO")
            for b_idx, b in enumerate(blocks):
                if "403" in b and "CACERES" in b.upper():
                    print(f"--- BLOQUE {b_idx} ---")
                    lines = [l.strip() for l in b.strip().split('\n') if l.strip()]
                    for j, line in enumerate(lines):
                        print(f"{j}: {line}")
            break
    doc.close()
