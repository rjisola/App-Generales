import fitz
import os

pdf_path = r"C:\Users\rjiso\OneDrive\Escritorio\2DA MARZO 2026\recibos-ordenados.pdf"

if not os.path.exists(pdf_path):
    print(f"File not found: {pdf_path}")
else:
    doc = fitz.open(pdf_path)
    # Analizar los primeros 2 recibos
    full_text = ""
    for i in range(min(5, len(doc))):
        full_text += doc[i].get_text()
    
    doc.close()
    
    parts = full_text.split("FIRMA DEL EMPLEADO")
    for i, part in enumerate(parts[:3]):
        print(f"\n--- BLOQUE {i} ---")
        lines = [l.strip() for l in part.strip().split('\n') if l.strip()]
        for j, line in enumerate(lines):
            print(f"{j}: {line}")
