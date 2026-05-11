
import fitz # PyMuPDF
import os

input_path = r"C:\Users\rjiso\OneDrive\Escritorio\Recibos_Diciembre2025-1eraEnero2026.pdf"
output_path = r"C:\Users\rjiso\OneDrive\Escritorio\Recibos_Diciembre2025-1eraEnero2026_reducido_v2.pdf"

try:
    initial_size = os.path.getsize(input_path) / (1024*1024)
    print(f"Tamaño inicial: {initial_size:.2f} MB")
    
    doc = fitz.open(input_path)
    # Optimización: eliminar duplicados, comprimir streams, eliminar basura
    doc.save(
        output_path, 
        garbage=4, 
        deflate=True, 
        clean=True
    )
    doc.close()
        
    final_size = os.path.getsize(output_path) / (1024*1024)
    reduction = ((initial_size - final_size) / initial_size) * 100
    
    print(f"Tamaño final: {final_size:.2f} MB")
    print(f"Reducción: {reduction:.1f}%")
    print(f"Archivo guardado en: {output_path}")

except Exception as e:
    print(f"Error: {e}")
