
import tkinter as tk
from tkinter import filedialog
import PyPDF2
import os
import sys

def main():
    root = tk.Tk()
    root.withdraw()

    print("--- EXTRACTOR DE TEXTO PDF ---")
    print("Seleccione el archivo PDF para analizar su contenido...")
    
    file_path = filedialog.askopenfilename(filetypes=[("PDF", "*.pdf")])
    if not file_path:
        print("No seleccionó archivo.")
        input("Presione ENTER para salir...")
        return

    try:
        reader = PyPDF2.PdfReader(file_path)
        print(f"\nArchivo: {os.path.basename(file_path)}")
        print(f"Páginas: {len(reader.pages)}")
        
        print("\n--- TEXTO DE LA PRIMERA PÁGINA ---\n")
        page_text = reader.pages[0].extract_text()
        print(page_text)
        print("\n----------------------------------\n")
        
        print("Copie el texto de arriba y péguelo en el chat.")
        
    except Exception as e:
        print(f"Error: {e}")

    input("\nPresione ENTER para salir...")

if __name__ == "__main__":
    main()
