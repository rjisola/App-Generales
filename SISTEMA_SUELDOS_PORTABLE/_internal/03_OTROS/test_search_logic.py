
import unicodedata
import re

def normalize_text(text):
    if not isinstance(text, str):
        text = str(text)
    text = ''.join(
        c for c in unicodedata.normalize('NFD', text)
        if unicodedata.category(c) != 'Mn'
    )
    text = re.sub(r'[^\w\s]', ' ', text)
    return ' '.join(text.lower().split())

def get_search_variations(full_name_norm):
    parts = full_name_norm.split()
    variations = [full_name_norm]
    if len(parts) >= 2:
        variations.append(f"{parts[0]} {parts[1]}")
        variations.append(f"{parts[1]} {parts[0]}")
        if len(parts) >= 3:
            variations.append(f"{' '.join(parts[1:])} {parts[0]}")
    return list(set(variations))

def test_search_logic(nombre_excel, texto_pdf):
    nombre_norm = normalize_text(nombre_excel)
    nombre_agresivo = nombre_norm.replace(" ", "")
    parts = nombre_norm.split()
    variations = get_search_variations(nombre_norm)
    
    texto_norm = normalize_text(texto_pdf)
    
    print(f"Buscando: '{nombre_excel}'")
    print(f"Variaciones: {variations}")
    print(f"Texto PDF Norm: '{texto_norm}'")
    
    # Pasada 1
    found1 = False
    for v in variations:
        if v in texto_norm:
            found1 = True
            break
    if found1: return "Pasada 1 (Exacta/Variación)"
    
    # Pasada 2
    if len(parts) >= 2:
        if all(p in texto_norm for p in parts[:2]):
            return "Pasada 2 (Partes)"
            
    # Pasada 3
    if nombre_agresivo in texto_norm.replace(" ", ""):
        return "Pasada 3 (Agresiva - Sin espacios)"
        
    return "No encontrado"

# Casos de prueba
casos = [
    ("PEREZ JUAN ALBERTO", "Recibo de sueldo de PEREZ JUAN ALBERTO"),
    ("PEREZ JUAN ALBERTO", "Firmado por JUAN PEREZ"),
    ("PEREZ JUAN ALBERTO", "P E R E Z  J U A N"),
    ("GOMEZ MARIA", "RECIBO MARIA GOMEZ"),
    ("LOPEZ CARLOS", "Carlos Lopez firma aqui")
]

for name, text in casos:
    result = test_search_logic(name, text)
    print(f"RESULTADO: {result}\n")
