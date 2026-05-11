import warnings
warnings.filterwarnings('ignore')
import google.generativeai as g
import json

with open(r'c:\Users\rjiso\OneDrive\Escritorio\PROCESAR ARCHIVO SUELDOS\03_OTROS\config_lector.json') as f:
    cfg = json.load(f)

g.configure(api_key=cfg['api_key'])

try:
    r = g.embed_content(
        model='models/gemini-embedding-001',
        content='Texto de prueba para el lector inteligente',
        task_type='RETRIEVAL_DOCUMENT'
    )
    print(f"SUCCESS - Embedding generado: {len(r['embedding'])} dimensiones")
except Exception as e:
    print(f"ERROR: {e}")
