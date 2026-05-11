import os
import sys
import threading
import customtkinter as ctk
from tkinter import filedialog, messagebox
from google import genai
from google.genai import types
import google.generativeai as genai_legacy  # Para embeddings (compatible con API Key)
import numpy as np
import PyPDF2
import docx
import pandas as pd
import json
import warnings
warnings.filterwarnings("ignore", category=FutureWarning)

# Importar módulo de posicionamiento compartido
try:
    script_dir = os.path.dirname(os.path.abspath(__file__))
    if script_dir not in sys.path:
        sys.path.insert(0, script_dir)
    
    # También agregar 03_OTROS (donde se encuentran los submódulos del sistema)
    others_dir = os.path.abspath(os.path.join(script_dir, "..", "03_OTROS"))
    if others_dir not in sys.path:
        sys.path.insert(0, others_dir)
        
    import modern_gui_components as mgc
    from icon_loader import set_window_icon, load_icon
    _has_mgc = True
except Exception:
    _has_mgc = False


# Configuración de apariencia (Heredada de mgc)
ctk.set_appearance_mode("Dark")
ctk.set_default_color_theme("blue")

# Configuración segura de Gemini API
CONFIG_FILE = "config_lector.json"

def load_config():
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r") as f:
                return json.load(f)
        except:
            return {}
    return {}

def save_config(api_key):
    with open(CONFIG_FILE, "w") as f:
        json.dump({"api_key": api_key}, f)

config = load_config()
API_KEY = config.get("api_key", os.environ.get("GEMINI_API_KEY", "")).strip()
client = None
if API_KEY:
    try:
        client = genai.Client(api_key=API_KEY)
        genai_legacy.configure(api_key=API_KEY)  # Para embeddings
    except Exception as e:
        print(f"Error inicializando cliente: {e}")

# Memoria de nuestra BD Vectorial
document_chunks = []
document_embeddings = []

def extract_text(filepath):
    """Extrae el contenido de texto dependiendo del tipo de archivo."""
    ext = filepath.lower()
    text = ""
    if ext.endswith(".txt"):
        with open(filepath, "r", encoding="utf-8", errors="ignore") as f:
            text = f.read()
    elif ext.endswith(".pdf"):
        with open(filepath, "rb") as f:
            reader = PyPDF2.PdfReader(f)
            for page in reader.pages:
                extracted = page.extract_text()
                if extracted:
                    text += extracted + "\n"
    elif ext.endswith(".docx"):
        doc = docx.Document(filepath)
        for para in doc.paragraphs:
            text += para.text + "\n"
    elif ext.endswith((".xlsx", ".xls")):
        df = pd.read_excel(filepath)
        text = df.to_string() 
    else:
        raise ValueError("Formato de archivo no soportado.")
    return text.strip()

def create_chunks(text, max_chars=800):
    """Segmenta el texto en fragmentos manejables."""
    words = text.split()
    chunks = []
    current_chunk = []
    current_len = 0
    
    for word in words:
        current_chunk.append(word)
        current_len += len(word) + 1
        if current_len >= max_chars:
            chunks.append(" ".join(current_chunk))
            current_chunk = []
            current_len = 0
            
    if current_chunk:
        chunks.append(" ".join(current_chunk))
    return chunks

def get_embedding(text):
    """Obtiene el vector numérico desde Gemini (usando SDK compatible)."""
    if not client: return None
    result = genai_legacy.embed_content(
        model="models/gemini-embedding-001",
        content=text,
        task_type="RETRIEVAL_DOCUMENT"
    )
    return result["embedding"]

def cosine_sim(a, b):
    """Calcula la similitud matemática entre dos vectores."""
    a, b = np.array(a), np.array(b)
    return np.dot(a, b) / (np.linalg.norm(a) * np.linalg.norm(b))

class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.title("Lector Inteligente - Gemini AI")
        self.configure(fg_color=mgc.COLORS['bg_primary'])

        # Posicionar (900x700 estándar)
        if _has_mgc:
            mgc.center_window(self, 900, 700)
        else:
            self.geometry("900x700")

        self.resizable(True, True)

        # Establecer icono de ventana
        set_window_icon(self, 'search')
        
        # Cargar iconos PNG
        self.icon_brain = load_icon('search', (64, 64)) 
        self.icon_folder = load_icon('folder', (24, 24))
        self.icon_key = load_icon('settings', (24, 24))
        self.icon_send = load_icon('check', (24, 24))

        # Layout Principal con Scroll
        self.main_container = ctk.CTkScrollableFrame(self, fg_color="transparent", corner_radius=0)
        self.main_container.pack(fill="both", expand=True, padx=30, pady=20)

        # Header
        mgc.create_header(self.main_container, "Lector Inteligente", 
                         "Carga documentos y haz preguntas usando Inteligencia Artificial", 
                         icon_image=self.icon_brain)

        # SECCIÓN SUPERIOR: CONFIGURACIÓN Y CARGA
        self.config_frame = ctk.CTkFrame(self.main_container, fg_color="transparent")
        self.config_frame.pack(fill="x", pady=(0, 15))
        
        # Card 1: API KEY
        card_api_outer, card_api_inner = mgc.create_card(self.config_frame, "1. Configuración de API", padding=15)
        card_api_outer.pack(side="left", fill="both", expand=True, padx=(0, 10))
        
        self.api_entry = ctk.CTkEntry(card_api_inner, placeholder_text="Gemini API Key...", show="*", height=32)
        self.api_entry.pack(fill="x", pady=(0, 10))
        if API_KEY: self.api_entry.insert(0, API_KEY)
        
        self.btn_save_api = mgc.create_button(card_api_inner, "Activar API", self.guardar_api_key, 
                                             color='blue', icon_image=self.icon_key)
        self.btn_save_api.pack(fill="x")

        # Card 2: CARGA DE DOCUMENTOS
        card_load_outer, card_load_inner = mgc.create_card(self.config_frame, "2. Base de Conocimiento", padding=15)
        card_load_outer.pack(side="left", fill="both", expand=True)

        self.lbl_status_files = ctk.CTkLabel(card_load_inner, text="No hay archivos cargados", 
                                            font=mgc.FONTS['small'], text_color=mgc.COLORS['text_secondary'])
        self.lbl_status_files.pack(pady=(0, 10))

        self.btn_cargar = mgc.create_large_button(card_load_inner, "CARGAR ARCHIVOS", self.cargar_archivo, 
                                                color='green', icon_image=self.icon_folder)
        self.btn_cargar.pack(fill="x")

        # Barra de progreso integrada (oculta por defecto)
        self.progress_container, self.progress_bar, self.progress_label_var = mgc.create_progress_section(self.main_container)

        # SECCIÓN DE CHAT
        card_chat_outer, card_chat_inner = mgc.create_card(self.main_container, "3. Consultas al Documento", padding=15)
        card_chat_outer.pack(fill="both", expand=True, pady=(0, 10))

        self.chat_box = ctk.CTkTextbox(card_chat_inner, font=mgc.FONTS['normal'], 
                                      fg_color="#f8fafc", text_color=mgc.COLORS['text_primary'],
                                      border_color=mgc.COLORS['border'], border_width=1)
        self.chat_box.pack(fill="both", expand=True, pady=(0, 10))
        self.chat_box.insert("0.0", "Bienvenido. Cargue un documento para comenzar a preguntar...")
        self.chat_box.configure(state="disabled")

        # Input de pregunta
        self.input_frame = ctk.CTkFrame(card_chat_inner, fg_color="transparent")
        self.input_frame.pack(fill="x")

        self.entry_pregunta = ctk.CTkEntry(self.input_frame, placeholder_text="Escriba aquí su pregunta sobre el archivo...", 
                                         height=40, font=mgc.FONTS['normal'])
        self.entry_pregunta.pack(side="left", fill="x", expand=True, padx=(0, 10))
        self.entry_pregunta.bind("<Return>", lambda e: self.buscar_respuesta())

        self.btn_buscar = mgc.create_button(self.input_frame, "Preguntar", self.buscar_respuesta, 
                                           color='blue', icon_image=self.icon_send, width=120, height=40)
        self.btn_buscar.pack(side="right")
        self.btn_buscar.configure(state="disabled")

        # Barra de estado
        self.status_frame, self.status_var = mgc.create_status_bar(self)

    def change_appearance_mode_event(self, new_appearance_mode: str):
        ctk.set_appearance_mode(new_appearance_mode)

    def guardar_api_key(self):
        new_key = self.api_entry.get().strip()
        if not new_key:
            messagebox.showwarning("Atención", "Por favor ingresa una clave válida.")
            return
        
        save_config(new_key)
        global API_KEY, client
        API_KEY = new_key
        try:
            client = genai.Client(api_key=API_KEY)
            genai_legacy.configure(api_key=API_KEY)
            self.status_var.set("✓ API Key activada correctamente")
            messagebox.showinfo("Éxito", "API Key guardada y activada correctamente.")
        except Exception as e:
            messagebox.showerror("Error", f"La clave parece ser inválida: {e}")

    def append_chat(self, msg, tag=None):
        self.chat_box.configure(state="normal")
        self.chat_box.insert("end", f"\n{msg}\n")
        self.chat_box.see("end")
        self.chat_box.configure(state="disabled")

    def cargar_archivo(self):
        if not API_KEY:
            messagebox.showerror("Error", "No se detectó la API Key en el sistema.")
            return

        files = filedialog.askopenfilenames(title="Seleccionar Documentos", 
                                            filetypes=[("Documentos", "*.pdf *.docx *.xlsx *.txt")])
        if not files: return

        # Ejecutar en un hilo separado para no congelar la UI
        threading.Thread(target=self._procesar_archivos, args=(files,), daemon=True).start()

    def _procesar_archivos(self, filepaths):
        global document_chunks, document_embeddings
        document_chunks, document_embeddings = [], []
        
        self.append_chat(f"\n[SISTEMA] Iniciando procesamiento de {len(filepaths)} archivo(s)...")
        # Mostrar barra de progreso
        self.progress_container.pack(fill="x", pady=10, before=self.chat_box.master)
        self.progress_bar['value'] = 0
        self.progress_label_var.set("Extrayendo texto...")
        
        try:
            full_text = ""
            for i, fp in enumerate(filepaths):
                nombre = os.path.basename(fp)
                self.append_chat(f"  > Leyendo: {nombre}")
                txt = extract_text(fp)
                if txt: full_text += f"\n[FUENTE: {nombre}]\n{txt}\n"
                val = int(((i+1) / (len(filepaths) * 2)) * 100)
                self.progress_bar['value'] = val

            if not full_text.strip():
                self.append_chat("[ERROR] No se pudo extraer texto utilizable.")
                return

            self.append_chat("[SISTEMA] Segmentando texto y generando vectores matemáticos...")
            self.progress_label_var.set("Generando Embeddings...")
            document_chunks = create_chunks(full_text)
            
            total = len(document_chunks)
            for i, chunk in enumerate(document_chunks):
                emb = get_embedding(chunk)
                document_embeddings.append(emb)
                val = 50 + int(((i+1) / (total * 2)) * 100)
                self.progress_bar['value'] = val
            
            self.append_chat(f"[*] ¡Éxito! Base de datos vectorial lista ({len(document_embeddings)} fragmentos).")
            self.lbl_status_files.configure(text=f"✓ {len(filepaths)} archivos cargados", text_color=mgc.COLORS['green'])
            self.btn_buscar.configure(state="normal")
            self.status_var.set(f"✓ {len(filepaths)} documentos listos para consulta")
            
        except Exception as e:
            self.append_chat(f"[FALLO] Error durante la carga: {e}")
            self.status_var.set("❌ Error en la carga")
        finally:
            self.progress_container.pack_forget()

    def buscar_respuesta(self):
        pregunta = self.entry_pregunta.get().strip()
        if not pregunta or not document_embeddings: return
        
        self.append_chat(f"\n[TU] >>> {pregunta}")
        self.entry_pregunta.delete(0, "end")
        
        try:
            q_emb = get_embedding(pregunta)
            best_score, best_chunk = -1.0, ""
            
            for i, doc_emb in enumerate(document_embeddings):
                score = cosine_sim(q_emb, doc_emb)
                if score > best_score:
                    best_score, best_chunk = score, document_chunks[i]
            
            confianza = int(best_score * 100)
            
            self.append_chat(f"[IA - Confianza: {confianza}%] Generando respuesta final...")
            
            # RAG: Generar respuesta con el contexto
            prompt = f"""
            Usa el siguiente contexto para responder la pregunta del usuario de forma profesional y precisa.
            Si la información no está en el contexto, indícalo amablemente.
            
            CONTEXTO:
            {best_chunk}
            
            PREGUNTA: {pregunta}
            """
            
            response = client.models.generate_content(
                model="gemini-2.0-flash",
                contents=prompt
            )
            
            self.append_chat("-" * 50)
            self.append_chat(response.text)
            self.append_chat("-" * 50)
            
        except Exception as e:
            self.append_chat(f"[ERROR] No se pudo procesar la búsqueda: {e}")

if __name__ == "__main__":
    app = App()
    app.mainloop()

