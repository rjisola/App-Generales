# -*- coding: utf-8 -*-
import tkinter as tk
from tkinter import filedialog, messagebox
import os
import sys
import json
import threading
import time

# Asegurar que el directorio de 03_OTROS esté en sys.path
# Asegurar que el directorio del script esté en sys.path para imports locales
script_dir = os.path.dirname(os.path.abspath(__file__))
if script_dir not in sys.path:
    sys.path.insert(0, script_dir)

# Agregar 03_OTROS
root_dir = os.path.dirname(script_dir)
others_dir = os.path.join(root_dir, "03_OTROS")
if others_dir not in sys.path:
    sys.path.insert(0, others_dir)

import modern_gui_components as mgc
import customtkinter as ctk
import logic_email
from icon_loader import set_window_icon, load_icon

CONFIG_FILE = os.path.join(others_dir, "config_email.json")

class EmailApp:
    def __init__(self, root):
        self.root = root
        self.root.title("📧 Envío Secuencial de Documentación")
        self.root.geometry("900x700")
        
        # Cargar configuración
        self.config = self.load_config()
        
        # Variables de la UI
        self.recipient_var = tk.StringVar(value=self.config.get("last_recipient", ""))
        self.subject_base_var = tk.StringVar(value=self.config.get("last_subject", "Sueldos"))
        self.folder_path_var = tk.StringVar(value="No seleccionada")
        self.folder_path = ""
        
        # Estilo Moderno
        mgc.center_window(self.root, 900, 700)
        set_window_icon(self.root, 'email')
        self.icon_email = load_icon('email', (64, 64))
        
        # Contenedor principal con scroll
        main_frame = mgc.create_main_container(self.root)
        
        # Header
        mgc.create_header(main_frame, "Envío Secuencial", "Envía archivos ZIP uno por uno con numeración", icon_image=self.icon_email)
        
        # SECCIÓN 1: CONFIGURACIÓN (Compacta)
        form_row = ctk.CTkFrame(main_frame, fg_color="transparent")
        form_row.pack(fill=tk.X, pady=(0, 10))

        # Card: Datos Destinatario
        card_dest_outer, card_dest_inner = mgc.create_card(form_row, "📩 Destinatario")
        card_dest_outer.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 10))
        self.create_label_entry(card_dest_inner, "Enviar a:", self.recipient_var).pack(fill=tk.X, pady=2)
        self.create_label_entry(card_dest_inner, "Asunto:", self.subject_base_var).pack(fill=tk.X, pady=2)
        
        # Card: Carpeta
        card_folder_outer, card_folder_inner = mgc.create_card(form_row, "📂 Archivos ZIP")
        card_folder_outer.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        mgc.create_button(card_folder_inner, "Seleccionar Carpeta", self.select_folder, color='purple', icon="📁").pack(pady=2)
        ctk.CTkLabel(card_folder_inner, textvariable=self.folder_path_var, font=mgc.FONTS['small'], text_color=mgc.COLORS['text_secondary']).pack()
        
        # SECCIÓN 2: CONTENIDO DE LOS CORREOS
        card_msg_outer, card_msg_inner = mgc.create_card(main_frame, "📝 Cuerpo del Mensaje", padding=15)
        card_msg_outer.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        msg_grid = ctk.CTkFrame(card_msg_inner, fg_color="transparent")
        msg_grid.pack(fill=tk.BOTH, expand=True)

        # Detalle arriba
        ctk.CTkLabel(msg_grid, text="Detalle para el PRIMER correo:", font=mgc.FONTS['small']).pack(anchor='w')
        self.detail_text = ctk.CTkTextbox(msg_grid, height=50, font=mgc.FONTS['normal'], border_width=1)
        self.detail_text.pack(fill=tk.X, pady=(2, 8))
        self.detail_text.insert("0.0", self.config.get("last_detail", ""))
        
        # Mensaje debajo
        ctk.CTkLabel(msg_grid, text="Mensaje General:", font=mgc.FONTS['small']).pack(anchor='w')
        self.message_text = ctk.CTkTextbox(msg_grid, height=80, font=mgc.FONTS['normal'], border_width=1)
        self.message_text.pack(fill=tk.BOTH, expand=True, pady=(2, 5))
        self.message_text.insert("0.0", self.config.get("last_message", ""))
        
        # SECCIÓN 3: ACCIÓN Y PROGRESO
        bottom_row = ctk.CTkFrame(main_frame, fg_color="transparent")
        bottom_row.pack(fill=tk.X, pady=(0, 5))

        self.send_btn = mgc.create_large_button(bottom_row, "Iniciar Envío Secuencial", self.start_send_thread, color='blue', icon="✈️")
        self.send_btn.pack(pady=(0, 10))
        
        # Barra de progreso
        self.progress_frame, self.progress_bar, self.progress_status = mgc.create_progress_section(bottom_row)
        self.progress_frame.pack(fill=tk.X)
        
        # Barra de estado
        self.status_frame, self.status_var = mgc.create_status_bar(self.root, "✓ Listo")

    def load_config(self):
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except: pass
        return {}

    def save_config(self):
        self.config["last_recipient"] = self.recipient_var.get()
        self.config["last_subject"] = self.subject_base_var.get()
        self.config["last_detail"] = self.detail_text.get("0.0", tk.END).strip()
        self.config["last_message"] = self.message_text.get("0.0", tk.END).strip()
        try:
            with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
                json.dump(self.config, f, indent=4)
            if hasattr(self, 'status_var'):
                self.status_var.set("✓ Configuración guardada")
        except: pass

    def create_label_entry(self, parent, label_text, string_var):
        frame = ctk.CTkFrame(parent, fg_color="transparent")
        ctk.CTkLabel(frame, text=label_text, font=mgc.FONTS['normal'], width=120, anchor="w").pack(side=tk.LEFT, padx=5)
        ctk.CTkEntry(frame, textvariable=string_var, font=mgc.FONTS['normal'], corner_radius=6).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        return frame

    def is_zip_file(self, filename):
        """Detecta .zip o archivos divididos .zip.001, .zip.002, etc."""
        f_lower = filename.lower()
        if f_lower.endswith('.zip'):
            return True
        if '.zip.' in f_lower:
            ext = f_lower.split('.zip.')[-1]
            if ext.isdigit():
                return True
        return False

    def select_folder(self):
        path = filedialog.askdirectory()
        if path:
            self.folder_path = os.path.normpath(path)
            try:
                files = [f for f in os.listdir(self.folder_path) if self.is_zip_file(f)]
                count = len(files)
                
                folder_name = os.path.basename(self.folder_path)
                self.folder_path_var.set(f"{folder_name} ({count} partes detectadas)")
                
                if count > 0:
                    self.status_var.set(f"✓ Seleccionada: {folder_name} — {count} archivos listos.")
                else:
                    self.status_var.set(f"⚠ Advertencia: No se encontraron archivos ZIP.")
                    messagebox.showwarning("Sin archivos", f"No se han detectado archivos compatibles en:\n{self.folder_path}")
            
            except Exception as e:
                self.status_var.set(f"✗ Error al acceder a la carpeta.")
                messagebox.showerror("Error de Acceso", f"No se pudo leer la carpeta:\n{str(e)}")

    def start_send_thread(self):
        if not self.folder_path or not os.path.isdir(self.folder_path):
            messagebox.showerror("Error", "Debe seleccionar una carpeta válida.")
            return
        
        recipient = self.recipient_var.get()
        if not recipient or "@" not in recipient:
            messagebox.showerror("Error", "Ingrese un destinatario válido.")
            return
            
        zip_files = sorted([f for f in os.listdir(self.folder_path) if self.is_zip_file(f)])
        if not zip_files:
            messagebox.showerror("Error", "No hay archivos ZIP.")
            return
            
        # Checkpoint logic
        checkpoint_file = os.path.join(self.folder_path, ".sending_state.json")
        sent_files = []
        if os.path.exists(checkpoint_file):
            try:
                with open(checkpoint_file, 'r', encoding='utf-8') as f:
                    sent_files = json.load(f).get("sent_files", [])
            except: pass
            
            if sent_files:
                if messagebox.askyesno("Reanudar Envío", f"Se detectó un lote interrumpido.\n\nYa se enviaron {len(sent_files)} archivos de {len(zip_files)}.\n¿Desea omitir los enviados y continuar?"):
                    pass
                else:
                    sent_files = [] # Empezar de cero
                    try: os.remove(checkpoint_file)
                    except: pass
        
        self.save_config()
        mgc.disable_button(self.send_btn)
        threading.Thread(target=self.sequential_send_process, args=(zip_files, sent_files, checkpoint_file), daemon=True).start()

    def sequential_send_process(self, zip_files, sent_files, checkpoint_file):
        batch_sender = None
        try:
            total = len(zip_files)
            pendientes = [f for f in zip_files if f not in sent_files]
            
            if not pendientes:
                self.root.after(0, lambda: self.finish_send(True, "Todos los archivos ya fueron enviados anteriormente."))
                try: os.remove(checkpoint_file)
                except: pass
                return

            total = len(zip_files)
            smtp_user = self.config.get("smtp_user")
            smtp_pass = self.config.get("smtp_password")
            recipient = self.recipient_var.get()
            subject_base = self.subject_base_var.get()
            detail = self.detail_text.get("0.0", tk.END).strip()
            general_msg = self.message_text.get("0.0", tk.END).strip()

            batch_sender = logic_email.GmailBatchSender(smtp_user, smtp_pass)
            
            for i, filename in enumerate(pendientes):
                num = sent_files.index(filename) + 1 if filename in sent_files else len(sent_files) + i + 1
                current_subject = f"{subject_base} - Parte {num}/{total}"
                file_path = os.path.join(self.folder_path, filename)
                
                # Construir cuerpo del correo
                body_parts = []
                if num == 1 and detail: body_parts.append(detail + "\n")
                body_parts.append(general_msg)
                if num == total: body_parts.append("\nFinalización de documentación.")
                final_body = "\n".join(body_parts)

                # Actualizar progreso en la UI
                self.root.after(0, lambda n=num, t=total, f=filename: self.update_progress(n, t, f, "Conectando y enviando..."))

                # ENVÍO (La nueva lógica abre y cierra la conexión aquí mismo)
                success, result = batch_sender.send_one(recipient, current_subject, final_body, [file_path])
                
                if not success:
                    self.root.after(0, lambda msg=result: self.finish_send(False, f"Fallo al enviar {filename}:\n{msg}"))
                    return # Salida inmediata del bucle si hay error crítico
                
                # Guardar checkpoint si el envío fue exitoso
                sent_files.append(filename)
                try:
                    with open(checkpoint_file, 'w', encoding='utf-8') as f:
                        json.dump({"sent_files": sent_files}, f)
                except: pass

                # Pausa de seguridad antispam
                if num < total:
                    for s in range(3, 0, -1):
                        self.root.after(0, lambda n=num, t=total, f=filename, sec=s: 
                                         self.progress_status.set(f"Archivo {n}/{t} enviado.\nEsperando {sec}s..."))
                        time.sleep(1)

            # Proceso completado exitosamente
            try: os.remove(checkpoint_file) # Limpiar checkpoint al finalizar 100%
            except: pass
            self.root.after(0, lambda: self.finish_send(True, f"Se enviaron exitosamente los {total} correos."))

        except Exception as e:
            self.root.after(0, lambda msg=str(e): self.finish_send(False, f"Error inesperado: {msg}"))

    def update_progress(self, current, total, filename, sub_message="Procesando..."):
        percent = (current / total) * 100
        self.progress_bar['value'] = percent
        self.progress_status.set(f"Archivo {current}/{total}: {filename}\n({sub_message})")
        self.status_var.set(f"⏳ {sub_message} ({current} de {total})")

    def finish_send(self, success, message):
        mgc.enable_button(self.send_btn)
        if success:
            self.progress_bar['value'] = 100
            self.progress_status.set("¡Todo enviado!")
            self.status_var.set("✓ Proceso completado.")
            messagebox.showinfo("Éxito", message)
        else:
            self.status_var.set("✗ Error en el proceso.")
            messagebox.showerror("Error", message)

if __name__ == "__main__":
    try:
        root = ctk.CTk()
        app = EmailApp(root)
        root.mainloop()
    except Exception as e:
        import traceback
        traceback.print_exc()
        try:
            import ctypes
            ctypes.windll.user32.MessageBoxW(0, f"Error fatal:\n{str(e)}", "Error", 0x10)
        except: pass