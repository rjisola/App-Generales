# -*- coding: utf-8 -*-
import smtplib
import os
import time
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication

class GmailBatchSender:
    def __init__(self, smtp_user, smtp_password):
        self.smtp_user = smtp_user
        self.smtp_password = smtp_password.replace(" ", "")

    def connect(self):
        """Simulamos la conexión inicial para mantener compatibilidad."""
        return True, "Listo"

    def disconnect(self):
        """Simulamos la desconexión."""
        pass

    def send_one(self, recipient, subject, body, attachment_paths):
        """Abre una conexión fresca, envía el correo y la cierra inmediatamente. Tiene reintentos."""
        max_retries = 2
        
        for attempt in range(max_retries):
            try:
                # 1. Conexión fresca. TIMEOUT AUMENTADO A 120 SEGS PARA ARCHIVOS PESADOS
                server = smtplib.SMTP("smtp.gmail.com", 587, timeout=120)
                server.ehlo()
                server.starttls()
                server.ehlo()
                server.login(self.smtp_user, self.smtp_password)

                # 2. Construir el mensaje
                msg = MIMEMultipart()
                msg['From'] = self.smtp_user
                msg['To'] = recipient
                msg['Subject'] = subject
                msg.attach(MIMEText(body, 'plain'))

                # 3. Adjuntar archivos
                for path in attachment_paths:
                    if os.path.exists(path):
                        with open(path, "rb") as f:
                            part = MIMEApplication(f.read(), Name=os.path.basename(path))
                        part['Content-Disposition'] = f'attachment; filename="{os.path.basename(path)}"'
                        msg.attach(part)
                
                # 4. Enviar y desconectar de forma segura
                server.send_message(msg)
                server.quit()
                
                return True, "Enviado"
                
            except smtplib.SMTPServerDisconnected as e:
                # Si el servidor corta, esperamos 3 segundos e intentamos de nuevo
                if attempt < max_retries - 1:
                    time.sleep(3)
                    continue
                return False, f"El servidor cerró la conexión. ¿El archivo pesa más de 25MB? (Límite Gmail)"
            except Exception as e:
                # Cualquier otro error (ej. contraseña mal) no lo reintenta, corta de una
                return False, f"Error en envío: {str(e)}"
        
        return False, "Fallaron todos los intentos de envío."

def send_email(smtp_user, smtp_password, recipient, subject, body, attachment_paths):
    """Función de compatibilidad para envíos simples."""
    sender = GmailBatchSender(smtp_user, smtp_password)
    return sender.send_one(recipient, subject, body, attachment_paths)