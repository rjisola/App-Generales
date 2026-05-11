# -*- coding: utf-8 -*-
import os

filepath = r"01_APLICACIONES\A-GENERAR_RECIBOS_CONTROL.pyw"

with open(filepath, 'r', encoding='utf-8') as f:
    content = f.read()

# 1. Cargar las 3 imágenes extra en __init__
init_old = "self.icon_main = load_icon('main', (64, 64))"
init_new = """self.icon_main = load_icon('main', (64, 64))
        self.icon_calendar_64 = load_icon('calendar', (64, 64))
        self.icon_chart_64 = load_icon('chart', (64, 64))
        self.icon_receipts_64 = load_icon('receipts', (64, 64))"""
content = content.replace(init_old, init_new)

# 2. Reemplazar la firma
def_old = "def _make_tool_card(self, parent, emoji, title, description, action, color):"
def_new = "def _make_tool_card(self, parent, icon_image, title, description, action, color):"
content = content.replace(def_old, def_new)

# 3. Cambiar Label text=emoji por image=icon_image
label_old = """icon_lbl = tk.Label(card_inner, text=emoji, font=('Segoe UI Emoji', 36),
                            bg=mgc.COLORS['bg_card'])"""
label_new = """icon_lbl = tk.Label(card_inner, image=icon_image, bg=mgc.COLORS['bg_card'])
        icon_lbl.image = icon_image # referenciar memoria"""
content = content.replace(label_old, label_new)

# 3.b Variante de label old si tenía el patch fg='white' (por si acaso no se restauro del todo o hay otra string)
label_old_w = """icon_lbl = tk.Label(card_inner, text=emoji, font=('Segoe UI Emoji', 36),
                            bg=mgc.COLORS['bg_card'], fg='white')"""
content = content.replace(label_old_w, label_new)


# 4. Actualizar Card 1
c1_old = 'emoji="\\U0001f4c5",'
c1_old2 = 'emoji="📅",'
c1_new = 'icon_image=self.icon_calendar_64,'
content = content.replace(c1_old, c1_new).replace(c1_old2, c1_new)

# 5. Actualizar Card 2
c2_old = 'emoji="\\U0001f4ca",'
c2_old2 = 'emoji="📊",'
c2_new = 'icon_image=self.icon_chart_64,'
content = content.replace(c2_old, c2_new).replace(c2_old2, c2_new)

# 6. Actualizar Card 3
c3_old = 'emoji="\\U0001f4c4",'
c3_old2 = 'emoji="📄",'
c3_new = 'icon_image=self.icon_receipts_64,'
content = content.replace(c3_old, c3_new).replace(c3_old2, c3_new)

# 7. IMPORTANTE: En el hover on_enter también se repite el error de bg='eef2ff' si lo usan los elementos de CTK.
# Por suerte el parche de CTkBaseClass lo previene, pero lo más sensato es no tocar los params obsoletos.
# El archivo restaurado YA TIENE config() en custom_tkinter. Con el parche ctk_config_patch de mgc debe funcionar sin crash.

with open(filepath, 'w', encoding='utf-8') as f:
    f.write(content)

print("PARCHE EXITOSO!")
