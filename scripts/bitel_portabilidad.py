import pandas as pd
import pyautogui as pg
import webbrowser as web
import time

# ============================================
# RPA - Contacto masivo por WhatsApp
# Operador: BITEL - Internet Móvil 4G/5G
# ============================================

data = pd.read_excel("../data/base_sample.xlsx")
print(f"Base cargada: {len(data)} contactos")
chrome_path = r'C:\Program Files\Google\Chrome\Application\chrome.exe'

mensaje = "Buenos días, tenemos una promoción para ti. 🥳%0A%0A"\
    "💛💚 Cámbiate a Bitel con Plan Ilimitado 69.90 y paga sólo el 50% 🤩 es decir S/34.90 🤩 por 12 meses.%0A%0A"\
    "✅📲Recibe 110GB en ALTA VELOCIDAD ➕ ILIMITADO para navegar en lo que quieras.🚀%0A"\
    "✅ Acceso a Paramount versión full y TV 360 ¡Totalmente GRATIS! 🤩 por 6 meses.%0A%0A"\
    "📍 Sólo pagas S/34.90 al recibir tu chip, que cubre 1 mes de adelanto.%0A%0A"\
    "🛵 El delivery es gratis, entregas a nivel nacional 🇵🇪. 🛵%0A"\
    "¡Promo exclusiva online whatsapp! 🤩%0A%0A"\
    "✅ Si deseas que podamos evaluarte escribe *'Si'*."

web.register('chrome', None, web.BackgroundBrowser(chrome_path))

for i in range(len(data)):
    celular = data.loc[i, 'Numero'].astype(str)  # Convertir número a string   
    web.get('chrome').open(f"https://web.whatsapp.com/send?phone={celular}&text={mensaje}")

    time.sleep(10)
    pg.click(669, 1001)

    time.sleep(1)
    pg.click(678,573)
    time.sleep(2)
    pg.click(216, 171)
    time.sleep(1)
    pg.press('enter')
    time.sleep(2)
    pg.press('enter')

    time.sleep(6)
    pg.hotkey('ctrl', 'w')
    time.sleep(1)