import pandas as pd
import pyautogui as pg
import webbrowser as web
import time

# ============================================
# RPA - Contacto masivo por WhatsApp
# Operador: WOW - Internet Hogar Fibra Óptica
# ============================================

data = pd.read_excel("../data/base_sample.xlsx")
print(f"Base cargada: {len(data)} contactos")
chrome_path = r'C:\Program Files\Google\Chrome\Application\chrome.exe'

mensaje = "*¡APROVECHA SUPER OFERTAS WOW 💜! 🥳*%0A"\
          "Le traemos la nueva oferta de Internet WOW 100% Fibra Óptica Hogar%0A%0A"\
          "✅ Incluye DIRECT TV GO ilimitado%0A"\
          "✅ CONTRATO 6 MESES%0A"\
          "🏡 Instalación ¡Totalmente GRATIS! 🤩%0A%0A"\
          "*✅ Si deseas que podamos evaluarte escribe 'Si*"

web.register('chrome', None, web.BackgroundBrowser(chrome_path))

for i in range(len(data)):
      celular = data.loc[i, 'Numero'].astype(str)  # Convertir número a string
      web.get('chrome').open(f"https://web.whatsapp.com/send?phone={celular}&text={mensaje}")

      time.sleep(10)
      pg.click(583, 821)

      time.sleep(1)
      pg.click(534,392)
      time.sleep(2)
      pg.click(209, 169)
      time.sleep(1)
      pg.press('enter')
      time.sleep(2)
      pg.press('enter')

      time.sleep(4)
      pg.hotkey('ctrl', 'w')
      time.sleep(1)
