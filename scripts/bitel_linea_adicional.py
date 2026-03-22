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

mensaje = "Buenos días tenemos una promoción para ti 🎉🥳%0A%0A"\
    "Por ser cliente BITEL 💛 no pierdas la oportunidad de obtener línea adicional con un descuento del 50% para ti y tu familia 🤩 PORTABILIDAD O LÍNEA NUEVA%0A%0A"\
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

    time.sleep(4)
    pg.hotkey('ctrl', 'w')
    time.sleep(1)
