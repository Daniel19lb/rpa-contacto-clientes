import pandas as pd
import pyautogui as pg
import webbrowser as web
import time

# ============================================
# RPA - Contacto masivo por WhatsApp
# Operador: CLARO - Internet Hogar
# ============================================

BASE_PATH = "../data/base_sample.xlsx"
CHROME_PATH = r'C:\Program Files\Google\Chrome\Application\chrome.exe'

MENSAJE = (
    "Tenemos esta OFERTA EXCLUSIVA para ti%0A"
    "Cambiate a CLARO Internet Hogar con plan de S/65%0A%0A"
    "Incluye CLARO VIDEO + PARAMOUNT por 48 meses sin costo adicional.%0A"
    "150Mbps de velocidad.%0A"
    "Instalacion Totalmente GRATIS!%0A%0A"
    "Si deseas que podamos evaluarte escribe 'Si'."
)

data = pd.read_excel(BASE_PATH)
print(f"Base cargada: {len(data)} contactos")

web.register('chrome', None, web.BackgroundBrowser(CHROME_PATH))

for _, row in data.iterrows():
    celular = str(row['Numero'])
    web.get('chrome').open(f"https://web.whatsapp.com/send?phone={celular}&text={MENSAJE}")
    time.sleep(10)
    pg.click(583, 821)
    time.sleep(1)
    pg.click(534, 392)
    time.sleep(2)
    pg.click(209, 169)
    time.sleep(1)
    pg.press('enter')
    time.sleep(2)
    pg.press('enter')
    time.sleep(4)
    pg.hotkey('ctrl', 'w')
    time.sleep(1)

print("Envio completado.")