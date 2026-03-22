import pandas as pd
import pyautogui as pg
import webbrowser as web
import time

# ============================================
# RPA - Contacto masivo por WhatsApp
# Operador: WIN - Internet Hogar Fibra Óptica
# ============================================

data = pd.read_excel("../data/base_sample.xlsx")
print(f"Base cargada: {len(data)} contactos")
CHROME_PATH = r'C:\Program Files\Google\Chrome\Application\chrome.exe'

mensaje = "Hola!, tenemos una promoción para ti. 🥳%0A%0A"\
          "🧡🧡 Cámbiate a Win con Plan 119 y paga S/54.90 🤩 por 3 meses.%0A%0A"\
          "🟠 Solicítalo ya con tu DNI O CARNET DE Extranjería.%0A"\
          "✅ ⚡Velocidad duplicada a 600Mbps por 6 meses Fibra Óptica. 🚀%0A"\
          "🏡 Instalación ¡Totalmente GRATIS! 🤩%0A"\
          "🎉 Podemos validar tu cobertura 🙌%0A%0A"\
          "✅ Si deseas que podamos evaluarte escribe 'Si'."

web.register('chrome', None, web.BackgroundBrowser(CHROME_PATH))

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

