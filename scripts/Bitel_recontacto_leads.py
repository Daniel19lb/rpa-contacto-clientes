import pandas as pd
import pyautogui as pg
import webbrowser as web
import time

# ============================================
# RPA - Contacto masivo por WhatsApp
# Operador: BITEL - Re-contacto de leads
# Selecciona mensaje segun plan del cliente
# ============================================

BASE_PATH = "../data/base_releads_sample.xlsx"
CHROME_PATH = r'C:\Program Files\Google\Chrome\Application\chrome.exe'

MENSAJES = {
    "79.9": (
        "Buenas tardes! Te saludamos de Bitel%0A"
        "Hacemos seguimiento a una orden que no pudo ser entregada exitosamente. "
        "Queremos saber el motivo y si le gustaria reagendar el envio de delivery%0A%0A"
        "Plan de Portabilidad Ilimitado S/79.90 de Bitel%0A"
        "Pagando S/39.90 por 125 GB en alta velocidad + ILIMITADO x 12 meses%0A"
        "Llamadas y SMS ILIMITADOS%0A%0A"
        "Pagando solo S/39.90 al momento de recibir el chip (Yape, Plin o efectivo)%0A"
        "DELIVERY TOTALMENTE GRATIS%0A%0A"
        "Si deseas reagendar el pedido escribe 'Si'."
    ),
    "69.9": (
        "Buenas tardes! Te saludamos de Bitel%0A"
        "Hacemos seguimiento a una orden que no pudo ser entregada exitosamente. "
        "Queremos saber el motivo y si le gustaria reagendar el envio de delivery%0A%0A"
        "Plan de Portabilidad Ilimitado S/69.90 de Bitel%0A"
        "Pagando S/34.90 por 110 GB en alta velocidad + ILIMITADO x 12 meses%0A"
        "Llamadas y SMS ILIMITADOS%0A%0A"
        "Pagando solo S/34.90 al momento de recibir el chip (Yape, Plin o efectivo)%0A"
        "DELIVERY TOTALMENTE GRATIS%0A%0A"
        "Si deseas reagendar el pedido escribe 'Si'."
    ),
    "55.9": (
        "Buenas tardes! Te saludamos de Bitel%0A"
        "Plan de Portabilidad Ilimitado S/55.90 de Bitel%0A"
        "Pagando S/27.90 por 60 GB en alta velocidad + ILIMITADO x 6 meses%0A"
        "DELIVERY TOTALMENTE GRATIS%0A%0A"
        "Si deseas reagendar el pedido escribe 'Si'."
    ),
    "49.9": (
        "Buenas tardes! Te saludamos de Bitel%0A"
        "Plan de Portabilidad Ilimitado S/49.90 de Bitel%0A"
        "Pagando S/24.90 por 35 GB en alta velocidad + ILIMITADO x 3 meses%0A"
        "DELIVERY TOTALMENTE GRATIS%0A%0A"
        "Si deseas reagendar el pedido escribe 'Si'."
    ),
    "39.9": (
        "Buenas tardes! Te saludamos de Bitel%0A"
        "Plan de Portabilidad Ilimitado S/39.90 de Bitel%0A"
        "Pagando S/39.90 por 30 GB en alta velocidad%0A"
        "DELIVERY TOTALMENTE GRATIS%0A%0A"
        "Si deseas reagendar el pedido escribe 'Si'."
    ),
    "29.9": (
        "Buenas tardes! Te saludamos de Bitel%0A"
        "Plan de Portabilidad Flash S/29.90 de Bitel%0A"
        "Pagando S/29.90 por 50 GB en alta velocidad x 6 meses%0A"
        "DELIVERY TOTALMENTE GRATIS%0A%0A"
        "Si deseas reagendar el pedido escribe 'Si'."
    )
}

data = pd.read_excel(BASE_PATH)
print(f"Base cargada: {len(data)} contactos")

web.register('chrome', None, web.BackgroundBrowser(CHROME_PATH))

for _, row in data.iterrows():
    celular = str(row['Numero'])
    plan = str(row['Plan'])

    if plan in MENSAJES:
        mensaje = MENSAJES[plan]
        web.get('chrome').open(f"https://web.whatsapp.com/send?phone={celular}&text={mensaje}")
        time.sleep(8)
        pg.press('enter')
        time.sleep(2)
        pg.hotkey('ctrl', 'w')
        time.sleep(1)
    else:
        print(f"Plan {plan} no reconocido para {celular} - omitiendo número")

print("Envio completado.")