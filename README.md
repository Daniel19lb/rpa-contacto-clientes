# RPA - Contacto Masivo por WhatsApp

Sistema de automatización RPA (Robotic Process Automation) desarrollado en Python para el envío masivo de mensajes comerciales via WhatsApp Web, orientado a campañas de ventas de servicios de telecomunicaciones.

## ⚙️ ¿Cómo funciona?

1. Lee una base de contactos desde Excel
2. Abre WhatsApp Web en Chrome por cada número
3. Envía el mensaje comercial del operador correspondiente
4. Cierra la pestaña y pasa al siguiente contacto

## 🛠️ Stack Tecnológico

- **Python** → Lenguaje principal
- **Pandas** → Lectura y procesamiento de base de contactos
- **PyAutoGUI** → Automatización de clicks y teclado (RPA)
- **Webbrowser** → Apertura automatizada de WhatsApp Web
- **openpyxl** → Lectura de archivos Excel

## 📁 Estructura del Proyecto
```
rpa-contacto-clientes/
│
├── scripts/
│   ├── claro_hogar.py              → Campaña CLARO Internet Hogar
│   ├── bitel_fibra_hogar.py        → Campaña BITEL Fibra Hogar (monitor principal)
│   ├── bitel_fibra_hogar_anydesk.py → Campaña BITEL Fibra Hogar (monitor reducido)
│   ├── bitel_linea_adicional.py    → Campaña BITEL Línea Adicional
│   ├── bitel_portabilidad.py       → Campaña BITEL Portabilidad (monitor principal)
│   ├── bitel_portabilidad_anydesk.py → Campaña BITEL Portabilidad (monitor reducido)
│   ├── bitel_recontacto_leads.py   → Re-contacto leads con mensaje por plan
│   ├── win_hogar.py                → Campaña WIN Internet Hogar (monitor principal)
│   ├── win_hogar_anydesk.py        → Campaña WIN Internet Hogar (monitor reducido)
│   └── wow_hogar_anydesk.py        → Campaña WOW Internet Hogar (monitor reducido)
│
└── data/
    ├── base_sample.xlsx            → Base de contactos de ejemplo (números ficticios)
    └── base_releads_sample.xlsx    → Base de re-leads con columna Plan (números ficticios)
```

## 🚀 Requisitos
```bash
pip install pandas pyautogui openpyxl
```

- Google Chrome instalado en `C:\Program Files\Google\Chrome\Application\chrome.exe`
- WhatsApp Web activo en Chrome con sesión iniciada

## ⚙️ Configuración antes de ejecutar

En cada script ajusta estas dos constantes según tu entorno:
```python
BASE_PATH = "../data/base_sample.xlsx"  # Ruta a tu base de contactos
CHROME_PATH = r'C:\Program Files\Google\Chrome\Application\chrome.exe'
```

## 📊 Variantes AnyDesk

Los scripts con sufijo `_anydesk` tienen coordenadas de clicks ajustadas para monitores de menor resolución — útil cuando se ejecuta remotamente via AnyDesk.

## ⚠️ Nota

Los números en `data/` son ficticios y solo sirven como referencia de la estructura esperada. El sistema fue desarrollado y utilizado en un entorno real de ventas de telecomunicaciones.

## 👤 Autor

**Herles Daniel Lopez Bancho**
Data Analyst BI | Estudiante Ing. de Sistemas
[LinkedIn](https://linkedin.com/in/daniel-lopez-572680194)
