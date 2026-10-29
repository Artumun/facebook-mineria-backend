import os
from dotenv import load_dotenv
from playwright.sync_api import sync_playwright

load_dotenv(override=True)
email = (os.getenv('FB_EMAIL') or '').strip()
password = (os.getenv('FB_PASSWORD') or '').strip()

with sync_playwright() as p:
    browser = p.chromium.launch(headless=False, args=['--disable-blink-features=AutomationControlled'])
    context = browser.new_context(locale='es-MX', viewport={'width': 1100, 'height': 800})
    page = context.new_page()

    page.goto('https://www.facebook.com/login')
    if email and password:
        page.fill('input[name="email"]', email)
        page.fill('input[name="pass"]', password)
        page.press('input[name="pass"]', 'Enter')

    print("Completa el código o verificación en la ventana de Facebook...")
    print("En cuanto entres al inicio de Facebook, la sesión se guardará sola.")

    # Espera hasta que pases el two_step_verification y aparezca la cookie c_user
    for _ in range(180):
        page.wait_for_timeout(1000)
        if any(c.get('name') == 'c_user' for c in context.cookies()):
            page.wait_for_timeout(3000)
            context.storage_state(path='fb_sesion_activa.json')
            print("¡Listo! Archivo fb_sesion_activa.json guardado con éxito.")
            break

    browser.close()