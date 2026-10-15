import os
import re
import io
import json
import pymysql
from dotenv import load_dotenv
from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
from playwright.sync_api import sync_playwright
from transformers import pipeline
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill

load_dotenv()

app = Flask(__name__)
CORS(app)

# Inicialización del clasificador NLP
print("Cargando modelo RoBERTa-Sentiment-ES...")
clasificador_ia = pipeline(
    "sentiment-analysis",
    model="pysentimiento/robertuito-sentiment-analysis"
)
print("Modelo cargado.")

MAPA_ETIQUETAS = {
    'POS': 'Positivo',
    'NEG': 'Negativo',
    'NEU': 'Neutro'
}

def get_db_connection():
    return pymysql.connect(
        host=os.getenv('DB_HOST', 'localhost'),
        user=os.getenv('DB_USER', 'root'),
        password=os.getenv('DB_PASSWORD', ''),
        database=os.getenv('DB_NAME', 'mineria_facebook_db'),
        cursorclass=pymysql.cursors.DictCursor
    )

def clasificar_lote_ia(lista_comentarios):
    if not lista_comentarios:
        return []

    textos = [c['comment'] for c in lista_comentarios]
    predicciones = clasificador_ia(
        textos,
        truncation=True,
        max_length=128,
        batch_size=16
    )

    procesados = []
    for item, pred in zip(lista_comentarios, predicciones):
        etiqueta = MAPA_ETIQUETAS.get(pred['label'], 'Neutro')
        procesados.append({
            'username': item['username'][:95],
            'comment': item['comment'],
            'sentiment': etiqueta
        })
    return procesados

def generar_resumen_ejecutivo(porc_pos, porc_neg, porc_neu, total_com):
    if total_com == 0:
        return "No se detectaron comentarios públicos visibles en esta publicación."

    if porc_pos >= porc_neg and porc_pos >= porc_neu:
        tendencia = f"predominantemente positiva ({porc_pos}%)"
        detalle = "El modelo neuronal detecta aceptación favorable hacia el contenido publicado."
    elif porc_neg > porc_pos and porc_neg >= porc_neu:
        tendencia = f"mayoritariamente negativa ({porc_neg}%)"
        detalle = "El modelo neuronal identifica quejas o posturas críticas en las respuestas."
    else:
        tendencia = f"principalmente neutral ({porc_neu}%)"
        detalle = "Las interacciones analizadas son en su mayoría informativas o sin carga emocional."

    return (
        f"Se analizaron {total_com} comentarios mediante el modelo RoBERTa-Sentiment-ES. "
        f"La publicación presenta una postura {tendencia}. {detalle} "
        f"Distribución complementaria: {porc_neg}% negativo y {porc_neu}% neutro."
    )

def parsear_numero_fb(texto_num):
    if not texto_num:
        return 0
    texto_num = str(texto_num).lower().strip()
    multiplicador = 1
    if 'mil' in texto_num or 'k' in texto_num:
        multiplicador = 1000
        texto_num = re.sub(r'(mil|k)', '', texto_num).strip()
    elif 'm' in texto_num:
        multiplicador = 1000000
        texto_num = texto_num.replace('m', '').strip()

    texto_num = texto_num.replace(',', '.')
    match = re.search(r'\d+(\.\d+)?', texto_num)
    if not match:
        return 0
    return int(float(match.group(0)) * multiplicador)

def recorrer_json_fb(nodo, acumulado, textos_vistos):
    if isinstance(nodo, dict):
        for clave_likes in ('reaction_count', 'reactors', 'likers', 'unified_reactors'):
            val = nodo.get(clave_likes)
            if isinstance(val, dict) and isinstance(val.get('count'), int):
                if val['count'] > acumulado['likes']:
                    acumulado['likes'] = val['count']

        if isinstance(nodo.get('i18n_reaction_count'), str):
            num = parsear_numero_fb(nodo['i18n_reaction_count'])
            if num > acumulado['likes']:
                acumulado['likes'] = num

        for clave_shares in ('share_count', 'reshares'):
            val = nodo.get(clave_shares)
            if isinstance(val, dict) and isinstance(val.get('count'), int):
                if val['count'] > acumulado['compartidos']:
                    acumulado['compartidos'] = val['count']

        if isinstance(nodo.get('i18n_share_count'), str):
            num = parsear_numero_fb(nodo['i18n_share_count'])
            if num > acumulado['compartidos']:
                acumulado['compartidos'] = num

        body = nodo.get('body')
        author = nodo.get('author')
        if isinstance(body, dict) and isinstance(author, dict):
            texto_com = (body.get('text') or '').strip()
            nombre_autor = (author.get('name') or 'Usuario').strip()
            if texto_com and len(texto_com) > 1 and texto_com not in textos_vistos:
                typename = nodo.get('__typename', '')
                if typename in ('Comment', '') or 'comment' in str(nodo.get('id', '')).lower():
                    textos_vistos.add(texto_com)
                    acumulado['comentarios'].append({
                        'username': nombre_autor[:95],
                        'comment': texto_com
                    })

        for valor in nodo.values():
            recorrer_json_fb(valor, acumulado, textos_vistos)

    elif isinstance(nodo, list):
        for item in nodo:
            recorrer_json_fb(item, acumulado, textos_vistos)

def ejecutar_intento_scraping(p, url, usar_sesion=True):
    acumulado = {'likes': 0, 'compartidos': 0, 'comentarios': []}
    textos_vistos = set()
    archivo_sesion = 'fb_guest_state.json'

    browser = p.chromium.launch(
        headless=True,
        args=[
            '--headless=new',
            '--disable-blink-features=AutomationControlled',
            '--disable-infobars',
            '--no-sandbox'
        ]
    )

    opciones_contexto = {
        'locale': 'es-MX',
        'user_agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36',
        'viewport': {'width': 1366, 'height': 900}
    }
    if usar_sesion and os.path.exists(archivo_sesion):
        opciones_contexto['storage_state'] = archivo_sesion

    context = browser.new_context(**opciones_contexto)
    context.add_init_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
    page = context.new_page()

    # Bloquear archivos multimedia pesados manteniendo scripts y XHR
    page.route(
        "**/*",
        lambda route: route.abort()
        if route.request.resource_type in ["image", "media", "font"]
        else route.continue_()
    )

    def interceptar_graphql(response):
        if 'graphql' in response.url and response.status == 200:
            try:
                for linea in response.text().splitlines():
                    linea = linea.strip()
                    if linea.startswith('{'):
                        recorrer_json_fb(json.loads(linea), acumulado, textos_vistos)
            except Exception:
                pass

    page.on("response", interceptar_graphql)

    try:
        page.goto(url, wait_until='domcontentloaded', timeout=20000)
        try:
            page.wait_for_selector('script[type="application/json"], div[role="article"]', timeout=3500)
        except Exception:
            page.wait_for_timeout(1500)

        # Cerrar modal de login superpuesto
        try:
            page.keyboard.press('Escape')
            botones_cerrar = page.locator('div[role="dialog"] div[aria-label="Cerrar"], div[role="dialog"] div[aria-label="Close"]')
            if botones_cerrar.count() > 0:
                botones_cerrar.last.click(timeout=1000)
                page.wait_for_timeout(400)
        except Exception:
            pass

        # Extraer bloques JSON iniciales
        scripts_json = page.locator('script[type="application/json"]').all_inner_texts()
        for raw_json in scripts_json:
            try:
                recorrer_json_fb(json.loads(raw_json), acumulado, textos_vistos)
            except Exception:
                continue

        # Posicionar cursor en el centro para hacer scroll dentro del modal de comentarios
        page.mouse.move(680, 500)
        for _ in range(4):
            page.mouse.wheel(0, 1200)
            page.wait_for_timeout(600)
            try:
                btn_mas = page.locator(
                    'span:has-text("Ver más comentarios"), '
                    'span:has-text("Ver comentarios anteriores"), '
                    'span:has-text("View more comments")'
                ).first
                if btn_mas.is_visible():
                    btn_mas.click(timeout=800, force=True)
                    page.wait_for_timeout(900)
            except Exception:
                pass

        # Lectura de comentarios en el DOM
        articulos = page.locator('div[role="article"]').all()
        for art in articulos:
            try:
                aria_label = art.get_attribute('aria-label') or ''
                if 'comentario de' in aria_label.lower() or 'comment by' in aria_label.lower():
                    autor = re.sub(r'^(Comentario de|Comment by)\s+', '', aria_label, flags=re.IGNORECASE).strip()
                    autor = re.sub(r'\s+(hace|about|a las|el\s+\d).*$', '', autor, flags=re.IGNORECASE).strip()

                    bloques_texto = art.locator('div[dir="auto"]').all_inner_texts()
                    texto_com = " ".join([t.strip() for t in bloques_texto if t.strip()])

                    if texto_com and texto_com not in textos_vistos:
                        textos_vistos.add(texto_com)
                        acumulado['comentarios'].append({
                            'username': (autor or 'Usuario')[:95],
                            'comment': texto_com
                        })
            except Exception:
                continue

        # Respaldo de métricas en DOM y etiquetas meta
        if acumulado['likes'] == 0 or acumulado['compartidos'] == 0:
            try:
                meta_desc = page.locator('meta[property="og:description"], meta[name="description"]').first.get_attribute('content') or ''
                m_likes_meta = re.search(r'([\d\.,]+\s*(?:mil|k|m)?)\s*(?:reacciones|likes|me gusta)', meta_desc, re.IGNORECASE)
                if m_likes_meta:
                    acumulado['likes'] = max(acumulado['likes'], parsear_numero_fb(m_likes_meta.group(1)))

                texto_pagina = page.inner_text('body')
                if acumulado['likes'] == 0:
                    m_reac = re.search(r'Todas las reacciones:\s*([\d\.,]+\s*(?:mil|k|m)?)', texto_pagina, re.IGNORECASE)
                    if m_reac:
                        acumulado['likes'] = parsear_numero_fb(m_reac.group(1))
                if acumulado['compartidos'] == 0:
                    m_comp = re.search(r'([\d\.,]+\s*(?:mil|k|m)?)\s*veces compartido', texto_pagina, re.IGNORECASE)
                    if m_comp:
                        acumulado['compartidos'] = parsear_numero_fb(m_comp.group(1))
            except Exception:
                pass

        if acumulado['likes'] > 0 or len(acumulado['comentarios']) > 0:
            context.storage_state(path=archivo_sesion)
        elif os.path.exists(archivo_sesion):
            os.remove(archivo_sesion)

        print(f"[Scraper FB] URL: {page.url} | Likes: {acumulado['likes']} | Shares: {acumulado['compartidos']} | Comentarios: {len(acumulado['comentarios'])}")

    except Exception as e:
        print(f"Error en intento de scraping: {e}")
    finally:
        browser.close()

    return acumulado['likes'], acumulado['compartidos'], acumulado['comentarios']

def extraer_datos_facebook(url):
    with sync_playwright() as p:
        likes, compartidos, comentarios = ejecutar_intento_scraping(p, url, usar_sesion=True)
        # Reintento inmediato con sesión limpia si Facebook bloqueó el primer acceso
        if likes == 0 and compartidos == 0 and len(comentarios) == 0:
            print("Reintentando extracción con sesión limpia...")
            likes, compartidos, comentarios = ejecutar_intento_scraping(p, url, usar_sesion=False)
        return likes, compartidos, comentarios

@app.route('/api/login', methods=['POST'])
def login():
    data = request.get_json() or {}
    correo = data.get('correo')
    password = data.get('password')

    if not correo or not password:
        return jsonify({'error': 'Credenciales incompletas'}), 400

    try:
        conn = get_db_connection()
        with conn.cursor() as cursor:
            sql = "SELECT id, nombre, correo, rol FROM usuarios WHERE correo = %s AND password = %s"
            cursor.execute(sql, (correo, password))
            user = cursor.fetchone()
        conn.close()

        if user:
            return jsonify({'mensaje': 'Autenticación exitosa', 'usuario': user}), 200
        return jsonify({'error': 'Usuario o contraseña incorrectos'}), 401
    except Exception as e:
        return jsonify({'error': f'Error en servidor: {str(e)}'}), 500

@app.route('/api/analizar', methods=['POST'])
def analizar():
    data = request.get_json() or {}
    url = data.get('url', '').strip()

    if not url or ('facebook.com' not in url and 'fb.watch' not in url and 'fb.com' not in url):
        return jsonify({'error': 'Ingrese una URL válida de Facebook'}), 400

    try:
        conn = get_db_connection()
        with conn.cursor() as cursor:
            cursor.execute(
                "SELECT * FROM analisis WHERE url_post = %s AND (total_comentarios > 0 OR total_likes > 0) ORDER BY id DESC LIMIT 1",
                (url[:495],)
            )
            analisis_previo = cursor.fetchone()

            if analisis_previo:
                cursor.execute(
                    "SELECT usuario AS username, comentario AS comment, sentimiento AS sentiment FROM comentarios WHERE id_analisis = %s",
                    (analisis_previo['id'],)
                )
                comentarios_guardados = cursor.fetchall()
                conn.close()
                return jsonify({
                    'id_analisis': analisis_previo['id'],
                    'url_post': analisis_previo['url_post'],
                    'likes': analisis_previo['total_likes'],
                    'compartidos': analisis_previo['total_compartidos'],
                    'total_comentarios': analisis_previo['total_comentarios'],
                    'porcentajes': {
                        'positivo': analisis_previo['positivos'],
                        'negativo': analisis_previo['negativos'],
                        'neutro': analisis_previo['neutros']
                    },
                    'resumen_ia': analisis_previo['resumen_ia'],
                    'comentarios': comentarios_guardados
                }), 200

        likes, compartidos, comentarios_raw = extraer_datos_facebook(url)

        if likes == 0 and compartidos == 0 and len(comentarios_raw) == 0:
            conn.close()
            return jsonify({
                'error': 'Facebook bloqueó temporalmente la lectura sin sesión. Intenta nuevamente en unos segundos.'
            }), 429

        comentarios_procesados = clasificar_lote_ia(comentarios_raw)

        conteo = {'Positivo': 0, 'Negativo': 0, 'Neutro': 0}
        for c in comentarios_procesados:
            conteo[c['sentiment']] += 1

        total_com = len(comentarios_procesados)
        porc_pos = round((conteo['Positivo'] / total_com) * 100) if total_com > 0 else 0
        porc_neg = round((conteo['Negativo'] / total_com) * 100) if total_com > 0 else 0
        porc_neu = (100 - porc_pos - porc_neg) if total_com > 0 else 0

        resumen_ia = generar_resumen_ejecutivo(porc_pos, porc_neg, porc_neu, total_com)

        with conn.cursor() as cursor:
            sql_analisis = """
                INSERT INTO analisis (url_post, total_likes, total_compartidos, total_comentarios, positivos, negativos, neutros, resumen_ia)
                VALUES (%s, %s, %s, %s, %s, %s, %s, %s)
            """
            cursor.execute(sql_analisis, (url[:495], likes, compartidos, total_com, porc_pos, porc_neg, porc_neu, resumen_ia))
            id_analisis = cursor.lastrowid

            if comentarios_procesados:
                sql_com = "INSERT INTO comentarios (id_analisis, usuario, comentario, sentimiento) VALUES (%s, %s, %s, %s)"
                for c in comentarios_procesados:
                    cursor.execute(sql_com, (id_analisis, c['username'], c['comment'], c['sentiment']))

            conn.commit()
        conn.close()

        return jsonify({
            'id_analisis': id_analisis,
            'url_post': url,
            'likes': likes,
            'compartidos': compartidos,
            'total_comentarios': total_com,
            'porcentajes': {
                'positivo': porc_pos,
                'negativo': porc_neg,
                'neutro': porc_neu
            },
            'resumen_ia': resumen_ia,
            'comentarios': comentarios_procesados
        }), 200

    except Exception as e:
        print(f"Error en /api/analizar: {e}")
        return jsonify({'error': str(e)}), 500

@app.route('/api/exportar/<int:id_analisis>', methods=['GET'])
def exportar_excel(id_analisis):
    try:
        conn = get_db_connection()
        with conn.cursor() as cursor:
            cursor.execute("SELECT * FROM analisis WHERE id = %s", (id_analisis,))
            analisis = cursor.fetchone()
            cursor.execute("SELECT usuario, comentario, sentimiento FROM comentarios WHERE id_analisis = %s", (id_analisis,))
            comentarios = cursor.fetchall()
        conn.close()

        if not analisis:
            return jsonify({'error': 'Registro no encontrado'}), 404

        wb = Workbook()
        ws = wb.active
        ws.title = "Reporte Facebook"

        header_fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
        header_font = Font(color="FFFFFF", bold=True)

        ws.append(["REPORTE DE ANÁLISIS DE PUBLICACIÓN - FACEBOOK"])
        ws["A1"].font = Font(size=14, bold=True, color="1E3A8A")
        ws.append(["URL:", analisis['url_post']])
        ws.append(["Fecha:", str(analisis['fecha_analisis'])])
        ws.append([])

        ws.append(["Métrica", "Total"])
        ws["A5"].fill, ws["B5"].fill = header_fill, header_fill
        ws["A5"].font, ws["B5"].font = header_font, header_font
        ws.append(["Reacciones", analisis['total_likes']])
        ws.append(["Compartidos", analisis['total_compartidos']])
        ws.append(["Comentarios", analisis['total_comentarios']])
        ws.append(["Positivos (%)", f"{analisis['positivos']}%"])
        ws.append(["Negativos (%)", f"{analisis['negativos']}%"])
        ws.append(["Neutros (%)", f"{analisis['neutros']}%"])
        ws.append([])

        ws.append(["Resumen de IA:"])
        ws["A13"].font = Font(bold=True)
        ws.append([analisis['resumen_ia']])
        ws.append([])

        ws.append(["Usuario", "Comentario", "Sentimiento"])
        for col in ["A16", "B16", "C16"]:
            ws[col].fill = header_fill
            ws[col].font = header_font

        for c in comentarios:
            ws.append([c['usuario'], c['comentario'], c['sentimiento']])

        ws.column_dimensions['A'].width = 25
        ws.column_dimensions['B'].width = 65
        ws.column_dimensions['C'].width = 18

        output = io.BytesIO()
        wb.save(output)
        output.seek(0)

        return send_file(
            output,
            as_attachment=True,
            download_name=f"reporte_facebook_{id_analisis}.xlsx",
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    port = int(os.getenv('PORT', 5000))
    app.run(debug=True, use_reloader=False, port=port)