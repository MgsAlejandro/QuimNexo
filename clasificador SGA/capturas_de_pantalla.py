# -*- coding: utf-8 -*-
"""Toma capturas de las pantallas de la aplicación con un navegador de verdad (Chromium, con Playwright), para comparar
el antes y el después de un cambio de apariencia. Abre su propia copia de la aplicación en otro puerto y con una carpeta
temporal de datos: no usa el puerto del usuario ni toca sus mediciones, elecciones o consultas.
Pantallas: inicio, paso 1 con producto, lista «qué revisar primero», configuración, paso 2 con una medición de PRUEBA,
paso 3 en los dos perfiles (arriba y abajo) y detalle técnico.
Uso:
    py "clasificador SGA/capturas_de_pantalla.py" CARPETA_DE_SALIDA            (1366×768 y 1920×1080)
    py "clasificador SGA/capturas_de_pantalla.py" CARPETA_DE_SALIDA --zoom     (además, 1366×768 con zoom del 150 %)
Requiere: pip install playwright  y  py -m playwright install chromium"""
import os, re, subprocess, sys, tempfile, time, urllib.request

AQUI = os.path.dirname(os.path.abspath(__file__))
RAIZ = os.path.abspath(os.path.join(AQUI, '..'))
PUERTO = 8517
LANZADOR = '''# Copia de trabajo para las capturas: la misma aplicación, con los datos del usuario apuntando a una carpeta temporal
import os, sys
APP = {app!r}
sys.path.insert(0, os.path.dirname(APP))
import almacen
almacen.AQUI = {datos!r}
almacen.CONSULTAS = os.path.join({datos!r}, 'consultas')
almacen.REGISTRO = os.path.join({datos!r}, 'consultas', 'registro.csv')
exec(compile(open(APP, encoding='utf-8').read(), APP, 'exec'), {{'__file__': APP, '__name__': '__main__'}})
'''


def _esperar_servidor(url, segundos=90):
    limite = time.time() + segundos
    while time.time() < limite:
        try:
            urllib.request.urlopen(url, timeout=2)
            return
        except Exception:
            time.sleep(1)
    raise SystemExit('La aplicación no arrancó a tiempo para las capturas.')


def _recorrido(pagina, destino, sufijo, problemas):
    """Recorre las pantallas y guarda una captura de cada una. Anota lo que no encuentre, sin detenerse."""
    pausa = lambda ms=2500: pagina.wait_for_timeout(ms)

    def foto(nombre):
        pagina.screenshot(path=os.path.join(destino, f'{nombre}_{sufijo}.png'))

    def paso(nombre, accion):
        try:
            accion()
        except Exception as error:
            problemas.append(f'{sufijo} · {nombre}: {type(error).__name__}: {str(error).splitlines()[0][:160]}')

    arriba = lambda: pagina.evaluate("() => { const m = document.querySelector('[data-testid=\"stMain\"]'); if (m) m.scrollTo(0, 0); window.scrollTo(0, 0); }")
    abajo = lambda px: pagina.evaluate("(px) => { const m = document.querySelector('[data-testid=\"stMain\"]'); if (m) m.scrollTo(0, px); }", px)
    boton = lambda inicio: pagina.get_by_role('button', name=re.compile('^' + re.escape(inicio))).first
    lateral = pagina.locator('section[data-testid="stSidebar"]')

    pagina.goto(f'http://localhost:{PUERTO}', wait_until='networkidle')
    pausa(4000)
    foto('01_inicio')

    def elegir_producto():
        caja = pagina.locator('input[role="combobox"][aria-label*="Producto"]').first
        caja.click()
        caja.type('02 ·')
        pagina.keyboard.press('Enter')
        pausa(4000)
    paso('elegir producto', elegir_producto)
    arriba()
    foto('02_paso1_producto')

    def lista():
        pagina.get_by_text(re.compile('^Qué revisar primero')).first.click()
        pausa(5000)
        pagina.get_by_text(re.compile('^Qué revisar primero')).first.scroll_into_view_if_needed()
        pausa(800)
    paso('lista de prioridad', lista)
    foto('03_lista_de_prioridad')
    paso('cerrar lista', lambda: (pagina.get_by_text(re.compile('^Qué revisar primero')).first.click(), pausa(1200)))

    def configuracion():
        lateral.get_by_text(re.compile('Configuración')).first.click()
        pausa(1500)
    paso('configuración', configuracion)
    foto('04_configuracion')
    paso('cerrar configuración', lambda: (pagina.keyboard.press('Escape'), pausa(800)))

    def exposicion():
        arriba()
        boton('Continuar').click()
        pausa(3500)
        casilla = pagina.locator('input[type="number"][aria-label^="Concentración medida"]').first
        casilla.click()
        casilla.fill('62')   # PRUEBA: 62 mg/m3 durante 8 h, contra un límite de 100 mg/m3
        pagina.keyboard.press('Enter')
        pausa(3000)
        arriba()
    paso('paso 2 con medición', exposicion)
    foto('05_paso2_exposicion')

    def resultado():
        boton('Continuar').click()
        pausa(5000)
        arriba()
    paso('paso 3', resultado)
    foto('06_paso3_profesional')
    abajo(900)
    pausa(800)
    foto('07_paso3_profesional_abajo')

    def trabajador():
        lateral.get_by_text('Trabajador/supervisor').first.click()
        pausa(3500)
        arriba()
    paso('perfil trabajador', trabajador)
    foto('08_paso3_trabajador')

    def detalle():
        lateral.get_by_text('Profesional de SST').first.click()
        pausa(3000)
        lateral.get_by_text('Ver el detalle técnico').first.click()
        pausa(4500)
        arriba()
    paso('detalle técnico', detalle)
    foto('09_detalle_tecnico')


def capturar(destino, con_zoom=False):
    """Devuelve la lista de problemas encontrados al recorrer (vacía si todo se pudo capturar)."""
    from playwright.sync_api import sync_playwright
    os.makedirs(destino, exist_ok=True)
    problemas = []
    with tempfile.TemporaryDirectory() as temporal:
        lanzador = os.path.join(AQUI, '_lanzador_de_capturas.py')   # junto a app.py, para que Streamlit sirva la carpeta static (la fuente)
        with open(lanzador, 'w', encoding='utf-8') as f:
            f.write(LANZADOR.format(app=os.path.join(AQUI, 'app.py'), datos=os.path.join(temporal, 'datos')))
        servidor = subprocess.Popen([sys.executable, '-m', 'streamlit', 'run', lanzador, '--server.port', str(PUERTO), '--server.headless', 'true',
                                     '--browser.gatherUsageStats', 'false'], cwd=RAIZ, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        try:
            _esperar_servidor(f'http://localhost:{PUERTO}')
            # (sufijo, ancho, alto, escala): el zoom del 150 % equivale a una ventana un tercio más angosta con los puntos 1,5 veces más grandes
            tamanos = [('1366x768', 1366, 768, 1), ('1920x1080', 1920, 1080, 1)] + ([('1366x768_zoom150', 911, 512, 1.5)] if con_zoom else [])
            with sync_playwright() as p:
                navegador = p.chromium.launch()
                for sufijo, ancho, alto, escala in tamanos:
                    contexto = navegador.new_context(viewport={'width': ancho, 'height': alto}, device_scale_factor=escala, locale='es-CO')
                    _recorrido(contexto.new_page(), destino, sufijo, problemas)
                    contexto.close()
                navegador.close()
        finally:
            if os.path.isfile(lanzador):
                os.remove(lanzador)
            servidor.terminate()
            try:
                servidor.wait(timeout=15)
            except Exception:
                servidor.kill()
    return problemas


if __name__ == '__main__':
    sys.stdout.reconfigure(encoding='utf-8')
    destino = sys.argv[1]
    problemas = capturar(destino, con_zoom='--zoom' in sys.argv)
    hechas = sorted(a for a in os.listdir(destino) if a.endswith('.png'))
    print(f'{len(hechas)} capturas en {destino}')
    print('Sin problemas en el recorrido.' if not problemas else 'Problemas en el recorrido:\n  ' + '\n  '.join(problemas))
    sys.exit(1 if problemas else 0)
