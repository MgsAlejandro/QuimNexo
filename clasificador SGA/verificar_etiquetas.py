# -*- coding: utf-8 -*-
"""Verificación de la etiqueta de uso en la mina en un navegador de verdad (Chromium, con Playwright).
Genera la etiqueta de todas las fichas con cinco capacidades de empaque, la abre y comprueba que cada una:
  - quede imprimible (cabe en su tamaño, sin desbordarse) y con letra no menor que la mínima configurada;
  - no traiga «...» ni corchetes, ni renglones repetidos (el nombre químico que viene de la ficha, en letra pequeña, sí puede traer corchetes);
  - lleve el nombre en español de la hoja Nombres_etiqueta;
  - tenga el tamaño de etiqueta que la hoja Etiqueta_tamano da para la capacidad convertida a litros.
Guarda cuatro capturas en clasificador SGA/salidas/capturas_etiqueta/.
Uso:  py "clasificador SGA/verificar_etiquetas.py"
Requiere: pip install playwright  y  py -m playwright install chromium"""
import os, re, sys

AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, AQUI)
from playwright.sync_api import sync_playwright
from tablas import cargar_tablas
from motor import clasificar, cargar_fichas, tamano_de_etiqueta
from etiqueta import ajustes
from etiqueta_html import _medida
from etiqueta_mina import contenido_de_etiqueta_mina
from etiqueta_mina_html import dibujar_etiqueta_mina

CAPACIDADES = ((500, 'mL'), (4, 'L'), (5, 'galón (US)'), (20, 'kg'), (200, 'L'))
CAPTURAS = {('14', 20, 'L'), ('04', 20, 'L'), ('27', 20, 'L'), ('02', 500, 'mL')}
DESTINO = os.path.join(AQUI, 'salidas', 'capturas_etiqueta')
MEDIR = """() => { const e = document.querySelector('.etiqueta'), i = e.querySelector('.interior'), r = e.getBoundingClientRect(), s = getComputedStyle(e);
  return { ajuste: document.documentElement.dataset.ajuste, letra: parseFloat(document.documentElement.dataset.letra), quitado: document.documentElement.dataset.quitado,
           ancho_mm: r.width / 96 * 25.4, alto_mm: r.height / 96 * 25.4,
           desborda: i.scrollHeight > e.clientHeight - parseFloat(s.paddingTop) - parseFloat(s.paddingBottom) + 0.5,
           texto: [...e.querySelectorAll('.interior > :not(.nombre-fds)')].map(x => x.innerText).join(' '), nombre: e.querySelector('.nombre').innerText, renglones: [...e.querySelectorAll('.hacer li')].map(x => x.innerText.trim()),
           pictogramas: e.querySelectorAll('.pictogramas > div').length, palabra: (e.querySelector('.palabra, .sin-peligros') || {}).innerText || '' }; }"""


def verificar(capturas=True):
    """Devuelve (filas del resumen por capacidad, lista de fallas, total de etiquetas)."""
    tablas = cargar_tablas()
    tablas['ajustes_etiqueta'] = {'proveedor_obligatorio': False, 'tamano_minimo_de_letra_pt': ajustes()['tamano_minimo_de_letra_pt']}
    minimo = tablas['ajustes_etiqueta']['tamano_minimo_de_letra_pt']
    fichas = cargar_fichas()
    os.makedirs(DESTINO, exist_ok=True)
    resumen, fallas, total = [], [], 0
    with sync_playwright() as p:
        navegador = p.chromium.launch()
        pagina = navegador.new_page(viewport={'width': 1300, 'height': 1000})
        for cantidad, unidad in CAPACIDADES:
            letras, quitando, completas = [], 0, 0
            for ficha in fichas:
                numero = ficha['archivo'].split('_')[0]
                r = clasificar(ficha, tablas)
                c = contenido_de_etiqueta_mina(ficha, r, {'capacidad': cantidad, 'unidad': unidad}, tablas)
                pagina.set_content(dibujar_etiqueta_mina(c, tablas), wait_until='load')
                m = pagina.evaluate(MEDIR)
                total += 1
                esperado = _medida(tamano_de_etiqueta(cantidad * tablas['unidades'][unidad]['litros'], tablas)['etiqueta_mm'], 'la etiqueta')
                problemas = []
                if m['ajuste'] not in ('cabe', 'cabe quitando contenido') or m['desborda']:
                    problemas.append(f"no queda imprimible ({m['ajuste']})")
                if m['letra'] < minimo:
                    problemas.append(f"letra de {m['letra']} pt")
                if re.search(r'\.\.\.|…|\[|\]', m['texto']):
                    problemas.append('trae «...» o corchetes')
                if len(set(m['renglones'])) != len(m['renglones']):
                    problemas.append('renglones repetidos')
                if m['nombre'].strip() != tablas['nombres_etiqueta'][numero]['nombre']:
                    problemas.append(f"nombre «{m['nombre']}»")
                if abs(m['ancho_mm'] - esperado[0]) > 0.6 or abs(m['alto_mm'] - esperado[1]) > 0.6:
                    problemas.append(f"mide {m['ancho_mm']:.0f} × {m['alto_mm']:.0f} mm y debía medir {esperado[0]:g} × {esperado[1]:g}")
                if m['pictogramas'] != len(c['pictogramas']) or not m['palabra'] or not (m['renglones'] or not c['renglones']):
                    problemas.append('falta contenido mínimo (pictogramas, palabra o un renglón)')
                if problemas:
                    fallas.append(f"{numero} con {cantidad} {unidad}: " + '; '.join(problemas))
                letras.append(m['letra'])
                quitando += bool(m['quitado'])
                completas += not m['quitado']
            litros = cantidad * tablas['unidades'][unidad]['litros']
            ancho, alto = _medida(tamano_de_etiqueta(litros, tablas)['etiqueta_mm'], 'la etiqueta')
            resumen.append((f'{cantidad} {unidad}', f'{litros:.3g} L'.replace('.', ','), f'{ancho:g} × {alto:g} mm', len(fichas), completas, quitando,
                            f'{min(letras):g} pt', f'{sum(letras) / len(letras):.1f} pt'.replace('.', ',')))
        if capturas:
            for numero, cantidad, unidad in sorted(CAPTURAS):
                ficha = next(f for f in fichas if f['archivo'].split('_')[0] == numero)
                c = contenido_de_etiqueta_mina(ficha, clasificar(ficha, tablas), {'capacidad': cantidad, 'unidad': unidad}, tablas)
                pagina.set_content(dibujar_etiqueta_mina(c, tablas), wait_until='load')
                pagina.screenshot(path=os.path.join(DESTINO, f"{numero}_{cantidad}_{re.sub(r'[^A-Za-z]', '', unidad)}.png"), full_page=True)
        navegador.close()
    return resumen, fallas, total


if __name__ == '__main__':
    sys.stdout.reconfigure(encoding='utf-8')
    resumen, fallas, total = verificar()
    print(f"{'Empaque':<16}{'En litros':<12}{'Etiqueta':<16}{'Fichas':>7}{'Cabe todo':>11}{'Quitando':>10}{'Letra mín.':>12}{'Letra media':>13}")
    for fila in resumen:
        print(f'{fila[0]:<16}{fila[1]:<12}{fila[2]:<16}{fila[3]:>7}{fila[4]:>11}{fila[5]:>10}{fila[6]:>12}{fila[7]:>13}')
    print(f"\n{total - len(fallas)} de {total} etiquetas bien." + ('' if not fallas else '\nFALLAS:\n  ' + '\n  '.join(fallas)))
    print(f'Capturas en {DESTINO}')
    sys.exit(1 if fallas else 0)
