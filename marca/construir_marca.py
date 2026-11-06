# -*- coding: utf-8 -*-
"""Genera los archivos de la marca QuimNexo (propuesta B, monograma N) en esta carpeta:
  simbolo.svg, simbolo_fondo_oscuro.svg, simbolo_una_tinta.svg        el símbolo solo
  logotipo_horizontal.svg, ..._fondo_oscuro.svg, ..._una_tinta.svg    símbolo y nombre, para la cabecera
  favicon.svg, favicon_32.png, favicon_64.png                         icono de la pestaña del navegador
  logotipo_horizontal.png                                             copia en PNG transparente, por si un programa no lee SVG
Todo es dibujo propio, con geometría simple y sin degradados. El nombre va en Inter (licencia OFL), que se incrusta dentro
de cada logotipo horizontal para que se vea igual en cualquier equipo, sin internet y sin tener la fuente instalada.
Uso:  py marca/construir_marca.py        (los PNG necesitan Playwright; los SVG no)"""
import base64, os, pathlib, sys

AQUI = os.path.dirname(os.path.abspath(__file__))
FUENTE = os.path.join(AQUI, '..', 'clasificador SGA', 'static', 'fuentes', 'InterVariable.woff2')
PETROLEO, ACCION, ACENTO, TINTA, CLARO, ACENTO_CLARO, PROFUNDO = '#0F3D4C', '#0B6E82', '#17A2B8', '#17232B', '#FFFFFF', '#5FD3E3', '#0A2A35'
ANCHO_DEL_NOMBRE = 196   # ancho del texto «QuimNexo» a 40 px en Inter; lo mide y lo corrige este mismo guion si hay navegador


def simbolo(ficha, letra, nodos):
    """Monograma N dentro de un cuadro redondeado: dos apoyos unidos por un trazo, con un nodo en cada extremo."""
    return (f'<rect x="2" y="2" width="60" height="60" rx="15" fill="{ficha}"/>'
            f'<path d="M20 46 V18 L44 46 V18" fill="none" stroke="{letra}" stroke-width="7" stroke-linecap="round" stroke-linejoin="round"/>'
            f'<circle cx="20" cy="18" r="5.500" fill="{nodos}"/><circle cx="44" cy="46" r="5.500" fill="{nodos}"/>')


VERSIONES = {'': (simbolo(PETROLEO, CLARO, ACENTO), PETROLEO, ACCION),
             '_fondo_oscuro': (simbolo(CLARO, PROFUNDO, ACCION), CLARO, ACENTO_CLARO),
             '_una_tinta': (simbolo(TINTA, CLARO, CLARO), TINTA, TINTA)}


def svg_del_simbolo(interior):
    return f'<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 64 64" role="img" aria-label="QuimNexo"><title>QuimNexo</title>{interior}</svg>\n'


def svg_del_logotipo(interior, color_quim, color_nexo, fuente_b64, ancho_del_nombre):
    ancho = 78 + ancho_del_nombre + 4
    return (f'<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 {ancho:g} 64" role="img" aria-label="QuimNexo"><title>QuimNexo</title>'
            f'<style>@font-face{{font-family:"InterQN";src:url("data:font/woff2;base64,{fuente_b64}") format("woff2");font-weight:100 900}}'
            f'.n{{font-family:"InterQN","Inter","Segoe UI",Arial,sans-serif;font-size:40px;letter-spacing:-0.8px}}</style>'
            f'{interior}<text class="n" x="78" y="46" fill="{color_quim}" font-weight="500">Quim<tspan fill="{color_nexo}" font-weight="800">Nexo</tspan></text></svg>\n')


def escribir(nombre, contenido):
    with open(os.path.join(AQUI, nombre), 'w', encoding='utf-8') as f:
        f.write(contenido)


def construir():
    with open(FUENTE, 'rb') as f:
        fuente_b64 = base64.b64encode(f.read()).decode('ascii')
    ancho = ANCHO_DEL_NOMBRE
    for sufijo, (interior, quim, nexo) in VERSIONES.items():
        escribir(f'simbolo{sufijo}.svg', svg_del_simbolo(interior))
        escribir(f'logotipo_horizontal{sufijo}.svg', svg_del_logotipo(interior, quim, nexo, fuente_b64, ancho))
    escribir('favicon.svg', svg_del_simbolo(VERSIONES[''][0]))
    try:
        from playwright.sync_api import sync_playwright
    except ImportError:
        return 'Sin Playwright: se generaron los SVG, pero no los PNG ni se midió el ancho del nombre.'
    with sync_playwright() as p:
        navegador = p.chromium.launch()
        pagina = navegador.new_page(viewport={'width': 800, 'height': 300})
        # 1. Medir el ancho real del nombre y rehacer los logotipos con ese ancho, para que no sobre ni falte lienzo
        pagina.goto(pathlib.Path(AQUI, 'logotipo_horizontal.svg').as_uri())
        pagina.wait_for_timeout(600)
        medido = pagina.evaluate("document.querySelector('text').getBBox().width")
        for sufijo, (interior, quim, nexo) in VERSIONES.items():
            escribir(f'logotipo_horizontal{sufijo}.svg', svg_del_logotipo(interior, quim, nexo, fuente_b64, round(medido + 1)))
        # 2. PNG transparentes
        for nombre, origen, ancho_px, alto_px in (('favicon_32.png', 'favicon.svg', 32, 32), ('favicon_64.png', 'favicon.svg', 64, 64),
                                                  ('logotipo_horizontal.png', 'logotipo_horizontal.svg', None, 128)):
            if ancho_px is None:
                ancho_px = round(alto_px * (78 + round(medido + 1) + 4) / 64)
            pagina.set_viewport_size({'width': ancho_px, 'height': alto_px})
            pagina.goto(pathlib.Path(AQUI, origen).as_uri())   # el SVG abierto solo ocupa toda la ventana, que tiene justo el tamaño pedido
            pagina.wait_for_timeout(500)
            pagina.screenshot(path=os.path.join(AQUI, nombre), omit_background=True)
        navegador.close()
    return f'Ancho medido del nombre: {medido:.1f} px. SVG y PNG generados.'


if __name__ == '__main__':
    sys.stdout.reconfigure(encoding='utf-8')
    print(construir())
    for a in sorted(os.listdir(AQUI)):
        if a.endswith(('.svg', '.png')) and not a.startswith('muestra'):
            print(f'  {a:<42}{os.path.getsize(os.path.join(AQUI, a)) / 1024:>8.1f} KB')
