# -*- coding: utf-8 -*-
"""Paso E1 del reporte y la etiqueta: pictogramas listos para incrustar.
Toma los SVG de plantillas/pictogramas (borde rojo, fondo blanco, símbolo negro) y entrega:
  - la versión de borde rojo, sin los metadatos que no hacen falta para imprimir;
  - la versión de borde negro, permitida cuando el producto no sale del lugar de trabajo (regla R09,
    Res. 773 de 2021, art. 12). Se obtiene cambiando el único rojo del dibujo, que es el rombo del borde.
No modifica los archivos originales. Si un SVG no tiene exactamente un rojo, se detiene: no recolorea a ciegas.
Uso de prueba:  py "clasificador SGA/pictogramas_svg.py"     (genera una hoja de muestra con los 18)
Requiere: solo Python."""
import os, re, sys

RAIZ = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..')
ORIGEN = os.path.join(RAIZ, 'plantillas', 'pictogramas')
SALIDAS = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'salidas')
ROJO = re.compile(r'#ff0000\b|#f00\b|(?<=["\':\s])red\b', re.I)
BORDES = {'rojo': None, 'negro': '#000000'}


class ErrorDePictograma(Exception):
    """Un SVG de pictograma no tiene la forma esperada."""


def svg_de_pictograma(codigo, borde='rojo', descripcion=None):
    """SVG del pictograma 'GHS02' como texto, listo para incrustar en una página.
    borde = 'rojo' (el de la norma) o 'negro' (uso dentro del lugar de trabajo, regla R09).
    El dibujo no trae ancho ni alto fijos: toma el tamaño del recuadro donde se ponga."""
    if borde not in BORDES:
        raise ErrorDePictograma(f'El borde debe ser «rojo» o «negro»; se recibió «{borde}».')
    ruta = os.path.join(ORIGEN, codigo + '.svg')
    if not os.path.isfile(ruta):
        raise ErrorDePictograma(f'No está el archivo {codigo}.svg en plantillas/pictogramas.')
    with open(ruta, encoding='utf-8') as f:
        svg = f.read()
    svg = re.sub(r'<\?xml.*?\?>|<!DOCTYPE.*?>|<!--.*?-->|<metadata>.*?</metadata>', '', svg, flags=re.S)
    raiz = re.search(r'<svg\b[^>]*>', svg)
    caja = re.search(r'viewBox="([^"]+)"', raiz.group(0)) if raiz else None
    if not caja:
        raise ErrorDePictograma(f'{codigo}.svg no declara su recuadro (viewBox); no se puede escalar con seguridad.')
    rojos = ROJO.findall(svg)
    if len(rojos) != 1:
        raise ErrorDePictograma(f'{codigo}.svg tiene {len(rojos)} zonas rojas y se esperaba una sola (el borde). Revisar el archivo.')
    if BORDES[borde]:
        svg = ROJO.sub(BORDES[borde], svg)
    titulo = f'<title>{codigo}{" — " + descripcion if descripcion else ""}</title>'
    nueva_raiz = f'<svg xmlns="http://www.w3.org/2000/svg" viewBox="{caja.group(1)}" role="img" aria-label="{codigo}">{titulo}'
    return (svg[:raiz.start()] + nueva_raiz + svg[raiz.end():]).strip()


if __name__ == '__main__':
    sys.stdout.reconfigure(encoding='utf-8')
    sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
    from tablas import cargar_tablas
    catalogo = cargar_tablas()['pictogramas']
    destino = os.path.join(SALIDAS, 'pictogramas')
    filas = []
    for codigo, datos in catalogo.items():
        celdas = ''
        for borde in BORDES:
            svg = svg_de_pictograma(codigo, borde, datos['simbolo'])
            os.makedirs(os.path.join(destino, borde), exist_ok=True)
            with open(os.path.join(destino, borde, codigo + '.svg'), 'w', encoding='utf-8') as f:
                f.write(svg)
            celdas += f'<td><div class="p">{svg}</div></td>'
            print(f"{codigo} borde {borde:<6} {len(svg):>6} caracteres | rojos que quedan: {len(ROJO.findall(svg))}")
        filas.append(f"<tr><th>{codigo}<br><span>{datos['simbolo']}</span></th>{celdas}</tr>")
    pagina = ('<!doctype html><html lang="es"><meta charset="utf-8"><title>Muestra de pictogramas</title>'
              '<style>body{font-family:Arial,sans-serif;margin:8mm}table{border-collapse:collapse}'
              'th,td{border:1px solid #999;padding:3mm;text-align:center;font-size:10pt}th span{font-weight:normal}'
              '.p{width:23mm;height:23mm;margin:auto}.p svg{width:100%;height:100%;display:block}</style>'
              '<h1 style="font-size:13pt">Pictogramas a 23 × 23 mm (envases de más de 3 y hasta 50 litros)</h1>'
              '<table><tr><th>Código</th><th>Borde rojo (norma)</th><th>Borde negro (uso interno, R09)</th></tr>'
              + ''.join(filas) + '</table></html>')
    with open(os.path.join(destino, 'muestra.html'), 'w', encoding='utf-8') as f:
        f.write(pagina)
    print(f'Muestra guardada en {os.path.join(destino, "muestra.html")}')
