# -*- coding: utf-8 -*-
"""Paso 1 de la extracción: lee una FDS en PDF y devuelve su texto por página,
sin los encabezados y pies que se repiten en cada hoja.
Uso de prueba:  py extraccion/leer_pdf.py FDS/15_Acetileno_74-86-2.pdf [página]
Requiere: pip install pdfplumber"""
import re, sys
from collections import Counter
import pdfplumber
import pypdfium2 as pdfium
from pdfplumber.utils.exceptions import PdfminerException

ZONA_ENCABEZADO = 8   # líneas del inicio de cada página donde se busca encabezado
ZONA_PIE = 5          # líneas del final de cada página donde se busca pie
MIN_REPETICION = 0.5  # una línea es encabezado o pie si se repite en esta fracción de páginas


def _clave(linea):
    """Forma comparable de una línea: sin números (cambian con la página) ni espacios dobles."""
    return re.sub(r'\s+', ' ', re.sub(r'\d+', '#', linea)).strip().lower()


def _es_separador(linea):
    return re.fullmatch(r'[\s_\-=.·]{5,}', linea) is not None


def _texto_por_pagina(ruta):
    """Líneas no vacías de cada página y el lector usado. Algunas FDS traen una
    protección que pdfplumber no abre (p. ej. la 23); esas se leen con pypdfium2."""
    try:
        with pdfplumber.open(ruta) as pdf:
            textos, lector = [p.extract_text() or '' for p in pdf.pages], 'pdfplumber'
    except PdfminerException:
        pdf = pdfium.PdfDocument(ruta)
        textos, lector = [pdf[i].get_textpage().get_text_bounded() for i in range(len(pdf))], 'pypdfium2'
    return [[l.rstrip() for l in t.splitlines() if l.strip()] for t in textos], lector


def leer_pdf(ruta):
    """Devuelve {'ruta', 'lector', 'paginas': [{'numero', 'lineas'}], 'quitadas': [(numero, línea)]}.
    'lineas' es el texto útil de la página; 'quitadas' deja constancia de lo que se retiró."""
    crudas, lector = _texto_por_pagina(ruta)

    def zonas(lineas):
        return lineas[:ZONA_ENCABEZADO], lineas[-ZONA_PIE:] if len(lineas) > ZONA_ENCABEZADO else []

    conteo = Counter()
    for lineas in crudas:
        arriba, abajo = zonas(lineas)
        conteo.update({_clave(l) for l in arriba + abajo})
    minimo = max(3, round(len(crudas) * MIN_REPETICION))
    repetidas = {c for c, n in conteo.items() if n >= minimo and c}

    paginas, quitadas = [], []
    for numero, lineas in enumerate(crudas, 1):
        n = len(lineas)
        utiles = []
        borde = True  # en la página 1: seguimos pegados al borde superior, sin contenido real todavía
        for i, linea in enumerate(lineas):
            en_encabezado = i < ZONA_ENCABEZADO
            en_pie = n > ZONA_ENCABEZADO and i >= n - ZONA_PIE
            repetida = _clave(linea) in repetidas
            # En la página 1 el contenido empieza arriba: una línea repetida que viene después
            # de contenido real (p. ej. el nombre del producto en la ficha 23) no es encabezado.
            if numero == 1 and en_encabezado and repetida and not borde:
                repetida = False
            if _es_separador(linea) or ((en_encabezado or en_pie) and repetida):
                quitadas.append((numero, linea))
            else:
                utiles.append(linea)
                borde = False
        paginas.append({'numero': numero, 'lineas': utiles})
    return {'ruta': ruta, 'lector': lector, 'paginas': paginas, 'quitadas': quitadas}


if __name__ == '__main__':
    sys.stdout.reconfigure(encoding='utf-8')
    doc = leer_pdf(sys.argv[1])
    ver = int(sys.argv[2]) if len(sys.argv) > 2 else 1
    total = sum(len(p['lineas']) for p in doc['paginas'])
    print(f"{doc['ruta']} (leído con {doc['lector']}): {len(doc['paginas'])} páginas, {total} líneas útiles, "
          f"{len(doc['quitadas'])} líneas quitadas")
    print('Encabezados y pies quitados (línea y veces):')
    for clave, veces in Counter(_clave(l) for _, l in doc['quitadas']).most_common():
        print(f'   {veces:>3} × {clave[:100]}')
    print(f'--- Página {ver} ya limpia ---')
    print('\n'.join(doc['paginas'][ver - 1]['lineas']))
