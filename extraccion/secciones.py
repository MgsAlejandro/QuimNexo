# -*- coding: utf-8 -*-
"""Paso 2 de la extracción: corta el texto de una FDS en sus 16 secciones.
No deduce nada: si una sección falta, está fuera de orden, se repite, viene vacía
o su título no corresponde a su número, la marca y la deja como está en la ficha.
Uso de prueba:
    py extraccion/secciones.py FDS                      (tabla de todas las fichas)
    py extraccion/secciones.py FDS/15_Acetileno_74-86-2.pdf [n]   (detalle; n = sección a ver)
Requiere: pip install pdfplumber openpyxl"""
import glob, os, re, sys, unicodedata
import openpyxl
from leer_pdf import leer_pdf

DICCIONARIO = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'matriz', 'frases_HP_es_en.xlsx')
# Palabras sin significado propio, que no sirven para reconocer un título
VACIAS = set('de del la las los el y e o u en que sobre al a un una por para the of and or on for to an in'.split())
UMBRAL = 0.5  # parte mínima del título de la ficha que debe coincidir con un título oficial
TITULO = re.compile(r'^\s*(?:(SECTION|SECCI[ÓO]N)\s*)?(\d{1,2})\s*([.:)\-–—]*)\s*(.*)$', re.I)

# Marcas del reporte
OK, FALTA, OTRA, NO_REC = 'OK', 'FALTA', 'TÍTULO DE OTRA SECCIÓN', 'TÍTULO NO RECONOCIDO'
LETRA = {FALTA: 'F', OTRA: 'T', NO_REC: 'N'}


def _palabras(titulo):
    """Raíces (6 letras) de las palabras de un título, sin tildes. Las palabras con guion
    se guardan separadas y unidas, para que 'fire-fighting' y 'firefighting' coincidan."""
    t = unicodedata.normalize('NFD', re.sub(r'\([^)]*\)', ' ', titulo.lower()))
    t = ''.join(ch for ch in t if unicodedata.category(ch) != 'Mn')
    palabras = re.findall(r'[a-z]+', t) + re.findall(r'[a-z]+', t.replace('-', ''))
    return {p[:6] for p in palabras if len(p) > 2 and p not in VACIAS}


def cargar_titulos_oficiales():
    """{número: raíces de sus títulos en español e inglés}, de la hoja Secciones_FDS."""
    hoja = openpyxl.load_workbook(DICCIONARIO, read_only=True)['Secciones_FDS']
    return {n: _palabras(es) | _palabras(en) for n, es, en, *_ in hoja.iter_rows(min_row=2, values_only=True)}


def _parecido(titulo, oficiales):
    """Para cada sección, qué parte del título de la ficha coincide con su título oficial.
    Una palabra que sale en varios títulos oficiales (p. ej. 'información') pesa menos."""
    peso = {}
    for raices in oficiales.values():
        for r in raices:
            peso[r] = peso.get(r, 0) + 1
    conocidas = [p for p in _palabras(titulo) if p in peso]
    total = sum(1 / peso[p] for p in conocidas)
    if not total:
        return {}
    return {n: sum(1 / peso[p] for p in conocidas if p in raices) / total for n, raices in oficiales.items()}


def _juzgar(numero, titulo, oficiales):
    """Estado del título frente a su número y, si es de otra sección, de cuál."""
    notas = _parecido(titulo, oficiales)
    mejor = max(notas.values(), default=0)
    if mejor < UMBRAL:
        return NO_REC, []
    if notas[numero] >= mejor - 1e-9:
        return OK, []
    return OTRA, [n for n, v in notas.items() if v >= mejor - 1e-9]


def cortar_secciones(doc, oficiales=None):
    """Recibe la salida de leer_pdf. Devuelve:
    'secciones': {1..16: {'estado', 'titulo', 'pagina', 'lineas': [(página, texto)], 'marcas': [..]}}
    'preambulo': líneas anteriores al primer título
    'avisos': hallazgos que no cambian el corte (títulos repetidos, candidatos descartados)."""
    oficiales = oficiales or cargar_titulos_oficiales()
    lineas = [(p['numero'], l) for p in doc['paginas'] for l in p['lineas']]

    candidatos = {}
    for pos, (pag, linea) in enumerate(lineas):
        m = TITULO.match(linea)
        if not m or len(linea) > 120:
            continue
        prefijo, numero, separador, titulo = m.group(1), int(m.group(2)), m.group(3), m.group(4).strip()
        if not 1 <= numero <= 16 or not titulo[:1].isalpha() or not (prefijo or separador):
            continue
        estado, de_otra = _juzgar(numero, titulo, oficiales)
        candidatos.setdefault(numero, []).append(
            {'pos': pos, 'pagina': pag, 'titulo': linea.strip(), 'estado': estado, 'de_otra': de_otra, 'prefijo': bool(prefijo)})

    secciones, avisos = {}, []
    for n in range(1, 17):
        todos = candidatos.get(n, [])
        reconocidos = [c for c in todos if c['estado'] != NO_REC]
        con_prefijo = [c for c in todos if c['estado'] == NO_REC and c['prefijo']]
        elegido = (reconocidos or con_prefijo or [None])[0]
        secciones[n] = {'estado': FALTA, 'titulo': None, 'pagina': None, 'lineas': [], 'marcas': [], 'pos': None}
        if elegido:
            secciones[n].update(estado=elegido['estado'], titulo=elegido['titulo'], pagina=elegido['pagina'], pos=elegido['pos'])
            if elegido['estado'] == OTRA:
                secciones[n]['marcas'].append('el título es el de la sección ' + ' o '.join(map(str, elegido['de_otra'])))
        for c in todos:
            if c is elegido:
                continue
            if c['estado'] != NO_REC:
                secciones[n]['marcas'].append(f"REPETIDA: otro título {n} en p. {c['pagina']}: «{c['titulo']}»")
            elif c['prefijo'] or not elegido:
                avisos.append(f"Sección {n}: línea que parece título y no se usó, p. {c['pagina']}: «{c['titulo']}»")

    presentes = sorted((s['pos'], n) for n, s in secciones.items() if s['pos'] is not None)
    for i, (pos, n) in enumerate(presentes):
        fin = presentes[i + 1][0] if i + 1 < len(presentes) else len(lineas)
        s = secciones[n]
        s['lineas'] = lineas[pos + 1:fin]
        if any(m < n and p > pos for p, m in presentes):
            s['marcas'].append('FUERA DE ORDEN: aparece antes que una sección de número menor')
        if len(s['lineas']) <= 1:
            contenido = f"«{s['lineas'][0][1]}»" if s['lineas'] else 'nada'
            s['marcas'].append(f'VACÍA: bajo el título solo hay {contenido}')
    preambulo = lineas[:presentes[0][0]] if presentes else lineas
    return {'secciones': secciones, 'preambulo': preambulo, 'avisos': avisos}


def _celda(s):
    """Página donde empieza la sección y letras de marca, para la tabla resumen."""
    letras = LETRA.get(s['estado'], '')
    for marca, letra in (('FUERA DE ORDEN', 'D'), ('REPETIDA', 'R'), ('VACÍA', 'V')):
        if any(m.startswith(marca) for m in s['marcas']):
            letras += letra
    return (str(s['pagina']) if s['pagina'] else '—') + letras


if __name__ == '__main__':
    sys.stdout.reconfigure(encoding='utf-8')
    oficiales = cargar_titulos_oficiales()
    destino = sys.argv[1]
    if os.path.isdir(destino):
        print('Página donde empieza cada sección. Marcas: F falta · T título de otra sección · '
              'N título no reconocido · D fuera de orden · R repetida · V vacía')
        print(f"{'ficha':<24}" + ''.join(f'{n:>5}' for n in range(1, 17)) + '   con marca')
        detalle = []
        for ruta in sorted(glob.glob(os.path.join(destino, '*.pdf'))):
            r = cortar_secciones(leer_pdf(ruta), oficiales)
            nombre = os.path.basename(ruta)[:-4]
            celdas = [_celda(r['secciones'][n]) for n in range(1, 17)]
            print(f'{nombre[:23]:<24}' + ''.join(f'{c:>5}' for c in celdas) + f"{sum(not c.isdigit() for c in celdas):>8}")
            for n, s in r['secciones'].items():
                if s['estado'] != OK or s['marcas']:
                    titulo = f"p. {s['pagina']} «{s['titulo']}»" if s['titulo'] else 'no hay título con ese número'
                    detalle.append(f"{nombre[:23]} · sección {n} · {s['estado']} · {titulo}" + ''.join(f'\n      - {m}' for m in s['marcas']))
            detalle += [f'{nombre[:23]} · aviso · {a}' for a in r['avisos']]
        print('\nDetalle de las marcas y avisos:')
        print('\n'.join(detalle) or '   (ninguno)')
    else:
        r = cortar_secciones(leer_pdf(destino), oficiales)
        for n, s in r['secciones'].items():
            print(f"{n:>2} {s['estado']:<24} p. {s['pagina'] or '—':<3} {len(s['lineas']):>4} líneas  {s['titulo'] or ''}")
            for m in s['marcas']:
                print(f'      - {m}')
        for a in r['avisos']:
            print('aviso:', a)
        if len(sys.argv) > 2:
            s = r['secciones'][int(sys.argv[2])]
            print(f"--- Sección {sys.argv[2]}: {s['titulo']} ---")
            print('\n'.join(f'p{pag:<3}{texto}' for pag, texto in s['lineas']))
