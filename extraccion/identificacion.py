# -*- coding: utf-8 -*-
"""Paso 3 de la extracción: identificación del producto (sección 1 de la FDS).
Saca nombre del producto, número CAS y proveedor (nombre, dirección, teléfono).
Cada dato lleva su página, la cita literal y un estado:
    encontrado     estaba bajo un rótulo claro de la ficha
    por revisar    se propone, pero la ficha no lo rotula con claridad
    no encontrado  la ficha no lo trae donde se buscó
Los datos del proveedor son una propuesta que el usuario confirma en la interfaz.
Uso de prueba:
    py extraccion/identificacion.py FDS                           (tabla de todas)
    py extraccion/identificacion.py FDS/14_Acido_sulfurico_7664-93-9.pdf   (detalle)
Requiere: pip install pdfplumber openpyxl"""
import glob, os, re, sys
import openpyxl, pdfplumber
from leer_pdf import leer_pdf
from secciones import cortar_secciones, cargar_titulos_oficiales

FICHAS = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'matriz', 'fichas_de_datos_de_seguridad.xlsx')
ENC, REV, NO = 'encontrado', 'por revisar', 'no encontrado'
NUM = r'(?:\d+(?:\.\d+)*\.?\s+)?'  # numeración de subsección: "1.1. ", "1.3 "
CAS = re.compile(r'(?<![\d-])(\d{2,7})-(\d{2})-(\d)(?![\d-])')

# Rótulos, en orden de preferencia. Son palabras para ubicar el dato, no criterios de la norma.
ROTULOS_NOMBRE = ['Product name', 'Product Description', 'Substance name', 'Trade name',
                  'Nombre del producto', 'Nombre comercial', 'Identificación del producto']
INICIO_PROVEEDOR = [r"Details of the supplier", r"Manufacturer or supplier'?s details", r"Manufacturer/Supplier",
                    r"Supplier", r"Company", r"Datos del proveedor", r"Detalles del proveedor",
                    r"Proveedor", r"Fabricante"]
# Líneas que solo rotulan y no son dato
SOLO_ROTULO = re.compile(r'^(Company|Supplier|Supplier Address|Manufacturer|Fabricante|Proveedor|'
                         r'(UK|EU) entity/business name)\s*:?$', re.I)
ROTULO_CON_VALOR = re.compile(r'^(Company|Manufacturer/Supplier|Supplier|Dirección|Address)\s*:?\s+(?=\S)', re.I)
FIN_PROVEEDOR = re.compile(r'^(' + NUM.replace('?:', '') + r')?(Emergency|N[úu]meros? (de )?tel[eé]fonos? de emergencia|Tel[eé]fono de emergencia|'
                           r'N[úu]mero telef[óo]nico de emergencia|Recommended use|Uses advised|Relevant identified)', re.I)
TELEFONO = re.compile(r'(?:\b(?:T\s?el[eé]fono|T\s?elephone|Tel|Tlf|T)\b\s*\.?\s*:?\s*)([+(]?\d[\d\s()+\-./]{5,}\d)', re.I)
NO_DIRECCION = re.compile(r'^(Fax|T\s?elefax|E-?mail|Correo|electr[óo]nico|Web|Sitio web|www\.|\S+@\S+)', re.I)


def _rotulo(texto):
    """Expresión que encuentra un rótulo aunque el PDF separe su primera letra ('P roduct name')."""
    return re.compile('^' + NUM + re.escape(texto[0]) + r'\s?' + re.escape(texto[1:]).replace(r'\ ', r'\s+') + r'\b\s*:?\s*(.*)$', re.I)


def cas_valido(cas):
    """Dígito de control del número CAS: descarta números con forma de CAS que no lo son."""
    m = CAS.fullmatch(cas)
    cifras = (m.group(1) + m.group(2))[::-1]
    return sum((i + 1) * int(c) for i, c in enumerate(cifras)) % 10 == int(m.group(3))


def _dato(valor, estado, pagina=None, cita=None, **extra):
    return dict(valor=valor, estado=estado, pagina=pagina, cita=cita, **extra)


def _nombre(lineas):
    for texto in ROTULOS_NOMBRE:
        patron = _rotulo(texto)
        for i, (pag, linea) in enumerate(lineas):
            m = patron.match(linea)
            if not m:
                continue
            if m.group(1).strip():
                return _dato(m.group(1).strip(), ENC, pag, linea)
            if i + 1 < len(lineas):  # el valor viene en la línea siguiente
                return _dato(lineas[i + 1][1].strip(), ENC, lineas[i + 1][0], f'{linea} / {lineas[i + 1][1]}')
    if lineas:  # sin rótulo: se propone la primera línea de la sección, para que alguien la revise
        return _dato(lineas[0][1].strip(), REV, lineas[0][0], lineas[0][1], nota='sin rótulo: es la primera línea de la sección 1')
    return _dato(None, NO)


def _cas_seccion1(lineas):
    for i, (pag, linea) in enumerate(lineas):
        if not re.search(r'\bCAS\b', linea):
            continue
        siguiente = lineas[i + 1][1] if i + 1 < len(lineas) else ''
        for texto in (linea, siguiente):
            for m in CAS.finditer(texto):
                if cas_valido(m.group(0)):
                    return _dato(m.group(0), ENC, pag, linea if texto is linea else f'{linea} / {siguiente}')
    return _dato(None, NO)


def _cas_seccion3(lineas):
    """Todos los CAS válidos que nombra la sección 3 (composición), con su línea."""
    vistos, salida = set(), []
    for pag, linea in lineas:
        for m in CAS.finditer(linea):
            if cas_valido(m.group(0)) and m.group(0) not in vistos:
                vistos.add(m.group(0))
                salida.append(_dato(m.group(0), ENC, pag, linea))
    return salida


def _separar_columnas(ruta, pagina, bloque):
    """Si el bloque trae dos empresas lado a lado (líneas como 'One Reagent Lane One Reagent Lane'),
    devuelve (columna izquierda, columna derecha) usando la posición de las palabras en la página."""
    def duplicada(t):
        p = t.split()
        return len(t) >= 10 and len(p) % 2 == 0 and p[:len(p) // 2] == p[len(p) // 2:]
    if not any(duplicada(t) for t in bloque):
        return None
    with pdfplumber.open(ruta) as pdf:
        palabras = pdf.pages[pagina - 1].extract_words()
    renglones = []  # palabras agrupadas por renglón (misma altura, con 3 puntos de tolerancia)
    for w in sorted(palabras, key=lambda w: w['top']):
        if renglones and w['top'] - renglones[-1][0]['top'] <= 3:
            renglones[-1].append(w)
        else:
            renglones.append([w])
    por_texto = {' '.join(w['text'] for w in sorted(r, key=lambda w: w['x0'])): sorted(r, key=lambda w: w['x0'])
                 for r in renglones}
    cortes = [por_texto[t][len(por_texto[t]) // 2]['x0'] for t in bloque if duplicada(t) and t in por_texto]
    if not cortes:
        return None
    corte = min(cortes) - 1
    izq, der = [], []
    for t in bloque:
        ws = por_texto.get(t)
        if ws is None:
            izq.append(t)
            continue
        izq.append(' '.join(w['text'] for w in ws if w['x0'] < corte))
        der.append(' '.join(w['text'] for w in ws if w['x0'] >= corte))
    return [t for t in izq if t], [t for t in der if t]


def _proveedor(lineas, ruta):
    inicio = None
    for texto in INICIO_PROVEEDOR:
        patron = re.compile('^' + NUM + texto + r'\b', re.I)
        inicio = next((i for i, (_, l) in enumerate(lineas) if patron.match(l)), None)
        if inicio is not None:
            break
    if inicio is None:
        return dict(nombre=None, direccion=None, telefono=None, estado=NO, pagina=None, cita=None, notas=[])

    pagina, primera = lineas[inicio]
    notas, bloque = [], []
    resto = ROTULO_CON_VALOR.sub('', re.sub('^' + NUM, '', primera))
    if resto != re.sub('^' + NUM, '', primera):      # "Company : Sigma-Aldrich Inc."
        bloque.append(resto)
    entidades = 0
    for _, linea in lineas[inicio + 1:inicio + 12]:
        if FIN_PROVEEDOR.match(linea) or re.match(r'^1\.\d', linea):
            break
        if re.search(r'entity/business name', linea, re.I):
            entidades += 1
            if entidades > 1:
                notas.append('la ficha nombra un segundo proveedor, que no se tomó: revisar')
                break
        bloque.append(linea)
    cita = ' / '.join([primera] + [l for _, l in lineas[inicio + 1:inicio + 1 + len(bloque)]])

    columnas = _separar_columnas(ruta, pagina, bloque)
    if columnas:
        bloque, derecha = columnas
        notas.append('la ficha trae dos empresas en columnas; se tomó la izquierda. La otra: ' + ', '.join(derecha))

    telefono = next((m.group(1).strip() for l in bloque for m in [TELEFONO.search(l)] if m), None)
    utiles = []
    for l in bloque:
        if SOLO_ROTULO.match(l):
            continue
        l = ROTULO_CON_VALOR.sub('', l)
        m = TELEFONO.search(l)
        if m:  # quita el teléfono y lo que le sigue (fax, etc.)
            l = l[:m.start()].rstrip(' -–.,')
        if l and not NO_DIRECCION.match(l):
            utiles.append(l)
    utiles = [u for u in utiles if not re.fullmatch(r'[\d\s]+', u)]  # restos de un número partido en dos líneas
    nombre = utiles[0] if utiles else None
    direccion = ', '.join(u.rstrip(',') for u in utiles[1:]) or None
    if nombre and not direccion:
        # Nombre y dirección en una sola línea: se parte en el primer guion o coma, y se avisa
        partes = re.split(r'\s*[-–]\s+|,\s+', nombre, maxsplit=1)
        if len(partes) == 2:
            nombre, direccion = partes
            notas.append('nombre y dirección venían en una sola línea; se separaron en el primer guion o coma: revisar')
    if not (nombre and direccion and telefono):
        notas.append('falta ' + ', '.join(c for c, v in (('nombre', nombre), ('dirección', direccion), ('teléfono', telefono)) if not v))
    return dict(nombre=nombre, direccion=direccion, telefono=telefono, estado=REV if notas else ENC,
                pagina=pagina, cita=cita, notas=notas)


def extraer_identificacion(doc, corte):
    """doc = salida de leer_pdf; corte = salida de cortar_secciones."""
    s1, s3 = corte['secciones'][1]['lineas'], corte['secciones'][3]['lineas']
    return {'nombre': _nombre(s1), 'cas_seccion1': _cas_seccion1(s1), 'cas_seccion3': _cas_seccion3(s3),
            'proveedor': _proveedor(s1, doc['ruta'])}


def cargar_excel_fichas():
    """{CAS: nombre de la sustancia} de la hoja 'busqueda 1 ' del Excel de fichas.
    Una celda puede traer varios CAS o un CAS dentro de un texto ('Mezcla (... 68937-41-7)')."""
    hoja = openpyxl.load_workbook(FICHAS, read_only=True, data_only=True)['busqueda 1 ']
    return {m.group(0): sustancia for sustancia, cas, *_ in hoja.iter_rows(min_row=2, values_only=True)
            for m in CAS.finditer(str(cas or ''))}


def cotejar_cas(ruta, ident, excel):
    """Compara el CAS de la ficha con el del nombre del archivo y con el Excel. No corrige nada."""
    m = CAS.search(os.path.basename(ruta))
    archivo = m.group(0) if m else None
    en_ficha = [ident['cas_seccion1']['valor']] if ident['cas_seccion1']['valor'] else []
    en_ficha += [c['valor'] for c in ident['cas_seccion3'] if c['valor'] not in en_ficha]
    if not archivo:
        veredicto = 'el nombre del archivo no trae CAS'
    elif not en_ficha:
        veredicto = 'LA FICHA NO TRAE CAS'
    elif ident['cas_seccion1']['valor'] == archivo:
        veredicto = 'coincide'
    elif archivo in en_ficha:
        veredicto = 'coincide (solo en sección 3)'
    else:
        veredicto = 'DIFIERE'
    en_excel = next((c for c in ([archivo] if archivo else []) + en_ficha if c in excel), None)
    # Hallazgos: se conservan y se muestran al usuario; pueden pesar en el índice de riesgo
    # y en las reacciones con otras sustancias. No se descartan ni se corrigen aquí.
    hallazgos = []
    if excel.get(en_excel):
        hallazgos.append(f"Producto de la ficha: «{ident['nombre']['valor']}». Sustancia en el Excel de fichas: «{excel[en_excel]}». "
                         'Verificar que la ficha corresponda al producto que se usa en la mina.')
    else:
        hallazgos.append('Ningún CAS de la ficha está en el Excel de fichas.')
    otros = [c for c in ident['cas_seccion3'] if c['valor'] != archivo]
    if otros:
        hallazgos.append('La sección 3 declara otros componentes: ' + '; '.join(f"{c['valor']} («{c['cita']}»)" for c in otros))
    if veredicto == 'DIFIERE':
        hallazgos.append(f"El CAS de la ficha ({', '.join(en_ficha)}) no es el del nombre del archivo ({archivo}).")
    return {'archivo': archivo, 'ficha': en_ficha, 'veredicto': veredicto,
            'excel': excel.get(en_excel), 'cas_excel': en_excel, 'hallazgos': hallazgos}


if __name__ == '__main__':
    sys.stdout.reconfigure(encoding='utf-8')
    oficiales, excel = cargar_titulos_oficiales(), cargar_excel_fichas()
    destino = sys.argv[1]
    rutas = sorted(glob.glob(os.path.join(destino, '*.pdf'))) if os.path.isdir(destino) else [destino]
    for ruta in rutas:
        doc = leer_pdf(ruta)
        ident = extraer_identificacion(doc, cortar_secciones(doc, oficiales))
        c, p, n = cotejar_cas(ruta, ident, excel), ident['proveedor'], ident['nombre']
        print(f'##### {os.path.basename(ruta)}')
        print(f"  Nombre     [{n['estado']}] {n['valor']}" + (f"   ({n['nota']})" if n.get('nota') else ''))
        print(f"  En Excel   {c['excel'] or '— (ningún CAS de la ficha está en el Excel de fichas)'}"
              + (f"   (por el CAS {c['cas_excel']})" if c['cas_excel'] and c['cas_excel'] != c['archivo'] else ''))
        print(f"  CAS        sección 1: {ident['cas_seccion1']['valor'] or '—'} | sección 3: "
              f"{', '.join(x['valor'] for x in ident['cas_seccion3']) or '—'} | archivo: {c['archivo'] or '—'} -> {c['veredicto']}")
        print(f"  Proveedor  [{p['estado']}] {p['nombre']}")
        print(f"  Dirección  {p['direccion']}")
        print(f"  Teléfono   {p['telefono']}")
        for nota in p['notas']:
            print(f'      - {nota}')
        for h in c['hallazgos']:
            print(f'  Hallazgo   {h}')
        if len(rutas) == 1:
            print(f"  Cita nombre (p. {n['pagina']}): {n['cita']}")
            print(f"  Cita proveedor (p. {p['pagina']}): {p['cita']}")
