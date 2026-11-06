# -*- coding: utf-8 -*-
"""Paso 6 de la extracción: clasificación que la FDS declara en la sección 2.
Saca, tal como las escribe la ficha, cada clase de peligro con su categoría, y la palabra
de advertencia. Sirve para saber la categoría de las frases H que cubren dos categorías
(p. ej. H272). Aquí no se clasifica: solo se transcribe lo declarado y se indica con qué
frases H de la misma ficha es compatible cada clase, según la clase y la categoría del SGA.
Uso de prueba:
    py extraccion/clasificacion.py FDS                     (resumen de todas)
    py extraccion/clasificacion.py FDS/02_Diesel_ACPM_68334-30-5.pdf   (detalle)
    py extraccion/clasificacion.py --examen FDS            (mide el enlace clase -> frase H)
Requiere: pip install pdfplumber openpyxl"""
import glob, os, re, sys, unicodedata
import openpyxl
from leer_pdf import leer_pdf
from secciones import cortar_secciones, cargar_titulos_oficiales, DICCIONARIO
from frases import cargar_diccionario, CODIGO, ENC, REV
from frases_texto import preparar, extraer_frases

NUM = r'(?:\d+(?:\.\d+)*\.?\s+)?'
FIN_BLOQUE = re.compile(r'^\s*' + NUM + r'(GHS[- ]?(US )?label|Label elements|Elementos de (la )?etiqueta|Signal word|'
                        r'Palabra de advertencia|Hazard pictograms|Pictogramas)', re.I)
# Categoría tal como la escriben las fichas
CATEGORIA = re.compile(r'\(?\b(?:Category|Categor[íi]a|Cat\.)\s*:?\s*(\d[A-C]?)\b\)?', re.I)
DIVISION = re.compile(r'\b(Divisi[óo]n|Division|Type|Tipo)\s+(\d\.\d|[A-G])\b', re.I)
GRUPO_GAS = re.compile(r'\b(Compressed gas|Refrigerated liquefied gas|Liquefied gas|Dissolved gas|Gas comprimido|'
                       r'Gas licuado refrigerado|Gas licuado|Gas disuelto)\b', re.I)
SIN_ROTULO = re.compile(r'^([A-Za-zÁÉÍÓÚáéíóúñ /]+?)\s+(\d[A-C]?)\s+[-–]\s')  # "Gas inflamable 1 - ..."
# Renglones que completan la clase anterior ("- repeated exposure", "hazard", "Target Organs - ...")
DETALLE = re.compile(r'^\s*([-–]\s*\S|[a-záéíóú]|Target Organs|[ÓO]rganos? diana)')
PALABRA = re.compile(r'^\s*(?:Signal word|Palabra de advertencia)(?:\s*\([^)]*\))?\s*:?\s*(.*)$', re.I)
# Las dos palabras de advertencia del SGA y su equivalente en español
PALABRAS = {'danger': 'Peligro', 'peligro': 'Peligro', 'warning': 'Atención', 'atención': 'Atención', 'atencion': 'Atención'}
VACIAS = set('de del la las los el y e o u en que sobre al a un una por para the of and or on for to an in chapter capitulo'.split())
ENLACE = 0.5  # parecido mínimo entre la clase de la ficha y la clase del SGA de una frase H


def _raices(texto):
    t = unicodedata.normalize('NFD', texto.lower())
    t = ''.join(ch for ch in t if unicodedata.category(ch) != 'Mn')
    return {p[:6] for p in re.findall(r'[a-z]+', t) if len(p) > 2 and p not in VACIAS}


def cargar_clases():
    """{código H: (raíces de su clase SGA en español e inglés, categorías SGA)} de la hoja Frases_H."""
    filas = list(openpyxl.load_workbook(DICCIONARIO, read_only=True)['Frases_H'].iter_rows(values_only=True))
    col = {nombre: i for i, nombre in enumerate(filas[0])}
    return {f[0]: (_raices(f[col['Clase de peligro (SGA)']]), _raices(f[col['Clase de peligro en inglés (SGA)']]),
                   str(f[col['Categoría (SGA)']])) for f in filas[1:]}


def _categoria(texto):
    """Categoría declarada en un texto: (valor literal, posición donde empieza, con rótulo claro)."""
    for patron, valor in ((CATEGORIA, lambda m: m.group(1)), (DIVISION, lambda m: f'{m.group(1)} {m.group(2)}'),
                          (GRUPO_GAS, lambda m: m.group(1))):
        m = patron.search(texto)
        if m:
            return valor(m), m.start(), True
    m = SIN_ROTULO.match(texto)
    return (m.group(2), m.end(1), False) if m else (None, None, False)


def extraer_clasificacion(corte):
    """Devuelve {'clases': [{clase, categoria, detalle, codigo_en_linea, estado, pagina, cita}],
                 'palabra_advertencia': {valor, literal, estado, pagina, cita}}."""
    lineas = corte['secciones'][2]['lineas']
    bloque = []
    for pag, linea in lineas:
        if FIN_BLOQUE.match(linea):
            break
        bloque.append((pag, linea))

    # 1) Renglones del bloque agrupados en entradas. En las fichas a dos columnas (clase | código y frase)
    #    se descarta desde el código hasta el fin del renglón y la categoría puede quedar en el renglón siguiente.
    entradas = []
    for pag, linea in bloque:
        m = CODIGO.search(linea)
        if m and m.group(1)[0] == 'H':
            entradas.append(dict(pagina=pag, partes=[linea[:m.start()]], cita=[linea], codigo=re.sub(r'\s+', '', m.group(1)), detalle=[]))
        elif entradas and entradas[-1]['codigo'] and _categoria(' '.join(entradas[-1]['partes']))[0] is None:
            entradas[-1]['partes'].append(linea)       # completa una entrada a dos columnas
            entradas[-1]['cita'].append(linea)
        elif _categoria(linea)[0] is not None:
            entradas.append(dict(pagina=pag, partes=[linea], cita=[linea], codigo=None, detalle=[]))
        elif entradas and DETALLE.match(linea) and not entradas[-1]['codigo']:
            entradas[-1]['detalle'].append(linea.strip())
            entradas[-1]['cita'].append(linea)

    clases = []
    for e in entradas:
        texto = ' '.join(p.strip() for p in e['partes'])
        valor, inicio, rotulada = _categoria(texto)
        if valor is None:
            continue
        clase = re.sub(r'\s+', ' ', texto[:inicio]).strip(' ,;:(-–')
        clases.append(dict(clase=clase, categoria=valor, detalle=' '.join(e['detalle']) or None, codigo_en_linea=e['codigo'],
                           estado=ENC if rotulada and clase else REV, pagina=e['pagina'], cita=' / '.join(e['cita'])))

    # 2) Palabra de advertencia declarada (puede venir en el renglón siguiente al rótulo)
    halladas = []
    for i, (pag, linea) in enumerate(lineas):
        m = PALABRA.match(linea)
        if m:
            literal = m.group(1).strip() or (lineas[i + 1][1].strip() if i + 1 < len(lineas) else '')
            halladas.append((literal, pag, linea if m.group(1).strip() else f'{linea} / {literal}'))
    validas = [(PALABRAS[l.lower()], l, p, c) for l, p, c in halladas if l.lower() in PALABRAS]
    if not halladas:
        palabra = dict(valor=None, literal=None, estado='no encontrado', pagina=None, cita=None)
    elif len({v[0] for v in validas}) == 1 and len(validas) == len(halladas):
        palabra = dict(valor=validas[0][0], literal=validas[0][1], estado=ENC, pagina=validas[0][2], cita=validas[0][3])
    else:  # no es una de las dos palabras del SGA, o la ficha trae dos distintas
        palabra = dict(valor=None, literal=' | '.join(l for l, _, _ in halladas), estado=REV,
                       pagina=halladas[0][1], cita=' | '.join(c for _, _, c in halladas))
    return {'clases': clases, 'palabra_advertencia': palabra}


def _compatibles(entrada, codigos_h, clases_sga):
    """Frases H de la ficha cuya clase del SGA se parece a la clase declarada y cuya categoría del SGA
    admite la categoría declarada. Es una consulta a la tabla del SGA, no una clasificación."""
    ficha = _raices(f"{entrada['clase']} {entrada['detalle'] or ''}")
    digito = re.match(r'\d', entrada['categoria'])
    puntajes = {}
    for codigo in codigos_h:
        base = codigo[:4]
        if base not in clases_sga or not ficha:
            continue
        es, en, categorias = clases_sga[base]
        if digito and digito.group(0) not in re.findall(r'\d', re.sub(r'\d\.\d', '', categorias)):
            continue
        if not digito and (re.search(r'\d', categorias) and not re.search(r'Divisi|Tipo', categorias)):
            continue  # la ficha da un grupo o división y el SGA pide número, o al revés
        parecido = max(2 * len(ficha & s) / (len(ficha) + len(s)) for s in (es, en) if s)
        if parecido >= ENLACE:
            puntajes[codigo] = parecido
    mejor = max(puntajes.values(), default=0)
    return [c for c, p in puntajes.items() if p >= mejor - 0.1]


def enlazar(clasificacion, frases, clases_sga):
    """Agrega a cada clase 'codigos_compatibles' y devuelve, por cada frase H de la ficha, la categoría
    declarada: {código: [categorías]}. Una frase sin clase declarada compatible queda con lista vacía."""
    codigos_h = [f['codigo'] for f in frases['H'] if f['codigo']]
    por_codigo = {c: [] for c in codigos_h}
    for e in clasificacion['clases']:
        e['codigos_compatibles'] = [e['codigo_en_linea']] if e['codigo_en_linea'] else _compatibles(e, codigos_h, clases_sga)
        for c in e['codigos_compatibles']:
            if c in por_codigo and e['categoria'] not in por_codigo[c]:
                por_codigo[c].append(e['categoria'])
    return por_codigo


if __name__ == '__main__':
    sys.stdout.reconfigure(encoding='utf-8')
    oficiales, diccionario, clases_sga = cargar_titulos_oficiales(), cargar_diccionario(), cargar_clases()
    indice = preparar(diccionario)
    modo_examen = sys.argv[1] == '--examen'
    destino = sys.argv[-1]
    rutas = sorted(glob.glob(os.path.join(destino, '*.pdf'))) if os.path.isdir(destino) else [destino]
    if modo_examen:
        # Entradas que traen el código en su renglón: se tapa y se mira si el enlace por clase lo recupera
        cuenta, fallos = {'acierta': 0, 'acierta, con otros': 0, 'no enlaza': 0, 'ERROR': 0}, []
        for ruta in rutas:
            corte = cortar_secciones(leer_pdf(ruta), oficiales)
            codigos_h = [f['codigo'] for f in extraer_frases(corte, diccionario, indice)['H'] if f['codigo']]
            for e in extraer_clasificacion(corte)['clases']:
                if not e['codigo_en_linea']:
                    continue
                dados = _compatibles(e, codigos_h, clases_sga)
                caso = ('acierta' if dados == [e['codigo_en_linea']] else 'acierta, con otros' if e['codigo_en_linea'] in dados
                        else 'no enlaza' if not dados else 'ERROR')
                cuenta[caso] += 1
                if caso != 'acierta':
                    fallos.append(f"  {caso:<19}{os.path.basename(ruta)[:14]:<15}{e['codigo_en_linea']:<6}«{e['clase']}» cat. {e['categoria']} -> {dados}")
        print(f'Clases con código en su renglón puestas a prueba: {sum(cuenta.values())}')
        print('\n'.join(f'  {k:<19}{v:>3}' for k, v in cuenta.items()))
        print('\n'.join(fallos))
        sys.exit()
    for ruta in rutas:
        corte = cortar_secciones(leer_pdf(ruta), oficiales)
        frases, clasif = extraer_frases(corte, diccionario, indice), extraer_clasificacion(corte)
        por_codigo = enlazar(clasif, frases, clases_sga)
        p = clasif['palabra_advertencia']
        sin_categoria = [c for c, cats in por_codigo.items() if not cats]
        if len(rutas) > 1:
            print(f"{os.path.basename(ruta)[:23]:<24}{len(clasif['clases']):>3} clases | palabra: {p['valor'] or p['literal'] or '—':<9}"
                  f"[{p['estado']}] | H sin categoría declarada: {', '.join(sin_categoria) or 'ninguna'}")
            continue
        print('Palabra de advertencia declarada: ' + (f"{p['valor'] or '?'} (la ficha dice «{p['literal']}», p. {p['pagina']}) [{p['estado']}]"
                                                      if p['literal'] else 'la ficha no la trae'))
        print(f"--- Clases declaradas: {len(clasif['clases'])} ---")
        for e in clasif['clases']:
            print(f"  p.{e['pagina']:<3}[{e['estado']}] {e['clase']} | categoría: {e['categoria']}"
                  + (f" | {e['detalle']}" if e['detalle'] else '') + f"  -> {', '.join(e['codigos_compatibles']) or 'sin frase H compatible en la ficha'}")
        print('--- Categoría declarada para cada frase H de la ficha ---')
        for c, cats in por_codigo.items():
            print(f"  {c:<8}{', '.join(cats) or 'SIN CLASE DECLARADA COMPATIBLE'}")
