# -*- coding: utf-8 -*-
"""Paso M1 del motor: carga de las tablas que gobiernan la clasificación.
El motor no lleva ningún valor normativo escrito en el código: niveles, palabras de advertencia,
pictogramas, prioridades, tamaños de etiqueta y límites se leen de los dos Excel del proyecto.
Si falta una hoja o una columna, se detiene y dice cuál; no sigue con datos a medias.
Uso de prueba:  py "clasificador SGA/tablas.py"
Requiere: pip install openpyxl"""
import hashlib, os, re, sys
import openpyxl

RAIZ = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..')
MATRIZ = os.path.join(RAIZ, 'matriz', 'matriz_decision_HP.xlsx')
LIMITES = os.path.join(RAIZ, 'tlv', 'tlv_limites.xlsx')
DICCIONARIO = os.path.join(RAIZ, 'matriz', 'frases_HP_es_en.xlsx')
FICHAS = os.path.join(RAIZ, 'matriz', 'fichas_de_datos_de_seguridad.xlsx')
PROTECCION = os.path.join(RAIZ, 'matriz', 'proteccion_seccion8.xlsx')
ES_CODIGO = re.compile(r'^[HP]\d{3}')
ES_PICTOGRAMA = re.compile(r'GHS\d{2}')


class ErrorDeTabla(Exception):
    """Una tabla no tiene la forma que el motor necesita."""


def _hoja(libro, ruta, nombre, columnas):
    """Filas de una hoja como diccionarios {columna: valor}, comprobando que estén las columnas pedidas."""
    archivo = os.path.basename(ruta)
    if nombre not in libro.sheetnames:
        raise ErrorDeTabla(f'En {archivo} falta la hoja «{nombre}». Hojas que hay: {", ".join(libro.sheetnames)}.')
    filas = list(libro[nombre].iter_rows(values_only=True))
    encabezado = [str(c).strip() if c is not None else '' for c in filas[0]]
    faltan = [c for c in columnas if c not in encabezado]
    if faltan:
        raise ErrorDeTabla(f'En {archivo}, hoja «{nombre}», falta la columna «{"», «".join(faltan)}». '
                           f'Columnas que hay: {", ".join(c for c in encabezado if c)}.')
    return [dict(zip(encabezado, f)) for f in filas[1:] if any(v is not None for v in f)]


def _palabra(celda):
    """Lee la celda 'Palabra de advertencia' de Matriz_H. Devuelve (palabra fija, {categoría: palabra}, nota).
    'Peligro' -> ('Peligro', {}, None)
    'Peligro (cat. 2) / Atención (cat. 3)' -> (None, {'2': 'Peligro', '3': 'Atención'}, None)
    '(sin palabra de advertencia)' -> (None, {}, 'sin palabra de advertencia')"""
    texto = str(celda or '').strip()
    if texto in ('Peligro', 'Atención'):
        return texto, {}, None
    por_categoria = {}
    for palabra, entre_parentesis in re.findall(r'(Peligro|Atención)\s*\(([^)]*)\)', texto):
        for categoria in re.findall(r'\b(\d[A-C]?|[A-G])\b', entre_parentesis.replace('cat.', '')):
            por_categoria[categoria] = palabra
    if por_categoria:
        return None, por_categoria, None
    return None, {}, texto.strip('()') or None


def _rango_litros(texto):
    """'Más de 3 y hasta 50 litros' -> (3.0, 50.0): más de 3 y hasta 50. None = sin límite por ese lado."""
    t = str(texto).lower()
    desde = re.search(r'más de\s+([\d.,]+)', t)
    hasta = re.search(r'hasta\s+([\d.,]+)', t)
    numero = lambda m: float(m.group(1).replace(',', '.')) if m else None
    return numero(desde), numero(hasta)


def _excel_de_fichas(ruta):
    """Datos descriptivos del Excel de fichas, por número CAS: uso en minería y reacciones.
    Son informativos (se muestran en el resultado); no entran a ningún cálculo. Una celda con dos CAS
    (cemento y óxido de calcio; acetileno y oxígeno) deja la misma fila bajo cada uno."""
    libro = openpyxl.load_workbook(ruta, read_only=True, data_only=True)
    try:
        por_cas = {}
        for f in _hoja(libro, ruta, 'busqueda 1 ', ['Sustancia', 'N° CAS', 'Uso en minería subterránea', 'Posibles reacciones']):
            for cas in re.findall(r'\d{2,7}-\d{2}-\d', str(f['N° CAS'] or '')):
                por_cas[cas] = dict(sustancia=f['Sustancia'], uso=f['Uso en minería subterránea'], reacciones=f['Posibles reacciones'],
                                    incompatibilidades=None, peligros=f.get('Peligros principales'),
                                    marco=f.get('Marco normativo colombiano aplicable'))
        for f in _hoja(libro, ruta, 'reaccio', ['Sustancia', 'CAS', 'Frases H (peligro)', 'Frases P clave', 'Uso en la mina', 'Incompatibilidades / reacciones']):
            for cas in re.findall(r'\d{2,7}-\d{2}-\d', str(f['CAS'] or '')):
                por_cas.setdefault(cas, dict(sustancia=f['Sustancia'], uso=f['Uso en la mina'], reacciones=None, incompatibilidades=None))
                por_cas[cas]['incompatibilidades'] = f['Incompatibilidades / reacciones']
                por_cas[cas]['frases_h'] = f['Frases H (peligro)']   # texto literal; aquí el autor anota "asfixiante simple"
                por_cas[cas]['frases_p'] = re.findall(r'P\d{3}(?:\s*\+\s*P\d{3})*', str(f['Frases P clave'] or ''))   # consejos P clave del inventario
        return por_cas
    finally:
        libro.close()


def cargar_tablas(matriz=MATRIZ, limites=LIMITES, diccionario=DICCIONARIO, fichas=FICHAS):
    """Devuelve un diccionario con todas las tablas, ya comprobadas."""
    libro = openpyxl.load_workbook(matriz, read_only=True, data_only=True)
    libro_tlv = openpyxl.load_workbook(limites, read_only=True, data_only=True)
    libro_sga = openpyxl.load_workbook(diccionario, read_only=True, data_only=True)
    try:
        t = _leer(libro, matriz, libro_tlv, limites)
        # Texto oficial en español de todos los consejos P del SGA, para los que Matriz_P no trae
        t['sga_p'] = {str(f['Código P']).strip(): f['Texto en español (SGA)']
                      for f in _hoja(libro_sga, diccionario, 'Frases_P', ['Código P', 'Texto en español (SGA)'])}
        # Clase de peligro de cada frase H y clases a las que el SGA asigna cada consejo P (anexo 3). Sirven solo para mostrar,
        # junto a cada peligro, los consejos de la ficha que le corresponden; no entran a ningún cálculo.
        capitulo = lambda texto: (re.search(r'cap[íi]tulo\s+(\d+\.\d+)', str(texto or '')) or [None, None])[1]
        t['sga_h_clase'] = {str(f['Código H']).strip(): dict(clase=re.sub(r'\s*\(cap[íi]tulo[^)]*\)', '', str(f['Clase de peligro (SGA)'] or '')),
                                                             capitulo=capitulo(f['Clase de peligro (SGA)']))
                            for f in _hoja(libro_sga, diccionario, 'Frases_H', ['Código H', 'Clase de peligro (SGA)'])}
        t['sga_p_clases'] = {}
        for f in _hoja(libro_sga, diccionario, 'P_por_clase', ['Código P', 'Capítulo del SGA', 'Categoría de peligro', 'Condiciones de uso']):
            t['sga_p_clases'].setdefault(str(f['Código P']).strip(), []).append(
                dict(capitulo=str(f['Capítulo del SGA']).strip(), categoria=f['Categoría de peligro'], condiciones=f['Condiciones de uso']))
        t['excel_fichas'] = _excel_de_fichas(fichas)
        # Los criterios de pertinencia solo pueden nombrar consejos P que existan en el SGA
        for criterio in t['pertinencia']:
            ajenos = [c for c in criterio['codigos'] if c not in t['sga_p']]
            if ajenos:
                raise ErrorDeTabla(f"En Pertinencia_mineria, {criterio['id']} nombra consejos que no están en el diccionario del SGA: {', '.join(ajenos)}.")
        t['proteccion'] = _proteccion()
        return t
    finally:  # se cierran siempre, para no dejar los Excel bloqueados
        libro.close()
        libro_tlv.close()
        libro_sga.close()


def _leer(libro, matriz, libro_tlv, limites):
    t = {}

    t['frases_h'] = {}
    for f in _hoja(libro, matriz, 'Matriz_H', ['Código H', 'Indicación de peligro', 'Tipo de peligro', 'Clase de peligro SGA', 'Categoría',
                                                'Palabra de advertencia', 'Pictograma', 'Nivel (0-4) — PROPUESTO', 'Nivel (texto)', 'Criterio del nivel']):
        codigo = str(f['Código H']).strip()
        if not ES_CODIGO.match(codigo):
            continue
        if not isinstance(f['Nivel (0-4) — PROPUESTO'], (int, float)):
            raise ErrorDeTabla(f'En Matriz_H, {codigo} no tiene un nivel numérico: «{f["Nivel (0-4) — PROPUESTO"]}».')
        fija, por_categoria, nota = _palabra(f['Palabra de advertencia'])
        t['frases_h'][codigo] = dict(
            texto=f['Indicación de peligro'], tipo=f['Tipo de peligro'], clase=f['Clase de peligro SGA'], categoria=str(f['Categoría']),
            palabra=fija, palabra_por_categoria=por_categoria, palabra_nota=nota, palabra_literal=f['Palabra de advertencia'],
            pictogramas=ES_PICTOGRAMA.findall(str(f['Pictograma'])), pictograma_literal=f['Pictograma'],
            nivel=int(f['Nivel (0-4) — PROPUESTO']), nivel_texto=f['Nivel (texto)'], criterio=f['Criterio del nivel'])

    t['niveles'] = {int(f['Nivel']): dict(nombre=f['Nombre'], criterio=f['Criterio (propuesto)'],
                                         accion=f['Acción genérica recomendada (propuesta, a validar)'])
                    for f in _hoja(libro, matriz, 'Niveles', ['Nivel', 'Nombre', 'Criterio (propuesto)',
                                                              'Acción genérica recomendada (propuesta, a validar)'])}

    # Prioridad de cada tipo de consejo P, por el primer dígito del código, y si va en la etiqueta por defecto (regla C05)
    t['prioridad_p'] = {}
    for f in _hoja(libro, matriz, 'Prioridad_P', ['Primer dígito', 'Tipo', 'Prioridad en etiqueta (1 = primero)', 'En la etiqueta por defecto',
                                                   'Sustento', 'Carácter']):
        if f['En la etiqueta por defecto'] not in ('Sí', 'No'):
            raise ErrorDeTabla(f'En Prioridad_P, «{f["Tipo"]}» tiene «En la etiqueta por defecto» = «{f["En la etiqueta por defecto"]}»; debe ser Sí o No.')
        t['prioridad_p'][str(f['Primer dígito']).strip()] = dict(
            tipo=f['Tipo'], prioridad=int(f['Prioridad en etiqueta (1 = primero)']), por_defecto=f['En la etiqueta por defecto'] == 'Sí',
            sustento=f['Sustento'], caracter=f['Carácter'])
    por_tipo = {v['tipo']: v['prioridad'] for v in t['prioridad_p'].values()}

    t['frases_p'], t['regla_p'] = {}, None
    for f in _hoja(libro, matriz, 'Matriz_P', ['Código P', 'Consejo de prudencia', 'Tipo', 'Prioridad en etiqueta (1 = primero)']):
        codigo = str(f['Código P']).strip()
        if codigo == 'Regla':
            t['regla_p'] = f['Consejo de prudencia']
        elif ES_CODIGO.match(codigo):
            prioridad = int(f['Prioridad en etiqueta (1 = primero)'])
            t['frases_p'][codigo] = dict(texto=f['Consejo de prudencia'], tipo=f['Tipo'], prioridad=prioridad)
            if por_tipo.get(f['Tipo']) != prioridad:
                raise ErrorDeTabla(f'Matriz_P da a {codigo} ({f["Tipo"]}) la prioridad {prioridad}, pero Prioridad_P dice {por_tipo.get(f["Tipo"])}.')

    # Cada regla guarda una huella de su texto: si el autor la cambia en el Excel, el motor puede avisar
    t['reglas'] = {str(f['ID']).strip(): dict(tema=f['Tema'], regla=f['Regla'], fuente=f['Fuente'], caracter=f['Carácter'],
                                             huella=hashlib.sha256(' '.join(str(f['Regla']).split()).encode('utf-8')).hexdigest()[:12])
                   for f in _hoja(libro, matriz, 'Reglas', ['ID', 'Tema', 'Regla', 'Fuente', 'Carácter'])}

    t['pictogramas'] = {str(f['Código']).strip(): dict(simbolo=f['Símbolo'], clases=f['Clases de peligro'])
                        for f in _hoja(libro, matriz, 'Pictogramas', ['Código', 'Símbolo', 'Clases de peligro'])}

    t['etiqueta_tamano'], t['etiqueta_fuente'] = [], None
    for f in _hoja(libro, matriz, 'Etiqueta_tamano', ['Capacidad del envase', 'Dimensiones de la etiqueta (mm)', 'Dimensiones de cada pictograma (mm)']):
        if str(f['Capacidad del envase']).strip() == 'Fuente':
            t['etiqueta_fuente'] = f['Dimensiones de cada pictograma (mm)'] or f['Dimensiones de la etiqueta (mm)']
            continue
        desde, hasta = _rango_litros(f['Capacidad del envase'])
        if desde is None and hasta is None:
            raise ErrorDeTabla(f'En Etiqueta_tamano no se pudo leer la capacidad «{f["Capacidad del envase"]}».')
        t['etiqueta_tamano'].append(dict(capacidad=f['Capacidad del envase'], mas_de=desde, hasta=hasta,
                                         etiqueta=f['Dimensiones de la etiqueta (mm)'], pictograma=f['Dimensiones de cada pictograma (mm)']))

    t['limites'] = []
    for f in _hoja(libro_tlv, limites, 'TLV', ['ID', 'N°', 'Sustancia', 'CAS', 'Agente que se mide', 'CAS del agente', 'Tipo de límite', 'Valor',
                                                'Unidad', 'Fracción o equivalencia', 'Notaciones', 'Fuente', 'Usar para Ir', 'Uso',
                                                'Mensaje para la herramienta']):
        if f['Usar para Ir'] not in ('Sí', 'No', 'Condicional'):
            raise ErrorDeTabla(f'En la hoja TLV, {f["ID"]} tiene «Usar para Ir» = «{f["Usar para Ir"]}»; debe ser Sí, No o Condicional.')
        if f['Usar para Ir'] != 'No' and not isinstance(f['Valor'], (int, float)):
            raise ErrorDeTabla(f'En la hoja TLV, {f["ID"]} está marcado «{f["Usar para Ir"]}» pero no tiene un valor numérico.')
        t['limites'].append(dict(id=f['ID'], numero=f['N°'], sustancia=f['Sustancia'], cas=re.findall(r'\d{2,7}-\d{2}-\d', str(f['CAS'] or '')),
                                 agente=f['Agente que se mide'], cas_agente=re.findall(r'\d{2,7}-\d{2}-\d', str(f['CAS del agente'] or '')),
                                 tipo=f['Tipo de límite'], valor=f['Valor'], unidad=f['Unidad'], fraccion=f['Fracción o equivalencia'],
                                 notaciones=f['Notaciones'], fuente=f['Fuente'], usar=f['Usar para Ir'], uso=f['Uso'],
                                 mensaje=f['Mensaje para la herramienta']))

    # Hoja Calculo: fórmulas y reglas del índice Ir, en texto. La jornada se lee de aquí, no del código.
    t['calculo'] = {str(f['Concepto']).strip(): dict(formula=f['Fórmula o regla'], aplicacion=f['Cómo se aplica'], fuente=f['Fuente'])
                    for f in _hoja(libro_tlv, limites, 'Calculo', ['Concepto', 'Fórmula o regla', 'Cómo se aplica', 'Fuente'])}
    jornada = re.search(r'(\d+(?:[.,]\d+)?)\s*horas', str(t['calculo'].get('Jornada laboral', {}).get('formula', '')))
    if not jornada:
        raise ErrorDeTabla('En la hoja Calculo no se pudo leer la jornada laboral (fila «Jornada laboral», por ejemplo «8 horas diarias»).')
    t['jornada_horas'] = float(jornada.group(1).replace(',', '.'))

    _leer_semaforo_y_seleccion(t, libro, matriz)
    _leer_etiqueta_de_mina(t, libro, matriz)

    # Coherencia entre hojas: lo que Matriz_H nombra tiene que existir en las otras
    for codigo, f in t['frases_h'].items():
        if f['nivel'] not in t['niveles']:
            raise ErrorDeTabla(f'Matriz_H asigna a {codigo} el nivel {f["nivel"]}, que no está en la hoja Niveles.')
        for p in f['pictogramas']:
            if p not in t['pictogramas']:
                raise ErrorDeTabla(f'Matriz_H asigna a {codigo} el pictograma {p}, que no está en la hoja Pictogramas.')
    return t


def _leer_semaforo_y_seleccion(t, libro, matriz):
    """Hojas de la matriz v0.3: semáforo del índice Ir (regla C06), selección de consejos P para minería (regla C07)
    y redacción en lenguaje claro de las frases H. Bandas, topes, pesos y mensajes salen de aquí, no del código."""
    numero = lambda v: float(v) if isinstance(v, (int, float)) else None
    t['semaforo'] = {}
    for f in _hoja(libro, matriz, 'Semaforo_Ir', ['Clave', 'Color en pantalla', 'Condición', 'Ir desde (incluido)', 'Ir hasta (sin incluir)',
                                                   'Nivel de peligro desde', 'Mensaje para el profesional de SST', 'Complemento si Ir > 1', 'Fuente', 'Estado']):
        if f['Clave'] in (None, 'Nota'):
            continue
        t['semaforo'][str(f['Clave']).strip()] = dict(
            color=str(f['Color en pantalla']).strip(), condicion=f['Condición'], desde=numero(f['Ir desde (incluido)']), hasta=numero(f['Ir hasta (sin incluir)']),
            nivel_desde=numero(f['Nivel de peligro desde']), mensaje=f['Mensaje para el profesional de SST'], complemento=f['Complemento si Ir > 1'],
            fuente=f['Fuente'], estado=f['Estado'])
    faltan = [c for c in ('verde', 'amarillo', 'rojo', 'rojo_alerta', 'alerta_ok', 'gris', 'gris_urgente', 'sin_limite') if c not in t['semaforo']]
    if faltan:
        raise ErrorDeTabla(f"En Semaforo_Ir faltan las filas: {', '.join(faltan)}.")
    bandas = [t['semaforo'][c] for c in ('verde', 'amarillo', 'rojo')]
    if any(b['desde'] is None for b in bandas) or bandas[0]['hasta'] != bandas[1]['desde'] or bandas[1]['hasta'] != bandas[2]['desde'] or bandas[2]['hasta'] is not None:
        raise ErrorDeTabla('En Semaforo_Ir las bandas verde, amarillo y rojo deben ir seguidas: cada una empieza donde termina la anterior, y la roja no tiene tope.')
    if t['semaforo']['gris_urgente']['nivel_desde'] is None:
        raise ErrorDeTabla('En Semaforo_Ir, la fila gris_urgente debe traer el «Nivel de peligro desde».')

    t['seleccion_p'] = {str(f['Tipo de consejo']).strip(): int(f['Máximo en la selección'])
                        for f in _hoja(libro, matriz, 'Seleccion_P', ['Tipo de consejo', 'Máximo en la selección'])}
    if 'Total' not in t['seleccion_p']:
        raise ErrorDeTabla('En Seleccion_P falta la fila «Total».')

    lista = lambda celda: [c.strip() for c in str(celda or '').split(',') if c.strip()]
    t['pertinencia'] = [dict(id=str(f['ID']).strip(), via=f['Vía o peligro bajo tierra'], peso=float(f['Peso']),
                             capitulos=lista(f['Capítulos del SGA (clases de peligro) que lo activan']), codigos=lista(f['Códigos P que responden']),
                             porque=f['Por qué importa en minería subterránea'], estado=f['Estado'])
                        for f in _hoja(libro, matriz, 'Pertinencia_mineria', ['ID', 'Vía o peligro bajo tierra', 'Peso', 'Capítulos del SGA (clases de peligro) que lo activan',
                                                                               'Códigos P que responden', 'Por qué importa en minería subterránea', 'Estado'])
                        if f['ID'] not in (None, 'Nota') and isinstance(f['Peso'], (int, float))]

    t['h_claro'] = {str(f['Código H']).strip(): dict(texto=f['En lenguaje claro'], estado=str(f['Estado'] or '').strip())
                    for f in _hoja(libro, matriz, 'H_lenguaje_claro', ['Código H', 'En lenguaje claro', 'Estado']) if ES_CODIGO.match(str(f['Código H'] or ''))}


def _leer_etiqueta_de_mina(t, libro, matriz):
    """Hojas de la matriz v0.4 para la etiqueta de uso en la mina (regla C08) y para las unidades de capacidad del empaque."""
    t['unidades'] = {}
    for f in _hoja(libro, matriz, 'Unidades_empaque', ['Unidad', 'Abreviatura', 'Litros por unidad', 'De dónde sale', 'Estado']):
        if not isinstance(f['Litros por unidad'], (int, float)) or f['Litros por unidad'] <= 0:
            raise ErrorDeTabla(f"En Unidades_empaque, «{f['Unidad']}» no tiene un número de litros por unidad mayor que cero.")
        t['unidades'][str(f['Unidad']).strip()] = dict(litros=float(f['Litros por unidad']), abreviatura=str(f['Abreviatura'] or f['Unidad']).strip(),
                                                       origen=f['De dónde sale'], estado=f['Estado'])
    if 'L' not in t['unidades']:
        raise ErrorDeTabla('En Unidades_empaque falta la unidad «L» (litros), que es la de la hoja Etiqueta_tamano.')

    t['p_corto'] = {str(f['Código P']).strip().replace(' ', ''): dict(texto=str(f['Texto corto para la etiqueta de uso en la mina']).strip(), estado=f['Estado'])
                    for f in _hoja(libro, matriz, 'P_corto', ['Código P', 'Texto corto para la etiqueta de uso en la mina', 'Estado'])
                    if ES_CODIGO.match(str(f['Código P'] or '')) and f['Texto corto para la etiqueta de uso en la mina']}
    malos = [c for c, v in t['p_corto'].items() if re.search(r'\.\.\.|…|\[|\]', v['texto'])]
    if malos:
        raise ErrorDeTabla(f"En P_corto hay textos con «...» o corchetes, que la etiqueta de uso en la mina no admite: {', '.join(malos)}.")

    t['etiqueta_mina'] = {str(f['Clave']).strip(): f['Valor'] for f in _hoja(libro, matriz, 'Etiqueta_mina', ['Clave', 'Valor', 'Para qué sirve', 'Estado']) if f['Clave']}
    faltan = [c for c in ('retiro_1', 'retiro_2', 'retiro_3', 'retiro_4', 'retiro_5', 'minimo_de_renglones', 'rotulo_peligro', 'rotulo_hacer', 'rotulo_epp',
                          'texto_sin_peligros', 'texto_asfixiante', 'texto_inventario', 'texto_no_es_res_773') if c not in t['etiqueta_mina']]
    if faltan:
        raise ErrorDeTabla(f"En Etiqueta_mina faltan las filas: {', '.join(faltan)}.")

    t['nombres_etiqueta'] = {str(f['N.º de ficha']).strip(): dict(nombre=str(f['Nombre en español para la etiqueta']).strip(), estado=f['Estado'])
                             for f in _hoja(libro, matriz, 'Nombres_etiqueta', ['N.º de ficha', 'Nombre en español para la etiqueta', 'Estado'])
                             if f['N.º de ficha'] not in (None, 'Nota') and f['Estado']}

    # Frases H que la propia matriz reconoce en el inventario para cada sustancia (hoja Aplicacion_30), por número CAS.
    # La etiqueta de uso en la mina las usa solo cuando la ficha no trae frases H en su sección 2.
    t['h_del_inventario'] = {}
    for f in _hoja(libro, matriz, 'Aplicacion_30', ['Sustancia', 'CAS', 'Frases H reconocidas']):
        codigos = [c for c in re.findall(r'H\d{3}', str(f['Frases H reconocidas'] or '')) if c in t['frases_h']]
        for cas in re.findall(r'\d{2,7}-\d{2}-\d', str(f['CAS'] or '')):
            t['h_del_inventario'][cas] = codigos


def _proteccion(ruta=PROTECCION):
    """Protección personal de la sección 8 de cada ficha, con su versión en español (matriz/proteccion_seccion8.xlsx).
    Es informativa. Si el archivo no está, la herramienta muestra el texto original de la ficha."""
    if not os.path.isfile(ruta):
        return {}
    libro = openpyxl.load_workbook(ruta, read_only=True, data_only=True)
    try:
        return {f['Ficha']: dict(pagina=f['Página'], idioma=f['Idioma del original'], original=f['Texto original de la ficha (sección 8)'],
                                 espanol=f['Protección personal en español'], estado=f['Estado'],
                                 # Columnas de la etiqueta de uso en la mina; si el archivo es anterior y no las trae, quedan vacías
                                 corto=f.get('EPP corto para etiqueta'), estado_corto=f.get('Estado del EPP corto'),
                                 iconos=[i.strip().lower() for i in str(f.get('Iconos de EPP') or '').split(',') if i.strip()])
                for f in _hoja(libro, ruta, 'Proteccion', ['Ficha', 'Página', 'Idioma del original', 'Texto original de la ficha (sección 8)',
                                                           'Protección personal en español', 'Estado']) if f['Ficha']}
    finally:
        libro.close()


if __name__ == '__main__':
    sys.stdout.reconfigure(encoding='utf-8')
    t = cargar_tablas()
    print(f"Matriz_H: {len(t['frases_h'])} frases H | niveles usados: {sorted({f['nivel'] for f in t['frases_h'].values()})}")
    print(f"Niveles: {', '.join(f'{n} {v['nombre']}' for n, v in t['niveles'].items())}")
    print(f"Matriz_P: {len(t['frases_p'])} consejos P | diccionario del SGA: {len(t['sga_p'])} consejos P con texto en español")
    print('Prioridad_P: ' + '; '.join(f"P{d}xx {v['tipo']} = {v['prioridad']} ({'por defecto' if v['por_defecto'] else 'disponible'})"
                                      for d, v in sorted(t['prioridad_p'].items(), key=lambda x: x[1]['prioridad'])))
    print(f"Reglas: {', '.join(t['reglas'])}")
    print(f"Pictogramas: {', '.join(t['pictogramas'])}")
    print('Etiqueta_tamano:')
    for e in t['etiqueta_tamano']:
        print(f"   más de {e['mas_de']} y hasta {e['hasta']} L -> etiqueta {e['etiqueta']} | pictograma {e['pictograma']}")
    usar = {}
    for l in t['limites']:
        usar[l['usar']] = usar.get(l['usar'], 0) + 1
    print(f"TLV: {len(t['limites'])} filas | Usar para Ir: {usar} | jornada (hoja Calculo): {t['jornada_horas']:g} horas")
    print('Frases H cuya palabra depende de la categoría:')
    for c, f in t['frases_h'].items():
        if f['palabra_por_categoria']:
            print(f"   {c}: {f['palabra_por_categoria']}   (celda: «{f['palabra_literal']}»)")
    print('Frases H sin palabra propia:')
    print('   ' + '; '.join(f"{c} ({f['palabra_nota']})" for c, f in t['frases_h'].items() if f['palabra_nota']))
    print('Frases H sin pictograma o con más de uno:')
    print('   sin: ' + ', '.join(c for c, f in t['frases_h'].items() if not f['pictogramas']))
    print('   varios: ' + ', '.join(f"{c} {f['pictogramas']}" for c, f in t['frases_h'].items() if len(f['pictogramas']) > 1))
    print('Semáforo del Ir: ' + '; '.join(f"{c} = {v['color']}" + (f" desde {v['desde']:g}" if v['desde'] is not None else '') for c, v in t['semaforo'].items()))
    print(f"Selección de consejos P: {t['seleccion_p']} | criterios de pertinencia: {', '.join(c['id'] for c in t['pertinencia'])}")
    print(f"Frases H en lenguaje claro: {len(t['h_claro'])} ({sum(v['estado'] == 'Validado' for v in t['h_claro'].values())} validadas) | protección personal: {len(t['proteccion'])} fichas")
    print(f"Unidades del empaque: {', '.join(f'{u} = {v['litros']:g} L' for u, v in t['unidades'].items())}")
    print(f"Etiqueta de uso en la mina: {len(t['p_corto'])} consejos con texto corto | {len(t['nombres_etiqueta'])} nombres en español | "
          f"{sum(bool(v) for v in t['h_del_inventario'].values())} sustancias con frases H en el inventario")
