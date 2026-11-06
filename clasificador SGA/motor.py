# -*- coding: utf-8 -*-
"""Motor de clasificación estático. Sin modelo de lenguaje: solo reglas fijas que leen las tablas Excel.
Cada resultado lleva su rastro: de qué fila de Matriz_H y de qué regla de la hoja Reglas salió.
Hecho hasta ahora: M2 (nivel de peligro), M3 (palabra de advertencia), M4 (pictogramas), M5 (frases de la etiqueta), M6 (tamaño de etiqueta:  --envase 20) y M7 (índice Ir:  --ir 04_ L-07=15x8).
Uso de prueba:
    py "clasificador SGA/motor.py"                      (las 32 fichas ya extraídas)
    py "clasificador SGA/motor.py" 14_Acido_sulfurico   (detalle de una, por el inicio de su nombre)
Requiere: pip install openpyxl, y haber corrido antes  py extraccion/extraer.py FDS"""
import glob, json, os, re, sys
from tablas import cargar_tablas, RAIZ

SALIDAS = os.path.join(RAIZ, 'extraccion', 'salidas')
# Reglas de la hoja Reglas que este código aplica, con la huella de su texto el día en que se programaron.
# Si el autor cambia la redacción de una regla en el Excel, la huella deja de coincidir y el motor avisa:
# hay que revisar que el código siga haciendo lo que la regla dice.
REGLAS_PROGRAMADAS = {
    'R01': 'b34a4f9b632b',   # nivel = máximo de los niveles de las frases H
    'C04': 'c3a61eb85177',   # los peligros ambientales entran al máximo
    'R02': '998489297657',   # "Peligro" desplaza a "Atención"
    'C02': '111c6bdef8c1',   # peor caso cuando la ficha no precisa
    'R03': 'a02ce1b5bb72',   # GHS06 desplaza a GHS07
    'R04': '8ef6d8b1ac9d',   # GHS05 desplaza a GHS07 por H315 y H319
    'R05': 'ff39115abafd',   # H334 desplaza a GHS07 por H317, H315 y H319
    'R06': '28bd649b0e16',   # con H314 se omite H318; con H410 se omite H400
    'R07': '8d85ddaf27d5',   # sin frases H: etiqueta de producto no peligroso
    'C05': '2ecc726ec803',   # qué consejos P van en la etiqueta por defecto
    'R10': 'b89f3490b249',   # contenido mínimo en envases muy pequeños
    'C01': '3c858d6dc7e7',   # asfixiantes simples: solo se anota
    'C06': 'b13164229d21',   # semáforo del índice Ir con las bandas de la hoja Semaforo_Ir
    'C08': '98ebac4a2fe3',   # etiqueta de uso en la mina
    'C07': 'efd23f50eb78',   # selección de consejos P para minería subterránea
}


def avisos_de_reglas(tablas):
    """Reglas que cambiaron o desaparecieron en el Excel desde que se programaron."""
    avisos = []
    for regla, huella in REGLAS_PROGRAMADAS.items():
        actual = tablas['reglas'].get(regla)
        if actual is None:
            avisos.append(f'La regla {regla} ya no está en la hoja Reglas, pero el motor la sigue aplicando: revisar.')
        elif actual['huella'] != huella:
            avisos.append(f'El texto de la regla {regla} cambió en la hoja Reglas: revisar que el motor siga aplicándola bien.')
    return avisos


def separar_codigos(codigos):
    """Códigos H tal como llegan -> códigos sencillos de Matriz_H, sin repetir.
    'H315+H320' se parte en sus dos códigos; 'H360D' se resuelve con el código base 'H360' (LEEME de la matriz)."""
    salida = []
    for codigo in codigos:
        for parte in codigo.replace(' ', '').split('+'):
            base = parte[:4]
            if base not in [b for b, _ in salida]:
                salida.append((base, parte))
    return salida


def nivel_de_peligro(codigos_h, tablas):
    """M2. Nivel de peligro de un conjunto de frases H.
    Devuelve {'nivel', 'nombre', 'accion', 'determinantes', 'detalle', 'fuera_de_matriz', 'reglas', 'notas'}."""
    detalle, fuera = [], []
    for base, escrito in separar_codigos(codigos_h):
        fila = tablas['frases_h'].get(base)
        if fila is None:
            fuera.append(escrito)   # no se le inventa un nivel a un código que la matriz no tiene
            continue
        detalle.append(dict(codigo=base, escrito=escrito, nivel=fila['nivel'], nivel_texto=fila['nivel_texto'], tipo=fila['tipo'],
                            criterio=fila['criterio'], texto=fila['texto']))
    nivel = max((d['nivel'] for d in detalle), default=0)    # R01; sin frases H: nivel 0 (LEEME, punto 4)
    determinantes = [d['codigo'] for d in detalle if d['nivel'] == nivel] if detalle else []
    notas = []
    # C04: los peligros ambientales entran al máximo. Como es una decisión pendiente del autor, se deja dicho
    # cuándo el nivel se debe solo a ellos y cuál sería sin contarlos.
    if determinantes and all(d['tipo'] == 'Ambiente' for d in detalle if d['codigo'] in determinantes):
        sin_ambiente = max((d['nivel'] for d in detalle if d['tipo'] != 'Ambiente'), default=0)
        notas.append(f"El nivel {nivel} se debe solo a peligros ambientales ({', '.join(determinantes)}). "
                     f"Sin contarlos sería {sin_ambiente} ({tablas['niveles'][sin_ambiente]['nombre']}). Regla C04, pendiente de decisión.")
    if fuera:
        notas.append(f"Códigos que no están en Matriz_H y no entraron al cálculo: {', '.join(fuera)}.")
    return dict(nivel=nivel, nombre=tablas['niveles'][nivel]['nombre'], accion=tablas['niveles'][nivel]['accion'],
                determinantes=determinantes, detalle=detalle, fuera_de_matriz=fuera,
                reglas=['R01'] + (['C04'] if any(d['tipo'] == 'Ambiente' for d in detalle) else []), notas=notas)


def _clave_de_categoria(declarada, claves):
    """Busca la categoría que declara la ficha ('3', '2A', 'Tipo C', 'Category 1B') entre las que distingue
    Matriz_H para esa frase ('2', '3', 'C'...). Devuelve la clave que corresponde, o None."""
    texto = re.sub(r'(?i)categor[íi]a|category|tipo|type', ' ', str(declarada))
    for ficha in re.findall(r'\b(\d[A-C]?|[A-G])\b', texto):
        for clave in claves:
            if ficha == clave or (clave.isdigit() and ficha[0] == clave):   # '2A' cuenta como categoría 2
                return clave
    return None


def palabra_de_advertencia(codigos_h, categorias, tablas):
    """M3. Palabra de advertencia de la etiqueta.
    'categorias' = {código H: [categorías que la ficha declara para esa frase]} (paso 6 de la extracción).
    Devuelve {'palabra', 'detalle', 'reglas', 'notas'}; 'palabra' es None si ninguna frase lleva palabra."""
    detalle, notas, reglas = [], [], ['R02']
    for base, escrito in separar_codigos(codigos_h):
        fila = tablas['frases_h'].get(base)
        if fila is None:
            continue
        palabra, origen = fila['palabra'], 'Matriz_H'
        if fila['palabra_por_categoria']:   # la misma frase lleva una u otra palabra según la categoría
            opciones = fila['palabra_por_categoria']
            declaradas = categorias.get(escrito) or categorias.get(base) or []
            halladas = {opciones[c] for c in (_clave_de_categoria(d, opciones) for d in declaradas) if c}
            if len(halladas) == 1:
                palabra, origen = halladas.pop(), f"Matriz_H, por la categoría {', '.join(declaradas)} que declara la ficha"
            else:
                # C02: sin categoría clara se toma el peor caso, y se deja dicho para que alguien lo revise
                palabra = 'Peligro' if 'Peligro' in opciones.values() else 'Atención'
                motivo = 'la ficha no declara la categoría' if not declaradas else f"la categoría declarada ({', '.join(declaradas)}) no decide"
                origen = f'peor caso (regla C02): {motivo}'
                notas.append(f"{base}: {motivo}; la matriz dice «{fila['palabra_literal']}» y se tomó «{palabra}». Revisar.")
                if 'C02' not in reglas:
                    reglas.append('C02')
        detalle.append(dict(codigo=base, palabra=palabra, origen=origen, sin_palabra=fila['palabra_nota']))
    palabras = {d['palabra'] for d in detalle if d['palabra']}
    # R02: si alguna frase lleva "Peligro", la etiqueta lleva "Peligro" y no se muestra "Atención"
    final = 'Peligro' if 'Peligro' in palabras else 'Atención' if 'Atención' in palabras else None
    return dict(palabra=final, detalle=detalle, reglas=reglas, notas=notas)


def pictogramas(codigos_h, tablas):
    """M4. Pictogramas de la etiqueta, con las precedencias de la hoja Reglas.
    Devuelve {'pictogramas': [{codigo, simbolo, por}], 'omitidos': [{codigo, frase, regla, motivo}], 'reglas'}.
    Cada pictograma lleva las frases H que lo producen ('por'); una precedencia quita motivos, y un pictograma
    que se queda sin motivos no se muestra."""
    presentes = [base for base, _ in separar_codigos(codigos_h) if base in tablas['frases_h']]
    motivos = {}
    for codigo in presentes:
        for p in tablas['frases_h'][codigo]['pictogramas']:
            motivos.setdefault(p, []).append(codigo)
    omitidos, reglas = [], []

    def quitar(pictograma, frases, regla, motivo):
        for frase in [f for f in motivos.get(pictograma, []) if frases is None or f in frases]:
            motivos[pictograma].remove(frase)
            omitidos.append(dict(codigo=pictograma, frase=frase, regla=regla, motivo=motivo))
            if regla not in reglas:
                reglas.append(regla)

    # R03: si aplica GHS06 (calavera), no se muestra GHS07 (exclamación)
    if motivos.get('GHS06'):
        quitar('GHS07', None, 'R03', f"aplica GHS06 por {', '.join(motivos['GHS06'])}")
    # R04: si aplica GHS05 (corrosión), no se muestra GHS07 por irritación cutánea u ocular (H315, H319)
    if motivos.get('GHS05'):
        quitar('GHS07', ('H315', 'H319'), 'R04', f"aplica GHS05 por {', '.join(motivos['GHS05'])}")
    # R05: si aplica GHS08 por sensibilización respiratoria (H334), no se muestra GHS07 por sensibilización
    #      cutánea ni por irritación (H317, H315, H319)
    if 'H334' in presentes:
        quitar('GHS07', ('H317', 'H315', 'H319'), 'R05', 'aplica GHS08 por H334')

    visibles = [dict(codigo=p, simbolo=tablas['pictogramas'][p]['simbolo'], por=motivos[p]) for p in sorted(motivos) if motivos[p]]
    sin_pictograma = [c for c in presentes if not tablas['frases_h'][c]['pictogramas']]
    return dict(pictogramas=visibles, omitidos=omitidos, reglas=reglas, sin_pictograma=sin_pictograma)


def tamano_de_etiqueta(litros, tablas):
    """M6. Tamaño de la etiqueta y de los pictogramas según la capacidad del envase, en litros.
    Devuelve {'capacidad', 'etiqueta', 'pictograma', 'etiqueta_mm', 'pictograma_mm', 'fuente', 'reglas', 'notas'}.
    Las medidas se devuelven tal como están escritas en la hoja Etiqueta_tamano y, aparte, como números
    (ancho, alto) en milímetros para quien dibuje la etiqueta: 'minimo' y, si la tabla lo da, 'recomendado'."""
    if not isinstance(litros, (int, float)) or litros <= 0:
        raise ValueError(f'La capacidad del empaque debe ser un número de litros mayor que cero; se recibió «{litros}».')
    fila = next((f for f in tablas['etiqueta_tamano']
                 if (f['mas_de'] is None or litros > f['mas_de']) and (f['hasta'] is None or litros <= f['hasta'])), None)
    if fila is None:
        raise ValueError(f'La hoja Etiqueta_tamano no tiene una fila para {litros} litros.')

    def medidas(texto):
        pares = [(float(a.replace(',', '.')), float(b.replace(',', '.'))) for a, b in re.findall(r'([\d.,]+)\s*[×xX]\s*([\d.,]+)', str(texto))]
        obligatorio = not re.match(r'\s*si es posible', str(texto), re.I)   # "Si es posible, al menos..." no es un mínimo exigible
        return dict(minimo=pares[0] if pares and obligatorio else None, recomendado=pares[-1] if pares and (len(pares) > 1 or not obligatorio) else None)

    notas, reglas = [], []
    # R10: los envases muy pequeños llevan un contenido mínimo. El umbral se lee del texto de la regla, no del código.
    r10 = tablas['reglas'].get('R10')
    umbral = re.search(r'menos de\s+([\d.,]+)\s*ml', r10['tema'], re.I) if r10 else None
    if umbral and litros * 1000 < float(umbral.group(1).replace(',', '.')):
        reglas.append('R10')
        notas.append(f"Empaque de menos de {umbral.group(1)} ml. {r10['regla']} ({r10['fuente']}).")
    elif r10 and not umbral:
        notas.append('No se pudo leer el umbral de la regla R10 en la hoja Reglas; no se aplicó.')
    if tablas['etiqueta_fuente']:
        notas.append(f"Fuente de la tabla: {tablas['etiqueta_fuente']}")
    return dict(litros=litros, capacidad=fila['capacidad'], etiqueta=fila['etiqueta'], pictograma=fila['pictograma'],
                etiqueta_mm=medidas(fila['etiqueta']), pictograma_mm=medidas(fila['pictograma']),
                fuente=tablas['etiqueta_fuente'], reglas=reglas, notas=notas)


# Cómo se obtiene la exposición comparable con cada tipo de límite (hojas Calculo y Calculadora de tlv_limites.xlsx):
#   'jornada': promedio ponderado en el tiempo, Σ(ci × ti) / jornada.   'directa': la concentración tal cual.
FORMA_DE_COMPARAR = {'TLV-TWA': 'jornada', 'Valor guía': 'jornada', 'TLV-STEL': 'directa', 'TLV-C': 'directa', 'Máximo permitido': 'directa'}


def _unidad(u):
    return str(u or '').replace('³', '3').replace(' ', '').replace('μ', 'µ').lower()


def limites_de_la_sustancia(cas, tablas):
    """Filas de la hoja TLV de una sustancia, por su número CAS (uno o varios)."""
    propios = {cas} if isinstance(cas, str) else set(cas or [])
    return [l for l in tablas['limites'] if set(l['cas']) & propios]


def indice_ir(filas, mediciones, tablas):
    """M7. Índice de riesgo Ir = exposición / TLV, para las filas de límites de una sustancia.
    'mediciones' = {ID de la fila: {'tareas': [(concentración, horas), ...], 'unidad': 'ppm'}}; 'unidad' es opcional.
    Devuelve {'limites': [una entrada por fila], 'ir_mayor', 'limite_del_mayor', 'supera', 'alertas', 'notas'}.
    No convierte unidades, no aplica la regla aditiva de mezclas y no usa bandas intermedias."""
    jornada = tablas['jornada_horas']
    entradas, notas, alertas = [], [], []
    for l in filas:
        e = dict(id=l['id'], agente=l['agente'], tipo=l['tipo'], valor=l['valor'], unidad=l['unidad'], usar=l['usar'], uso=l['uso'],
                 mensaje=l['mensaje'], fraccion=l['fraccion'], exposicion=None, ir=None, supera=None, estado=None)
        medicion = mediciones.get(l['id'])
        tareas = (medicion or {}).get('tareas') or []
        forma = FORMA_DE_COMPARAR.get(l['tipo'])
        problema = None
        if medicion:
            if not tareas or any(not isinstance(v, (int, float)) or v < 0 for c, h in tareas for v in (c, h)):
                problema = 'la medición debe traer concentraciones y horas que sean números no negativos'
            elif medicion.get('unidad') and _unidad(medicion['unidad']) != _unidad(l['unidad']):
                problema = f"la medición está en {medicion['unidad']} y el límite en {l['unidad']}; no se convierten unidades"
            elif forma is None:
                problema = f"el tipo de límite «{l['tipo']}» no tiene forma de cálculo definida en la hoja Calculo"
            elif forma == 'jornada' and sum(h for _, h in tareas) > jornada:
                problema = (f"las horas suman {sum(h for _, h in tareas):g} y la jornada es de {jornada:g}; para turnos más largos falta "
                            'incorporar la corrección del Decreto 1886, art. 39, parágrafo 2 (hoja Calculo)')
        if problema:
            e['estado'] = 'medición no válida: ' + problema
        elif medicion:
            # Promedio ponderado por jornada para TWA y valor guía; concentración más alta para STEL, techo y máximo
            e['exposicion'] = sum(c * h for c, h in tareas) / jornada if forma == 'jornada' else max(c for c, _ in tareas)
            if l['usar'] == 'No':
                if l['uso'] == 'Alerta de seguridad' and isinstance(l['valor'], (int, float)):
                    e['supera'] = e['exposicion'] > l['valor']
                    e['estado'] = 'alerta de seguridad (no es un Ir)'
                    if e['supera']:
                        alertas.append(f"{l['agente']}: {e['exposicion']:g} {l['unidad']} supera el máximo permitido de {l['valor']:g} {l['unidad']} "
                                       f"({l['fraccion']}). {l['mensaje']}")
                else:
                    e['exposicion'] = None
                    e['estado'] = 'no se calcula Ir con este límite'
            else:
                e['ir'] = round(e['exposicion'] / l['valor'], 9)   # evita que un error de redondeo del computador pase de 1
                e['supera'] = e['ir'] > 1          # hoja Calculo: Ir > 1 se supera; Ir ≤ 1 no se supera
                e['estado'] = 'calculado' + (' (condicional: ver mensaje)' if l['usar'] == 'Condicional' else '')
        else:
            e['estado'] = 'no se calcula Ir con este límite' if l['usar'] == 'No' else 'sin medición'
        entradas.append(e)

    calculados = [e for e in entradas if e['ir'] is not None]
    mayor = max(calculados, key=lambda e: e['ir'], default=None)
    if len({e['agente'] for e in calculados}) > 1:
        notas.append('Hay mediciones de varios agentes: se reporta el mayor Ir por componente (hoja Decisiones, caso 4). '
                     'La regla aditiva de mezclas no se aplica porque la hoja Calculo la marca sin cotejar.')
    if not filas:
        notas.append('La sustancia no tiene filas en la hoja TLV: no se puede calcular Ir.')
    return dict(limites=entradas, ir_mayor=mayor['ir'] if mayor else None, limite_del_mayor=mayor['id'] if mayor else None,
                supera=mayor['supera'] if mayor else None, alertas=alertas, notas=notas,
                interpretacion=tablas['calculo']['Interpretación']['formula'] if 'Interpretación' in tablas['calculo'] else None)


def semaforo_ir(ir, nivel, tablas):
    """Regla C06. Color del riesgo por exposición, a partir del resultado de indice_ir y del nivel de peligro.
    Las bandas, el umbral de nivel y los mensajes se leen de la hoja Semaforo_Ir. Con varios agentes manda el mayor Ir.
    Devuelve {'clave', 'color', 'mensaje', 'ir', 'limite', 'regla', 'fila', 'condicion', 'estado_del_mensaje'}."""
    s = tablas['semaforo']
    medibles = [l for l in ir['limites'] if isinstance(l['valor'], (int, float)) and (l['usar'] != 'No' or l['uso'] == 'Alerta de seguridad')]
    if ir['alertas']:
        clave = 'rojo_alerta'
    elif ir['ir_mayor'] is not None:
        clave = next(c for c in ('verde', 'amarillo', 'rojo') if ir['ir_mayor'] >= s[c]['desde'] and (s[c]['hasta'] is None or ir['ir_mayor'] < s[c]['hasta']))
    elif any((l['estado'] or '').startswith('alerta') for l in ir['limites']):
        clave = 'alerta_ok'
    elif not medibles:
        clave = 'sin_limite'
    else:
        clave = 'gris_urgente' if nivel >= s['gris_urgente']['nivel_desde'] else 'gris'
    fila = s[clave]
    mensaje = fila['mensaje']
    if clave == 'rojo' and ir['supera'] and fila['complemento']:   # hoja Calculo: Ir > 1 supera el valor límite
        mensaje = f"{mensaje} {fila['complemento']}"
    limite = next((l for l in ir['limites'] if l['id'] == ir['limite_del_mayor']), None)
    return dict(clave=clave, color=fila['color'], mensaje=mensaje, ir=ir['ir_mayor'],
                limite=dict(id=limite['id'], agente=limite['agente'], tipo=limite['tipo'], valor=limite['valor'], unidad=limite['unidad']) if limite else None,
                regla='C06', fila=f'Semaforo_Ir, fila «{clave}»', condicion=fila['condicion'], estado_del_mensaje=fila['estado'])


def seleccion_de_consejos(consejos, codigos_h, tablas):
    """Regla C07. De los consejos P de la ficha (salida 'P' de frases_de_etiqueta) elige los más pertinentes para
    minería subterránea, hasta el máximo por tipo de la hoja Seleccion_P. El puntaje de cada consejo es la suma de
    los pesos de los criterios de Pertinencia_mineria que están activos (la ficha trae una frase H de sus capítulos
    del SGA) y que nombran ese consejo. No se rellena un tipo con otro. La etiqueta no usa esta selección.
    Devuelve {'elegidos': [...], 'disponibles': n, 'regla': 'C07'}; cada elegido lleva su rastro en 'criterio'."""
    presentes = [c for c in dict.fromkeys(parte[:4] for h in codigos_h for parte in h.replace(' ', '').split('+')) if c in tablas['sga_h_clase']]
    activos = []
    for criterio in tablas['pertinencia']:
        por = [h for h in presentes if tablas['sga_h_clase'][h]['capitulo'] in criterio['capitulos']]
        if por:
            activos.append((criterio, por))

    def nombra(criterio, codigo):
        partes = set(codigo.split('+'))
        return any(set(c.split('+')) <= partes for c in criterio['codigos'])

    puntuados = []
    for f in consejos:
        if not f['texto']:
            continue
        aplican = [(c, por) for c, por in activos if nombra(c, f['codigo'])]
        puntuados.append(dict(codigo=f['codigo'], tipo=f['tipo'], texto=f['texto'], completar=f['completar'], texto_ficha=f['texto_ficha'],
                              puntaje=sum(c['peso'] for c, _ in aplican),
                              criterio='; '.join(f"{c['id']} {c['via']} (peso {c['peso']:g}, por {', '.join(por)})" for c, por in aplican)
                                       or 'Sin criterio de pertinencia: entra porque sobra cupo de su tipo.',
                              filas=[f"Pertinencia_mineria, fila {c['id']}" for c, _ in aplican]))
    elegidos = []
    for tipo, maximo in tablas['seleccion_p'].items():
        if tipo != 'Total':
            del_tipo = sorted((x for x in puntuados if x['tipo'] == tipo), key=lambda x: (-x['puntaje'], x['codigo']))
            elegidos += del_tipo[:maximo]
    elegidos = sorted(elegidos, key=lambda x: (-x['puntaje'], x['codigo']))[:tablas['seleccion_p']['Total']]   # tope general
    orden = {v['tipo']: v['prioridad'] for v in tablas['prioridad_p'].values()}
    elegidos.sort(key=lambda x: (orden.get(x['tipo'], 9), -x['puntaje'], x['codigo']))
    return dict(elegidos=elegidos, disponibles=len(puntuados), regla='C07')


def cas_de_la_ficha(resultado):
    """CAS con que la ficha se busca en la tabla de límites: el del nombre del archivo o, si no lo trae, los que declara."""
    c = resultado['cotejo_cas']
    return [c['archivo']] if c['archivo'] else list(c['ficha'])


HAY_QUE_COMPLETAR = re.compile(r'\.\.\.|…')


def frases_de_etiqueta(frases_h, frases_p, tablas):
    """M5. Frases H y consejos P para la etiqueta, en español y ordenados.
    'frases_h' y 'frases_p' = [(código, texto tal como lo escribe la ficha)].
    Devuelve {'H': [...], 'P': [...], 'reglas', 'notas'}. Cada frase trae 'incluida' (si va en la etiqueta por
    defecto) y, cuando no va, el 'motivo'. La selección final de los consejos P es del usuario (regla C05)."""
    notas, reglas = [], []
    texto_ficha_h = {}
    for codigo, texto in frases_h:
        for parte in codigo.replace(' ', '').split('+'):
            texto_ficha_h.setdefault(parte[:4], (parte, texto))
    presentes = [c for c in texto_ficha_h if c in tablas['frases_h']]

    # R06: si aplica H314 se puede omitir H318; si aplica H410 se puede omitir H400
    omitir = {}
    for manda, cede in (('H314', 'H318'), ('H410', 'H400')):
        if manda in presentes and cede in presentes:
            omitir[cede] = f'se omite porque aplica {manda} (regla R06)'
            reglas.append('R06') if 'R06' not in reglas else None
    lista_h = []
    for codigo in presentes:
        escrito, texto = texto_ficha_h[codigo]
        lista_h.append(dict(codigo=codigo, escrito=escrito, texto=tablas['frases_h'][codigo]['texto'], texto_ficha=texto,
                            incluida=codigo not in omitir, motivo=omitir.get(codigo)))
    fuera = [texto_ficha_h[c][0] for c in texto_ficha_h if c not in tablas['frases_h']]
    if fuera:
        notas.append(f"Frases H que no están en Matriz_H y no van en la etiqueta: {', '.join(fuera)}.")

    lista_p, vistos = [], set()
    for codigo, texto in frases_p:
        codigo = codigo.replace(' ', '')
        if codigo in vistos:
            continue
        vistos.add(codigo)
        escala = tablas['prioridad_p'].get(codigo[1])
        if escala is None:
            notas.append(f'{codigo}: su tipo no está en la hoja Prioridad_P; no va en la etiqueta.')
            continue
        # Texto en español: el de Matriz_P si lo tiene (fuente de verdad del proyecto); si no, el oficial del SGA
        if codigo in tablas['frases_p']:
            oficial, fuente = tablas['frases_p'][codigo]['texto'], 'Matriz_P'
        elif codigo in tablas['sga_p']:
            oficial, fuente = tablas['sga_p'][codigo], 'SGA rev. 6, anexo 3'
        else:
            oficial, fuente = None, None
            notas.append(f'{codigo}: no tiene texto en español en Matriz_P ni en el diccionario del SGA.')
        lista_p.append(dict(codigo=codigo, tipo=escala['tipo'], prioridad=escala['prioridad'], texto=oficial, fuente_texto=fuente,
                            texto_ficha=texto, completar=bool(oficial and HAY_QUE_COMPLETAR.search(oficial)),
                            incluida=escala['por_defecto'] and oficial is not None,
                            motivo=None if escala['por_defecto'] else f"tipo {escala['tipo'].lower()}: no va por defecto; el usuario puede marcarlo (regla C05)"))
    lista_p.sort(key=lambda f: (f['prioridad'], f['codigo']))
    if lista_p:
        reglas.append('C05')
    if not lista_h:   # R07: producto sin frases H
        reglas.append('R07')
        notas.append('Sin frases H: etiqueta de producto no peligroso, con identificación, proveedor y consejos de prudencia (regla R07).')
    return dict(H=lista_h, P=lista_p, reglas=reglas, notas=notas)


# Solo asfixiantes SIMPLES (los que desplazan el oxígeno). Un asfixiante químico, como el CO, no entra en la regla C01.
ASFIXIANTE = re.compile(r'asfixiante simple|simple asphyxiant|displace oxygen|desplaza\w* (el |al )?ox[íi]geno', re.I)


def clasificar(ficha, tablas, envase_litros=None, mediciones=None):
    """M8. Resultado completo del motor para una ficha ya extraída.
    'envase_litros' y 'mediciones' son opcionales: los da el usuario en la interfaz.
    Junta nivel, palabra, pictogramas, frases de etiqueta, tamaño, límites e Ir, con su rastro, y conserva
    los hallazgos de la extracción. No combina nivel e Ir en una valoración final: eso está pendiente del autor."""
    entran, h_por_revisar = frases_h_de_la_ficha(ficha)
    pares_h, pares_p, p_por_revisar = frases_para_etiqueta_de_la_ficha(ficha)
    cas = cas_de_la_ficha(ficha)
    filas = limites_de_la_sustancia(cas, tablas)
    excel = next((tablas['excel_fichas'][c] for c in cas if c in tablas['excel_fichas']), None)
    palabra = palabra_de_advertencia(entran, ficha['clasificacion']['categoria_por_frase_H'], tablas)
    declarada = ficha['clasificacion']['palabra_advertencia']['valor']

    notas = []
    # C01: asfixiante simple sin frase H propia. La hoja Reglas dice que en esta versión no se aplica al cálculo; solo se anota.
    indicios = [f"Tabla de límites ({l['id']}): {l['mensaje']}" for l in filas if l['mensaje'] and ASFIXIANTE.search(l['mensaje'])]
    indicios += [f"Ficha: {h['texto']}" for h in ficha['hallazgos'] if ASFIXIANTE.search(h['texto'])]
    if excel and ASFIXIANTE.search(str(excel.get('frases_h') or '')):
        indicios.append(f"Excel de fichas (hoja reaccio): {excel['frases_h']}")
    if indicios and 'C01' in tablas['reglas']:
        notas.append(dict(regla='C01', texto=f"Asfixiante simple. {tablas['reglas']['C01']['regla']}", indicios=indicios))
    # Frases cuyo código confirmó el usuario en la interfaz: entran al cálculo, y se deja constancia
    confirmadas = [f for t in 'HP' for f in ficha['frases'][t] if f.get('metodo') == 'confirmada por el usuario']
    if confirmadas:
        notas.append(dict(regla=None, texto='Códigos confirmados por el usuario para frases que la extracción dejó «por revisar». Entran al cálculo.',
                          indicios=[f"{f['codigo']} para «{f['texto_ficha']}» (confirmado el {f.get('confirmada') or 'sin fecha'})" for f in confirmadas]))
    if declarada and declarada != palabra['palabra']:
        notas.append(dict(regla='R02', texto=f"La ficha declara la palabra «{declarada}» y el motor obtiene «{palabra['palabra']}». En la etiqueta va la del motor.", indicios=[]))
    if not filas:
        notas.append(dict(regla=None, texto='La sustancia no tiene filas en la hoja TLV (no hay número CAS en común): no se puede calcular Ir.', indicios=[]))

    apartados = ficha['reacciones']['apartados']
    nivel = nivel_de_peligro(entran, tablas)
    etiqueta = frases_de_etiqueta(pares_h, pares_p, tablas)
    ir = indice_ir(filas, mediciones or {}, tablas)
    return {
        'archivo': ficha['archivo'],
        'producto': ficha['identificacion']['nombre']['valor'],
        'cas': cas,
        'sustancia_en_excel': excel['sustancia'] if excel else None,
        'uso_en_mineria': excel['uso'] if excel else None,
        'proveedor': {c: ficha['identificacion']['proveedor'][c] for c in ('nombre', 'direccion', 'telefono', 'estado')},
        'nivel': nivel,
        'palabra_de_advertencia': dict(palabra, declarada_en_la_ficha=declarada),
        'pictogramas': pictogramas(entran, tablas),
        'etiqueta': etiqueta,
        'tamano_de_etiqueta': tamano_de_etiqueta(envase_litros, tablas) if envase_litros is not None else None,
        'indice_ir': ir,
        'semaforo': semaforo_ir(ir, nivel['nivel'], tablas),
        'seleccion_p': seleccion_de_consejos(etiqueta['P'], [f['codigo'] for f in etiqueta['H'] if f['incluida']], tablas),
        'reacciones': {'ficha': {k: dict(nombre=v['nombre'], texto=v['texto'], estado=v['estado']) for k, v in apartados.items()},
                       'excel_reacciones': excel['reacciones'] if excel else None,
                       'excel_incompatibilidades': excel['incompatibilidades'] if excel else None},
        'valoracion_final': None,   # pendiente del autor: cómo se combina el nivel de peligro con el índice Ir
        'notas': notas,
        'por_revisar': {'frases_h': h_por_revisar, 'consejos_p': p_por_revisar},
        'hallazgos_de_la_ficha': ficha['hallazgos'],   # se conservan tal como salieron de la extracción
        'avisos_de_reglas': avisos_de_reglas(tablas),
    }


def frases_h_de_la_ficha(resultado):
    """De un resultado de extracción: (códigos H que entran al cálculo, frases H por revisar que no entran)."""
    entran = [f['codigo'] for f in resultado['frases']['H'] if f['estado'] == 'encontrado']
    por_revisar = [f"«{f['texto_ficha']}»" + (f" ({f['codigo']})" if f.get('codigo') else '') for f in resultado['frases']['H']
                   if f['estado'] != 'encontrado']
    return entran, por_revisar


def frases_para_etiqueta_de_la_ficha(resultado):
    """De un resultado de extracción: (frases H, consejos P, consejos P por revisar), como pares (código, texto de la ficha).
    Solo entran las frases con estado 'encontrado'."""
    pares = lambda tipo: [(f['codigo'], f['texto_ficha']) for f in resultado['frases'][tipo] if f['estado'] == 'encontrado']
    p_por_revisar = [f"«{f['texto_ficha']}»" + (f" ({f['codigo']})" if f.get('codigo') else '') for f in resultado['frases']['P']
                     if f['estado'] != 'encontrado']
    return pares('H'), pares('P'), p_por_revisar


def cargar_fichas(prefijo=''):
    rutas = sorted(glob.glob(os.path.join(SALIDAS, prefijo + '*.json')))
    return [json.load(open(r, encoding='utf-8')) for r in rutas]


if __name__ == '__main__':
    sys.stdout.reconfigure(encoding='utf-8')
    tablas = cargar_tablas()
    for aviso in avisos_de_reglas(tablas):
        print('AVISO:', aviso)
    if '--envase' in sys.argv:   # M6: tamaño de etiqueta para una capacidad en litros, p. ej.  --envase 20
        t = tamano_de_etiqueta(float(sys.argv[sys.argv.index('--envase') + 1].replace(',', '.')), tablas)
        print(f"Envase de {t['litros']} L -> fila «{t['capacidad']}»")
        print(f"   Etiqueta:   {t['etiqueta']}   (mm: {t['etiqueta_mm']})")
        print(f"   Pictograma: {t['pictograma']}   (mm: {t['pictograma_mm']})")
        for nota in t['notas']:
            print('   Nota:', nota)
        sys.exit()
    if '--ir' in sys.argv:
        # M7: límites e índice Ir de una ficha.  --ir 04_  [L-07=15x8  L-07=30x2+10x6:ppm ...]
        #     ID=concentración x horas; varias tareas con "+"; unidad opcional tras ":"
        argumentos = sys.argv[sys.argv.index('--ir') + 1:]
        ficha = cargar_fichas(argumentos[0])[0]
        mediciones = {}
        for a in argumentos[1:]:
            identificador, resto = a.split('=')
            valores, _, unidad = resto.partition(':')
            mediciones[identificador] = {'tareas': [tuple(float(x.replace(',', '.')) for x in t.lower().split('x')) for t in valores.split('+')],
                                         'unidad': unidad or None}
        r = indice_ir(limites_de_la_sustancia(cas_de_la_ficha(ficha), tablas), mediciones, tablas)
        print(f"{ficha['archivo']}   (jornada: {tablas['jornada_horas']:g} horas)")
        for e in r['limites']:
            limite = f"{e['tipo']} {e['valor']:g} {e['unidad']}" if isinstance(e['valor'], (int, float)) else e['tipo']
            calculo = f" | exposición {e['exposicion']:g} {e['unidad']}" if e['exposicion'] is not None else ''
            calculo += f" | Ir = {e['ir']:.3g} -> {'SUPERA el valor límite' if e['supera'] else 'no supera el valor límite'}" if e['ir'] is not None else ''
            print(f"   {e['id']}  {str(e['agente'])[:38]:<39}{limite:<24}Usar: {e['usar']:<12}{e['estado']}{calculo}")
            if e['mensaje']:
                print(f"         mensaje: {e['mensaje']}")
        if r['ir_mayor'] is not None:
            print(f"Ir mayor: {r['ir_mayor']:.3g} (límite {r['limite_del_mayor']}) -> {'SUPERA el valor límite' if r['supera'] else 'no supera el valor límite'}"
                  f"   [{r['interpretacion']}]")
        for a in r['alertas']:
            print('ALERTA DE SEGURIDAD:', a)
        for nota in r['notas']:
            print('Nota:', nota)
        sys.exit()
    solo_etiqueta = '--etiqueta' in sys.argv   # resumen de las frases de la etiqueta (M5) en vez del de nivel, palabra y pictogramas
    if solo_etiqueta:
        sys.argv.remove('--etiqueta')
    fichas = cargar_fichas(sys.argv[1] if len(sys.argv) > 1 else '')
    if len(fichas) != 1 and solo_etiqueta:
        print('H: frases en la etiqueta / omitidas por R06.  P: en la etiqueta por defecto / disponibles para marcar / con "..." por completar / por revisar')
        print(f"{'ficha':<26}{'H etiq.':>8}{'H omit.':>8}{'P def.':>8}{'P disp.':>8}{'P compl.':>9}{'P rev.':>7}  notas")
    elif len(fichas) != 1:
        print(f"{'ficha':<24}{'nivel':<14}{'determinantes':<19}{'palabra':<10}{'la ficha dice':<15}{'pictogramas':<27}notas")
    for ficha in fichas:
        entran, por_revisar = frases_h_de_la_ficha(ficha)
        pares_h, pares_p, p_por_revisar = frases_para_etiqueta_de_la_ficha(ficha)
        e = frases_de_etiqueta(pares_h, pares_p, tablas)
        if len(fichas) != 1 and solo_etiqueta:
            print(f"{ficha['archivo'][:25]:<26}{sum(f['incluida'] for f in e['H']):>8}{sum(not f['incluida'] for f in e['H']):>8}"
                  f"{sum(f['incluida'] for f in e['P']):>8}{sum(not f['incluida'] for f in e['P']):>8}"
                  f"{sum(f['completar'] and f['incluida'] for f in e['P']):>9}{len(p_por_revisar):>7}  {' '.join(x for x in e['notas'] if 'R07' not in x)}"
                  + ('sin frases H (R07)' if 'R07' in e['reglas'] else ''))
            continue
        n = nivel_de_peligro(entran, tablas)
        p = palabra_de_advertencia(entran, ficha['clasificacion']['categoria_por_frase_H'], tablas)
        g = pictogramas(entran, tablas)
        declarada = ficha['clasificacion']['palabra_advertencia']['valor']
        # La palabra que declara la ficha es informativa: se compara, pero la de la etiqueta es la del motor
        comparacion = 'la ficha no declara palabra' if not declarada else 'coincide con la ficha' if declarada == p['palabra'] \
            else f'DIFIERE de la ficha, que declara «{declarada}»'
        if len(fichas) != 1:
            notas = ([f'{len(por_revisar)} frase(s) H por revisar, fuera del cálculo'] if por_revisar else []) + [x.split('.')[0] for x in n['notas']] + p['notas']
            if declarada and declarada != p['palabra']:
                notas.insert(0, 'PALABRA DIFIERE')
            notas += [f"sin {o['codigo']} por {o['frase']} ({o['regla']})" for o in g['omitidos']]
            print(f"{ficha['archivo'][:23]:<24}{str(n['nivel']) + ' ' + n['nombre']:<14}{', '.join(n['determinantes'])[:18]:<19}"
                  f"{p['palabra'] or '—':<10}{declarada or '—':<15}{' '.join(x['codigo'][3:] for x in g['pictogramas']) or '—':<27}{'; '.join(notas)}")
            continue
        print(f"{ficha['archivo']}\nNivel de peligro: {n['nivel']} — {n['nombre']}   (reglas: {', '.join(n['reglas'])})")
        print(f"Frases determinantes: {', '.join(n['determinantes']) or 'ninguna (sin frases H)'}")
        print(f"Acción genérica (hoja Niveles): {n['accion']}")
        print('Rastro, frase por frase (Matriz_H):')
        for d in n['detalle']:
            marca = ' <- determina' if d['codigo'] in n['determinantes'] else ''
            sufijo = f" (escrito {d['escrito']})" if d['escrito'] != d['codigo'] else ''
            print(f"   {d['codigo']}{sufijo}  nivel {d['nivel']} {d['nivel_texto']:<9} {d['tipo']:<9} {d['texto']}{marca}")
        for nota in n['notas']:
            print('Nota:', nota)
        print(f"Palabra de advertencia: {p['palabra'] or 'ninguna'}   (reglas: {', '.join(p['reglas'])}) — {comparacion}")
        for d in p['detalle']:
            print(f"   {d['codigo']}  {d['palabra'] or '(' + (d['sin_palabra'] or 'sin palabra') + ')':<28} {d['origen']}")
        for nota in p['notas']:
            print('Nota:', nota)
        print(f"Pictogramas: {', '.join(x['codigo'] for x in g['pictogramas']) or 'ninguno'}"
              + (f"   (reglas: {', '.join(g['reglas'])})" if g['reglas'] else ''))
        for x in g['pictogramas']:
            print(f"   {x['codigo']}  {x['simbolo']:<28} por {', '.join(x['por'])}")
        for o in g['omitidos']:
            print(f"   No se muestra {o['codigo']} por {o['frase']}: {o['motivo']} (regla {o['regla']})")
        if g['sin_pictograma']:
            print(f"   Frases que no llevan pictograma en la matriz: {', '.join(g['sin_pictograma'])}")
        if por_revisar:
            print('Frases H por revisar, que NO entraron al cálculo: ' + '; '.join(por_revisar))
        print(f"\nFrases para la etiqueta   (reglas: {', '.join(e['reglas'])})")
        print('Indicaciones de peligro (texto de Matriz_H | texto original de la ficha):')
        for f in e['H']:
            print(f"   {'[x]' if f['incluida'] else '[ ]'} {f['codigo']}  {f['texto']}" + (f"   <- {f['motivo']}" if f['motivo'] else ''))
            print(f"            ficha: {f['texto_ficha']}")
        print('Consejos de prudencia, por prioridad ([x] va por defecto; [ ] disponible para marcar; * hay que completar los "..."):')
        for f in e['P']:
            print(f"   {'[x]' if f['incluida'] else '[ ]'} {f['prioridad']} {f['tipo']:<15}{f['codigo']:<17}{'*' if f['completar'] else ' '} {f['texto']}   [{f['fuente_texto']}]")
            if f['completar']:
                print(f"{'':<44}ficha: {f['texto_ficha']}")
        for nota in e['notas']:
            print('Nota:', nota)
        if p_por_revisar:
            print('Consejos P por revisar, que NO van en la etiqueta: ' + '; '.join(p_por_revisar))
