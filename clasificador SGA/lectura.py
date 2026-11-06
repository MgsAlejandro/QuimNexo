# -*- coding: utf-8 -*-
"""Lectura práctica del resultado: ordena lo que ya extrajo la ficha y lo que ya calculó el motor para que lo
entienda un trabajador, un supervisor o un profesional de Seguridad y Salud en el Trabajo.
No clasifica, no calcula y no redacta recomendaciones propias. Todo lo que entrega sale de:
  - las frases H y P de la ficha (con su página) y su apartado de protección personal,
  - el resultado del motor (nivel de peligro, semáforo del índice Ir, selección de consejos para minería),
  - las tablas del proyecto (Matriz_H, Niveles, Reglas, Semaforo_Ir, H_lenguaje_claro, TLV, Calculo, proteccion_seccion8).
Lo único escrito aquí son rótulos y explicaciones de la herramienta. Ningún umbral está en este archivo: las bandas
del semáforo que nombran las explicaciones se leen de la tabla.
La interfaz (app.py) y el informe (reporte_html.py) usan esta misma lectura, para no decir cosas distintas.
Uso de prueba:  py "clasificador SGA/lectura.py" 02_"""
import re, sys

NOMBRE = 'QuimNexo'
VERSION = '2.03'   # se cambia al crear una versión nueva con crear_version.py
DESCRIPCION = ('Herramienta de apoyo para consultar Fichas de Datos de Seguridad sin tener que leerlas enteras: muestra los peligros del producto, '
               'el riesgo por exposición y qué hacer en la minería subterránea de carbón.')
CONTEXTO = 'Pensada para pequeñas y medianas minas subterráneas de carbón de la provincia de Sugamuxi (Boyacá, Colombia).'
NO_REEMPLAZA = ('No reemplaza la Ficha de Datos de Seguridad del fabricante, las mediciones de exposición, la evaluación del profesional '
                'de Seguridad y Salud en el Trabajo ni las obligaciones legales de la empresa.')
PERFILES = ('Profesional de SST', 'Trabajador/supervisor')
# Orden de la lista «qué revisar primero» (decisión del autor): rojo, amarillo, verde y gris; dentro de cada color, por nivel de peligro
ORDEN_DE_COLORES = ('rojo_alerta', 'rojo', 'amarillo', 'verde', 'alerta_ok', 'gris_urgente', 'gris', 'sin_limite')

# Rótulos de presentación para los tres tipos de peligro del SGA (columna «Tipo de peligro» de Matriz_H)
TIPOS = {'Físico': 'Incendio, explosión o presión', 'Salud': 'Salud', 'Ambiente': 'Ambiente'}
# Rótulos de presentación para los cinco tipos de consejo de prudencia del SGA (hoja Prioridad_P)
MOMENTOS = {'Prevención': 'Antes y durante el trabajo (prevención)',
            'Intervención': 'Si ocurre un accidente (qué hacer)',
            'Almacenamiento': 'Al guardar el producto (almacenamiento)',
            'Eliminación': 'Al desechar el producto o el empaque (eliminación)',
            'General': 'Generales'}
# Vía de entrada al cuerpo que nombra el propio texto de la frase H. No se deduce nada que la frase no diga.
VIAS = (('Inhalación (al respirarlo)', re.compile(r'inhal|respirator|somnolencia|vértigo', re.I)),
        ('Contacto con la piel', re.compile(r'\bpiel\b|cutáne', re.I)),
        ('Contacto con los ojos', re.compile(r'\bojos\b|ocular', re.I)),
        ('Ingestión (al tragarlo)', re.compile(r'ingesti|ingerir', re.I)))
SECCIONES_CLAVE = {'1': 'Identificación', '2': 'Identificación de peligros', '8': 'Controles de exposición y protección personal',
                   '10': 'Estabilidad y reactividad'}


def _num(valor):
    """Número con coma decimal."""
    return f'{valor:.3g}'.replace('.', ',')


def segun_cuantos(n, uno, varios):
    """El texto en singular si n es 1; si no, en plural."""
    return uno if n == 1 else varios


def bandas_en_texto(tablas):
    """Las tres bandas del semáforo, leídas de la hoja Semaforo_Ir, en una frase."""
    s = tablas['semaforo']
    return (f"verde si es menor que {_num(s['amarillo']['desde'])}; amarillo de {_num(s['amarillo']['desde'])} a menos de {_num(s['rojo']['desde'])}; "
            f"rojo desde {_num(s['rojo']['desde'])}, que es el nivel de intervención")


def glosario(tablas):
    """Términos técnicos que la herramienta no puede evitar, explicados en una línea."""
    return (
        ('FDS', 'Ficha de Datos de Seguridad: documento del fabricante con los peligros del producto y cómo manejarlo. Tiene 16 secciones.'),
        ('SGA', 'Sistema Globalmente Armonizado: reglas internacionales, adoptadas en Colombia, para clasificar y etiquetar productos químicos.'),
        ('Indicación de peligro (frase H)', 'Frase oficial que dice qué daño puede causar el producto. Lleva un código que empieza por H.'),
        ('Consejo de prudencia (frase P)', 'Frase oficial que dice qué hacer para evitar el daño o responder a un accidente. Su código empieza por P.'),
        ('Peligro del producto (nivel de peligro)', 'Escala de 0 a 4 de este proyecto que resume la gravedad de las frases H. Depende solo del producto. Es una propuesta en validación.'),
        ('Valor límite (TLV)', 'Concentración en el aire que la mayoría de los trabajadores puede respirar durante la jornada sin efectos adversos, según la ACGIH. '
                               'Es el límite que ordena usar la norma colombiana.'),
        ('Índice Ir', 'Exposición medida dividida por el valor límite. Si pasa de 1, se supera el límite.'),
        ('Riesgo por exposición (semáforo)', f'Color que resume el índice Ir: {bandas_en_texto(tablas)}. Sin medición es gris. Depende de lo que se midió en el puesto de trabajo.'),
        ('Nivel de intervención', f"Valor del índice Ir ({_num(tablas['semaforo']['rojo']['desde'])}) desde el cual la herramienta pide parar la operación, antes de llegar al valor límite."),
        ('Palabra de advertencia', '«Peligro» para los peligros más graves y «Atención» para los menos graves.'),
        ('Pictograma', 'Símbolo dentro de un rombo de borde rojo que avisa el tipo de peligro.'),
        ('Número CAS', 'Número único que identifica una sustancia química en todo el mundo.'),
        ('Asfixiante simple', 'Gas que no es tóxico por sí mismo pero desplaza el oxígeno del aire.'),
    )


def _paginas(ficha, tipo):
    """{código: página de la ficha donde está la frase}. Un código combinado (H302+H312) deja la página en cada parte."""
    paginas = {}
    for f in ficha['frases'][tipo]:
        if f.get('codigo') and f['estado'] == 'encontrado':
            for parte in [f['codigo'].replace(' ', '')] + f['codigo'].replace(' ', '').split('+'):
                paginas.setdefault(parte if tipo == 'P' else parte[:4], f.get('pagina'))
    return paginas


def _clases_del_consejo(codigo, tablas):
    """Filas de la hoja P_por_clase para un consejo; si es combinado y no está como tal, las de sus partes."""
    filas = tablas['sga_p_clases'].get(codigo)
    if filas is None:
        filas = [x for parte in codigo.split('+') for x in tablas['sga_p_clases'].get(parte, [])]
    return filas


def texto_del_indice(semaforo):
    """Línea pequeña bajo el semáforo: «Ir = 0,62 · TLV-TWA 15 ppm»."""
    if semaforo['ir'] is None or not semaforo['limite']:
        return None
    l = semaforo['limite']
    return f"Ir = {_num(semaforo['ir'])} · {l['tipo']} {l['valor']:g} {l['unidad']}"


def construir_lectura(ficha, r, tablas):
    """ficha = resultado de la extracción; r = salida de motor.clasificar.
    Devuelve un diccionario con lo que muestran la pantalla y el informe."""
    n, e, ir = r['nivel'], r['etiqueta'], r['indice_ir']
    pag_h, pag_p = _paginas(ficha, 'H'), _paginas(ficha, 'P')
    nueva = ficha.get('origen') == 'nueva'
    en_etiqueta = {f['codigo'] for f in e['H'] if f['incluida']}

    # --- Peligros: una entrada por frase H que entró al cálculo, de mayor a menor nivel ---
    peligros = []
    for d in n['detalle']:
        sga = tablas['sga_h_clase'].get(d['codigo'], {})
        claro = tablas['h_claro'].get(d['codigo'])
        peligros.append(dict(codigo=d['codigo'], texto=d['texto'], tipo=d['tipo'], nivel=d['nivel'], nivel_texto=d['nivel_texto'],
                             # El texto en lenguaje claro solo se usa cuando el autor lo marcó «Validado» en la hoja H_lenguaje_claro
                             claro=claro['texto'] if claro and claro['estado'] == 'Validado' and claro['texto'] else None,
                             criterio=d['criterio'], clase=tablas['frases_h'][d['codigo']]['clase'], capitulo=sga.get('capitulo'),
                             pagina=pag_h.get(d['codigo']), determinante=d['codigo'] in n['determinantes'], visible=d['codigo'] in en_etiqueta))
    peligros.sort(key=lambda p: (-p['nivel'], p['codigo']))
    por_capitulo = {}
    for p in peligros:
        if p['capitulo']:
            por_capitulo.setdefault(p['capitulo'], []).append(p['codigo'])
    visibles = [p for p in peligros if p['visible']]   # todas las de la ficha, menos las que la regla R06 permite omitir

    # --- Todos los consejos P de la ficha, por momento, con el peligro que controlan según el SGA (para el detalle técnico) ---
    elegidos = {x['codigo']: x for x in r['seleccion_p']['elegidos']}
    medidas = {}
    for f in e['P']:
        if not f['texto']:
            continue
        filas = [x for x in _clases_del_consejo(f['codigo'], tablas) if x['capitulo'] in por_capitulo]
        medidas.setdefault(f['tipo'], []).append(dict(
            codigo=f['codigo'], texto=f['texto'], completar=f['completar'], pagina=pag_p.get(f['codigo']),
            controla=list(dict.fromkeys(h for x in filas for h in por_capitulo[x['capitulo']])),
            condiciones=list(dict.fromkeys(x['condiciones'].lstrip('–- ').strip() for x in filas if x['condiciones'])),
            en_etiqueta=f['incluida'], seleccionado=f['codigo'] in elegidos))
    orden = sorted(medidas, key=lambda tipo: next((v['prioridad'] for v in tablas['prioridad_p'].values() if v['tipo'] == tipo), 9))
    grupos_de_medidas = [dict(tipo=t, titulo=MOMENTOS.get(t, t), consejos=medidas[t]) for t in orden]

    # --- Selección para minería (regla C07): cada consejo con su página y el criterio por el que se eligió ---
    seleccion = [dict(x, pagina=pag_p.get(x['codigo'])) for x in r['seleccion_p']['elegidos']]

    # --- Protección personal (sección 8): versión en español de la tabla, o el texto original si la ficha no está en ella ---
    tabla, leida = tablas['proteccion'].get(ficha['archivo']), ficha.get('proteccion') or {}
    if tabla and tabla['espanol'] and tabla['original'] == leida.get('texto'):
        proteccion = dict(texto=tabla['espanol'], pagina=tabla['pagina'], estado=tabla['estado'], original=False)
    elif leida.get('texto'):
        proteccion = dict(texto=leida['texto'], pagina=leida.get('pagina'), estado='Texto original de la ficha, sin traducir', original=True)
    else:
        proteccion = dict(texto=None, pagina=ficha['secciones'].get('8', {}).get('pagina'), estado='No se encontró en la ficha', original=False)

    # --- Lo que falta o hay que verificar ---
    pendientes = []
    if not r['producto']:
        pendientes.append('La ficha no dejó claro el nombre del producto. Verifíquelo en la sección 1 de la FDS.')
    faltan = [f"{k} ({v})" for k, v in SECCIONES_CLAVE.items() if ficha['secciones'].get(k, {}).get('estado') != 'OK']
    if faltan:
        pendientes.append('No se reconocieron bien estas secciones de la ficha: ' + '; '.join(faltan) + '. Lo que dependa de ellas puede estar incompleto.')
    if not peligros:
        pendientes.append('No se encontraron indicaciones de peligro (frases H) en la ficha. Puede ser un producto no clasificado como peligroso, '
                          'o una ficha que las escribe de forma que la herramienta no reconoció. Confírmelo en la sección 2 de la FDS.')
    por_confirmar = len(r['por_revisar']['frases_h']) + len(r['por_revisar']['consejos_p'])
    if r['por_revisar']['frases_h']:
        cuantas = len(r['por_revisar']['frases_h'])
        pendientes.append(f"{cuantas} " + segun_cuantos(cuantas, 'indicación de peligro de la ficha quedó sin código seguro y NO entró al nivel: ',
                                                  'indicaciones de peligro de la ficha quedaron sin código seguro y NO entraron al nivel: ')
                          + 'el peligro real puede ser mayor que el mostrado. ' + segun_cuantos(cuantas, 'Confírmela', 'Confírmelas') + ' en el detalle técnico.')
    if r['por_revisar']['consejos_p']:
        cuantas = len(r['por_revisar']['consejos_p'])
        pendientes.append(f"{cuantas} " + segun_cuantos(cuantas, 'consejo de prudencia de la ficha quedó sin código seguro y no se muestra. Revíselo',
                                                  'consejos de prudencia de la ficha quedaron sin código seguro y no se muestran. Revíselos')
                          + ' en el detalle técnico.')
    if peligros and not e['P']:
        pendientes.append('La ficha no trae consejos de prudencia (frases P). La herramienta no los inventa: consulte al proveedor o busque una FDS actualizada.')
    por_completar = [x['codigo'] for x in seleccion if x['completar']]
    if por_completar:
        pendientes.append('Consejos cuyo texto oficial trae «...»: hay que completarlos con el dato que da el fabricante en la ficha: ' + ', '.join(por_completar) + '.')
    declarada = r['palabra_de_advertencia'].get('declarada_en_la_ficha')
    if declarada and declarada != r['palabra_de_advertencia']['palabra']:
        pendientes.append(f"La ficha declara la palabra de advertencia «{declarada}» y las tablas dan «{r['palabra_de_advertencia']['palabra'] or 'ninguna'}». "
                          'Revise la clasificación en la sección 2.')
    if proteccion['texto'] and not proteccion['original'] and proteccion['estado'] != 'Validado':
        pendientes.append('La protección personal se muestra con una traducción resumida que el autor aún no ha validado: compárela con la sección 8 de la ficha.')
    if nueva:
        pendientes.append('Ficha cargada por el usuario: su lectura automática no ha sido revisada por una persona. Compárela con el PDF.')
    if r['hallazgos_de_la_ficha']:
        cuantas = len(r['hallazgos_de_la_ficha'])
        pendientes.append(f"La lectura de la ficha dejó {cuantas} " + segun_cuantos(cuantas, 'observación', 'observaciones')
                          + ' (por ejemplo, diferencias con el Excel de fichas). ' + segun_cuantos(cuantas, 'Está', 'Están') + ' en el detalle técnico.')

    excel = next((tablas['excel_fichas'][c] for c in r['cas'] if c in tablas['excel_fichas']), None)
    asfixiante = any(nota['regla'] == 'C01' for nota in r['notas'])
    lectura = dict(
        nombre=r['sustancia_en_excel'] or r['producto'] or r['archivo'], producto=r['producto'], cas=r['cas'], archivo=r['archivo'], nueva=nueva,
        peligros=peligros, visibles=visibles,
        grupos_de_peligros=[dict(tipo=t, titulo=TIPOS[t], peligros=[p for p in visibles if p['tipo'] == t]) for t in TIPOS if any(p['tipo'] == t for p in visibles)],
        vias=[nombre for nombre, patron in VIAS if any(patron.search(p['texto'] or '') for p in peligros if p['tipo'] == 'Salud')],
        asfixiante=asfixiante, medidas=grupos_de_medidas,
        que_hacer=[x for x in seleccion if x['tipo'] != 'Intervención'], accidente=[x for x in seleccion if x['tipo'] == 'Intervención'],
        consejos_disponibles=r['seleccion_p']['disponibles'], proteccion=proteccion,
        indice=texto_del_indice(r['semaforo']), accion_del_nivel=n['accion'], pendientes=pendientes, por_confirmar=por_confirmar,
        marco_normativo=(excel or {}).get('marco'),
        referencias=[dict(seccion=k, titulo=v, pagina=ficha['secciones'].get(k, {}).get('pagina'),
                          estado=ficha['secciones'].get(k, {}).get('estado') or 'no encontrada') for k, v in SECCIONES_CLAVE.items()])
    lectura['recomendaciones'] = recomendaciones_para_sst(r, lectura)
    lectura['factores'] = _factores(r, lectura)
    lectura['explicacion_del_nivel'] = _explicar_nivel(lectura, r, tablas)
    lectura['explicacion_del_ir'] = _explicar_ir(r, tablas)
    return lectura


def recomendaciones_para_sst(r, lectura, maximo=3):
    """Acción recomendada para el profesional de Seguridad y Salud en el Trabajo: a lo sumo tres puntos, del más urgente
    al menos. Cada uno sale de un dato concreto del resultado y lo dice ('porque'). No repite el mensaje del semáforo.
    La acción del nivel viene de la hoja Niveles; lo demás son verificaciones de lo que falta."""
    n = r['nivel']
    lista = []
    if r['por_revisar']['frases_h']:
        cuantas = len(r['por_revisar']['frases_h'])
        lista.append(dict(que='Confirme en el detalle técnico ' + segun_cuantos(cuantas, 'la indicación de peligro que quedó sin código seguro.',
                                                                                f'las {cuantas} indicaciones de peligro que quedaron sin código seguro.'),
                          porque='Mientras no se confirmen, no cuentan, y el nivel de peligro real puede ser mayor que el mostrado.'))
    if not lectura['peligros']:
        lista.append(dict(que='Compruebe en la sección 2 de la ficha que el producto no está clasificado como peligroso y que la ficha es la vigente del proveedor.',
                          porque='La herramienta no encontró indicaciones de peligro; puede ser cierto, o puede ser una ficha antigua o mal leída.'))
    else:
        lista.append(dict(que=n['accion'], porque=f"Es la acción de la tabla de niveles para el nivel de peligro {n['nivel']} ({n['nombre']}), "
                                                  f"que este producto alcanza por {', '.join(n['determinantes'])}. Propuesta en validación."))
    if lectura['asfixiante']:
        lista.append(dict(que='Incluya este gas en el control de la atmósfera de la mina (oxígeno), aunque no suba el nivel de peligro.',
                          porque='Es un asfixiante simple: desplaza el oxígeno y el SGA no le asigna una indicación de peligro por eso.'))
    if lectura['peligros'] and not lectura['medidas']:
        lista.append(dict(que='Pida al proveedor una ficha completa y actualizada.', porque='La ficha no trae consejos de prudencia y la herramienta no los inventa.'))
    if lectura['nueva']:
        lista.append(dict(que='Compare los datos leídos con el PDF de la ficha.',
                          porque='Es una ficha cargada por el usuario y su lectura automática no ha sido revisada por una persona.'))
    return lista[:maximo]


def _factores(r, lectura):
    """Datos que sostienen los dos indicadores, para el detalle técnico."""
    n, s = r['nivel'], r['semaforo']
    return [('Peligro del producto', f"Nivel {n['nivel']} de 4 ({n['nombre']})"
             + (f", por {', '.join(n['determinantes'])}" if n['determinantes'] else ', porque la ficha no trae indicaciones de peligro')
             + '. Regla R01, tabla Matriz_H.'),
            ('Riesgo por exposición', f"{s['color']}: {s['condicion']}" + (f" ({lectura['indice']})" if lectura['indice'] else '') + f". Regla {s['regla']}, {s['fila']}."),
            ('Consejos seleccionados', f"{len(r['seleccion_p']['elegidos'])} de los {r['seleccion_p']['disponibles']} que trae la ficha. Regla C07, hojas Seleccion_P y Pertinencia_mineria."),
            ('Información por verificar', f"{len(lectura['pendientes'])} " + segun_cuantos(len(lectura['pendientes']), 'punto.', 'puntos.'))]


def _explicar_nivel(lectura, r, tablas):
    """Seis preguntas sobre el nivel de peligro. Las respuestas citan las tablas; no añaden criterios."""
    n = r['nivel']
    regla = tablas['reglas'].get('R01', {})
    usados = '; '.join(f"{p['codigo']} «{p['texto']}»" + (f" (p. {p['pagina']} de la ficha)" if p['pagina'] else '') for p in lectura['peligros'])
    motivos = '; '.join(f"{p['codigo']} tiene nivel {p['nivel']} ({p['nivel_texto']}) porque: {p['criterio']}"
                        for p in lectura['peligros'] if p['determinante'])
    escala = '; '.join(f"{k} = {v['nombre']}" for k, v in sorted(tablas['niveles'].items()))
    return [
        ('¿Qué representa?',
         f"El peligro del producto ({escala}) resume qué tan grave es el daño que puede causar por sus propiedades, según las indicaciones de peligro "
         'de su ficha. No depende de cuánto se use ni de cuánto se mida: para eso está el semáforo del riesgo por exposición.'),
        ('¿Qué información se usó?',
         (f'Las indicaciones de peligro (frases H) leídas en la sección 2 de la ficha: {usados}.' if lectura['peligros']
          else 'La ficha no trae indicaciones de peligro reconocidas; por eso el nivel es 0.')),
        ('¿Cómo se obtuvo?',
         f"Cada frase H tiene un nivel en la tabla Matriz_H del proyecto. Regla R01 ({regla.get('caracter', '—')}): «{regla.get('regla', '—')}» "
         'El cálculo lo hace un programa fijo que lee esa tabla; ningún modelo de inteligencia artificial interviene.'),
        ('¿Por qué este producto quedó en este nivel?',
         (f"Quedó en nivel {n['nivel']} ({n['nombre']}) por la frase de mayor nivel. {motivos}." if motivos
          else f"Quedó en nivel {n['nivel']} ({n['nombre']}) porque no hay frases H que lo suban.")),
        ('¿Qué falta para mejorar la valoración?',
         'Los niveles 0 a 4 son una propuesta del trabajo de grado que aún está en validación. '
         + (f"En esta ficha hay {len(lectura['pendientes'])} " + segun_cuantos(len(lectura['pendientes']), 'punto', 'puntos') + ' por verificar.' if lectura['pendientes'] else '')),
        ('¿Qué decisiones apoya y cuáles no?',
         'Sirve para ordenar qué productos revisar primero, preparar la etiqueta, orientar la capacitación y decidir dónde conviene medir. '
         'No sirve, por sí solo, para afirmar que una labor es segura o insegura, ni para reemplazar la matriz de peligros de la empresa '
         'o el criterio del profesional de Seguridad y Salud en el Trabajo.'),
    ]


def _explicar_ir(r, tablas):
    """Seis preguntas sobre el índice Ir y su semáforo. Fórmulas de la hoja Calculo; bandas de la hoja Semaforo_Ir."""
    c, ir, s = tablas['calculo'], r['indice_ir'], tablas['semaforo']
    formula = (c.get('Índice de riesgo') or {}).get('formula') or 'Ir = exposición / TLV'
    twa = (c.get('Exposición promedio de la jornada (TWA)') or {}).get('formula')
    usables = [l for l in ir['limites'] if l['usar'] != 'No' and isinstance(l['valor'], (int, float))]
    limites = '; '.join(f"{l['agente']}: {l['tipo']} {l['valor']:g} {l['unidad']}" for l in usables)
    return [
        ('¿Qué representa?',
         'El índice Ir compara la concentración del producto que se midió en el aire del puesto de trabajo con el valor límite permitido '
         '(TLV de la ACGIH, que es el que ordena la norma colombiana). El semáforo lo convierte en un color.'),
        ('¿Qué información se usó?',
         (f'Los valores límite de la tabla del proyecto para este producto ({limites}) y la medición más reciente que usted guardó.' if usables
          else 'Este producto no tiene en la tabla del proyecto un valor límite con el que se pueda calcular el índice.')
         + ' La ficha de datos de seguridad no trae mediciones: esas las aporta la mina.'),
        ('¿Cómo se obtuvo?',
         f"{formula}." + (f" Para el promedio de la jornada: {twa} (ci = concentración de cada tarea, ti = horas de esa tarea)." if twa else '')
         + f" Jornada de {tablas['jornada_horas']:g} horas. Con varios agentes o varios límites, manda el mayor índice."),
        ('¿Cómo se lee?',
         f"El semáforo es {bandas_en_texto(tablas)}. La norma solo fija que por encima de 1 se supera el valor límite; las bandas son una decisión del autor "
         f"de este trabajo ({s['rojo']['fuente']}) para actuar antes de llegar al límite. Una alerta de seguridad superada también da rojo. "
         f"Sin medición el semáforo es gris; desde el nivel de peligro {s['gris_urgente']['nivel_desde']:g} pide medir con urgencia."),
        ('¿Qué falta para mejorar la valoración?',
         'Mediciones representativas, hechas con método validado y en las mismas unidades del límite. La herramienta no convierte unidades, '
         'no corrige jornadas de más de 8 horas y no suma el efecto de varias sustancias a la vez.'),
        ('¿Qué decisiones apoya y cuáles no?',
         'Apoya la decisión de parar, de reforzar controles y de programar nuevas mediciones. Un semáforo verde no garantiza ausencia de daño: no cubre el '
         'contacto con la piel, los peligros de incendio o explosión ni la sensibilidad de cada persona. Por eso el peligro del producto se muestra al lado, '
         'con el mismo tamaño.'),
    ]


def prioridad(resultados):
    """Ordena los productos para la lista «qué revisar primero»: por color del semáforo (rojo, amarillo, verde, gris) y,
    dentro de cada color, por nivel de peligro de mayor a menor. 'resultados' son salidas de motor.clasificar, o
    diccionarios con al menos 'semaforo', 'nivel' y un nombre."""
    return sorted(resultados, key=lambda r: (ORDEN_DE_COLORES.index(r['semaforo']['clave']), -r['nivel']['nivel'],
                                             (r.get('sustancia_en_excel') or r.get('producto') or r['archivo']).lower()))


if __name__ == '__main__':
    sys.stdout.reconfigure(encoding='utf-8')
    from tablas import cargar_tablas
    from motor import clasificar, cargar_fichas
    tablas = cargar_tablas()
    for ficha in cargar_fichas(sys.argv[1] if len(sys.argv) > 1 else '02_'):
        r = clasificar(ficha, tablas)
        x = construir_lectura(ficha, r, tablas)
        print(f"{x['nombre']}\n  Peligro: nivel {r['nivel']['nivel']} | Riesgo por exposición: {r['semaforo']['color']} ({r['semaforo']['clave']}) {x['indice'] or ''}")
        print('  Qué puede pasar: ' + '; '.join(f"{p['codigo']} {p['claro'] or p['texto']}" for p in x['visibles']))
        for z in x['que_hacer'] + x['accidente']:
            print(f"  {z['tipo']:<14} {z['codigo']:<16} p.{z['pagina']} | {z['criterio']}")
        print(f"  Protección personal (p. {x['proteccion']['pagina']}, {x['proteccion']['estado']}): {(x['proteccion']['texto'] or '—')[:110]}")
        for z in x['recomendaciones']:
            print('  ACCIÓN:', z['que'])
        print(f"  Por verificar: {len(x['pendientes'])}")
