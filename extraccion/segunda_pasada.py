# -*- coding: utf-8 -*-
"""Paso 10 de la extracción: segunda pasada OPCIONAL con modelo de lenguaje.
Solo mira lo que los pasos 3 y 5 dejaron 'por revisar' y solo extrae; nunca clasifica:
  - Frases sin código: el modelo elige un código del diccionario del SGA o 'NINGUNO'.
    El programa descarta cualquier código que no esté en el diccionario.
  - Nombre del producto y proveedor dudosos: el modelo los señala en la sección 1.
    El programa descarta todo dato que no aparezca literal en esa sección.
Lo aceptado queda marcado método 'LLM'. Sin credenciales, el módulo no hace nada y la
extracción sigue igual, con esos datos 'por revisar'.
Funciona con dos proveedores, a elección en configuracion.json: Anthropic (Claude) u OpenAI.
Los controles son los mismos con cualquiera de los dos. Al proveedor elegido se le envían
únicamente los fragmentos dudosos de la ficha.
Se usa desde extraer.py:
    py extraccion/extraer.py FDS --llm             (proveedor de configuracion.json)
    py extraccion/extraer.py FDS --llm openai      (o: --llm anthropic)
Requiere la librería y la credencial del proveedor elegido:
    Anthropic: pip install anthropic   y la variable de entorno ANTHROPIC_API_KEY
    OpenAI:    pip install openai      y la variable de entorno OPENAI_API_KEY"""
import json, os, re, unicodedata

CONFIGURACION = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'configuracion.json')
PROVEEDORES = ('anthropic', 'openai')
ENC, REV = 'encontrado', 'por revisar'
NINGUNO = 'NINGUNO'

INSTRUCCIONES_FRASES = """Vas a ayudar a transcribir Fichas de Datos de Seguridad (FDS) de productos químicos.
Recibirás frases copiadas de la sección 2 de una FDS que no traen su código, y la lista oficial de \
{nombre} del Sistema Globalmente Armonizado (SGA, sexta edición revisada) con su código y su texto en español e inglés.

Para cada frase, indica el código de la lista cuyo texto oficial dice lo mismo que la frase. Los fabricantes \
reescriben, abrevian o traducen las frases, así que no hace falta que coincidan palabra por palabra, pero sí en \
significado: mismo peligro o misma medida, misma gravedad y misma vía. En el texto oficial, "..." es un dato que \
completa el fabricante, "/" separa alternativas y lo que va entre paréntesis son instrucciones.

Responde {ninguno} cuando la frase no corresponda a ninguna de la lista, cuando sea más amplia o más estrecha que \
la oficial de modo que cambie su sentido, o cuando dudes entre dos códigos. Es preferible {ninguno} a un código dudoso: \
una persona revisará esas frases después.

Compara solo el texto de la frase con los textos oficiales. No uses lo que sepas del producto químico ni decidas \
qué peligro tiene: eso lo resuelve otro proceso con reglas fijas.

Lista oficial de {nombre}:
{tabla}"""

INSTRUCCIONES_IDENTIFICACION = """Vas a ayudar a transcribir la sección 1 de una Ficha de Datos de Seguridad (FDS).
Recibirás el texto de esa sección tal como salió del PDF; a veces dos columnas quedan mezcladas en un mismo renglón.

Señala el nombre del producto y los datos del proveedor o fabricante (nombre de la empresa, dirección y teléfono \
de contacto, no el de emergencias). Cada dato debe ser una copia exacta de un fragmento del texto, sin corregir, \
traducir ni completar nada: un programa comprobará que aparezca literal y lo descartará si no. Para la dirección, \
si ocupa varios renglones, sepáralos con ", ". Si un dato no está en el texto, déjalo vacío."""

ESQUEMA_IDENTIFICACION = {
    'type': 'object', 'additionalProperties': False,
    'required': ['nombre_producto', 'proveedor_nombre', 'proveedor_direccion', 'proveedor_telefono'],
    'properties': {c: {'type': 'string'} for c in ('nombre_producto', 'proveedor_nombre', 'proveedor_direccion', 'proveedor_telefono')},
}


def leer_configuracion():
    """Proveedor y modelos de configuracion.json: {'proveedor': ..., 'modelos': {proveedor: modelo}}."""
    with open(CONFIGURACION, encoding='utf-8') as f:
        return json.load(f)['modelo_de_lenguaje']


def crear_consulta(proveedor=None):
    """Prepara la consulta al proveedor elegido. Devuelve (consulta, origen, motivo):
    'consulta(instrucciones, mensaje, esquema)' -> (datos, aviso); 'origen' = {proveedor, modelo}.
    Si falta la librería o la credencial, consulta es None y 'motivo' lo explica (la extracción sigue sin esta pasada)."""
    configuracion = leer_configuracion()
    proveedor = proveedor or configuracion['proveedor']
    if proveedor not in PROVEEDORES:
        return None, None, f"proveedor desconocido: «{proveedor}» (opciones: {', '.join(PROVEEDORES)})"
    modelo = configuracion['modelos'][proveedor]
    origen = {'proveedor': proveedor, 'modelo': modelo}
    try:
        if proveedor == 'anthropic':
            import anthropic
            cliente = anthropic.Anthropic()   # toma la credencial del entorno; no se escribe en el código
            return (lambda i, m, e: _preguntar_anthropic(cliente, modelo, i, m, e)), origen, None
        import openai
        cliente = openai.OpenAI()             # ídem, de la variable OPENAI_API_KEY
        return (lambda i, m, e: _preguntar_openai(cliente, modelo, i, m, e)), origen, None
    except ImportError:
        return None, origen, f'no está instalada la librería {proveedor} (pip install {proveedor})'
    except Exception as e:  # falta de credencial: cada librería la señala con su propia clase de error
        return None, origen, f'no hay credencial para {proveedor}: {e}'


def _leer_json(texto):
    try:
        return json.loads(texto), None
    except (TypeError, ValueError):
        return None, 'la respuesta no vino en el formato pedido'


def _preguntar_openai(cliente, modelo, instrucciones, mensaje, esquema):
    """Una consulta a OpenAI (Responses API) con respuesta en JSON que cumple 'esquema'. Devuelve (datos, aviso)."""
    import openai
    try:
        respuesta = cliente.responses.create(
            model=modelo,
            input=[{'role': 'system', 'content': instrucciones}, {'role': 'user', 'content': mensaje}],
            text={'format': {'type': 'json_schema', 'name': 'extraccion_fds', 'schema': esquema, 'strict': True}},
        )
    except openai.AuthenticationError:
        return None, 'la API rechazó la credencial'
    except openai.RateLimitError:
        return None, 'la API pidió esperar (límite de uso o de saldo); intente más tarde'
    except openai.APIStatusError as e:
        return None, f'la API devolvió un error ({e.status_code})'
    except openai.APIConnectionError:
        return None, 'no hubo conexión con la API'
    if respuesta.status == 'incomplete':
        return None, f'el modelo no completó la respuesta ({respuesta.incomplete_details.reason})'
    for salida in respuesta.output:   # un rechazo viene como contenido de tipo "refusal" y no cumple el esquema
        for parte in getattr(salida, 'content', None) or []:
            if parte.type == 'refusal':
                return None, 'el modelo declinó responder'
    return _leer_json(respuesta.output_text)


def _preguntar_anthropic(cliente, modelo, instrucciones, mensaje, esquema):
    """Una consulta a Claude con respuesta en JSON que cumple 'esquema'. Devuelve (datos, aviso)."""
    import anthropic
    try:
        respuesta = cliente.beta.messages.create(
            model=modelo,
            max_tokens=16000,
            # Si el modelo declina por un filtro de seguridad (posible con textos de explosivos o tóxicos),
            # la API reintenta con el modelo de respaldo que ella misma elige.
            betas=['server-side-fallback-2026-07-01'],
            fallbacks='default',
            cache_control={'type': 'ephemeral'},  # las instrucciones y la lista del SGA se repiten entre fichas
            output_config={'effort': 'medium', 'format': {'type': 'json_schema', 'schema': esquema}},
            system=instrucciones,
            messages=[{'role': 'user', 'content': mensaje}],
        )
    except anthropic.AuthenticationError:
        return None, 'la API rechazó la credencial'
    except anthropic.RateLimitError:
        return None, 'la API pidió esperar (límite de uso); intente más tarde'
    except anthropic.APIStatusError as e:
        return None, f'la API devolvió un error ({e.status_code})'
    except anthropic.APIConnectionError:
        return None, 'no hubo conexión con la API'
    if respuesta.stop_reason != 'end_turn':
        return None, f'el modelo no completó la respuesta ({respuesta.stop_reason})'
    return _leer_json(next((b.text for b in respuesta.content if b.type == 'text'), None))


def _compacto(texto):
    """Texto sin espacios, comas ni mayúsculas, para comprobar que un dato esté literal en la ficha."""
    t = unicodedata.normalize('NFKC', texto or '').lower()
    return re.sub(r'[\s,]+', '', t)


def _frases(resultado, recursos, consulta, avisos):
    frases = resultado['frases']
    for tipo, nombre in (('H', 'indicaciones de peligro (frases H)'), ('P', 'consejos de prudencia (frases P)')):
        # Pendientes: frases por texto sin código asignado, y frases de ese bloque que no se parecieron a ninguna
        dudosas = [f for f in frases[tipo] if f['estado'] == REV and not f['codigo']]
        sueltas = [s for s in frases['sin_codigo'] if s['bloque'] == tipo]
        pendientes = dudosas + sueltas
        if not pendientes:
            continue
        tabla = recursos.diccionario[tipo]
        instrucciones = INSTRUCCIONES_FRASES.format(nombre=nombre, ninguno=NINGUNO,
                                                    tabla='\n'.join(f'{c} | {es} | {en}' for c, (es, en) in tabla.items()))
        esquema = {'type': 'object', 'additionalProperties': False, 'required': ['asignaciones'], 'properties': {'asignaciones': {
            'type': 'array', 'items': {'type': 'object', 'additionalProperties': False, 'required': ['numero', 'codigo', 'motivo'],
                                       'properties': {'numero': {'type': 'integer'}, 'codigo': {'type': 'string', 'enum': list(tabla) + [NINGUNO]},
                                                      'motivo': {'type': 'string'}}}}}}
        mensaje = 'Frases de la ficha:\n' + '\n'.join(f"{i}. {p['texto_ficha']}" for i, p in enumerate(pendientes, 1))
        datos, aviso = consulta(instrucciones, mensaje, esquema)
        if datos is None:
            avisos.append(f'Frases {tipo}: {aviso}')
            continue
        for a in datos.get('asignaciones', []):
            numero, codigo = a.get('numero'), a.get('codigo')
            if not isinstance(numero, int) or not 1 <= numero <= len(pendientes):
                continue
            frase = pendientes[numero - 1]
            if codigo == NINGUNO:
                continue
            if codigo not in tabla:   # control: el código tiene que existir en el diccionario del SGA
                avisos.append(f"Frases {tipo}: se descartó «{codigo}» para «{frase['texto_ficha']}» porque no está en el diccionario")
                continue
            if frase in sueltas:
                frases['sin_codigo'].remove(frase)
                frases[tipo].append(frase)
                frase.pop('bloque')
            idioma = 0 if re.search(r'[áéíóúñ]| de | la | el ', f" {frase['texto_ficha'].lower()} ") else 1
            frase.update(codigo=codigo, estado=ENC, metodo='LLM', motivo_llm=a.get('motivo'), texto_sga=tabla[codigo][idioma])
        if frases['metodo'][tipo] == 'ninguna' and frases[tipo]:
            frases['metodo'][tipo] = 'texto'


def _identificacion(resultado, corte, consulta, avisos):
    ident = resultado['identificacion']
    nombre, proveedor = ident['nombre'], ident['proveedor']
    if nombre['estado'] == ENC and proveedor['estado'] == ENC:
        return
    seccion = '\n'.join(l for _, l in corte['secciones'][1]['lineas'])
    if not seccion.strip():
        return
    datos, aviso = consulta(INSTRUCCIONES_IDENTIFICACION, 'Texto de la sección 1:\n' + seccion, ESQUEMA_IDENTIFICACION)
    if datos is None:
        avisos.append(f'Identificación: {aviso}')
        return
    literal = _compacto(seccion)

    def aceptado(campo):
        valor = (datos.get(campo) or '').strip()
        if not valor:
            return None
        if _compacto(valor) not in literal:   # control: el dato tiene que estar literal en la sección 1
            avisos.append(f'Identificación: se descartó «{valor}» ({campo}) porque no aparece literal en la sección 1')
            return None
        return valor

    if nombre['estado'] != ENC:
        valor = aceptado('nombre_producto')
        if valor:
            nombre.update(valor=valor, estado=ENC, metodo='LLM', cita=valor)
            nombre.pop('nota', None)
    if proveedor['estado'] != ENC:
        nuevos = {c: aceptado('proveedor_' + c) for c in ('nombre', 'direccion', 'telefono')}
        if any(nuevos.values()):
            proveedor['propuesta_anterior'] = {c: proveedor[c] for c in ('nombre', 'direccion', 'telefono')}
            proveedor.update({c: v for c, v in nuevos.items() if v})
            proveedor['metodo'] = 'LLM'
            faltan = [c for c in ('nombre', 'direccion', 'telefono') if not proveedor[c]]
            proveedor['estado'] = REV if faltan else ENC
            proveedor['notas'] = ['falta ' + ', '.join(faltan).replace('direccion', 'dirección').replace('telefono', 'teléfono')] if faltan else []


def aplicar(resultado, corte, recursos, consulta, origen=None):
    """Completa 'resultado' con lo que el modelo proponga y pase los controles.
    'origen' = {proveedor, modelo}: se guarda con el resultado, para la trazabilidad.
    'consulta(instrucciones, mensaje, esquema)' devuelve (datos, aviso); se pasa como argumento para poder
    probar los controles sin llamar a la API. Devuelve los avisos de la pasada."""
    avisos = []
    _frases(resultado, recursos, consulta, avisos)
    _identificacion(resultado, corte, consulta, avisos)
    resultado['segunda_pasada'] = {**(origen or {'proveedor': None, 'modelo': None}), 'avisos': avisos}
    return avisos
