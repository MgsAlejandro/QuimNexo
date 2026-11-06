# -*- coding: utf-8 -*-
"""Interfaz de QuimNexo (Streamlit; antes QuimRiesgo Minero). Corre en este equipo y no decide nada: llama a la extracción,
al motor de clasificación y a la lectura práctica, y muestra lo que entregan con las piezas de presentacion.py.
La consulta se hace en tres pasos:  1 Producto -> 2 Exposición -> 3 Resultado
Hay dos perfiles (profesional de SST; trabajador o supervisor): recorren los mismos pasos y solo cambia cuánto
muestra el resultado. El detalle técnico no es un paso: se abre desde el panel izquierdo.
El informe de consulta y la etiqueta los dibujan reporte_html.py y etiqueta_html.py; aquí solo se les llama.
Para abrirla: doble clic en "Abrir QuimNexo.bat", o
    py -m streamlit run "clasificador SGA/app.py"
Requiere: pip install streamlit pandas openpyxl pdfplumber, y haber corrido antes  py extraccion/extraer.py FDS"""
import json, os, re, sys
import pandas as pd
import streamlit as st

AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, AQUI)
import almacen
import presentacion as ver
from presentacion import t, num
from tablas import cargar_tablas, ErrorDeTabla, MATRIZ, LIMITES, DICCIONARIO, FICHAS, PROTECCION
from motor import (clasificar, cargar_fichas, limites_de_la_sustancia, cas_de_la_ficha, FORMA_DE_COMPARAR,
                   SALIDAS as SALIDAS_EXTRACCION)
from lectura import (construir_lectura, prioridad as ordenar_por_prioridad, glosario, segun_cuantos, DESCRIPCION, CONTEXTO, NO_REEMPLAZA, VERSION,
                     PERFILES)
from reporte_html import _versiones, dibujar_reporte
from etiqueta import contenido_de_etiqueta, capacidad_del_empaque, numero
from etiqueta_html import dibujar_etiqueta, _medida
from etiqueta_mina import contenido_de_etiqueta_mina
from etiqueta_mina_html import dibujar_etiqueta_mina
from confirmaciones import (leer_confirmaciones, aplicar as aplicar_confirmaciones, opciones as opciones_de_confirmacion,
                            confirmar as confirmar_frase, deshacer as deshacer_confirmacion, clave_de_frase)
from fichas_nuevas import (procesar_fds_nueva, cargar_fichas_nuevas, rutas_de_resultados, resumen_de_extraccion, retirar_ficha_nueva,
                           ErrorDeFicha, Recursos)
import segunda_pasada   # está en la carpeta extraccion/, que fichas_nuevas ya agregó a la ruta de búsqueda

CONFIGURACION = os.path.join(AQUI, '..', 'configuracion.json')
VARIABLE_DE_CLAVE = {'anthropic': 'ANTHROPIC_API_KEY', 'openai': 'OPENAI_API_KEY'}
INICIO = -1
MAX_TAREAS = 4
# Los tres pasos de la consulta: (nombre corto, color de identificación, qué se hace). Los colores son solo para ubicarse.
PASOS = (('Producto', '#0F3D4C', 'Elija el producto o cargue su Ficha de Datos de Seguridad.'),
         ('Exposición', '#0F3D4C', 'Escriba la medición del puesto de trabajo, o indique que no la tiene.'),
         ('Resultado', '#0F3D4C', 'Peligro del producto, riesgo por exposición y qué hacer.'))

# Qué es obligatorio en cada paso, para decirlo en la navegación
REQUISITOS = ('Obligatorio: elegir un producto.', 'Obligatorio: la medición, o marcar «No tengo medición».', '')

# Apariencia: la paleta, la letra y los estilos de la pantalla están en presentacion.py (bloque «Identidad de la pantalla») y en
# .streamlit/config.toml. El informe y las etiquetas no los usan.
st.set_page_config(page_title=ver.NOMBRE_EN_PANTALLA, page_icon=ver.FAVICON, layout='wide')

# Streamlit olvida lo escrito en una casilla cuando esa casilla no está en pantalla. Como aquí cada paso muestra solo
# sus casillas, al empezar se "tocan" los valores guardados para que sigan vivos al cambiar de paso.
# (Las claves de los botones de confirmar frases, con '|conf|', no se pueden tocar y no hace falta.)
for _clave in list(st.session_state):
    if _clave in ('sustancia', 'paso', 'ver_detalle', 'perfil') or ('|' in _clave and '|conf|' not in _clave):
        st.session_state[_clave] = st.session_state[_clave]
st.session_state.setdefault('paso', 1)
st.session_state.setdefault('perfil', PERFILES[0])

# Modo oscuro (interruptor en el panel izquierdo). Streamlit manda el tema al empezar cada dibujo de la pantalla:
# si el interruptor lo cambió, se vuelve a dibujar para que el cambio se vea de una vez.
if 'modo_oscuro' not in st.session_state:
    st.session_state['modo_oscuro'] = ver.tema_es_oscuro()
st.session_state['modo_oscuro'] = st.session_state['modo_oscuro']   # que no se olvide si el interruptor sale de pantalla
OSCURO = st.session_state['modo_oscuro']
if ver.poner_tema(OSCURO):
    st.rerun()
st.logo(ver.LOGOTIPO_OSCURO if OSCURO else ver.LOGOTIPO, size='large', icon_image=ver.SIMBOLO_OSCURO if OSCURO else ver.SIMBOLO)
st.markdown(f'<style>{ver.CSS}{ver.CSS_DE_PANTALLA}{ver.CSS_MODO_OSCURO if OSCURO else ""}</style>', unsafe_allow_html=True)


def leer_configuracion():
    with open(CONFIGURACION, encoding='utf-8') as f:
        return json.load(f)


def guardar_configuracion(configuracion):
    """Guarda configuracion.json. Nunca contiene claves: esas viven solo en la memoria de la aplicación."""
    with open(CONFIGURACION, 'w', encoding='utf-8') as f:
        json.dump(configuracion, f, ensure_ascii=False, indent=1)


def _huella_de_archivos(rutas):
    """Fechas de modificación: si el autor cambia una tabla, la aplicación la vuelve a leer."""
    return tuple(os.path.getmtime(r) for r in rutas if os.path.exists(r))


@st.cache_resource(show_spinner='Leyendo las tablas…')
def tablas_cargadas(huella):
    return cargar_tablas()


@st.cache_data(show_spinner=False)
def fichas_extraidas(huella):
    return cargar_fichas() + cargar_fichas_nuevas()   # las fichas originales y, después, las que cargó el usuario


@st.cache_resource(show_spinner=False)
def recursos_de_extraccion(huella):
    return Recursos()


def html(texto):
    st.markdown(texto, unsafe_allow_html=True)


_numero_de_tarjeta = iter(range(10000))   # solo para darle a cada tarjeta un nombre propio y poder darle estilo; se reinicia en cada pasada


def tarjeta(titulo, contenido='', destino=st):
    """Recuadro con título. Devuelve el recuadro, por si hay que ponerle casillas dentro."""
    recuadro = destino.container(border=True, key=f'tarjeta_{next(_numero_de_tarjeta)}')
    recuadro.markdown(f'<div class="qr-h3">{titulo}</div>{contenido}', unsafe_allow_html=True)
    return recuadro


def tabla(filas, columnas):
    """Tabla con el texto completo de cada celda (las frases son largas y no deben cortarse)."""
    datos = pd.DataFrame([[('' if c is None else c) for c in f] for f in filas], columns=columnas)
    if datos[columnas[0]].duplicated().any():   # p. ej. dos filas del mismo paso: se numeran para no perder ninguna
        datos.insert(0, 'N.º', range(1, len(datos) + 1))
        columnas = ['N.º'] + list(columnas)
    st.table(datos.set_index(columnas[0]))


def validacion(titulo, falta, donde, porque, solucion):
    """Mensaje de validación: qué falta, dónde, por qué se necesita y qué hacer."""
    return (f'<div class="qr-validacion"><b>{t(titulo)}</b><div><b>Qué falta:</b> {t(falta)}</div><div><b>Dónde:</b> {t(donde)}</div>'
            f'<div><b>Por qué se necesita:</b> {t(porque)}</div><div><b>Qué hacer:</b> {t(solucion)}</div></div>')


# ---------- Carga ----------
try:
    tablas = tablas_cargadas(_huella_de_archivos((MATRIZ, LIMITES, DICCIONARIO, FICHAS, PROTECCION)))
except ErrorDeTabla as error:
    st.error(f'No se pudieron leer las tablas del proyecto. {error}')
    st.stop()
except Exception as error:   # un Excel abierto en otro programa, dañado o movido
    st.error(f'No se pudieron abrir las tablas del proyecto ({error}). Cierre los archivos de Excel si los tiene abiertos y vuelva a cargar la página.')
    st.stop()
archivos = sorted(os.listdir(SALIDAS_EXTRACCION)) if os.path.isdir(SALIDAS_EXTRACCION) else []
fichas = fichas_extraidas(_huella_de_archivos([os.path.join(SALIDAS_EXTRACCION, a) for a in archivos if a.endswith('.json')]
                                              + rutas_de_resultados()))
if not fichas:
    st.error('Todavía no hay fichas procesadas. Ejecute primero:  py extraccion/extraer.py FDS')
    st.stop()


def nombre_en_lista(i):
    if i == INICIO:
        return 'Elija un producto…'
    ficha = fichas[i]
    cas = [ficha['cotejo_cas']['archivo']] if ficha['cotejo_cas']['archivo'] else ficha['cotejo_cas']['ficha']
    excel = next((tablas['excel_fichas'][c]['sustancia'] for c in cas if c in tablas['excel_fichas']), None)
    numero = 'Nueva' if ficha.get('origen') == 'nueva' else ficha['archivo'].split('_')[0]
    return f"{numero} · {excel or ficha['identificacion']['nombre']['valor'] or ficha['archivo']}"


def cuenta_de_fichas():
    """Texto del pie: cuántas fichas hay. Un número con letra (08a, 08b) son dos fichas del mismo renglón del inventario."""
    propias = [f['archivo'].split('_')[0] for f in fichas if f.get('origen') != 'nueva']
    dobles = sorted({re.sub(r'\D', '', n) for n in propias if re.search(r'\d[a-z]$', n)})
    cargadas = len(fichas) - len(propias)
    return (f"{len(propias)} fichas del proyecto" + (f" (la {' y la '.join(dobles)} tiene{'n' if len(dobles) > 1 else ''} dos)" if dobles else '')
            + (f" y {cargadas} cargada{'s' if cargadas > 1 else ''} por usted" if cargadas else ''))


# Una ficha recién cargada queda elegida y la consulta sigue en el paso 2; si se quitó una, se vuelve al paso 1
if 'ir_a' in st.session_state:
    destino = st.session_state.pop('ir_a')
    st.session_state['sustancia'] = next((i for i, f in enumerate(fichas) if f['archivo'] == destino and f.get('origen') == 'nueva'), INICIO)
    st.session_state['paso'] = 2 if st.session_state['sustancia'] != INICIO else 1
if st.session_state.get('sustancia', INICIO) >= len(fichas):
    st.session_state['sustancia'] = INICIO
elegida = st.session_state.setdefault('sustancia', INICIO)


def borrar_mediciones_escritas():
    """Deja en blanco las casillas de medición de todos los productos: cada consulta empieza como una medición nueva.
    No toca lo de la etiqueta ('etq') ni las frases confirmadas ('conf')."""
    for _k in list(st.session_state):
        partes = _k.split('|') if isinstance(_k, str) else []
        if len(partes) >= 2 and partes[1] not in ('etq', 'conf') and (partes[1] == 'sin_medicion' or (len(partes) >= 3 and partes[2] in ('n', 'c', 'h'))):
            del st.session_state[_k]


# Consultar un producto (otro, o el mismo otra vez después de «Empezar otra consulta») es una medición nueva
_archivo_elegido = fichas[elegida]['archivo'] if elegida != INICIO else None
if _archivo_elegido is not None and st.session_state.get('producto_en_consulta') != _archivo_elegido:
    borrar_mediciones_escritas()
st.session_state['producto_en_consulta'] = _archivo_elegido
if elegida == INICIO:
    st.session_state['paso'] = 1          # sin producto no hay más pasos: se valida aquí y no solo en el botón
    st.session_state['ver_detalle'] = False
ficha_extraida = fichas[elegida] if elegida != INICIO else None
es_sst = st.session_state['perfil'] == PERFILES[0]


def clave(*partes):
    """Nombre interno de una casilla; lleva la ficha, para que los datos de un producto no pasen a otro."""
    return '|'.join((ficha_extraida['archivo'],) + tuple(str(p) for p in partes))


def mediciones_escritas(limites):
    """Mediciones que están escritas en las casillas del paso 2. Si se marcó «No tengo medición», ninguna."""
    if st.session_state.get(clave('sin_medicion')):
        return {}
    mediciones = {}
    for l in limites:
        if FORMA_DE_COMPARAR.get(l['tipo']) == 'jornada':
            tareas = [(st.session_state.get(clave(l['id'], 'c', i)), st.session_state.get(clave(l['id'], 'h', i), tablas['jornada_horas']))
                      for i in range(int(st.session_state.get(clave(l['id'], 'n'), 1)))]
            tareas = [(c, h) for c, h in tareas if c is not None and h is not None]
        else:
            c = st.session_state.get(clave(l['id'], 'c', 0))
            tareas = [(c, 0.0)] if c is not None else []
        if tareas:
            mediciones[l['id']] = {'tareas': tareas, 'unidad': l['unidad']}
    return mediciones


r = x = None
if ficha_extraida is not None:
    try:
        # Las frases cuyo código confirmó el usuario entran al cálculo
        ficha = aplicar_confirmaciones(ficha_extraida, leer_confirmaciones(ficha_extraida['archivo']), tablas)
        limites = limites_de_la_sustancia(cas_de_la_ficha(ficha), tablas)
        escritas = mediciones_escritas(limites)
        # El resultado usa solo lo escrito en esta consulta. Las mediciones guardadas no se cargan: cada consulta es una medición nueva
        r = clasificar(ficha, tablas, mediciones=escritas)
        x = construir_lectura(ficha, r, tablas)
        if r['indice_ir']['ir_mayor'] is not None or r['indice_ir']['alertas'] or r['semaforo']['clave'] == 'alerta_ok':
            x['fecha_de_medicion'] = 'esta consulta'
        n, p, g, e, ir, sem = r['nivel'], r['palabra_de_advertencia'], r['pictogramas'], r['etiqueta'], r['indice_ir'], r['semaforo']
        medibles = [l for l in ir['limites'] if isinstance(l['valor'], (int, float)) and FORMA_DE_COMPARAR.get(l['tipo'])
                    and (l['usar'] != 'No' or l['uso'] == 'Alerta de seguridad')]
    except Exception as error:
        st.error(f'No se pudo analizar «{ficha_extraida.get("archivo", "esta ficha")}» ({type(error).__name__}: {error}). '
                 'Elija otro producto para continuar.')
        st.session_state.update(sustancia=INICIO, paso=1)
        r = x = ficha_extraida = None


# ---------- Validación: qué falta para poder seguir ----------
def faltantes(paso):
    """Lo indispensable que falta en un paso. Solo bloquea lo que de verdad impide seguir."""
    if paso == 1 and x is None:
        return [('No ha elegido ningún producto.', 'Casilla «Producto», en este mismo paso.',
                 'Sin una Ficha de Datos de Seguridad no hay nada que revisar ni clasificar.',
                 'Elija un producto de la lista o cargue el PDF de la ficha y oprima «Analizar la ficha».')]
    if paso == 2 and x is not None:
        malas = [(f"La medición de «{l['agente']}» no se puede usar: {l['estado'].split(': ', 1)[-1]}.",
                  f"Casillas «Concentración medida» y «Horas de exposición» de {l['agente']}.",
                  'Con una medición mal escrita, el semáforo saldría equivocado.',
                  'Corrija los números, o marque «No tengo medición».') for l in ir['limites'] if (l['estado'] or '').startswith('medición no válida')]
        if malas:
            return malas
        if medibles and not escritas and not st.session_state.get(clave('sin_medicion')):
            return [('No ha escrito ninguna medición ni ha indicado que no la tiene.', 'Casilla «Concentración medida» o casilla «No tengo medición».',
                     'El riesgo por exposición se calcula con la medición; si no la hay, la herramienta debe saberlo para decirlo en el resultado.',
                     'Escriba la concentración medida, o marque «No tengo medición» para continuar sin ella.')]
    return []


def ir_al_paso(destino):
    """Avanza o retrocede. Al avanzar se revisan los pasos intermedios, para que no haya atajos. La medición no se guarda aquí."""
    actual = st.session_state['paso']
    for paso in range(actual, destino):
        pendientes = faltantes(paso)
        if pendientes:
            st.session_state.update(paso=paso, validacion=pendientes)
            return
    st.session_state.pop('validacion', None)
    st.session_state['paso'] = max(1, min(len(PASOS), destino))


def botones_de_paso():
    """Atrás y Continuar, al final de cada paso."""
    paso = st.session_state['paso']
    st.write('')
    atras, _, adelante = st.columns([1, 2, 1.3])
    if paso > 1:
        atras.button(f'← Atrás: {PASOS[paso - 2][0]}', on_click=ir_al_paso, args=(paso - 1,), width='stretch', key='boton_atras')
    if paso < len(PASOS):
        adelante.button(f'Continuar: {PASOS[paso][0]} →', on_click=ir_al_paso, args=(paso + 1,), type='primary', width='stretch', key='boton_continuar')
    else:
        adelante.button('Empezar otra consulta', on_click=lambda: (borrar_mediciones_escritas(), st.session_state.update(sustancia=INICIO, paso=1, ver_detalle=False)), width='stretch', key='boton_otra')


def aviso_de_frases_por_revisar():
    """Una línea: hay frases de la ficha que no entraron porque su código no quedó claro."""
    if x is not None and x['por_confirmar']:
        n = x['por_confirmar']
        html(f'<div class="qr-aviso">{n} ' + segun_cuantos(n, 'frase de la ficha quedó sin código seguro y no entra al resultado. Puede confirmarla',
                                                           'frases de la ficha quedaron sin código seguro y no entran al resultado. Puede confirmarlas')
             + ' en el detalle técnico (panel izquierdo).</div>')


# ---------- Panel izquierdo: perfil, dónde voy, detalle técnico y configuración ----------
def configuracion():
    datos = leer_configuracion()
    etq, llm = datos['etiqueta'], datos['modelo_de_lenguaje']
    st.markdown('**Etiqueta**')
    letra = st.number_input('Tamaño mínimo de letra (puntos)', 5.0, 14.0, float(etq['tamano_minimo_de_letra_pt']), 0.5, format='%g',
                            help='Si el texto no cabe con esta letra, la etiqueta no se imprime. Las tablas no fijan un mínimo; es un ajuste suyo.')
    exigir = st.checkbox('Exigir el proveedor en la etiqueta', bool(etq['proveedor_obligatorio']),
                         help='La Resolución 773 de 2021, art. 7.2, lo incluye en el contenido mínimo. Sin marcar, la etiqueta puede ir sin proveedor y lleva un aviso.')
    st.markdown('**Ayuda de inteligencia artificial** (opcional, solo para leer fichas nuevas)')
    proveedor = st.selectbox('Proveedor', list(VARIABLE_DE_CLAVE), index=list(VARIABLE_DE_CLAVE).index(llm['proveedor']))
    modelo = st.text_input('Modelo', llm['modelos'][proveedor], help='Nombre del modelo tal como lo llama el proveedor.')
    if st.button('Guardar configuración', width='stretch'):
        etq.update(tamano_minimo_de_letra_pt=letra, proveedor_obligatorio=exigir)
        llm['proveedor'] = proveedor
        llm['modelos'][proveedor] = modelo.strip() or llm['modelos'][proveedor]
        guardar_configuracion(datos)
        st.session_state['aviso'] = 'Configuración guardada.'
        st.rerun()
    variable = VARIABLE_DE_CLAVE[proveedor]
    st.caption(f"Clave de {proveedor}: {'hay una en uso' if os.environ.get(variable) else 'no hay ninguna'}. "
               'La clave no se guarda en ningún archivo: vale mientras la aplicación esté abierta.')
    clave_escrita = st.text_input(f'Clave de acceso de {proveedor}', type='password', key=f'clave_api_{proveedor}')
    if st.button('Usar esta clave mientras la aplicación esté abierta', width='stretch', disabled=not clave_escrita):
        os.environ[variable] = clave_escrita.strip()   # solo en la memoria de este programa; no se escribe en disco
        st.session_state['aviso'] = f'Clave de {proveedor} en uso. No se guardó en ningún archivo.'
        st.rerun()


with st.sidebar:
    html(f'<div class="qr-nota" style="margin-bottom:.8rem">{t(ver.LEMA)}.</div>')
    if st.session_state['paso'] > 1:   # en el paso 1 el perfil se elige en la pantalla principal
        st.radio('Perfil', PERFILES, key='perfil')
    if x is not None:
        html(f'<div class="qr-nota">Producto en consulta</div><div class="qr-producto" style="font-weight:700;color:#0b2a4f;margin-bottom:.5rem">{t(x["nombre"])}</div>')
    html(''.join(f'<div class="qr-lateral {"activa" if i == st.session_state["paso"] else "hecha" if i < st.session_state["paso"] else ""}" style="--c:{color}">'
                 f'{"✓" if i < st.session_state["paso"] else i} · {nombre}</div>' for i, (nombre, color, _) in enumerate(PASOS, 1)))
    st.divider()
    html('<div class="qr-h4" style="margin-top:0">Detalle técnico</div>')
    st.toggle('Ver el detalle técnico', key='ver_detalle', disabled=x is None,
              help='De dónde sale cada dato: todas las frases, criterios de selección, método, límites, historial de mediciones, reglas y tablas. '
                   'Aquí también se confirman las frases dudosas. No hace falta para completar la consulta.')
    if x is None:
        st.caption('Disponible cuando elija un producto.')
    st.divider()
    if 'aviso' in st.session_state:
        st.info(st.session_state.pop('aviso'))
    st.toggle('🌙 Modo oscuro', key='modo_oscuro', help='Fondo oscuro para la pantalla. El informe y las etiquetas se siguen viendo e imprimiendo en blanco.')
    with st.popover('⚙️ Configuración', width='stretch'):
        configuracion()
    html(f'<div class="qr-nota" style="margin-top:1rem">Versión {VERSION} · {cuenta_de_fichas()}. Funciona en este equipo; nada sale de él.</div>')


# ---------- Paso 1: Producto ----------
def analizar(subida, con_modelo):
    """Procesa el PDF cargado. Cualquier fallo se informa en español y no deja nada a medias."""
    try:
        consulta = origen = None
        if con_modelo:
            consulta, origen, motivo = segunda_pasada.crear_consulta()
            if consulta is None:
                st.session_state['aviso'] = f'No se usó la ayuda de inteligencia artificial: {motivo}. La ficha se analizó sin ella.'
        with st.spinner('Leyendo la ficha…'):
            nueva = procesar_fds_nueva(subida.name, subida.getvalue(),
                                       recursos_de_extraccion(_huella_de_archivos((MATRIZ, DICCIONARIO, FICHAS))), consulta, origen)
        st.session_state['ultima_carga'] = (nueva['archivo'], resumen_de_extraccion(nueva))
        st.session_state['ir_a'] = nueva['archivo']
        st.rerun()
    except ErrorDeFicha as error:
        html(ver.caja('mal', f'<b>No se pudo analizar el archivo.</b> {t(error)}'))
    except Exception as error:
        html(ver.caja('mal', f'<b>Ocurrió un error y no se guardó nada.</b> Revise que el PDF abra bien y vuelva a intentarlo. '
                             f'<span class="qr-pag">{t(type(error).__name__)}: {t(error)}</span>'))


def nueva_consulta(actual):
    """Quita de la lista la ficha cargada que se está viendo, para empezar con otra. No borra: mueve."""
    if st.session_state.get('por_quitar') != actual:
        if st.button('Nueva consulta: quitar esta ficha cargada', help='La quita de la lista para empezar con otra. No la borra.'):
            st.session_state['por_quitar'] = actual
            st.rerun()
        return
    st.warning(f'¿Quitar «{actual}» de la lista? No se borra: se mueve a la carpeta FDS_nuevas/retiradas, junto con lo que usted eligió para su etiqueta.')
    si, no, _ = st.columns([1, 1, 3])
    if si.button('Sí, quitar', width='stretch'):
        try:
            retirar_ficha_nueva(actual)
            st.session_state.pop('ultima_carga', None)
            st.session_state['aviso'] = f'«{actual}» se quitó de la lista. Quedó guardada en FDS_nuevas/retiradas.'
        except ErrorDeFicha as error:
            st.session_state['aviso'] = str(error)
        st.session_state.pop('por_quitar', None)
        st.session_state['ir_a'] = None
        st.rerun()
    if no.button('Cancelar', width='stretch'):
        st.session_state.pop('por_quitar', None)
        st.rerun()


@st.cache_data(show_spinner='Clasificando los productos…')
def panorama(huella_tablas, huella_fichas, huella_mediciones):
    """Resultado del motor para todas las fichas, cada una con su medición guardada más reciente, para «qué revisar primero»."""
    resultados = []
    for f in fichas:
        ultima = almacen.ultima_medicion(f['archivo'])
        res = clasificar(aplicar_confirmaciones(f, leer_confirmaciones(f['archivo']), tablas), tablas, mediciones=(ultima or {}).get('mediciones') or {})
        resultados.append(dict(archivo=res['archivo'], producto=res['producto'], sustancia_en_excel=res['sustancia_en_excel'], nivel=res['nivel'],
                               semaforo=res['semaforo'], fecha=ultima['fecha'] if ultima else None, nueva=f.get('origen') == 'nueva'))
    return resultados


def lista_de_revision():
    carpeta = os.path.join(almacen.AQUI, 'mediciones')
    guardadas = [os.path.join(carpeta, a) for a in sorted(os.listdir(carpeta))] if os.path.isdir(carpeta) else []
    ordenados = ordenar_por_prioridad(panorama(_huella_de_archivos((MATRIZ, LIMITES, DICCIONARIO, FICHAS)), tuple(f['archivo'] for f in fichas),
                                               _huella_de_archivos(guardadas)))
    filas = ''.join(
        f"<tr><td>{i}</td><td><b>{t(q['sustancia_en_excel'] or q['producto'] or q['archivo'])}</b>"
        + ('<br><span class="qr-pag">Ficha cargada por el usuario</span>' if q['nueva'] else '') + '</td>'
        f"<td>{ver.chip_de_semaforo(q['semaforo'])}" + (f" <span class='qr-pag'>Ir = {num(q['semaforo']['ir'])}</span>" if q['semaforo']['ir'] is not None else '')
        + f"</td><td>{ver.chip_de_nivel(q['nivel'])}</td><td>{t(q['fecha'], '<span class=qr-pag>Sin medición guardada</span>')}</td></tr>"
        for i, q in enumerate(ordenados, 1))
    html('<div class="qr-desliza"><table class="qr-tabla"><tr><th>N.º</th><th>Producto</th><th>Riesgo por exposición</th><th>Peligro del producto</th>'
         f'<th>Última medición</th></tr>{filas}</table></div>')


def paso_producto():
    st.radio('¿Quién consulta?', PERFILES, key='perfil', horizontal=True,
             help='Los dos perfiles hacen los mismos tres pasos. El profesional de SST ve además la acción recomendada y lo que falta verificar.')
    izquierda, derecha = st.columns(2, gap='medium')
    with tarjeta('Elija un producto de la lista', destino=izquierda):
        st.selectbox('Producto', [INICIO] + list(range(len(fichas))), format_func=nombre_en_lista, key='sustancia',
                     help='Son las fichas que ya están procesadas en este equipo.')
        if x is not None:
            html(ver.caja('bien', f'<b>Producto elegido:</b> {t(x["nombre"])}. Oprima «Continuar».'))
            aviso_de_frases_por_revisar()
            if x['nueva']:
                nueva_consulta(ficha_extraida['archivo'])
    with tarjeta('O cargue una ficha nueva (PDF)', destino=derecha):
        subida = st.file_uploader('Ficha de Datos de Seguridad en PDF', type=['pdf'],
                                  help='Debe tener texto que se pueda seleccionar (no sirve una foto ni un escaneo). Se guarda en la carpeta FDS_nuevas, '
                                       'aparte de las fichas originales.')
        con_modelo = st.checkbox('Usar ayuda de inteligencia artificial para lo dudoso (opcional)', key='con_modelo',
                                 help='Solo propone el código de las frases que quedaron dudosas, y el programa descarta lo que no pase sus controles. '
                                      'Nunca clasifica. Envía al proveedor elegido solo esos fragmentos. Requiere una clave, que se pone en Configuración.')
        if st.button('Analizar la ficha', type='primary', width='stretch', disabled=subida is None):
            analizar(subida, con_modelo)
        if subida is None:
            st.caption('Primero busque el archivo; después se activa el botón.')

    with st.expander('Qué revisar primero: todos los productos'):
        html('<div class="qr-nota">Orden: primero por el color del riesgo por exposición (rojo, amarillo, verde y, al final, sin medición) y, dentro de cada '
             'color, por el peligro del producto de mayor a menor. Se usa la medición guardada más reciente de cada producto.</div>')
        try:
            lista_de_revision()
        except Exception as error:
            st.error(f'No se pudo armar la lista de productos ({type(error).__name__}: {error}). Los productos se pueden consultar uno por uno.')
    guardadas = almacen.consultas()
    with st.expander(f'Consultas guardadas · {len(guardadas)}'):
        if not guardadas:
            st.caption('Aquí aparecerán los informes que usted guarde al final del paso «Resultado».')
        else:
            tabla([(c.get('Fecha'), c.get('Producto'), c.get('Nivel de peligro'), c.get('Riesgo por exposición') or c.get('Prioridad'), c.get('Índice Ir') or '—')
                   for c in guardadas], ['Fecha', 'Producto', 'Peligro del producto', 'Riesgo por exposición', 'Índice Ir'])
            cual = st.selectbox('Abrir un informe guardado', range(len(guardadas)), format_func=lambda i: f"{guardadas[i].get('Fecha')} · {guardadas[i].get('Producto')}")
            try:
                st.download_button('Descargar ese informe', almacen.leer_informe(guardadas[cual]['Informe']), file_name=guardadas[cual]['Informe'], mime='text/html')
            except OSError as error:
                st.error(f'No se pudo abrir el informe guardado ({error}).')
            st.caption('Los informes y el registro (registro.csv, que se abre en Excel) están en la carpeta «clasificador SGA/consultas».')
    with st.expander('Qué hace esta herramienta y palabras que usa'):
        html(f'<div>{t(DESCRIPCION)} {t(CONTEXTO)} {t(NO_REEMPLAZA)}</div>'
             + ver.lista(f'<b>{t(termino)}.</b> {t(explicacion)}' for termino, explicacion in glosario(tablas)))


# ---------- Paso 2: Exposición (solo la medición) ----------
def paso_exposicion():
    jornada = tablas['jornada_horas']
    if not medibles:
        html(ver.caja('info', '<b>Este producto no tiene valor límite en la tabla del proyecto.</b> No hay con qué comparar una medición, así que el semáforo '
                              'del riesgo por exposición no aplica. Oprima «Continuar».'))
        return
    st.caption('Cada consulta es una medición nueva: las casillas empiezan en blanco. La medición solo se guarda si al final oprime «Guardar en el historial».')
    sin_medicion = st.checkbox('No tengo medición', key=clave('sin_medicion'),
                               help='Puede continuar sin medición. El resultado lo dirá y recomendará medir.')
    for l in medibles:
        forma = FORMA_DE_COMPARAR.get(l['tipo'])
        with st.container(border=True):
            html(f"<div class='qr-h4' style='margin-top:0'>{t(l['agente'])}</div><div><span class='qr-cod'>{l['id']}</span>Valor límite: <b>{t(l['tipo'])} "
                 f"{l['valor']:g} {t(l['unidad'])}</b>" + (f" · {t(l['fraccion'])}" if l['fraccion'] else '') + '</div>')
            if l['mensaje']:
                st.caption(('Condición: ' if l['usar'] == 'Condicional' else '') + l['mensaje'])
            if forma == 'jornada':
                st.session_state.setdefault(clave(l['id'], 'n'), 1)   # el valor inicial va por aquí, no en la casilla, porque se conserva entre pasos
                st.number_input('¿En cuántas tareas de la jornada hay exposición?', 1, MAX_TAREAS, key=clave(l['id'], 'n'), disabled=sin_medicion,
                                help='Si la exposición cambia durante la jornada, escriba cada tarea con su concentración y sus horas.')
                for i in range(int(st.session_state.get(clave(l['id'], 'n'), 1))):
                    a, b = st.columns(2)
                    a.number_input(f"Concentración medida ({l['unidad']})" + (f', tarea {i + 1}' if i else ''), min_value=0.0, value=None,
                                   format='%g', placeholder='sin medición', key=clave(l['id'], 'c', i), disabled=sin_medicion)
                    st.session_state.setdefault(clave(l['id'], 'h', i), jornada)
                    b.number_input('Horas de exposición' + (f', tarea {i + 1}' if i else ''), min_value=0.0, max_value=24.0,
                                   format='%g', key=clave(l['id'], 'h', i), disabled=sin_medicion)
            else:
                que = 'promedio de 15 minutos' if l['tipo'] == 'TLV-STEL' else 'máxima instantánea' if l['tipo'] == 'TLV-C' else 'medida'
                st.number_input(f"Concentración {que} ({l['unidad']})", min_value=0.0, value=None, format='%g', placeholder='sin medición',
                                key=clave(l['id'], 'c', 0), disabled=sin_medicion)
            if (l['estado'] or '').startswith('medición no válida') and not sin_medicion:
                html(ver.caja('mal', '<b>Corrija esta medición para continuar.</b> ' + t(l['estado'].capitalize()) + '.'))
    st.caption(f'La herramienta divide la exposición por el valor límite (índice Ir), con jornada de {jornada:g} horas. El color se ve en el paso siguiente.')


# ---------- Paso 3: Resultado, según el perfil ----------
def paso_resultado():
    html(ver.indicadores(r, x))
    for alerta in ir['alertas']:
        html(ver.caja('mal', f'<b>Alerta de seguridad:</b> {t(alerta)}'))
    if es_sst:
        html(f'<div class="qr-caja qr-accion"><div class="qr-h3" style="margin-bottom:.2rem">Acción recomendada</div>{ver.accion(x)}</div>')
    aviso_de_frases_por_revisar()
    izquierda, derecha = st.columns(2, gap='medium')
    if x['que_hacer'] or x['accidente']:
        tarjeta('Qué hacer', (ver.consejos(x['que_hacer']) if x['que_hacer'] else '<div class="qr-nota">La ficha no trae consejos de prevención ni de almacenamiento.</div>')
                + '<div class="qr-h4">Protección personal</div>' + ver.proteccion(x), izquierda)
        if x['accidente']:
            tarjeta('Si hay un accidente', ver.consejos(x['accidente']), izquierda)
    else:
        tarjeta('Qué hacer', (ver.caja('pend', 'La ficha no trae consejos de prudencia. La herramienta no los inventa: pida al proveedor una ficha completa.')
                              if x['peligros'] else '<div class="qr-nota">Sin indicaciones de peligro, la ficha no da consejos de prudencia.</div>')
                + '<div class="qr-h4">Protección personal</div>' + ver.proteccion(x), izquierda)
    tarjeta('Qué puede pasar', ver.frases_h(x), derecha)
    if es_sst:
        with st.expander(f'Por verificar · {len(x["pendientes"])}'):
            html(ver.lista(t(i) for i in x['pendientes']) if x['pendientes'] else '<div>No quedó nada pendiente en la lectura de esta ficha.</div>')
        st.button('Abrir el detalle técnico', on_click=lambda: st.session_state.update(ver_detalle=True), key='boton_detalle',
                  help='Todas las frases, los criterios de selección, el método, los límites, el historial de mediciones, las reglas y las tablas.')

    st.divider()
    html('<div class="qr-seccion">Documentos</div>')
    guardadas = preparar_etiqueta()
    lado_informe, lado_etiqueta = st.columns(2, gap='medium')
    informe = None
    with tarjeta('Informe de consulta', '<div class="qr-nota">Una hoja carta con este resultado.</div>', lado_informe):
        try:
            con_anexo = es_sst and st.checkbox('Agregar el anexo técnico al informe', key=clave('anexo'))
            informe = dibujar_reporte(r, tablas, x, detalle=bool(con_anexo))
            st.download_button('Descargar informe', informe, file_name=os.path.splitext(r['archivo'])[0] + '_informe.html', mime='text/html', type='primary',
                               width='stretch', help='Se abre en el navegador y trae su botón para imprimir o guardar como PDF.')
            if st.button('Guardar en el historial', width='stretch',
                         help='Guarda una copia del informe en la carpeta de consultas, la anota en el registro y guarda la medición con su fecha.'):
                almacen.guardar_consulta(r, x, informe)
                if escritas:
                    almacen.guardar_medicion(ficha['archivo'], escritas)
                st.session_state['aviso_informe'] = 'Consulta guardada. Aparece en el paso «Producto», en «Consultas guardadas».'
            if 'aviso_informe' in st.session_state:
                html(ver.caja('bien', st.session_state.pop('aviso_informe')))
        except Exception as error:
            st.error(f'No se pudo generar el informe ({type(error).__name__}: {error}).')
    with tarjeta('Etiqueta del empaque', '<div class="qr-nota">Etiqueta corta de uso en la mina. Su tamaño depende de la capacidad del empaque.</div>', lado_etiqueta):
        a, b = st.columns([1.3, 1])
        a.number_input('Capacidad del empaque', min_value=0.0, value=None, format='%g', placeholder='por ejemplo 20', key=clave('etq', 'capacidad'),
                       help='De la capacidad depende el tamaño de la etiqueta y de los pictogramas. «Empaque» es lo que la Resolución 773 de 2021 (art. 11) llama envase.')
        b.selectbox('Unidad', list(tablas['unidades']), key=clave('etq', 'unidad'),
                    help='La tabla de tamaños de etiqueta está en litros: la capacidad se convierte con la hoja Unidades_empaque. Para g y kg se toma 1 kg = 1 L.')
        zona_de_descarga = st.container()   # aquí queda el botón de la etiqueta; se llena más abajo, cuando ya se conocen todas las elecciones
    if informe:
        with st.expander('Ver el informe antes de descargarlo'):
            st.iframe(informe, height=900)
    with st.expander('Etiqueta: vista previa y opciones (nombre, proveedor, etiqueta completa)'):
        documento_etiqueta(guardadas, zona_de_descarga)


def preparar_etiqueta():
    """Lo que el usuario decidió antes para la etiqueta de este producto se pone en las casillas la primera vez. Devuelve lo guardado.
    Una elección guardada por una versión anterior solo trae 'envase_litros': se lee como litros."""
    guardadas = almacen.leer('elecciones', ficha['archivo'])
    if clave('etq', 'capacidad') not in st.session_state:
        prov = guardadas.get('proveedor') or {}
        vieja = guardadas.get('capacidad') in (None, '')
        iniciales = {'capacidad': guardadas.get('envase_litros') if vieja else guardadas.get('capacidad'),
                     'unidad': guardadas['unidad'] if not vieja and guardadas.get('unidad') in tablas['unidades'] else 'L',
                     'nombre': guardadas.get('nombre') or '', 'completa': False, 'interno': bool(guardadas.get('uso_interno')),
                     'borde': guardadas.get('borde') or 'rojo', 'componentes': guardadas.get('componentes') or '',
                     'cantidad': guardadas.get('cantidad') or '', 'lote': guardadas.get('lote') or '',
                     'prov_nombre': prov.get('nombre') or '', 'prov_direccion': prov.get('direccion') or '', 'prov_telefono': prov.get('telefono') or ''}
        for campo, valor_inicial in iniciales.items():
            st.session_state[clave('etq', campo)] = valor_inicial
        marcados_antes = guardadas.get('consejos_p')
        de_la_seleccion = {z['codigo'] for z in r['seleccion_p']['elegidos']}   # en la etiqueta completa vienen marcados los de la selección para minería (regla C05)
        for f in e['P']:
            st.session_state[clave('etq', 'p', f['codigo'])] = f['codigo'] in de_la_seleccion if marcados_antes is None else f['codigo'] in marcados_antes
            st.session_state[clave('etq', 'texto', f['codigo'])] = (guardadas.get('textos') or {}).get(f['codigo'], '')
    return guardadas


def documento_etiqueta(guardadas, zona_de_descarga):
    """Vista previa y opciones de la etiqueta de uso en la mina, que es la que se descarga con el botón de arriba ('zona_de_descarga').
    La etiqueta completa de la Resolución 773, con todas sus frases, queda como opción secundaria."""
    valor = lambda campo: st.session_state.get(clave('etq', campo))
    formulario, vista = st.columns([1, 1.2], gap='medium')
    with formulario:
        with tarjeta('Datos opcionales'):
            por_defecto = contenido_de_etiqueta_mina(ficha, r, {}, tablas)
            st.text_input('Nombre del producto en la etiqueta', key=clave('etq', 'nombre'), placeholder=por_defecto['nombre'],
                          help=f"Déjelo vacío para usar «{por_defecto['nombre']}» ({por_defecto['origen_del_nombre']}). Lo que escriba se guarda para este producto.")
            sugerido = r['proveedor']
            if sugerido.get('nombre'):
                st.caption(f"Proveedor que aparece en la ficha: {sugerido['nombre']} · {sugerido.get('direccion') or '—'} · {sugerido.get('telefono') or '—'}. "
                           'No se pone solo: escriba el proveedor real si quiere incluirlo.')
            st.text_input('Proveedor: nombre', key=clave('etq', 'prov_nombre'))
            a, b = st.columns(2)
            a.text_input('Proveedor: dirección', key=clave('etq', 'prov_direccion'))
            b.text_input('Proveedor: teléfono', key=clave('etq', 'prov_telefono'))
            a, b = st.columns(2)
            a.text_input('Cantidad', key=clave('etq', 'cantidad'))
            b.text_input('Lote', key=clave('etq', 'lote'))
        completa = st.toggle('Etiqueta completa (Res. 773)', key=clave('etq', 'completa'),
                             help='Opción secundaria: la etiqueta con todas las indicaciones de peligro y los consejos de prudencia oficiales, que usted marca.')
        if completa:
            with tarjeta('Etiqueta completa · uso y borde'):
                st.checkbox('El producto no sale del lugar de trabajo (uso interno)', key=clave('etq', 'interno'))
                interno = st.session_state[clave('etq', 'interno')]
                if not interno:
                    st.session_state[clave('etq', 'borde')] = 'rojo'   # el borde negro solo se admite con uso interno (regla R09)
                st.radio('Borde de los pictogramas', ['rojo', 'negro'], horizontal=True, key=clave('etq', 'borde'), disabled=not interno,
                         help='Rojo es el de la norma. El negro solo se admite si el producto no sale del lugar de trabajo (Resolución 773 de 2021, art. 12).')
                st.text_input('Componentes (para mezclas)', key=clave('etq', 'componentes'),
                              help='La norma pide indicar los componentes que aportan ciertos peligros (Resolución 773 de 2021, art. 10). Escríbalos usted.')
            with tarjeta('Etiqueta completa · consejos de prudencia'):
                if e['P']:
                    st.caption('Vienen marcados los cinco de la selección para minería. Marque los demás que quiera incluir; si el texto no cabe, deje menos '
                               '(Resolución 773 de 2021, art. 7, parágrafo 2). La selección es suya.')

                    def marcar(cuales):
                        elegidos = {z['codigo'] for z in r['seleccion_p']['elegidos']}
                        for f in e['P']:
                            st.session_state[clave('etq', 'p', f['codigo'])] = bool(f['texto']) and (cuales == 'todos' and f['incluida'] or cuales == 'mineria' and f['codigo'] in elegidos)
                    a, b, c = st.columns(3)
                    a.button('Los de minería', on_click=marcar, args=('mineria',), width='stretch', help='Los cinco de la selección para minería.')
                    b.button('Todos', on_click=marcar, args=('todos',), width='stretch', help='Prevención, intervención y almacenamiento.')
                    c.button('Ninguno', on_click=marcar, args=('ninguno',), width='stretch')
                    for f in e['P']:
                        st.checkbox(f"**{f['codigo']}** · {f['tipo']}: {f['texto'] or 'sin texto en español en las tablas'}", key=clave('etq', 'p', f['codigo']))
                        if f['completar'] and st.session_state[clave('etq', 'p', f['codigo'])]:
                            st.text_input(f"Texto completo de {f['codigo']} (reemplaza los «...»)", key=clave('etq', 'texto', f['codigo']),
                                          placeholder=f['texto'], help=f"Texto original de la ficha: {f['texto_ficha']}")
                else:
                    st.caption('La ficha no trae consejos de prudencia. La herramienta no los inventa.')

    # Elecciones tal como quedaron en las casillas. 'envase_litros' se conserva (ya convertido a litros) para las versiones anteriores.
    elecciones = {
        'capacidad': valor('capacidad') or None,
        'unidad': valor('unidad') or 'L',
        'nombre': (valor('nombre') or '').strip(),
        'uso_interno': bool(valor('interno')),
        'borde': valor('borde') or 'rojo',
        'consejos_p': [f['codigo'] for f in e['P'] if st.session_state.get(clave('etq', 'p', f['codigo']))],
        'textos': {f['codigo']: st.session_state.get(clave('etq', 'texto', f['codigo']), '').strip() for f in e['P']
                   if st.session_state.get(clave('etq', 'texto', f['codigo']), '').strip()},
        'componentes': (valor('componentes') or '').strip(),
        'proveedor': {c: (valor('prov_' + c) or '').strip() for c in ('nombre', 'direccion', 'telefono')},
        'cantidad': (valor('cantidad') or '').strip(),
        'lote': (valor('lote') or '').strip(),
    }
    litros, capacidad = capacidad_del_empaque(elecciones, tablas)
    elecciones['envase_litros'] = litros
    # Se guarda solo cuando el usuario ya empezó la etiqueta (dio la capacidad) o ya había algo guardado
    if elecciones != guardadas and (guardadas or litros):
        almacen.guardar('elecciones', ficha['archivo'], elecciones)

    base = os.path.splitext(r['archivo'])[0]
    with tarjeta('Vista previa a tamaño real', destino=vista):
        try:
            mina = contenido_de_etiqueta_mina(ficha, r, elecciones, tablas)
            if not mina['tamano']:
                html(validacion('Falta un dato para ver la etiqueta', 'La capacidad del empaque.', 'Casilla «Capacidad del empaque», arriba, en «Documentos».',
                                'La norma fija el tamaño de la etiqueta y de los pictogramas según la capacidad del empaque.',
                                'Escriba la capacidad (por ejemplo 20), elija la unidad y oprima Enter.'))
                zona_de_descarga.download_button('Descargar etiqueta para imprimir', '', file_name='etiqueta.html', disabled=True, width='stretch', key=clave('etq', 'boton_sin_capacidad'))
                zona_de_descarga.markdown('<div class="qr-aviso">Para activar la descarga, escriba la capacidad del empaque y oprima Enter.</div>', unsafe_allow_html=True)
            else:
                ancho, alto = _medida(mina['tamano']['etiqueta_mm'], 'la etiqueta')
                pagina = dibujar_etiqueta_mina(mina, tablas)
                zona_de_descarga.download_button('Descargar etiqueta para imprimir', pagina, type='primary', width='stretch', mime='text/html',
                                                 file_name=f"{base}_etiqueta_{re.sub(r'[^0-9A-Za-z]+', '_', capacidad).strip('_')}.html",
                                                 help='Etiqueta de uso en la mina. Se abre en el navegador y trae su botón «Imprimir etiqueta». Imprima en hoja carta, al 100 %, sin «ajustar a la página».')
                zona_de_descarga.caption((f'{capacidad} = {numero(litros)} L → ' if elecciones['unidad'] != 'L' else f'{capacidad} → ') + f'etiqueta de {ancho:g} × {alto:g} mm.')
                if mina['origen_de_los_peligros'] == 'inventario':
                    zona_de_descarga.markdown(f'<div class="qr-aviso">{t(tablas["etiqueta_mina"]["texto_inventario"])}</div>', unsafe_allow_html=True)
                st.caption(f'Etiqueta de uso en la mina · empaque de {capacidad} · {ancho:g} × {alto:g} mm. Siempre se puede imprimir: si algo no cabe, se quita y el panel lo dice.')
                st.iframe(pagina, height=int(alto * 3.78) + 520)
        except Exception as error:
            st.error(f'No se pudo armar la etiqueta ({type(error).__name__}: {error}).')
            zona_de_descarga.error('No se pudo armar la etiqueta. Abra «Etiqueta: vista previa y opciones» para ver el motivo.')
        if completa:
            html('<div class="qr-h3" style="margin-top:1rem">Etiqueta completa (Res. 773)</div>')
            try:
                contenido = contenido_de_etiqueta(r, elecciones, tablas)
                if contenido['tamano']:
                    tam = contenido['tamano']
                    st.caption(f"{contenido['tipo'].capitalize()} · empaque de {contenido['capacidad']} · etiqueta de {tam['etiqueta']} mm · pictogramas de {tam['pictograma']} mm")
                    for falta in contenido['faltantes']:
                        html(ver.caja('mal', f'<b>Falta:</b> {t(falta)}'))
                    pagina_completa = dibujar_etiqueta(contenido, tablas)
                    alto_mm = (tam['etiqueta_mm']['minimo'] or tam['etiqueta_mm']['recomendado'])[1]
                    st.iframe(pagina_completa, height=int(alto_mm * 3.78) + 620)
                    st.download_button('Descargar etiqueta completa (Res. 773)', pagina_completa, mime='text/html', disabled=not contenido['lista_para_imprimir'],
                                       file_name=f"{base}_etiqueta_completa_{re.sub(r'[^0-9A-Za-z]+', '_', contenido['capacidad']).strip('_')}.html",
                                       help='Trae todas las indicaciones de peligro y los consejos marcados. Si el texto no cabe en el tamaño de la norma, avisa y no se imprime recortada.')
                    st.caption('Los renglones resaltados traen «...» o un texto entre corchetes: complételos o desmárquelos. Eso no afecta la etiqueta de uso en la mina.')
                else:
                    st.caption('Escriba la capacidad del empaque para verla.')
            except Exception as error:
                st.error(f'No se pudo armar la etiqueta completa ({type(error).__name__}: {error}).')


# ---------- Detalle técnico: se abre desde el panel izquierdo; no es un paso ----------
def frases_por_confirmar():
    """Frases de la ficha sin código seguro, con la opción de confirmarlo. Devuelve cuántas son."""
    dudosas = [(tipo, f) for tipo in 'HP' for f in ficha_extraida['frases'][tipo] if f['estado'] != 'encontrado']
    if not dudosas:
        return 0
    st.write('La herramienta no pudo asignarles un código con seguridad. Mientras estén sin confirmar, no cuentan para el nivel, la selección ni la etiqueta. '
             'Usted puede confirmar el código; solo se ofrecen los que la propia lectura propuso.')
    hechas = leer_confirmaciones(ficha_extraida['archivo'])
    for i, (tipo, f) in enumerate(dudosas):
        with st.container(border=True):
            st.markdown(f"**{'Indicación de peligro' if tipo == 'H' else 'Consejo de prudencia'}** en la ficha (p. {f['pagina']}): «{f['texto_ficha']}»")
            hecha = hechas.get(clave_de_frase(tipo, f))
            ofrecidas = opciones_de_confirmacion(tipo, f, tablas)
            if hecha and hecha.get('codigos') in [o['codigos'] for o in ofrecidas]:
                st.success(f"Confirmada por el usuario el {hecha.get('fecha')}: {', '.join(hecha['codigos'])}. Ya entra al cálculo.")
                if st.button('Deshacer', key=clave('conf', 'deshacer', i)):
                    deshacer_confirmacion(ficha_extraida['archivo'], tipo, f)
                    st.rerun()
            elif ofrecidas:
                elegido = st.radio('Código que corresponde', range(len(ofrecidas)), format_func=lambda k, o=ofrecidas: o[k]['rotulo'],
                                   key=clave('conf', 'opcion', i))
                if st.button('Confirmar', key=clave('conf', 'confirmar', i)):
                    confirmar_frase(ficha_extraida['archivo'], tipo, f, ofrecidas[elegido]['codigos'])
                    st.rerun()
            else:
                st.caption('La lectura no propuso ningún código para esta frase. Sigue por confirmar.')
    return len(dudosas)


def detalle_tecnico():
    html('<div class="qr-cabeza" style="--c:#44505c"><small>Fuera de los pasos</small><b>Detalle técnico</b>'
         '<span>De dónde sale cada dato. No hace falta para completar la consulta.</span></div>')
    st.button(f"← Volver al paso {st.session_state['paso']}: {PASOS[st.session_state['paso'] - 1][0]}", on_click=lambda: st.session_state.update(ver_detalle=False),
              type='primary', key='boton_volver')
    html(ver.factores(x['factores']))
    with st.expander(f'Frases de la ficha por confirmar · {x["por_confirmar"]}', expanded=bool(x['por_confirmar'])):
        if not frases_por_confirmar():
            st.write('No hay frases por confirmar en esta ficha.')
    with st.expander(f'Clasificación de peligros, frase por frase · {len(x["peligros"])}'):
        etiqueta_h = {f['codigo']: f for f in e['H']}
        if x['peligros']:
            tabla([(q['codigo'], q['texto'], etiqueta_h.get(q['codigo'], {}).get('texto_ficha'), q['clase'], f"{q['nivel']} {q['nivel_texto']}",
                    'Sí' if q['visible'] else f"No: {etiqueta_h.get(q['codigo'], {}).get('motivo')}", q['pagina'])
                   for q in x['peligros']],
                  ['Código', 'Indicación de peligro', 'Texto original de la ficha', 'Clase de peligro', 'Nivel', 'Se muestra', 'Página'])
        else:
            st.write('La ficha no trae indicaciones de peligro: producto no peligroso para efectos de la etiqueta (regla R07).')
        for o in g['omitidos']:
            st.caption(f"No se muestra {o['codigo']} por {o['frase']}: {o['motivo']} (regla {o['regla']}).")
        for nota in r['notas']:
            html(ver.caja('pend', (f"<b>Regla {nota['regla']}.</b> " if nota['regla'] else '') + t(nota['texto'])
                          + ''.join(f'<br>· {t(i)}' for i in nota['indicios'])))
        html('<div class="qr-h4">Cómo se obtuvo el nivel de peligro</div>' + ver.preguntas(x['explicacion_del_nivel']))
    elegidos_n = len(x['que_hacer']) + len(x['accidente'])
    with st.expander(f'Consejos de prudencia: los {x["consejos_disponibles"]} de la ficha y por qué se '
                     + segun_cuantos(elegidos_n, 'eligió este', f'eligieron estos {elegidos_n}')):
        elegidos = {z['codigo']: z for z in x['que_hacer'] + x['accidente']}
        if x['medidas']:
            tabla([(m['codigo'], m['texto'] + (' (completar «...»)' if m['completar'] else ''), grupo['tipo'], ', '.join(m['controla']),
                    f"Sí, puntaje {elegidos[m['codigo']]['puntaje']:g}: {elegidos[m['codigo']]['criterio']}" if m['codigo'] in elegidos else 'No',
                    'Sí' if m['en_etiqueta'] else 'Opcional', m['pagina']) for grupo in x['medidas'] for m in grupo['consejos']],
                  ['Código', 'Consejo de prudencia', 'Tipo', 'Peligro que controla (SGA)', 'Seleccionado para minería', 'En la etiqueta', 'Página'])
            st.caption('Regla C07. Máximo por tipo: ' + ', '.join(f'{k.lower()} {v}' for k, v in tablas['seleccion_p'].items() if v) + '. Criterios y pesos: hoja '
                       'Pertinencia_mineria (borrador por validar del autor). La etiqueta no usa esta selección.')
        else:
            st.write('La ficha no trae consejos de prudencia. La herramienta no los inventa.')
    with st.expander('Riesgo por exposición: semáforo, límites y método'):
        html(f'<div><b>{t(sem["color"])}</b>: {t(sem["condicion"])}. Regla {sem["regla"]}, {t(sem["fila"])}.'
             + (f' Estado del mensaje: {t(sem["estado_del_mensaje"])}.' if 'validar' in str(sem['estado_del_mensaje']).lower() else '') + '</div>')
        html(ver.tabla_de_semaforo(tablas['semaforo']))
        if ir['limites']:
            tabla([(l['id'], l['agente'], f"{l['tipo']} {l['valor']:g} {l['unidad']}" if isinstance(l['valor'], (int, float)) else l['tipo'], l['usar'],
                    '' if l['ir'] is None else num(l['ir']), l['estado'], l['mensaje'] or '') for l in ir['limites']],
                  ['N.º', 'Agente', 'Valor límite', '¿Se usa para Ir?', 'Ir', 'Estado', 'Nota de la tabla'])
        for nota in ir['notas']:
            st.caption(nota)
        html(ver.preguntas(x['explicacion_del_ir']))
    historial = almacen.mediciones_de(ficha['archivo'])
    with st.expander(f'Historial de mediciones · {len(historial)}'):
        if historial:
            tabla([(m['fecha'], '; '.join(f"{k}: " + ' + '.join(f'{c:g} × {h:g} h' if h else f'{c:g}' for c, h in v['tareas']) + f" {v.get('unidad') or ''}"
                                         for k, v in m['mediciones'].items()), 'Sí' if m is historial[-1] else '') for m in reversed(historial)],
                  ['Fecha', 'Medición (límite: concentración × horas)', 'La más reciente'])
            st.caption('Son las mediciones guardadas con «Guardar en el historial», en la carpeta «clasificador SGA/mediciones», un archivo por producto. '
                       'La más reciente es la que usa la lista «Qué revisar primero». No se cargan en una consulta nueva.')
        else:
            st.write('No hay mediciones guardadas de este producto.')
    with st.expander('Protección personal: texto original de la sección 8'):
        original = (ficha.get('proteccion') or {}).get('texto')
        html(f'<div class="qr-original">{t(original, "La extracción no encontró este apartado.")}</div><div class="qr-h4">Versión que se muestra</div>' + ver.proteccion(x)
             + '<div class="qr-nota">Las traducciones se validan y se corrigen en matriz/proteccion_seccion8.xlsx.</div>')
    with st.expander('Estabilidad y reactividad (sección 10) e incompatibilidades'):
        html(ver.incompatibilidades(r))
        if r['reacciones']['ficha']:
            tabla([(a['nombre'], a['texto'] if a['estado'] == 'encontrado' else f"{a['estado']}" + (f": {a['texto']}" if a['texto'] else ''))
                   for a in r['reacciones']['ficha'].values()], ['Apartado', 'Texto de la ficha, sin traducir'])
        if x['marco_normativo']:
            st.caption(f"Marco normativo colombiano anotado en el inventario: {x['marco_normativo']}")
    with st.expander(f'Observaciones de la lectura de la ficha · {len(r["hallazgos_de_la_ficha"])}'):
        st.caption('Se conservan tal como salieron. Pueden pesar en la valoración del riesgo y en las reacciones con otras sustancias.')
        if r['hallazgos_de_la_ficha']:
            tabla([(i['paso'], i['texto']) for i in r['hallazgos_de_la_ficha']], ['Paso', 'Observación'])
        else:
            st.write('Ninguna.')
    with st.expander('Reglas aplicadas y tablas usadas'):
        usadas = list(dict.fromkeys(n['reglas'] + p['reglas'] + g['reglas'] + e['reglas'] + [i['regla'] for i in r['notas'] if i['regla']]
                                    + [sem['regla'], r['seleccion_p']['regla']]))
        tabla([(i, tablas['reglas'][i]['tema'], tablas['reglas'][i]['regla'], tablas['reglas'][i]['fuente'], tablas['reglas'][i]['caracter'])
               for i in usadas if i in tablas['reglas']], ['Regla', 'Tema', 'Texto de la regla', 'Fuente', 'Carácter'])
        tabla(_versiones(), ['Tabla', 'Archivo', 'Versión declarada', 'Fecha del archivo'])
        st.caption('El nivel, el semáforo, la palabra de advertencia, los pictogramas y la selección de consejos los calcula un programa fijo que lee estas tablas. '
                   'Ningún modelo de inteligencia artificial interviene en la clasificación.')


# ---------- Pantalla ----------
if x is not None:
    for aviso in r['avisos_de_reglas']:
        st.error(aviso)
if x is not None and st.session_state.get('ver_detalle'):
    html(f'<div class="qr-marca">Detalle técnico</div><div class="qr-titulo">{t(x["nombre"])}</div>')
    detalle_tecnico()
    st.stop()

paso = st.session_state['paso']
nombre_del_paso, color, que_se_hace = PASOS[paso - 1]
if x is None:
    html(f'<div class="qr-marca">{t(ver.LEMA)}</div><div class="qr-titulo">Consulta de un producto químico</div><div class="qr-sub">{t(DESCRIPCION)}</div>')
else:
    html(f'<div class="qr-marca">{"Ficha cargada por el usuario" if x["nueva"] else "Producto del inventario"} · {t(st.session_state["perfil"])}</div>'
         f'<div class="qr-titulo">{t(x["nombre"])}</div><div class="qr-sub">Nombre en la ficha: {t(r["producto"])} · CAS: {t(", ".join(r["cas"]))}</div>')
html('<div class="qr-ruta">' + ''.join(
    f'<div class="qr-etapa {"activa" if i == paso else "hecha" if i < paso else ""}" style="--c:{c}"><i>{"✓" if i < paso else i}</i>'
    f'<span>{nombre}<small>{"Usted está aquí" if i == paso else "Completado" if i < paso else "Pendiente"}</small>{f"<em>{t(REQUISITOS[i - 1])}</em>" if REQUISITOS[i - 1] else ""}</span></div>'
    for i, (nombre, c, _) in enumerate(PASOS, 1)) + '</div>')
if paso < len(PASOS):   # en el resultado no se repite el encabezado: los dos indicadores deben verse sin bajar
    html(f'<div class="qr-cabeza" style="--c:{color}"><small>Paso {paso} de {len(PASOS)}</small><b>{nombre_del_paso}</b><span>{que_se_hace}</span></div>')
for pendiente in st.session_state.pop('validacion', None) or []:
    html(validacion('No se puede continuar todavía', *pendiente))
(paso_producto, paso_exposicion, paso_resultado)[paso - 1]()
botones_de_paso()
