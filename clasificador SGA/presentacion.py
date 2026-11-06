# -*- coding: utf-8 -*-
"""Piezas de presentación que comparten la pantalla (app.py) y el informe (reporte_html.py): colores, estilos y
fragmentos de HTML. Así un mismo dato se ve igual en los dos lados y se escribe una sola vez.
Aquí no se decide ni se calcula nada: cada función recibe lo que ya armó lectura.py o el motor y lo dibuja.
Los colores son de presentación. El nivel de peligro va en tonos de azul; verde, amarillo y rojo se reservan para
el semáforo del riesgo por exposición, cuyo color lo decide el motor con la hoja Semaforo_Ir."""
import base64
from html import escape
from pictogramas_svg import svg_de_pictograma

COLOR_NIVEL = {0: '#eef1f5', 1: '#dbe9f9', 2: '#b7d3f2', 3: '#1f5fa8', 4: '#0b2a4f'}   # fondo
TINTA_NIVEL = {0: '#3c4a5c', 1: '#123a63', 2: '#0b2a4f', 3: '#ffffff', 4: '#ffffff'}   # texto
# Cómo se pinta cada color que nombra la hoja Semaforo_Ir: (luz, fondo de la tarjeta, texto)
COLOR_SEMAFORO = {'Verde': ('#1a7f37', '#eaf6ed', '#0f4d21'), 'Amarillo': ('#e3b100', '#fff8dc', '#5c4600'),
                  'Rojo': ('#b3261e', '#fdecec', '#7a1410'), 'Gris': ('#8a94a0', '#f2f4f7', '#3c4a5c')}

CSS = """
.qr-marca { font-size: .78rem; font-weight: 700; letter-spacing: .08em; text-transform: uppercase; color: #1f5fa8; }
.qr-titulo { font-size: 2.05rem; font-weight: 700; letter-spacing: -0.02em; line-height: 1.12; margin: .15rem 0 .3rem 0; color: #0b2a4f; }
.qr-sub { color: #5b6672; font-size: .95rem; margin-bottom: .6rem; }
.qr-par { display: grid; grid-template-columns: 1fr 1fr; gap: 14px; margin: .4rem 0 .7rem 0; }
.qr-ind { border-radius: 16px; padding: 14px 16px; border: 2px solid #dfe5ec; background: #ffffff; }
.qr-ind > small { display: block; font-size: .78rem; font-weight: 800; letter-spacing: .07em; text-transform: uppercase; color: #5b6672; margin-bottom: .4rem; }
.qr-luz { display: flex; gap: 14px; align-items: center; }
.qr-bombillo { flex: none; width: 4.2rem; height: 4.2rem; border-radius: 50%; display: flex; align-items: center; justify-content: center; font-weight: 800; font-size: 1.9rem; }
.qr-luz b { display: block; font-size: 1.55rem; line-height: 1.1; letter-spacing: -.01em; }
.qr-luz span { display: block; font-size: .95rem; line-height: 1.35; margin-top: .15rem; }
.qr-ind em { display: block; font-style: normal; font-size: .82rem; color: #5b6672; margin-top: .45rem; }
.qr-pictos { display: flex; gap: 6px; flex-wrap: wrap; margin-top: .45rem; } .qr-pictos img { width: 44px; height: 44px; }
.qr-nota { color: #5b6672; font-size: .86rem; line-height: 1.4; }
.qr-h3 { font-size: 1.05rem; font-weight: 700; margin: 0 0 .35rem 0; color: #0b2a4f; }
.qr-h4 { font-size: .9rem; font-weight: 700; margin: .7rem 0 .2rem 0; color: #1f5fa8; }
.qr-lista { margin: 0; padding-left: 1.1rem; } .qr-lista li { margin-bottom: .32rem; line-height: 1.38; }
.qr-cod { display: inline-block; font-size: .72rem; font-weight: 700; color: #0b2a4f; background: #e8f1fb; border-radius: 6px; padding: 1px 7px; margin-right: 6px; white-space: nowrap; }
.qr-pag { font-size: .76rem; color: #6b7682; }
.qr-caja { border-radius: 12px; padding: 11px 14px; margin: .45rem 0; line-height: 1.4; font-size: .93rem; }
.qr-accion { background: #e8f1fb; border: 1px solid #b7d3f2; color: #0b2a4f; }
.qr-pend { background: #fff6dc; border: 1px solid #f2d27a; color: #6b4a00; }
.qr-mal { background: #fdecec; border: 1px solid #f1b3b3; color: #8b1a1a; }
.qr-bien { background: #e9f5ec; border: 1px solid #b6dabf; color: #14532d; }
.qr-info { background: #f2f5f9; border: 1px solid #d5dde6; color: #2b3a4d; }
table.qr-tabla { width: 100%; border-collapse: collapse; font-size: .88rem; }
table.qr-tabla th { text-align: left; font-size: .74rem; letter-spacing: .04em; text-transform: uppercase; color: #5b6672; padding: 8px 10px; border-bottom: 1px solid #cfd8e2; }
table.qr-tabla td { padding: 8px 10px; border-bottom: 1px solid #e8edf2; vertical-align: top; line-height: 1.35; overflow-wrap: anywhere; }
.qr-chip { display: inline-block; padding: 2px 10px; border-radius: 999px; font-weight: 700; font-size: .8rem; white-space: nowrap; }
.qr-original { color: #555; font-style: italic; }
.qr-desliza { overflow-x: auto; }
@media (max-width: 760px) { .qr-par { grid-template-columns: 1fr; } }
"""


def t(valor, vacio='—'):
    """Texto seguro para poner dentro de HTML."""
    return escape(str(valor)) if valor not in (None, '') else vacio


def num(valor, cifras=3):
    """Número con coma decimal, como se escribe en Colombia."""
    return f'{valor:.{cifras}g}'.replace('.', ',')


def caja(clase, contenido):
    return f'<div class="qr-caja qr-{clase}">{contenido}</div>'


def lista(elementos):
    return '<ul class="qr-lista">' + ''.join(f'<li>{e}</li>' for e in elementos) + '</ul>'


def pagina(numero, seccion=None):
    if not numero:
        return ''
    return f' <span class="qr-pag">FDS{", secc. " + str(seccion) if seccion else ""}, p. {numero}</span>'


def chip_de_nivel(n):
    """Nivel de peligro en pequeño, para listas."""
    return (f'<span class="qr-chip" style="background:{COLOR_NIVEL.get(n["nivel"], "#fff")};color:{TINTA_NIVEL.get(n["nivel"], "#111")}">'
            f'Nivel {n["nivel"]} · {t(n["nombre"])}</span>')


def chip_de_semaforo(semaforo):
    """Color del riesgo por exposición en pequeño, para listas."""
    luz, _, _ = COLOR_SEMAFORO.get(semaforo['color'], COLOR_SEMAFORO['Gris'])
    rotulo = semaforo['color'] if semaforo['ir'] is not None or semaforo['clave'] in ('rojo_alerta', 'alerta_ok') else 'Sin medición'
    return (f'<span class="qr-chip" style="background:{luz};color:{"#1d2530" if semaforo["color"] == "Amarillo" else "#fff"}">{t(rotulo)}</span>')


def tabla_de_semaforo(semaforo):
    """Filas de la hoja Semaforo_Ir, cada una pintada con su color para reconocerla de un vistazo."""
    filas = ''
    for clave, v in semaforo.items():
        luz, fondo, tinta = COLOR_SEMAFORO.get(v['color'], COLOR_SEMAFORO['Gris'])
        estilo = f'background:{fondo};color:{tinta}'
        filas += (f'<tr style="{estilo}"><td style="border-left:6px solid {luz};white-space:nowrap;overflow-wrap:normal;{estilo}">{t(clave)}</td><td style="white-space:nowrap;overflow-wrap:normal;{estilo}"><b>{t(v["color"])}</b></td>'
                  f'<td style="{estilo}">{t(v["condicion"])}</td><td style="{estilo}">{"" if v["desde"] is None else num(v["desde"])}</td>'
                  f'<td style="{estilo}">{"" if v["hasta"] is None else num(v["hasta"])}</td><td style="{estilo}">{t(v["mensaje"])}</td></tr>')
    return ('<div class="qr-desliza"><table class="qr-tabla"><tr><th>Fila</th><th>Color</th><th>Condición</th><th>Ir desde</th><th>Ir hasta</th>'
            f'<th>Mensaje</th></tr>{filas}</table></div>')


def imagen_de_pictograma(p):
    datos = base64.b64encode(svg_de_pictograma(p['codigo'], 'rojo', p['simbolo']).encode('utf-8')).decode('ascii')
    return f'<img src="data:image/svg+xml;base64,{datos}" alt="{t(p["simbolo"])}" title="{p["codigo"]} · {t(p["simbolo"])}">'


def indicadores(r, x):
    """Los dos indicadores, lado a lado y del mismo tamaño: peligro del producto y riesgo por exposición."""
    n, s = r['nivel'], r['semaforo']
    pictos = r['pictogramas']['pictogramas']
    peligro = (f'<div class="qr-ind"><small>Peligro del producto</small><div class="qr-luz">'
               f'<div class="qr-bombillo" style="background:{COLOR_NIVEL.get(n["nivel"], "#fff")};color:{TINTA_NIVEL.get(n["nivel"], "#111")}">{n["nivel"]}</div>'
               f'<div><b style="color:#0b2a4f">{t(n["nombre"])}</b><span>Nivel {n["nivel"]} de 4'
               + (f' · palabra de advertencia: <b style="display:inline;font-size:inherit">{t(r["palabra_de_advertencia"]["palabra"])}</b>'
                  if r['palabra_de_advertencia']['palabra'] else '') + '</span></div></div>'
               + ('<div class="qr-pictos">' + ''.join(imagen_de_pictograma(p) for p in pictos) + '</div>' if pictos else '')
               + '<em>Depende solo del producto. ' + (f'Lo definen {t(", ".join(n["determinantes"]))}.' if n['determinantes'] else 'La ficha no trae indicaciones de peligro.')
               + '</em></div>')
    luz, fondo, tinta = COLOR_SEMAFORO.get(s['color'], COLOR_SEMAFORO['Gris'])
    sin_medir = s['ir'] is None and s['clave'] not in ('rojo_alerta', 'alerta_ok')
    titulo = ('No aplica' if s['clave'] == 'sin_limite' else 'Sin medición') if sin_medir else s['color']
    riesgo = (f'<div class="qr-ind" style="background:{fondo};border-color:{luz}"><small>Riesgo por exposición</small><div class="qr-luz">'
              f'<div class="qr-bombillo" style="background:{luz}"></div>'
              f'<div><b style="color:{tinta}">{t(titulo)}</b><span style="color:{tinta}">{t(s["mensaje"])}</span></div></div>'
              f'<em>{t(x["indice"], "Depende de lo que se mida en el puesto de trabajo.")}'
              + (f' · Medición del {t(x["fecha_de_medicion"])}.' if x.get('fecha_de_medicion') else '') + '</em></div>')
    return f'<div class="qr-par">{peligro}{riesgo}</div>'


def frases_h(x):
    """«Qué puede pasar»: todas las indicaciones de peligro de la ficha. En lenguaje claro si el autor ya lo validó; si no, el texto oficial."""
    if not x['visibles']:
        return caja('info', 'La ficha no trae indicaciones de peligro reconocidas. Con esa información el producto no se clasifica como peligroso.')
    h = lista((t(p['claro']) + f' <span class="qr-pag">Texto oficial: {t(p["texto"])}.</span>' if p['claro'] else t(p['texto']))
              + f' <span class="qr-pag">{p["codigo"]}{" · p. " + str(p["pagina"]) if p["pagina"] else ""}</span>' for p in x['visibles'])
    if x['asfixiante']:
        h += caja('pend', '<b>Asfixiante simple.</b> Desplaza el oxígeno del aire, aunque no tenga una frase de peligro por eso.')
    return h


def consejos(elementos, con_criterio=False):
    """Consejos de prudencia elegidos para minería, con su código, su página y, si se pide, el criterio por el que se eligieron."""
    return lista(t(z['texto']) + (' <b>(completar «...» con el dato de la ficha)</b>' if z['completar'] else '')
                 + f' <span class="qr-pag">{z["codigo"]}{" · p. " + str(z["pagina"]) if z["pagina"] else ""}'
                 + (f' · {t(z["criterio"])}' if con_criterio else '') + '</span>' for z in elementos)


def proteccion(x):
    """Protección personal de la sección 8, en una línea, con su página y el estado de la traducción."""
    p = x['proteccion']
    if not p['texto']:
        return f'<div>La ficha no dio este apartado de forma reconocible. Consulte su sección 8{", página " + str(p["pagina"]) if p["pagina"] else ""}.</div>'
    return (f'<div class="{"qr-original" if p["original"] else ""}">{t(p["texto"])}</div>'
            f'<div class="qr-pag">Sección 8 de la ficha{", p. " + str(p["pagina"]) if p["pagina"] else ""} · {t(p["estado"])}.</div>')


def accion(x):
    """Acción recomendada para el profesional de SST: a lo sumo tres puntos, cada uno con su porqué."""
    return '<ol class="qr-lista">' + ''.join(f'<li><b>{t(z["que"])}</b> <span class="qr-pag">{t(z["porque"])}</span></li>' for z in x['recomendaciones']) + '</ol>'


def factores(pares):
    return ('<div class="qr-desliza"><table class="qr-tabla"><tr><th>Dato</th><th>Qué se encontró y de dónde sale</th></tr>'
            + ''.join(f'<tr><td><b>{t(nombre)}</b></td><td>{t(valor)}</td></tr>' for nombre, valor in pares) + '</table></div>')


def preguntas(pares):
    return ''.join(f'<div class="qr-h4">{t(p)}</div><div style="line-height:1.45">{t(r)}</div>' for p, r in pares)


def incompatibilidades(r):
    """Con qué no mezclar ni almacenar: en español si el inventario lo trae; si no, el texto original de la ficha."""
    rx = r['reacciones']
    if rx['excel_reacciones'] or rx['excel_incompatibilidades']:
        return (''.join(f'<div style="margin-bottom:.35rem">{t(i)}</div>' for i in (rx['excel_reacciones'], rx['excel_incompatibilidades']) if i)
                + '<div class="qr-pag">Fuente: inventario del proyecto.</div>')
    incompatibles = rx['ficha'].get('materiales_incompatibles')
    if incompatibles and incompatibles['estado'] == 'encontrado':
        return (f'<div>Materiales incompatibles según la ficha (texto original, sin traducir):</div>'
                f'<div class="qr-original">{t(incompatibles["texto"])}</div>')
    return caja('pend', 'La ficha no dio esta información de forma reconocible. Consulte su sección 10.')


# =====================================================================================================================
# IDENTIDAD DE LA PANTALLA (QuimNexo). Todo lo que sigue lo usa solo la pantalla (app.py).
# El informe de consulta y las etiquetas NO lo usan: ellos siguen con el bloque CSS y los colores de arriba, sin cambios,
# y con el nombre que tiene lectura.NOMBRE. Por eso el nombre de la pantalla y su paleta viven aquí, aparte.
# =====================================================================================================================
import os as _os

NOMBRE_EN_PANTALLA = 'QuimNexo'
LEMA = 'De la ficha de seguridad a la decisión preventiva'
_MARCA = _os.path.join(_os.path.dirname(_os.path.abspath(__file__)), '..', 'marca')
LOGOTIPO = _os.path.join(_MARCA, 'logotipo_horizontal.svg')
SIMBOLO = _os.path.join(_MARCA, 'simbolo.svg')
FAVICON = _os.path.join(_MARCA, 'favicon_64.png')

# Paleta de la marca, en un solo lugar. El verde, el ámbar y el rojo no están aquí a propósito: quedan reservados para el
# semáforo del índice Ir y los avisos (COLOR_SEMAFORO y las cajas de arriba). Los niveles de peligro siguen en tonos de azul.
PALETA = {
    '--qn-marca': '#0F3D4C',            # azul petróleo oscuro: marca, títulos
    '--qn-marca-profundo': '#0A2A35',   # paso activo y fondos oscuros
    '--qn-accion': '#0B6E82',           # turquesa oscuro: botones, enlaces, navegación
    '--qn-acento': '#17A2B8',           # turquesa: solo detalles gráficos grandes
    '--qn-accion-suave': '#E3F1F4',     # fondo de lo seleccionado
    '--qn-texto': '#17232B',
    '--qn-texto-suave': '#44525C',
    '--qn-borde': '#C9D3DA',
    '--qn-borde-de-casilla': '#7C8B96',
    '--qn-fondo': '#F4F6F8',
    '--qn-superficie': '#FFFFFF',
}

# Escala de letra (con la letra base de 18 px que fija .streamlit/config.toml): título 36, sección 26, tarjeta 21,
# texto, formularios y botones 18, tablas 17, metadatos 16. Interlineado de 1,5. Nada baja de 16 px.
CSS_DE_PANTALLA = ':root {' + ' '.join(f'{k}: {v};' for k, v in PALETA.items()) + '}' + """
.stApp, .stMarkdown, .stMarkdown p, .stMarkdown li { line-height: 1.5; color: var(--qn-texto); }
.block-container { padding-top: 2.2rem; padding-bottom: 4rem; max-width: 1240px; }
section[data-testid="stSidebar"] { background: var(--qn-superficie); border-right: 1px solid var(--qn-borde); }
/* Tarjetas: superficie blanca sobre el fondo gris, sin borde ni sombra */
[class*="st-key-tarjeta_"] { border-color: transparent !important; border-radius: 16px !important; background: var(--qn-superficie); box-shadow: none; }
/* Botones cómodos, con el rótulo completo */
.stButton button, .stDownloadButton button, [data-testid="stPopover"] button { min-height: 3rem; font-weight: 600; white-space: normal; height: auto; padding-top: .5rem; padding-bottom: .5rem; }
.stButton button p, .stDownloadButton button p, [data-testid="stPopover"] button p { font-size: 1rem; font-weight: 600; line-height: 1.3; }
/* Formularios: rótulos visibles y casillas con borde que se distingue */
[data-testid="stWidgetLabel"] p, label p { font-size: 1rem !important; font-weight: 600; color: var(--qn-texto); }
[data-testid="stCaptionContainer"] p { font-size: .89rem !important; color: var(--qn-texto-suave) !important; line-height: 1.5; }
input, textarea, [data-baseweb="select"] div { font-size: 1rem !important; }
[data-baseweb="input"], [data-baseweb="base-input"], [data-baseweb="select"] > div { border-color: var(--qn-borde-de-casilla) !important; }
[data-testid="stExpander"] details { border-color: var(--qn-borde); border-radius: 14px; background: var(--qn-superficie); }
[data-testid="stExpander"] summary p { font-size: 1rem !important; font-weight: 600; }
[data-testid="stTable"] td, [data-testid="stTable"] th { font-size: .94rem; line-height: 1.45; }
/* Textos que Streamlit trae en inglés en el cargador de archivos y en las casillas: se reemplazan por español */
[data-testid="stFileUploader"] button [data-testid="stMarkdownContainer"] p { font-size: 0 !important; }
[data-testid="stFileUploader"] button [data-testid="stMarkdownContainer"] p::after { content: "Buscar archivo"; font-size: 1rem; }
[data-testid="stFileUploaderDropzoneInstructions"], [data-testid="stFileUploaderDropzoneInstructions"] * { font-size: 0 !important; }
[data-testid="stFileUploaderDropzoneInstructions"]::after { content: "Archivo PDF de hasta 200 MB"; font-size: .89rem; color: var(--qn-texto-suave); }
[data-testid="InputInstructions"] { display: none; }
/* Piezas propias */
.qr-marca { font-size: .89rem; font-weight: 700; letter-spacing: .08em; text-transform: uppercase; color: var(--qn-accion); }
.qr-titulo { font-size: 2rem; font-weight: 700; letter-spacing: -.02em; line-height: 1.15; margin: .1rem 0 .25rem 0; color: var(--qn-marca); }
.qr-sub { font-size: 1rem; color: var(--qn-texto-suave); margin-bottom: .8rem; line-height: 1.5; }
.qr-h3 { font-size: 1.17rem; font-weight: 700; margin: 0 0 .45rem 0; color: var(--qn-marca); line-height: 1.3; }
.qr-h4 { font-size: 1rem; font-weight: 700; margin: .9rem 0 .25rem 0; color: var(--qn-accion); }
.qr-seccion { font-size: 1.45rem; font-weight: 700; color: var(--qn-marca); letter-spacing: -.01em; margin: 1.6rem 0 .6rem 0; line-height: 1.25; }
.qr-nota { font-size: .89rem !important; color: var(--qn-texto-suave); line-height: 1.5; }
.qr-pag { font-size: .89rem !important; color: var(--qn-texto-suave); }
.qr-cod { font-size: .86rem !important; background: var(--qn-accion-suave); color: var(--qn-marca); }
.qr-lista li { margin-bottom: .45rem; line-height: 1.5; }
.qr-caja { font-size: 1rem; line-height: 1.5; border-radius: 12px; padding: .8rem 1rem; }
.qr-accion { background: var(--qn-accion-suave); border: 0; border-left: 6px solid var(--qn-accion); border-radius: 0 12px 12px 0; color: var(--qn-texto); }
.qr-info { background: var(--qn-superficie); border: 1px solid var(--qn-borde); color: var(--qn-texto); }
/* Los dos indicadores: arriba, grandes y del mismo tamaño */
.qr-par { gap: 18px; margin: .6rem 0 1rem 0; }
.qr-ind { padding: 1.1rem 1.3rem; border-radius: 18px; }
.qr-ind > small { font-size: .89rem !important; color: var(--qn-texto-suave); }
.qr-bombillo { width: 5rem; height: 5rem; font-size: 2.4rem !important; }
.qr-luz b { font-size: 2rem !important; } .qr-luz span { font-size: 1rem !important; line-height: 1.45; }
.qr-luz span b { font-size: inherit !important; display: inline; }
.qr-ind em { font-size: .89rem !important; color: var(--qn-texto-suave); }
.qr-pictos img { width: 52px; height: 52px; }
table.qr-tabla { font-size: .94rem; } table.qr-tabla th { font-size: .86rem !important; color: var(--qn-texto-suave); } table.qr-tabla td { font-size: .94rem !important; line-height: 1.45; }
.qr-chip { font-size: .89rem; padding: 3px 12px; }
/* Navegación de tres pasos: dónde estoy, qué falta y qué es obligatorio */
.qr-ruta { display: grid; grid-template-columns: repeat(auto-fit, minmax(11.5rem, 1fr)); gap: 10px; margin: .4rem 0 1.2rem 0; align-items: start; }
.qr-etapa { display: flex; align-items: flex-start; gap: 12px; padding: .7rem .9rem; border-radius: 14px; background: var(--qn-superficie); border: 2px solid var(--qn-borde); color: var(--qn-texto-suave); }
.qr-etapa i { flex: none; width: 2.2rem; height: 2.2rem; border-radius: 50%; display: flex; align-items: center; justify-content: center; font-style: normal; font-weight: 800;
              font-size: 1rem !important; background: var(--qn-fondo); color: var(--qn-texto-suave); border: 2px solid var(--qn-borde-de-casilla); }
.qr-etapa span { font-size: 1.06rem !important; font-weight: 700; line-height: 1.25; }
.qr-etapa small { display: block; font-size: .89rem; font-weight: 600; }
.qr-etapa em { display: block; font-style: normal; font-weight: 400; font-size: .89rem; line-height: 1.35; margin-top: .1rem; }
.qr-etapa.hecha em { display: none; }   /* lo obligatorio de un paso ya completado no hace falta repetirlo */
.qr-etapa.hecha { border-color: var(--qn-accion); color: var(--qn-accion); } .qr-etapa.hecha i { background: var(--qn-accion); color: #fff; border-color: var(--qn-accion); }
.qr-etapa.activa { border-color: var(--qn-marca-profundo); background: var(--qn-marca-profundo); color: #fff; } .qr-etapa.activa i { background: #fff; color: var(--qn-marca-profundo); border-color: #fff; }
.qr-cabeza { margin: 0 0 1.1rem 0; }
.qr-cabeza small { display: block; font-size: .89rem !important; font-weight: 700; letter-spacing: .08em; text-transform: uppercase; color: var(--qn-accion); }
.qr-cabeza b { display: block; font-size: 1.45rem !important; line-height: 1.25; color: var(--qn-marca); letter-spacing: -.01em; }
.qr-cabeza span { font-size: 1rem !important; color: var(--qn-texto-suave); }
.qr-lateral { display: flex; align-items: center; gap: 8px; padding: .4rem .6rem; border-radius: 10px; font-weight: 600; color: var(--qn-texto-suave); font-size: 1rem !important; border-left: 5px solid transparent; }
.qr-lateral.activa { background: var(--qn-accion-suave); color: var(--qn-marca); border-left-color: var(--qn-accion); } .qr-lateral.hecha { color: var(--qn-accion); }
.qr-validacion { border: 3px solid #b3261e; background: #fdecec; color: #5f1410; border-radius: 14px; padding: .9rem 1.1rem; margin: .6rem 0; font-size: 1rem; line-height: 1.5; }
.qr-validacion > b { font-size: 1.17rem; display: block; margin-bottom: .3rem; }
.qr-aviso { border-left: 6px solid #e0a100; background: #fff6dc; color: #5c4600; border-radius: 0 10px 10px 0; padding: .6rem 1rem; margin: .5rem 0; font-size: 1rem; line-height: 1.5; }
@media (max-width: 900px) { .qr-titulo { font-size: 1.6rem; } }
"""

# ---------- Modo oscuro (versión 2.01) ----------
# Solo cambia la pantalla. El informe y las etiquetas se siguen viendo e imprimiendo en blanco, porque son documentos.
# Los colores del semáforo y de los niveles de peligro no cambian: son información, no decoración.
LOGOTIPO_OSCURO = _os.path.join(_MARCA, 'logotipo_horizontal_fondo_oscuro.svg')
SIMBOLO_OSCURO = _os.path.join(_MARCA, 'simbolo_fondo_oscuro.svg')

# Tema de Streamlit para el modo oscuro (el claro está en .streamlit/config.toml). Misma marca, tonos aclarados para que se lean.
TEMA_OSCURO = {
    'base': 'dark',
    'primaryColor': '#2FB3C8',
    'backgroundColor': '#0E1A20',
    'secondaryBackgroundColor': '#16262E',
    'textColor': '#E6EDF1',
    'linkColor': '#5CC8DA',
    'borderColor': '#34474F',
}
PALETA_OSCURA = {
    '--qn-marca': '#CFE6EC',
    '--qn-marca-profundo': '#0B6E82',
    '--qn-accion': '#5CC8DA',
    '--qn-acento': '#17A2B8',
    '--qn-accion-suave': '#163A44',
    '--qn-texto': '#E6EDF1',
    '--qn-texto-suave': '#A9B7BF',
    '--qn-borde': '#34474F',
    '--qn-borde-de-casilla': '#6E828D',
    '--qn-fondo': '#0E1A20',
    '--qn-superficie': '#16262E',
}
# Piezas que en el bloque CSS de arriba tienen colores claros fijos (lo comparte el informe, por eso no se tocan allá)
CSS_MODO_OSCURO = ':root {' + ' '.join(f'{k}: {v};' for k, v in PALETA_OSCURA.items()) + '}' + """
.qr-ind { background: var(--qn-superficie); border-color: var(--qn-borde); }
.qr-ind:first-child b[style] { color: var(--qn-marca) !important; }
.qr-ind[style] > small, .qr-ind[style] em { color: #44525C !important; }   /* la tarjeta del semáforo conserva su fondo claro */
.qr-pend { background: #3a3012; border-color: #8a6d1a; color: #f5dc8e; }
.qr-mal, .qr-validacion { background: #3d1715; border-color: #b3261e; color: #f6c2be; }
.qr-bien { background: #12301c; border-color: #2f7d47; color: #b9e5c5; }
.qr-aviso { background: #3a3012; color: #f5dc8e; }
.qr-original { color: var(--qn-texto-suave); }
table.qr-tabla th { border-bottom-color: var(--qn-borde); } table.qr-tabla td { border-bottom-color: #22343C; }
.qr-producto { color: var(--qn-marca) !important; }
"""


def tema_es_oscuro():
    """Si la aplicación ya está en modo oscuro (el tema vale para todas las ventanas abiertas de este programa)."""
    try:
        from streamlit import config
        return config.get_option('theme.base') == 'dark'
    except Exception:
        return False


def poner_tema(oscuro):
    """Cambia el tema de Streamlit entre claro y oscuro. Devuelve True si hubo cambio (hay que volver a dibujar la pantalla).
    Usa la configuración interna de Streamlit: si una versión futura la cambia, la pantalla sigue en claro y no se rompe nada."""
    try:
        import tomllib
        from streamlit import config
        with open(_os.path.join(_os.path.dirname(_os.path.abspath(__file__)), '..', '.streamlit', 'config.toml'), 'rb') as f:
            claro = tomllib.load(f)['theme']   # el tema claro se lee siempre de su archivo, no se copia aquí
        deseado = TEMA_OSCURO if oscuro else {k: claro.get(k) for k in TEMA_OSCURO}
        if all(config.get_option(f'theme.{k}') == v for k, v in deseado.items()):
            return False
        for k, v in deseado.items():
            config.set_option(f'theme.{k}', v)
        return True
    except Exception:
        return False
