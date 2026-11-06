# -*- coding: utf-8 -*-
"""Informe de consulta de un producto: documento administrativo de UNA hoja carta, para archivar, imprimir o
guardar como PDF desde el navegador. Es un documento distinto de la etiqueta (etiqueta_html.py).
Trae: producto y fecha, los dos indicadores (peligro del producto y riesgo por exposición, con su mensaje y el
número del índice Ir), las frases H, los consejos P seleccionados para minería, la protección personal, la acción
recomendada y la fuente de los datos. Si se pide, se le agrega un anexo técnico en hojas aparte.
No decide ni recalcula nada: dibuja lo que entregan motor.py y lectura.py, con las piezas de presentacion.py.
Uso de prueba:
    py "clasificador SGA/reporte_html.py" 14_
    py "clasificador SGA/reporte_html.py" 14_ --anexo
    py "clasificador SGA/reporte_html.py" 14_ --medicion L-21=0.3x8
Requiere: haber corrido antes  py extraccion/extraer.py FDS"""
import os, sys, time
from html import escape
import openpyxl
import presentacion as ver
from presentacion import t as _t, num
from tablas import cargar_tablas, MATRIZ, LIMITES, DICCIONARIO, FICHAS
from motor import clasificar, cargar_fichas
from lectura import construir_lectura, segun_cuantos, NOMBRE, NO_REEMPLAZA, VERSION

SALIDAS = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'salidas')

ESTILO = """
*{box-sizing:border-box}
html{font-size:11.5px}
body{font-family:-apple-system,'Segoe UI',Helvetica,Arial,sans-serif;line-height:1.32;color:#1d2530;margin:0 auto;padding:8mm;max-width:200mm;background:#fff}
h2{font-size:1.08rem;margin:.7rem 0 .25rem 0;color:#0b2a4f;padding-bottom:.15rem;border-bottom:2px solid #1f5fa8}
h3{font-size:1rem;margin:.9rem 0 .25rem 0;color:#0b2a4f}p{margin:0 0 .3rem 0}h1,p,li{overflow-wrap:break-word}
.qr-titulo{font-size:1.55rem;margin:0}.qr-par{margin:.5rem 0 .3rem 0;gap:10px}.qr-ind{padding:9px 12px;border-radius:10px}
.qr-bombillo{width:3.2rem;height:3.2rem;font-size:1.5rem}.qr-luz b{font-size:1.3rem}.qr-pictos img{width:34px;height:34px}
.qr-lista li{margin-bottom:.16rem;line-height:1.3}.qr-caja{padding:6px 10px;margin:.25rem 0}
.datos{color:#5b6672;margin:.1rem 0 0 0}
.columnas{display:grid;grid-template-columns:1fr 1fr;gap:0 7mm}
table{border-collapse:collapse;width:100%;margin:.2rem 0 .5rem 0}
th,td{border-bottom:1px solid #d5dde6;padding:4px 6px;vertical-align:top;text-align:left;overflow-wrap:anywhere}
th{background:#e8f1fb;color:#0b2a4f;font-weight:700;font-size:.85rem}
.chico{font-size:.88rem}table.chico td:first-child{white-space:nowrap;overflow-wrap:normal}.sub{color:#5b6672}.original{color:#444;font-style:italic}.no{color:#777}.supera{color:#b3261e;font-weight:700}
.pendiente{border-left:4px solid #e0a100;background:#fff6dc;padding:6px 10px;margin:6px 0;border-radius:0 6px 6px 0}
.alerta{border-left:4px solid #b3261e;background:#fdecec;color:#8b1a1a;padding:6px 10px;margin:6px 0;font-weight:700;border-radius:0 6px 6px 0}
.firmas{display:grid;grid-template-columns:1fr 1fr;gap:10mm;margin-top:9mm}.firmas div{border-top:1px solid #5b6672;padding-top:2px;color:#5b6672;font-size:.85rem}
.anexo{break-before:page}
dl{margin:0}dt{font-weight:700;margin-top:.4rem;color:#0b2a4f}dd{margin:0 0 .25rem 0}
.pie{margin-top:3mm;padding-top:1.5mm;border-top:1px solid #d5dde6;color:#5b6672;font-size:.8rem}
.boton{font:inherit;font-weight:600;padding:8px 20px;border-radius:999px;border:0;background:#1f5fa8;color:#fff;cursor:pointer;font-size:13px}
@page{size:letter portrait;margin:10mm}
@media print{body{padding:0;max-width:none}h2,h3{break-after:avoid}tr,li,.qr-ind,.firmas{break-inside:avoid}.no-imprimir{display:none}
 .qr-ind,.qr-bombillo,.qr-caja,.qr-chip,.qr-cod,th{-webkit-print-color-adjust:exact;print-color-adjust:exact}}
@media screen and (max-width:700px){body{padding:5mm}.columnas,.firmas{grid-template-columns:1fr}}
"""


def _versiones():
    """Versión declarada en la hoja LEEME de cada tabla y fecha de su archivo, para dejar constancia."""
    filas = []
    for nombre, ruta in (('Matriz de decisión', MATRIZ), ('Tabla de límites', LIMITES), ('Diccionario del SGA', DICCIONARIO), ('Excel de fichas', FICHAS)):
        version = ''
        libro = openpyxl.load_workbook(ruta, read_only=True, data_only=True)
        try:
            if 'LEEME' in libro.sheetnames:
                version = next((str(f[1]) for f in libro['LEEME'].iter_rows(values_only=True) if f and f[0] == 'Versión'), '')
        finally:
            libro.close()
        filas.append((nombre, os.path.basename(ruta), version.split('—')[0].strip()[:20], time.strftime('%Y-%m-%d %H:%M', time.localtime(os.path.getmtime(ruta)))))
    return filas


def _tabla(encabezado, filas, clase=''):
    if not filas:
        return ''
    cuerpo = ''.join('<tr>' + ''.join(f'<td>{c}</td>' for c in f) + '</tr>' for f in filas)
    return f'<table class="{clase}"><tr>' + ''.join(f'<th>{escape(h)}</th>' for h in encabezado) + f'</tr>{cuerpo}</table>'


def _preguntas(pares):
    return '<dl>' + ''.join(f'<dt>{_t(p)}</dt><dd>{_t(r)}</dd>' for p, r in pares) + '</dl>'


def _informe(r, x):
    """Cuerpo del informe de una hoja. x = lectura práctica."""
    s = r['semaforo']
    h = [ver.indicadores(r, x)]
    for a in r['indice_ir']['alertas']:
        h.append(f'<div class="alerta">Alerta de seguridad: {_t(a)}</div>')
    h.append('<div class="columnas"><div><h2>Qué puede pasar</h2>' + ver.frases_h(x) + '</div><div><h2>Qué hacer</h2>')
    if x['que_hacer'] or x['accidente']:
        h.append(ver.consejos(x['que_hacer']) if x['que_hacer'] else '<p class="sub">La ficha no trae consejos de prevención ni de almacenamiento.</p>')
        if x['accidente']:
            h.append('<p><b>Si hay un accidente:</b></p>' + ver.consejos(x['accidente']))
    elif x['peligros']:
        h.append(ver.caja('pend', 'La ficha no trae consejos de prudencia. La herramienta no los inventa: pida al proveedor una ficha completa y actualizada.'))
    else:
        h.append('<p class="sub">Sin indicaciones de peligro, la ficha no da consejos de prudencia.</p>')
    h.append('</div></div>')
    h.append('<h2>Protección personal</h2>' + ver.proteccion(x))
    h.append('<h2>Acción recomendada</h2>' + ver.accion(x))
    paginas = ' · '.join(f'sección {i["seccion"]}: ' + (f'p. {i["pagina"]}' if i['pagina'] else 'no encontrada') for i in x['referencias'])
    h.append('<h2>Fuente de los datos</h2><p class="chico">'
             f'Ficha de Datos de Seguridad «{_t(r["archivo"])}» ({paginas}). Clasificación según el SGA (Decreto 1496 de 2018, Resolución 773 de 2021) con las tablas del '
             f'proyecto; valores límite: TLV de la ACGIH. Semáforo: regla {s["regla"]}, {_t(s["fila"])}; consejos: regla C07, {len(x["que_hacer"]) + len(x["accidente"])} de los '
             f'{x["consejos_disponibles"]} que trae la ficha (los demás están en la ficha y en la etiqueta). El nivel de peligro, las bandas del semáforo y la selección de '
             f'consejos son propuestas en validación. Por verificar: {len(x["pendientes"])} {segun_cuantos(len(x["pendientes"]), "punto, detallado", "puntos, detallados")} en la aplicación. {_t(NO_REEMPLAZA)}</p>')
    h.append('<div class="firmas"><div>Consultó (nombre y cargo)</div><div>Revisó (responsable de Seguridad y Salud en el Trabajo)</div></div>')
    return h


def _anexo(r, x, tablas):
    """Anexo técnico opcional, en hojas aparte: de dónde sale cada dato."""
    n, p, g, e, ir = r['nivel'], r['palabra_de_advertencia'], r['pictogramas'], r['etiqueta'], r['indice_ir']
    h = ['<div class="anexo"><div class="qr-marca">Anexo técnico</div>']

    h.append('<h2>A. Clasificación de peligros, frase por frase</h2>')
    etiqueta_h = {f['codigo']: f for f in e['H']}
    h.append(_tabla(('Código', 'Indicación de peligro', 'Texto original de la ficha', 'Clase de peligro', 'Nivel', 'En la etiqueta', 'Página'),
                    [(q['codigo'], _t(q['texto']), f'<span class="original">{_t(etiqueta_h.get(q["codigo"], {}).get("texto_ficha"))}</span>', _t(q['clase']),
                      f"{q['nivel']} {_t(q['nivel_texto'])}",
                      'Sí' if etiqueta_h.get(q['codigo'], {}).get('incluida') else f"No: {_t(etiqueta_h.get(q['codigo'], {}).get('motivo'))}",
                      _t(q['pagina'])) for q in x['peligros']], 'chico')
             or '<p>La ficha no trae frases H: producto no peligroso para efectos de la etiqueta (regla R07).</p>')
    if g['pictogramas']:
        h.append('<p class="chico">Pictogramas: ' + '; '.join(f'{i["codigo"]} {_t(i["simbolo"])} (por {", ".join(i["por"])})' for i in g['pictogramas']) + '.</p>')
    for o in g['omitidos']:
        h.append(f'<p class="chico">No se muestra {o["codigo"]} por {o["frase"]}: {_t(o["motivo"])} (regla {o["regla"]}).</p>')
    for nota in n['notas'] + p['notas'] + e['notas']:
        h.append(f'<p class="chico">{_t(nota)}</p>')
    for nota in r['notas']:
        h.append(f'<div class="pendiente">{("Regla " + nota["regla"] + ". ") if nota["regla"] else ""}{_t(nota["texto"])}'
                 + ''.join(f'<br><span class="chico">· {_t(i)}</span>' for i in nota['indicios']) + '</div>')
    h.append('<h3>Cómo se obtuvo el nivel de peligro</h3>' + _preguntas(x['explicacion_del_nivel']))

    h.append('<h2>B. Consejos de prudencia de la ficha y selección para minería</h2>')
    elegidos = {z['codigo']: z for z in x['que_hacer'] + x['accidente']}
    filas = [(m['codigo'], _t(m['texto']) + (' <b>(completar «...»)</b>' if m['completar'] else ''), _t(grupo['tipo']), _t(', '.join(m['controla'])),
              (f"Sí ({elegidos[m['codigo']]['puntaje']:g}): {_t(elegidos[m['codigo']]['criterio'])}" if m['codigo'] in elegidos else 'No'),
              'Sí' if m['en_etiqueta'] else 'Opcional', _t(m['pagina'])) for grupo in x['medidas'] for m in grupo['consejos']]
    h.append(_tabla(('Código', 'Consejo de prudencia', 'Tipo', 'Peligro que controla (SGA)', 'Seleccionado para minería (puntaje y criterio)', 'En la etiqueta', 'Página'),
                    filas, 'chico') or '<p>La ficha no trae consejos de prudencia. La herramienta no los inventa.</p>')

    h.append('<h2>C. Valores límite, índice Ir y semáforo</h2>')
    if not ir['limites']:
        h.append('<p>El producto no tiene filas en la tabla de límites: no se puede calcular el índice Ir.</p>')
    else:
        def celda_ir(l):
            if l['ir'] is not None:
                return f'Ir = {num(l["ir"])}' + (' (supera el valor límite)' if l['supera'] else '')
            if (l['estado'] or '').startswith('alerta'):
                return f'<span class="{"supera" if l["supera"] else ""}">{"Supera" if l["supera"] else "No supera"} el máximo permitido (no es un Ir)</span>'
            return f'<span class="no">{_t(l["estado"])}</span>'
        h.append(_tabla(('N.º', 'Agente que se mide', 'Valor límite', '¿Se usa para Ir?', 'Exposición', 'Resultado', 'Nota de la tabla'),
                        [(l['id'], _t(l['agente']), f"{_t(l['tipo'])} {l['valor']:g} {_t(l['unidad'])}" if isinstance(l['valor'], (int, float)) else _t(l['tipo']),
                          _t(l['usar']), f"{l['exposicion']:g} {_t(l['unidad'])}" if l['exposicion'] is not None else '—', celda_ir(l), _t(l['mensaje'], ''))
                         for l in ir['limites']], 'chico'))
        for nota in ir['notas']:
            h.append(f'<p class="chico">{_t(nota)}</p>')
    s = r['semaforo']
    h.append(f'<p class="chico">Semáforo: {_t(s["color"])} ({_t(s["condicion"])}). Regla {s["regla"]}, {_t(s["fila"])}.'
             + (f' Estado del mensaje: {_t(s["estado_del_mensaje"])}.' if 'validar' in str(s['estado_del_mensaje']).lower() else '')
             + (f' Medición del {_t(x["fecha_de_medicion"])}.' if x.get('fecha_de_medicion') else '') + '</p>')
    h.append('<h3>Cómo se calcula y cómo se lee el índice Ir</h3>' + _preguntas(x['explicacion_del_ir']))

    h.append('<h2>D. Estabilidad y reactividad según la ficha (sección 10, texto original)</h2>')
    h.append(_tabla(('Apartado', 'Texto de la ficha, sin traducir'),
                    [(_t(a['nombre']), f'<span class="original">{_t(a["texto"])}</span>' if a['estado'] == 'encontrado'
                      else f'<span class="no">{_t(a["estado"])}' + (f': {_t(a["texto"])}' if a['texto'] else '') + '</span>')
                     for a in r['reacciones']['ficha'].values()], 'chico') or '<p>La ficha no trae la sección 10.</p>')
    if x['marco_normativo']:
        h.append(f'<h3>Marco normativo colombiano anotado en el Excel de fichas</h3><p>{_t(x["marco_normativo"])}</p>')

    if x['pendientes']:
        h.append('<h2>E. Pendiente de verificar</h2>' + ver.lista(_t(i) for i in x['pendientes']))
    sin_codigo = r['por_revisar']['frases_h'] + r['por_revisar']['consejos_p']
    if sin_codigo:
        h.append('<h3>Frases de la ficha sin código seguro</h3><p>No entraron al nivel ni a la selección. Se citan como están en la ficha.</p><ul>'
                 + ''.join(f'<li><span class="original">{_t(i)}</span></li>' for i in sin_codigo) + '</ul>')
    if r['hallazgos_de_la_ficha']:
        h.append('<h3>Observaciones de la lectura de la ficha</h3>')
        h.append(_tabla(('Paso', 'Observación'), [(_t(i['paso']), _t(i['texto'])) for i in r['hallazgos_de_la_ficha']], 'chico'))

    usadas = list(dict.fromkeys(n['reglas'] + p['reglas'] + g['reglas'] + e['reglas'] + [i['regla'] for i in r['notas'] if i['regla']]
                                + [s['regla'], r['seleccion_p']['regla']]))
    h.append('<h2>F. Reglas aplicadas y tablas usadas</h2>')
    h.append(_tabla(('Regla', 'Tema', 'Texto de la regla', 'Fuente', 'Carácter'),
                    [(i, _t(tablas['reglas'][i]['tema']), _t(tablas['reglas'][i]['regla']), _t(tablas['reglas'][i]['fuente']), _t(tablas['reglas'][i]['caracter']))
                     for i in usadas if i in tablas['reglas']], 'chico'))
    for aviso in r['avisos_de_reglas']:
        h.append(f'<div class="alerta">{_t(aviso)}</div>')
    h.append(_tabla(('Tabla', 'Archivo', 'Versión declarada', 'Fecha del archivo'), [tuple(_t(c) for c in f) for f in _versiones()], 'chico'))
    h.append('</div>')
    return h


def dibujar_reporte(r, tablas, lectura, detalle=False, fecha=None):
    """r = salida de motor.clasificar; lectura = salida de lectura.construir_lectura.
    Devuelve la página HTML del informe de consulta (una hoja). detalle=True le agrega el anexo técnico."""
    x = lectura
    fecha = fecha or time.strftime('%Y-%m-%d %H:%M')
    h = ['<p class="no-imprimir"><button class="boton" onclick="window.print()">Imprimir o guardar como PDF</button></p>',
         f'<div class="qr-marca">{escape(NOMBRE)} · Informe de consulta</div>', f'<div class="qr-titulo">{_t(x["nombre"])}</div>',
         f'<p class="datos">Consulta del {_t(fecha)} · Producto según la ficha: {_t(r["producto"])} · CAS: {_t(", ".join(r["cas"]))}'
         + (' · Ficha cargada por el usuario' if x['nueva'] else '') + '</p>']
    h += _informe(r, x)
    h.append(f'<div class="pie">{escape(NOMBRE)}, versión {VERSION}. El nivel, el semáforo, la palabra de advertencia, los pictogramas y la selección de consejos '
             'los calcula un programa fijo que lee las tablas del proyecto; ningún modelo de inteligencia artificial interviene en la clasificación.</div>')
    if detalle:
        h += _anexo(r, x, tablas)
    return (f'<!doctype html><html lang="es"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">'
            f'<title>Informe de consulta · {_t(x["nombre"])}</title><style>{ver.CSS}{ESTILO}</style></head><body>{"".join(h)}</body></html>')


def guardar_reporte(html, r, sufijo=''):
    destino = os.path.join(SALIDAS, 'reportes')
    os.makedirs(destino, exist_ok=True)
    ruta = os.path.join(destino, os.path.splitext(r['archivo'])[0] + sufijo + '.html')
    with open(ruta, 'w', encoding='utf-8') as f:
        f.write(html)
    return ruta


def leer_mediciones(argumentos):
    """--medicion L-21=0.3x8  o  L-02=150x2+60x6:mg/m3   ->  {ID: {'tareas': [(c, h)], 'unidad': ...}}"""
    mediciones = {}
    for i, a in enumerate(argumentos):
        if a == '--medicion':
            identificador, resto = argumentos[i + 1].split('=')
            valores, _, unidad = resto.partition(':')
            mediciones[identificador] = {'tareas': [tuple(float(v.replace(',', '.')) for v in tarea.lower().split('x')) for tarea in valores.split('+')],
                                         'unidad': unidad or None}
    return mediciones


if __name__ == '__main__':
    sys.stdout.reconfigure(encoding='utf-8')
    argumentos = sys.argv[1:]
    tablas = cargar_tablas()
    prefijo = argumentos[0] if argumentos and not argumentos[0].startswith('--') else ''
    con_anexo = '--anexo' in argumentos
    for ficha in cargar_fichas(prefijo):
        r = clasificar(ficha, tablas, None, leer_mediciones(argumentos))
        html = dibujar_reporte(r, tablas, construir_lectura(ficha, r, tablas), detalle=con_anexo)
        print('Informe guardado en', guardar_reporte(html, r, '_con_anexo' if con_anexo else ''))
