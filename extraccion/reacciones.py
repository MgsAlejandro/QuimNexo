# -*- coding: utf-8 -*-
"""Paso 8 de la extracción: estabilidad, reacciones e incompatibilidades (sección 10 de la FDS).
Separa la sección en sus apartados y guarda el texto literal de cada uno. No interpreta ni
resume: qué reacciona con qué lo dice la ficha, y aquí solo se transcribe y se ordena.
Cuando la ficha dice que no tiene información, el apartado se marca 'la ficha no informa'.
Uso de prueba:
    py extraccion/reacciones.py FDS                        (resumen de todas)
    py extraccion/reacciones.py FDS/15_Acetileno_74-86-2.pdf    (detalle)
Requiere: pip install pdfplumber openpyxl"""
import glob, os, re, sys
from leer_pdf import leer_pdf
from secciones import cortar_secciones, cargar_titulos_oficiales

ENC, NO_INFORMA, SIN_TEXTO = 'encontrado', 'la ficha no informa', 'rótulo sin texto propio'
# Apartados de la sección 10 y rótulos con que las fichas los nombran (los más largos primero).
# Son palabras para ubicar el texto; el contenido se copia literal.
APARTADOS = [
    ('productos_descomposicion', 'Productos de descomposición peligrosos',
     [r'Hazardous decomposition(?:\s+products)?', r'Productos? de descomposici[óo]n(?:\s+peligros[oa]s)?']),
    ('polimerizacion', 'Polimerización peligrosa',
     [r'Hazardous polymerization', r'Polimerizaci[óo]n peligrosa']),
    ('reacciones_peligrosas', 'Posibilidad de reacciones peligrosas',
     [r'Possibility of hazardous reac(?:tions?)?-?', r'Hazardous reactions', r'Posibilidad de reacci[óo]n(?:es)? peligrosas?']),
    ('condiciones_a_evitar', 'Condiciones que deben evitarse',
     [r'Conditions to (?:be )?av[io]{2}de?d?', r'Condiciones (?:a evitar|que deben evitarse)']),
    ('materiales_incompatibles', 'Materiales incompatibles',
     [r'Incompatible materials', r'Materials to avoid', r'Substances to be avoided',
      r'Materiales incompatibles', r'Incompatibilidad(?:es)?']),
    ('estabilidad', 'Estabilidad química',
     [r'Chemical stability', r'Stability', r'Estabilidad qu[íi]mica', r'Estabilidad']),
    ('reactividad', 'Reactividad',
     [r'Reactive hazard', r'Reactivity', r'Reactividad']),
]
ORDEN = ['reactividad', 'estabilidad', 'reacciones_peligrosas', 'polimerizacion', 'condiciones_a_evitar',
         'materiales_incompatibles', 'productos_descomposicion']
NOMBRE = {clave: nombre for clave, nombre, _ in APARTADOS}
# Numeración ("10.1.", "a)") y rótulo; admite que el PDF separe la primera letra ("R eactivity")
PREFIJO = r'^\s*(?:\d+(?:\.\d+)*\.?\s+|[a-z]\)\s*)?'
ROTULOS = [(clave, re.compile(PREFIJO + '(?:' + '|'.join(p[0] + r'\s?' + p[1:] for p in patrones) + r')\s*[:.]?\s*(.*)$', re.I))
           for clave, _, patrones in APARTADOS]
SIN_DATO = re.compile(r'^(none(\s+known)?|no (data|information|additional information)(\s+available)?|not (available|determined|applicable)|'
                      r'no (determinado|disponible|aplica|hay datos)|ningun[oa]( conocid[oa])?|sin datos|'
                      r'no further (information|informations|requirements))\b', re.I)


def extraer_reacciones(corte):
    """corte = salida de cortar_secciones. Devuelve
    {'apartados': {clave: {nombre, texto, estado, pagina, cita}}, 'sin_rotulo': [(página, línea)]}."""
    apartados, sin_rotulo, actual = {}, [], None
    for pagina, linea in corte['secciones'][10]['lineas']:
        hallado = next(((clave, patron.match(linea)) for clave, patron in ROTULOS if patron.match(linea)), None)
        if hallado:
            actual, m = hallado[0], hallado[1]
            a = apartados.setdefault(actual, dict(nombre=NOMBRE[actual], partes=[], pagina=pagina, cita=[]))
            a['cita'].append(linea)
            if m.group(1).strip():
                a['partes'].append(m.group(1).strip())
        elif actual:
            apartados[actual]['partes'].append(linea.strip())
            apartados[actual]['cita'].append(linea)
        else:
            sin_rotulo.append((pagina, linea))
    for a in apartados.values():
        a['texto'] = ' '.join(a.pop('partes')) or None
        a['cita'] = ' / '.join(a['cita'])
        # Un rótulo sin texto propio (p. ej. 'Reactividad' seguido de subapartados) no es lo mismo que 'no informa'
        a['estado'] = SIN_TEXTO if not a['texto'] else NO_INFORMA if SIN_DATO.match(a['texto']) else ENC
    return {'apartados': {c: apartados[c] for c in ORDEN if c in apartados}, 'sin_rotulo': sin_rotulo}


if __name__ == '__main__':
    sys.stdout.reconfigure(encoding='utf-8')
    oficiales = cargar_titulos_oficiales()
    destino = sys.argv[1]
    rutas = sorted(glob.glob(os.path.join(destino, '*.pdf'))) if os.path.isdir(destino) else [destino]
    if len(rutas) > 1:
        sigla = dict(zip(ORDEN, ['Rea', 'Est', 'RxP', 'Pol', 'Con', 'Inc', 'Des']))
        print('Apartados: Rea reactividad · Est estabilidad · RxP reacciones peligrosas · Pol polimerización · '
              'Con condiciones a evitar · Inc materiales incompatibles · Des productos de descomposición')
        print('Marcas: sí = trae texto · no = la ficha dice que no informa · rót = rótulo sin texto propio · — = la ficha no trae ese apartado')
        print(f"{'ficha':<24}" + ''.join(f'{sigla[c]:>5}' for c in ORDEN) + '  sin rótulo | materiales incompatibles')
    for ruta in rutas:
        r = extraer_reacciones(cortar_secciones(leer_pdf(ruta), oficiales))
        ap = r['apartados']
        if len(rutas) > 1:
            marcas = ''.join(f"{'—' if c not in ap else 'sí' if ap[c]['estado'] == ENC else 'no' if ap[c]['estado'] == NO_INFORMA else 'rót':>5}" for c in ORDEN)
            inc = ap.get('materiales_incompatibles', {}).get('texto') or ''
            print(f"{os.path.basename(ruta)[:23]:<24}{marcas}{len(r['sin_rotulo']):>12} | {inc[:95]}")
            continue
        for clave, a in ap.items():
            print(f"[{a['estado']}] {a['nombre']} (p. {a['pagina']})\n    {a['texto'] or '—'}")
        if r['sin_rotulo']:
            print('Líneas de la sección 10 sin rótulo:')
            print('\n'.join(f'    p.{p} {l}' for p, l in r['sin_rotulo']))
