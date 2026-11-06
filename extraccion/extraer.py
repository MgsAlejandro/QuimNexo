# -*- coding: utf-8 -*-
"""Extracción completa de una FDS: corre los pasos 1 a 8 y guarda un archivo de resultado
por ficha en extraccion/salidas/ (formato JSON, que es lo que leerá el motor de clasificación).
Aquí no se clasifica: el resultado solo contiene lo que la ficha declara, con página y cita,
y los hallazgos, que se conservan para que el usuario los tenga a la vista.
Uso:
    py extraccion/extraer.py FDS                          (todas las fichas de la carpeta)
    py extraccion/extraer.py FDS --llm                    (con segunda pasada de modelo de lenguaje; ver segunda_pasada.py)
    py extraccion/extraer.py FDS --llm openai             (eligiendo el proveedor: anthropic u openai)
    py extraccion/extraer.py FDS/15_Acetileno_74-86-2.pdf
Requiere: pip install pdfplumber openpyxl"""
import glob, json, os, sys
from leer_pdf import leer_pdf
from secciones import cortar_secciones, cargar_titulos_oficiales, OK
from identificacion import extraer_identificacion, cargar_excel_fichas, cotejar_cas
from frases import cargar_diccionario
from frases_texto import preparar, extraer_frases
from clasificacion import extraer_clasificacion, enlazar, cargar_clases
from limites import extraer_limites
from reacciones import extraer_reacciones, NO_INFORMA
from proteccion import extraer_proteccion
import segunda_pasada

SALIDAS = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'salidas')
REV = 'por revisar'


class Recursos:
    """Tablas que se leen una sola vez: títulos de sección, diccionario del SGA y Excel de fichas."""
    def __init__(self):
        self.oficiales = cargar_titulos_oficiales()
        self.diccionario = cargar_diccionario()
        self.indice = preparar(self.diccionario)
        self.clases_sga = cargar_clases()
        self.excel = cargar_excel_fichas()


def _hallazgos(r):
    """Lista única de hallazgos y pendientes de la ficha, cada uno con el paso de donde sale."""
    h = []
    for n, s in r['secciones'].items():
        if s['estado'] != OK or s['marcas']:
            h.append(('Secciones', f"Sección {n}: {s['estado']}" + ''.join(f'; {m}' for m in s['marcas'])))
    i = r['identificacion']
    if i['nombre']['estado'] != 'encontrado':
        h.append(('Identificación', f"Nombre del producto {i['nombre']['estado']}: «{i['nombre']['valor']}»"))
    for dato, x in (('Nombre del producto', i['nombre']), ('Proveedor', i['proveedor'])):
        if x.get('metodo') == 'LLM':
            h.append(('Identificación', f'{dato} obtenido con modelo de lenguaje: verificar'))
    if i['proveedor']['estado'] != 'encontrado':
        h.append(('Identificación', f"Proveedor {i['proveedor']['estado']}" + ''.join(f'; {n}' for n in i['proveedor']['notas'])))
    h += [('Identificación', x) for x in r['cotejo_cas']['hallazgos']]
    f = r['frases']
    for tipo in 'HP':
        for x in f[tipo]:
            if x['estado'] == REV:
                candidatos = ', '.join(f'{c} ({p:.2f})' for c, p in x.get('candidatos', []))
                h.append(('Frases', f"Frase {tipo} por revisar: «{x['texto_ficha']}»" + (f' [{x["codigo"]}: {x["nota"]}]' if x.get('nota') else '')
                          + (f'. Candidatos: {candidatos}' if candidatos else '')))
            elif x.get('metodo') == 'LLM':
                h.append(('Frases', f"Código {x['codigo']} asignado con modelo de lenguaje a «{x['texto_ficha']}»: verificar"))
    h += [('Frases', f"La ficha declara: «{d['texto_ficha']}»") for d in f['declaraciones']]
    h += [('Frases', f"Frase sin código del SGA: «{s['texto_ficha']}»") for s in f['sin_codigo']]
    h += [('Frases', f"Frase de otro sistema: {o['codigo']} «{o['texto_ficha']}»") for o in f['otras_frases']]
    h += [('Frases', f"Códigos en la sección {n} que no están en la sección 2 (no se tomaron): {', '.join(c)}")
          for n, c in f['fuera_de_seccion_2'].items()]
    if not f['H'] and not f['declaraciones']:
        h.append(('Frases', 'La sección 2 no trae frases H ni declara que el producto no sea peligroso.'))
    c = r['clasificacion']
    h += [('Clasificación', f"Clase declarada sin frase H en la ficha: «{e['clase']}», categoría {e['categoria']}")
          for e in c['clases'] if not e['codigos_compatibles']]
    h += [('Clasificación', f'Frase {cod} sin clase declarada compatible: no se pudo leer su categoría')
          for cod, cats in c['categoria_por_frase_H'].items() if not cats]
    h += [('Límites', f"La sección 8 dice: «{x['texto_ficha']}»") for x in r['limites']['hallazgos']]
    if r['limites']['sin_limites']:
        h.append(('Límites', 'La sección 8 no declara valores límite.'))
    inc = r['reacciones']['apartados'].get('materiales_incompatibles')
    if inc is None:
        h.append(('Reacciones', 'La sección 10 no trae el apartado de materiales incompatibles.'))
    elif inc['estado'] == NO_INFORMA:
        h.append(('Reacciones', f"Materiales incompatibles: la ficha no informa («{inc['texto']}»)."))
    h += [('Segunda pasada', a) for a in r.get('segunda_pasada', {}).get('avisos', [])]
    return [dict(paso=p, texto=t) for p, t in h]


def extraer_ficha(ruta, recursos, consulta=None, origen=None):
    """Resultado completo de una ficha (diccionario listo para guardar).
    'consulta' activa la segunda pasada con modelo de lenguaje (paso 10); sin ella no se usa ningún modelo."""
    doc = leer_pdf(ruta)
    corte = cortar_secciones(doc, recursos.oficiales)
    identificacion = extraer_identificacion(doc, corte)
    frases = extraer_frases(corte, recursos.diccionario, recursos.indice)
    clasificacion = extraer_clasificacion(corte)
    clasificacion['categoria_por_frase_H'] = enlazar(clasificacion, frases, recursos.clases_sga)
    r = {
        'archivo': os.path.basename(ruta),
        'lector': doc['lector'],
        'paginas': len(doc['paginas']),
        'secciones': {n: dict(estado=s['estado'], titulo=s['titulo'], pagina=s['pagina'], lineas=len(s['lineas']), marcas=s['marcas'])
                      for n, s in corte['secciones'].items()},
        'identificacion': identificacion,
        'cotejo_cas': cotejar_cas(ruta, identificacion, recursos.excel),
        'frases': frases,
        'clasificacion': clasificacion,
        'limites': extraer_limites(doc, corte),
        'reacciones': extraer_reacciones(corte),
    }
    if consulta:
        segunda_pasada.aplicar(r, corte, recursos, consulta, origen)
        r['clasificacion']['categoria_por_frase_H'] = enlazar(clasificacion, frases, recursos.clases_sga)
    r['hallazgos'] = _hallazgos(r)
    return r


def guardar(resultado):
    os.makedirs(SALIDAS, exist_ok=True)
    destino = os.path.join(SALIDAS, os.path.splitext(resultado['archivo'])[0] + '.json')
    with open(destino, 'w', encoding='utf-8') as f:
        json.dump(resultado, f, ensure_ascii=False, indent=1)
    return destino


if __name__ == '__main__':
    sys.stdout.reconfigure(encoding='utf-8')
    recursos = Recursos()
    consulta = origen = None
    if '--llm' in sys.argv:
        i = sys.argv.index('--llm')
        elegido = sys.argv[i + 1] if len(sys.argv) > i + 1 and sys.argv[i + 1] in segunda_pasada.PROVEEDORES else None
        del sys.argv[i:i + (2 if elegido else 1)]
        consulta, origen, motivo = segunda_pasada.crear_consulta(elegido)
        if consulta:
            print(f"Segunda pasada activada: {origen['proveedor']}, modelo {origen['modelo']}.")
        else:
            print(f'Segunda pasada NO activada: {motivo}. Se sigue sin modelo de lenguaje.')
    destino = sys.argv[1]
    rutas = sorted(glob.glob(os.path.join(destino, '*.pdf'))) if os.path.isdir(destino) else [destino]
    print(f"{'ficha':<26}{'H':>3}{'P':>4}{'clases':>7}{'límites':>8}{'hallazgos':>10}")
    for ruta in rutas:
        r = extraer_ficha(ruta, recursos, consulta, origen)
        guardar(r)
        print(f"{r['archivo'][:25]:<26}{sum(f['estado'] == 'encontrado' for f in r['frases']['H']):>3}"
              f"{sum(f['estado'] == 'encontrado' for f in r['frases']['P']):>4}{len(r['clasificacion']['clases']):>7}"
              f"{len(r['limites']['limites']):>8}{len(r['hallazgos']):>10}")
    print(f'{len(rutas)} resultado(s) guardado(s) en {SALIDAS}')
    if len(rutas) == 1:
        print('Hallazgos y pendientes de esta ficha:')
        print('\n'.join(f"  [{h['paso']}] {h['texto']}" for h in r['hallazgos']) or '  ninguno')
