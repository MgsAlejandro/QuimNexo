# -*- coding: utf-8 -*-
"""Paso I4 de la interfaz: procesar una FDS nueva que el usuario carga en PDF.
El PDF se guarda en FDS_nuevas/ (aparte de las 32 fichas originales de FDS/), se le corre la misma
extracción (pasos 1 a 8) y el resultado queda en FDS_nuevas/resultados/. Después el motor la clasifica
como a cualquier otra. Aquí no se decide nada.
Uso de prueba:  py "clasificador SGA/fichas_nuevas.py" ruta/al/archivo.pdf
Requiere: pip install pdfplumber openpyxl"""
import glob, json, os, re, shutil, sys, time

AQUI = os.path.dirname(os.path.abspath(__file__))
RAIZ = os.path.join(AQUI, '..')
sys.path.insert(0, os.path.join(RAIZ, 'extraccion'))
from extraer import Recursos, extraer_ficha   # noqa: E402  (la extracción vive en la carpeta extraccion/)

CARPETA = os.path.join(RAIZ, 'FDS_nuevas')
RESULTADOS = os.path.join(CARPETA, 'resultados')


class ErrorDeFicha(Exception):
    """El archivo cargado no se puede procesar como FDS."""


def nombre_seguro(nombre):
    """Nombre de archivo sin rutas ni caracteres raros, terminado en .pdf."""
    base = os.path.basename(str(nombre).replace('\\', '/'))
    base = re.sub(r'[^\w\-. ()]', '_', os.path.splitext(base)[0]).strip(' .') or 'ficha'
    return base[:120] + '.pdf'


def procesar_fds_nueva(nombre, datos, recursos, consulta=None, origen=None):
    """Guarda el PDF, lo extrae y guarda su resultado. Devuelve el resultado de la extracción.
    'datos' son los bytes del archivo. Si ya había una ficha nueva con ese nombre, se reemplaza."""
    if not datos or not datos.lstrip()[:5].startswith(b'%PDF'):
        raise ErrorDeFicha('El archivo no es un PDF.')
    os.makedirs(RESULTADOS, exist_ok=True)
    archivo = nombre_seguro(nombre)
    ruta = os.path.join(CARPETA, archivo)
    with open(ruta, 'wb') as f:
        f.write(datos)
    try:
        resultado = extraer_ficha(ruta, recursos, consulta, origen)
    except Exception as error:   # un PDF dañado o protegido puede fallar de muchas formas; se informa y no se deja a medias
        os.remove(ruta)
        raise ErrorDeFicha(f'No se pudo leer el PDF: {error}') from error
    if not any(s['lineas'] for s in resultado['secciones'].values()):
        os.remove(ruta)
        raise ErrorDeFicha('El PDF no tiene texto extraíble (parece una imagen escaneada) o no se reconoció ninguna sección de FDS.')
    resultado['origen'] = 'nueva'
    with open(os.path.join(RESULTADOS, os.path.splitext(archivo)[0] + '.json'), 'w', encoding='utf-8') as f:
        json.dump(resultado, f, ensure_ascii=False, indent=1)
    return resultado


def retirar_ficha_nueva(archivo):
    """Quita de la lista una ficha cargada por el usuario. No borra nada: mueve su PDF y su resultado a
    FDS_nuevas/retiradas/, con la fecha y la hora, para poder recuperarlos. Devuelve la carpeta de destino."""
    base = os.path.splitext(nombre_seguro(archivo))[0]
    origenes = [os.path.join(CARPETA, base + '.pdf'), os.path.join(RESULTADOS, base + '.json'),
                os.path.join(AQUI, 'elecciones', base + '.json'), os.path.join(AQUI, 'confirmaciones', base + '.json')]
    existentes = [r for r in origenes if os.path.isfile(r)]
    if not os.path.isfile(origenes[1]):
        raise ErrorDeFicha(f'«{archivo}» no es una ficha cargada por el usuario; las 32 originales no se pueden quitar desde aquí.')
    destino = os.path.join(CARPETA, 'retiradas', f"{time.strftime('%Y-%m-%d_%H%M%S')}_{base}")
    os.makedirs(destino, exist_ok=True)
    for ruta in existentes:
        sufijo = {'elecciones': '_elecciones', 'confirmaciones': '_confirmaciones'}.get(os.path.basename(os.path.dirname(ruta)), '')
        nombre, extension = os.path.splitext(os.path.basename(ruta))
        shutil.move(ruta, os.path.join(destino, nombre + sufijo + extension))
    return destino


def rutas_de_resultados():
    return sorted(glob.glob(os.path.join(RESULTADOS, '*.json')))


def cargar_fichas_nuevas():
    fichas = []
    for ruta in rutas_de_resultados():
        with open(ruta, encoding='utf-8') as f:
            fichas.append(json.load(f))
    return fichas


def resumen_de_extraccion(resultado):
    """Texto corto de lo que se extrajo, para mostrar al usuario después de cargar."""
    marcadas = [f"sección {n}: {s['estado']}" + ''.join(f' ({m.split(":")[0]})' for m in s['marcas'])
                for n, s in resultado['secciones'].items() if s['estado'] != 'OK' or s['marcas']]
    h = [f for f in resultado['frases']['H']]
    p = [f for f in resultado['frases']['P']]
    return {
        'producto': resultado['identificacion']['nombre']['valor'],
        'cas': resultado['cotejo_cas']['ficha'],
        'secciones_con_marca': marcadas,
        'frases_h': [f['codigo'] for f in h if f['estado'] == 'encontrado'],
        'consejos_p': sum(f['estado'] == 'encontrado' for f in p),
        'por_revisar': sum(f['estado'] != 'encontrado' for f in h + p),
        'metodo_h': resultado['frases']['metodo']['H'],
        'limites': len(resultado['limites']['limites']),
        'hallazgos': len(resultado['hallazgos']),
    }


if __name__ == '__main__':
    sys.stdout.reconfigure(encoding='utf-8')
    origen = sys.argv[1]
    with open(origen, 'rb') as f:
        r = procesar_fds_nueva(sys.argv[2] if len(sys.argv) > 2 else os.path.basename(origen), f.read(), Recursos())
    print('Guardada como', r['archivo'])
    for k, v in resumen_de_extraccion(r).items():
        print(f'   {k}: {v}')
