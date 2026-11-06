# -*- coding: utf-8 -*-
"""Lo que el usuario guarda en este equipo. Nada sale de él y aquí no hay claves.
  elecciones/   lo que eligió para la etiqueta de cada producto
  mediciones/   las mediciones de exposición de cada producto, cada una con su fecha y hora (se usa la más reciente)
  consultas/    historial: una copia de cada informe de consulta que decidió guardar, y registro.csv con una fila por consulta
  contexto/     datos de la mina que versiones anteriores pedían (actividad, tarea…). Ya no se piden; lo guardado se conserva
Todo son archivos sueltos (JSON, HTML y CSV) que se pueden abrir, copiar o borrar a mano."""
import csv, json, os, re, time

AQUI = os.path.dirname(os.path.abspath(__file__))
CONSULTAS = os.path.join(AQUI, 'consultas')
REGISTRO = os.path.join(CONSULTAS, 'registro.csv')
COLUMNAS = ['Fecha', 'Producto', 'CAS', 'Ficha', 'Nivel de peligro', 'Riesgo por exposición', 'Índice Ir', 'Por verificar', 'Informe']


def _ruta(carpeta, archivo):
    return os.path.join(AQUI, carpeta, os.path.splitext(archivo)[0] + '.json')


def leer(carpeta, archivo, vacio=None):
    """Datos guardados de una ficha en una de las carpetas; {} (o 'vacio') si no hay."""
    ruta = _ruta(carpeta, archivo)
    if os.path.isfile(ruta):
        try:
            with open(ruta, encoding='utf-8') as f:
                return json.load(f)
        except (OSError, ValueError):   # archivo dañado o a medio escribir: se ignora en vez de tumbar la aplicación
            pass
    return {} if vacio is None else vacio


def guardar(carpeta, archivo, datos):
    os.makedirs(os.path.join(AQUI, carpeta), exist_ok=True)
    with open(_ruta(carpeta, archivo), 'w', encoding='utf-8') as f:
        json.dump(datos, f, ensure_ascii=False, indent=1)


# ---------- Mediciones de exposición, con historial ----------
def mediciones_de(archivo):
    """Historial de mediciones de un producto, de la más antigua a la más reciente.
    Cada registro: {'fecha': 'AAAA-MM-DD HH:MM', 'mediciones': {ID del límite: {'tareas': [[concentración, horas], ...], 'unidad': ...}}}."""
    historial = leer('mediciones', archivo, vacio=[])
    return historial if isinstance(historial, list) else []


def ultima_medicion(archivo):
    """La medición más reciente de un producto, o None."""
    historial = mediciones_de(archivo)
    return historial[-1] if historial else None


def guardar_medicion(archivo, mediciones):
    """Agrega una medición al historial del producto, con la fecha y la hora de este equipo. No guarda una igual a la última."""
    limpia = {k: {'tareas': [list(tarea) for tarea in v['tareas']], 'unidad': v.get('unidad')} for k, v in mediciones.items()}
    historial = mediciones_de(archivo)
    if not limpia or (historial and historial[-1]['mediciones'] == limpia):
        return historial[-1] if historial else None
    historial.append({'fecha': time.strftime('%Y-%m-%d %H:%M'), 'mediciones': limpia})
    guardar('mediciones', archivo, historial)
    return historial[-1]


# ---------- Historial de consultas ----------
def guardar_consulta(r, lectura, html):
    """Guarda el informe de consulta en el historial y anota una fila en registro.csv. Devuelve la ruta del informe."""
    os.makedirs(CONSULTAS, exist_ok=True)
    ahora = time.localtime()
    base = re.sub(r'[^\w\-]', '_', os.path.splitext(r['archivo'])[0])[:60]
    nombre = f"{time.strftime('%Y-%m-%d_%H%M%S', ahora)}_{base}.html"
    with open(os.path.join(CONSULTAS, nombre), 'w', encoding='utf-8') as f:
        f.write(html)
    ir = r['indice_ir']['ir_mayor']
    fila = [time.strftime('%Y-%m-%d %H:%M', ahora), lectura['nombre'], ', '.join(r['cas']), r['archivo'],
            f"{r['nivel']['nivel']} {r['nivel']['nombre']}", r['semaforo']['color'] if ir is not None or r['semaforo']['clave'] in ('rojo_alerta', 'alerta_ok') else 'Sin medición',
            '' if ir is None else f'{ir:.3g}'.replace('.', ','), len(lectura['pendientes']), nombre]
    nuevo = not os.path.isfile(REGISTRO)
    with open(REGISTRO, 'a', encoding='utf-8-sig', newline='') as f:   # con marca para que Excel lea bien las tildes
        escritor = csv.writer(f, delimiter=';')
        if nuevo:
            escritor.writerow(COLUMNAS)
        escritor.writerow(fila)
    return os.path.join(CONSULTAS, nombre)


def consultas():
    """Filas del historial, de la más reciente a la más antigua. Solo las que conservan su informe."""
    if not os.path.isfile(REGISTRO):
        return []
    try:
        with open(REGISTRO, encoding='utf-8-sig', newline='') as f:
            filas = list(csv.DictReader(f, delimiter=';'))
    except (OSError, csv.Error):
        return []
    return [f for f in reversed(filas) if f.get('Informe') and os.path.isfile(os.path.join(CONSULTAS, f['Informe']))]


def leer_informe(nombre):
    """Contenido de un informe del historial. El nombre se toma tal cual del registro, sin rutas."""
    with open(os.path.join(CONSULTAS, os.path.basename(nombre)), encoding='utf-8') as f:
        return f.read()
