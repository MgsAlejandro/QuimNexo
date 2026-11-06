# -*- coding: utf-8 -*-
"""Paso I4b de la interfaz: confirmación, por el usuario, del código de una frase "por revisar".
La extracción deja "por revisar" las frases cuyo código no quedó claro, y por eso no entran al nivel ni a la
etiqueta. Aquí el usuario, como experto, puede confirmar el código. Solo puede elegir entre lo que la propia
extracción propuso: los candidatos por parecido, o los códigos por separado de una combinación que el SGA no lista.
La frase confirmada entra al cálculo marcada "confirmada por el usuario", con la fecha. Se puede deshacer.
Las confirmaciones se guardan por ficha en "clasificador SGA/confirmaciones/". El resultado de la extracción no se modifica."""
import copy, json, os, time

CARPETA = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'confirmaciones')
CONFIRMADA = 'confirmada por el usuario'


def clave_de_frase(tipo, frase):
    return f"{tipo}|{frase['texto_ficha']}"


def _existe(tipo, codigo, tablas):
    return codigo[:4] in tablas['frases_h'] if tipo == 'H' else (codigo in tablas['frases_p'] or codigo in tablas['sga_p'])


def _texto(tipo, codigo, tablas):
    if tipo == 'H':
        return tablas['frases_h'][codigo[:4]]['texto']
    return tablas['frases_p'][codigo]['texto'] if codigo in tablas['frases_p'] else tablas['sga_p'][codigo]


def opciones(tipo, frase, tablas):
    """Lo que el usuario puede confirmar para una frase por revisar: [{'codigos': [...], 'rotulo': texto}].
    Nunca se ofrece un código que no exista en las tablas."""
    salida = []
    for codigo, parecido in frase.get('candidatos') or []:
        if _existe(tipo, codigo, tablas):
            salida.append({'codigos': [codigo], 'rotulo': f'{codigo} — {_texto(tipo, codigo, tablas)} (parecido {parecido:.2f})'})
    partes = (frase.get('codigo') or '').split('+')
    if not salida and len(partes) > 1 and all(_existe(tipo, p, tablas) for p in partes):
        grupos = _agrupar(tipo, partes, tablas)   # combinaciones oficiales más largas, de izquierda a derecha
        salida.append({'codigos': grupos, 'rotulo': 'Aceptar como: ' + ', '.join(grupos)
                       + ' (el SGA no lista la combinación completa; se usan sus partes oficiales)'})
    return salida


def _agrupar(tipo, partes, tablas):
    """Parte una combinación que el SGA no lista en las combinaciones oficiales más largas que sí están en las tablas,
    de izquierda a derecha; lo que no tiene combinación oficial queda suelto. Ejemplo: P304+P340+P312 -> P304+P340, P312.
    Decisión del autor (17 de octubre de 2026): así no queda suelto un encabezado como P304 «EN CASO DE INHALACIÓN:»."""
    grupos, i = [], 0
    while i < len(partes):
        j = next(j for j in range(len(partes), i, -1) if j - i == 1 or _combinacion_oficial(tipo, '+'.join(partes[i:j]), tablas))
        grupos.append('+'.join(partes[i:j]))
        i = j
    return grupos


def _combinacion_oficial(tipo, codigo, tablas):
    """La combinación está, tal cual, en las tablas (no basta con que exista su primer código)."""
    return codigo in tablas['frases_h'] if tipo == 'H' else (codigo in tablas['frases_p'] or codigo in tablas['sga_p'])


def leer_confirmaciones(archivo):
    ruta = os.path.join(CARPETA, os.path.splitext(archivo)[0] + '.json')
    if not os.path.isfile(ruta):
        return {}
    with open(ruta, encoding='utf-8') as f:
        return json.load(f)


def guardar_confirmaciones(archivo, confirmaciones):
    os.makedirs(CARPETA, exist_ok=True)
    ruta = os.path.join(CARPETA, os.path.splitext(archivo)[0] + '.json')
    if not confirmaciones:
        if os.path.isfile(ruta):
            os.remove(ruta)   # no queda nada confirmado: el archivo sobra
        return
    with open(ruta, 'w', encoding='utf-8') as f:
        json.dump(confirmaciones, f, ensure_ascii=False, indent=1)


def confirmar(archivo, tipo, frase, codigos):
    confirmaciones = leer_confirmaciones(archivo)
    confirmaciones[clave_de_frase(tipo, frase)] = {'codigos': list(codigos), 'fecha': time.strftime('%Y-%m-%d %H:%M')}
    guardar_confirmaciones(archivo, confirmaciones)


def deshacer(archivo, tipo, frase):
    confirmaciones = leer_confirmaciones(archivo)
    confirmaciones.pop(clave_de_frase(tipo, frase), None)
    guardar_confirmaciones(archivo, confirmaciones)


def aplicar(ficha, confirmaciones, tablas):
    """Copia de la ficha con las frases confirmadas ya dentro del cálculo. La ficha original no cambia.
    Una confirmación solo vale si coincide con una de las opciones que la extracción ofrece para esa frase."""
    if not confirmaciones:
        return ficha
    nueva = copy.deepcopy(ficha)
    for tipo in 'HP':
        frases = []
        for frase in nueva['frases'][tipo]:
            dato = confirmaciones.get(clave_de_frase(tipo, frase)) if frase['estado'] != 'encontrado' else None
            validas = [o['codigos'] for o in opciones(tipo, frase, tablas)]
            if not dato or dato.get('codigos') not in validas:
                frases.append(frase)
                continue
            for codigo in dato['codigos']:
                frases.append(dict(frase, codigo=codigo, codigo_base=codigo, estado='encontrado', metodo=CONFIRMADA,
                                   confirmada=dato.get('fecha'), nota=None))
        nueva['frases'][tipo] = frases
    return nueva
