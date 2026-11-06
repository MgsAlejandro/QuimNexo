# -*- coding: utf-8 -*-
"""Prueba de los controles de la segunda pasada SIN llamar a la API.
Se le dan al módulo respuestas simuladas, unas correctas y otras tramposas, y se comprueba
que acepte solo las que pasan los controles: el código debe existir en el diccionario del SGA
y los datos de identificación deben aparecer literales en la sección 1 de la ficha.
Uso:  py extraccion/probar_segunda_pasada.py"""
import glob, os, sys
from extraer import Recursos, extraer_ficha

FDS = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'FDS')


def simulada(frases=None, identificacion=None):
    """Devuelve una función 'consulta' que responde lo que se le indique en vez de llamar a la API."""
    def consulta(instrucciones, mensaje, esquema):
        if 'asignaciones' in esquema['properties']:
            return ({'asignaciones': frases}, None) if frases is not None else (None, 'sin respuesta simulada')
        return (identificacion, None) if identificacion is not None else (None, 'sin respuesta simulada')
    return consulta


def ficha(prefijo):
    return glob.glob(os.path.join(FDS, prefijo + '*.pdf'))[0]


def mostrar(titulo, r):
    print(f'\n=== {titulo} ===')
    for f in r['frases']['H']:
        if f.get('metodo') == 'LLM' or f['estado'] != 'encontrado':
            print(f"  Frase H: {f['codigo'] or '?'} [{f['estado']}, método {f.get('metodo', 'texto')}] «{f['texto_ficha']}»")
    n, p = r['identificacion']['nombre'], r['identificacion']['proveedor']
    print(f"  Nombre: {n['valor']} [{n['estado']}, método {n.get('metodo', 'rótulo')}]")
    print(f"  Proveedor: {p['nombre']} | {p['direccion']} | {p['telefono']} [{p['estado']}, método {p.get('metodo', 'rótulo')}]")
    for a in r['segunda_pasada']['avisos']:
        print(f'  AVISO: {a}')


if __name__ == '__main__':
    sys.stdout.reconfigure(encoding='utf-8')
    recursos = Recursos()

    r = extraer_ficha(ficha('06_'), recursos, simulada(frases=[{'numero': 1, 'codigo': 'H361', 'motivo': 'misma frase, sin "fertility"'}]))
    mostrar('06 Estireno · respuesta correcta (H361): debe aceptarse', r)
    assert any(f['codigo'] == 'H361' and f['metodo'] == 'LLM' for f in r['frases']['H'])

    r = extraer_ficha(ficha('06_'), recursos, simulada(frases=[{'numero': 1, 'codigo': 'H999', 'motivo': 'código inventado'}]))
    mostrar('06 Estireno · código que no existe (H999): debe descartarse', r)
    assert not any(f['codigo'] == 'H999' for f in r['frases']['H'])

    r = extraer_ficha(ficha('06_'), recursos, simulada(frases=[{'numero': 1, 'codigo': 'NINGUNO', 'motivo': 'duda'}]))
    mostrar('06 Estireno · el modelo responde NINGUNO: la frase sigue por revisar', r)
    assert any(f['estado'] == 'por revisar' for f in r['frases']['H'])

    buena = {'nombre_producto': 'CALCIUM NITRATE SOLUTION APPROX.45%', 'proveedor_nombre': 'POCH SA',
             'proveedor_direccion': '44-101 Gliwice, Sowinskiego Str. 11', 'proveedor_telefono': '+48 32 23-92-381'}
    r = extraer_ficha(ficha('23_'), recursos, simulada(identificacion=buena))
    mostrar('23 Nitrato de calcio · datos que sí están en la ficha: deben aceptarse', r)
    assert r['identificacion']['proveedor']['nombre'] == 'POCH SA' and r['identificacion']['proveedor']['estado'] == 'encontrado'

    mala = dict(buena, proveedor_nombre='POCH S.A. de Polonia', proveedor_telefono='+48 32 000-00-00')
    r = extraer_ficha(ficha('23_'), recursos, simulada(identificacion=mala))
    mostrar('23 Nitrato de calcio · nombre y teléfono inventados: deben descartarse', r)
    assert r['identificacion']['proveedor']['nombre'] != 'POCH S.A. de Polonia'

    print('\nLos cinco controles se comportaron como se esperaba.')
