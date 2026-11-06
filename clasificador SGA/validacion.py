# -*- coding: utf-8 -*-
"""Paso M8 del motor: resultado por ficha y validación contra la hoja Aplicacion_30 de la matriz.
Hace tres cosas y no corrige ninguna tabla:
  1) Clasifica las fichas ya extraídas y guarda un resultado por ficha en "clasificador SGA/salidas/".
  2) Prueba A, de las reglas: le da al motor las mismas frases H que usó la matriz en Aplicacion_30 y comprueba
     que salgan el mismo nivel, las mismas frases determinantes, la misma palabra y los mismos pictogramas.
     Si coincide, el motor aplica las reglas igual que el generador de la matriz.
  3) Prueba B, de las fichas: compara lo que da el motor con las frases de cada FDS contra lo que la matriz
     obtuvo con las frases del Excel de fichas. Las diferencias aquí no son errores del motor: son diferencias
     entre la FDS y el Excel, para que el autor las revise.
Todo queda en "clasificador SGA/salidas/validacion_motor.xlsx".
Uso:  py "clasificador SGA/validacion.py"
Requiere: pip install openpyxl, y haber corrido antes  py extraccion/extraer.py FDS"""
import json, os, re, sys
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from tablas import cargar_tablas, MATRIZ
from motor import (nivel_de_peligro, palabra_de_advertencia, pictogramas, clasificar, cargar_fichas, avisos_de_reglas)

SALIDAS = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'salidas')
CAS = re.compile(r'\d{2,7}-\d{2}-\d')


def cargar_aplicacion_30():
    libro = openpyxl.load_workbook(MATRIZ, read_only=True, data_only=True)
    try:
        filas = list(libro['Aplicacion_30'].iter_rows(values_only=True))
    finally:
        libro.close()
    col = {n: i for i, n in enumerate(filas[0])}
    return [dict(sustancia=f[col['Sustancia']], cas=CAS.findall(str(f[col['CAS']] or '')), frases=re.findall(r'H\d{3}', str(f[col['Frases H reconocidas']] or '')),
                 nivel=int(f[col['Nivel matriz']]), determinantes=re.findall(r'H\d{3}', str(f[col['Frase(s) determinante(s)']] or '')),
                 palabra=str(f[col['Palabra de advertencia']]), pictogramas=re.findall(r'GHS\d{2}', str(f[col['Pictogramas (con precedencia)']] or '')),
                 observaciones=f[col['Observaciones']]) for f in filas[1:] if f[col['Sustancia']]]


def _palabra_comparable(del_motor, de_la_matriz):
    """La matriz escribe 'Sin palabra de advertencia' y, cuando depende de la categoría, 'Peligro o Atención según
    categoría (ver FDS)'. En ese último caso el motor, sin categoría, aplica el peor caso (C02): es el mismo caso."""
    if de_la_matriz.startswith('Peligro o Atención'):
        return 'C02' in del_motor['reglas']
    if de_la_matriz.startswith('Sin palabra'):
        return del_motor['palabra'] is None
    return del_motor['palabra'] == de_la_matriz


def prueba_de_reglas(aplicacion, tablas):
    """Prueba A: mismas frases H que Aplicacion_30 -> ¿mismo resultado?"""
    filas = []
    for a in aplicacion:
        n = nivel_de_peligro(a['frases'], tablas)
        p = palabra_de_advertencia(a['frases'], {}, tablas)
        g = [x['codigo'] for x in pictogramas(a['frases'], tablas)['pictogramas']]
        iguales = dict(nivel=n['nivel'] == a['nivel'], determinantes=sorted(n['determinantes']) == sorted(a['determinantes']),
                       palabra=_palabra_comparable(p, a['palabra']), pictogramas=g == sorted(a['pictogramas']))
        filas.append(dict(sustancia=a['sustancia'], frases=a['frases'], matriz=a,
                          motor=dict(nivel=n['nivel'], determinantes=n['determinantes'],
                                     palabra=(p['palabra'] or 'ninguna') + (' (peor caso, C02)' if 'C02' in p['reglas'] else ''), pictogramas=g),
                          iguales=iguales, coincide=all(iguales.values())))
    return filas


def prueba_de_fichas(resultados, aplicacion):
    """Prueba B: motor con las frases de la FDS frente a la matriz con las frases del Excel."""
    filas = []
    for r in resultados:
        a = next((x for x in aplicacion if set(x['cas']) & set(r['cas'])), None)
        entran = [d['codigo'] for d in r['nivel']['detalle']]
        fila = dict(archivo=r['archivo'], producto=r['producto'], nivel=r['nivel']['nivel'], nombre=r['nivel']['nombre'],
                    determinantes=r['nivel']['determinantes'], palabra=r['palabra_de_advertencia']['palabra'],
                    pictogramas=[x['codigo'] for x in r['pictogramas']['pictogramas']], frases=entran, matriz=a)
        if a is None:
            fila['resultado'] = 'Sin fila en Aplicacion_30 (no hay CAS en común)'
        elif fila['nivel'] == a['nivel']:
            fila['resultado'] = 'Mismo nivel'
        else:
            fila['resultado'] = f"Nivel distinto: FDS {fila['nivel']}, Excel {a['nivel']}"
        if a is not None:
            fila['solo_fds'] = sorted(set(entran) - set(a['frases']))
            fila['solo_excel'] = sorted(set(a['frases']) - set(entran))
            fila['compartida'] = len(a['cas']) > 1
        filas.append(fila)
    return filas


def escribir(resultados, reglas, fichas, destino):
    AR = 'Arial'
    rojo, ambar, verde = (PatternFill('solid', fgColor=c) for c in ('FF7C80', 'FFE699', 'C6E0B4'))
    wb = openpyxl.Workbook()

    def hoja(nombre, encabezado, filas, anchos, primera=False):
        ws = wb.active if primera else wb.create_sheet()
        ws.title = nombre
        ws.append(encabezado)
        for f in filas:
            ws.append([', '.join(map(str, x)) if isinstance(x, (list, tuple)) else x for x in f])
        for fila in ws.iter_rows():
            for c in fila:
                c.font = Font(name=AR, size=10)
                c.alignment = Alignment(wrap_text=True, vertical='top')
        for c in ws[1]:
            c.font = Font(name=AR, bold=True, color='FFFFFF')
            c.fill = PatternFill('solid', fgColor='1F3864')
        for i, a in enumerate(anchos, 1):
            ws.column_dimensions[get_column_letter(i)].width = a
        ws.freeze_panes = 'B2'
        return ws

    hoja('LEEME', ('Tema', 'Detalle'), [
        ('Qué es', 'Validación del motor de clasificación estático. Lo genera "clasificador SGA/validacion.py"; se puede regenerar cuando cambien las fichas o las tablas.'),
        ('Qué NO hace', 'No corrige nada: ni la matriz, ni el Excel de fichas, ni lo extraído de las FDS.'),
        ('Hoja Prueba_reglas', 'Al motor se le dan las mismas frases H que la matriz usó en su hoja Aplicacion_30. Si nivel, frases determinantes, palabra y pictogramas coinciden en todas las filas, el motor aplica las reglas igual que construir_matriz_HP.py. Una diferencia aquí sí sería un defecto del motor o una regla no escrita en la hoja Reglas.'),
        ('Hoja Fichas_vs_Excel', 'Motor con las frases H de cada FDS, frente a la matriz con las frases H del Excel de fichas. Una diferencia aquí NO es un defecto del motor: indica que la FDS y el Excel no traen las mismas frases. La decisión es del autor.'),
        ('Hoja Resultado_32', 'Resultado del motor para cada ficha: nivel, palabra, pictogramas, frases de la etiqueta, límites disponibles y notas.'),
        ('Palabra «peor caso»', 'Cuando una frase lleva "Peligro" o "Atención" según la categoría y no se conoce la categoría, la matriz escribe "Peligro o Atención según categoría (ver FDS)" y el motor toma el peor caso (regla C02). Se cuentan como el mismo caso. Con la FDS real el motor sí usa la categoría declarada.'),
        ('Lo que el motor no decide', 'La valoración final que combina nivel de peligro e índice Ir. Está pendiente de definición por el autor; el motor entrega ambos por separado.'),
    ], (26, 130), primera=True)

    ws = hoja('Prueba_reglas', ('Sustancia (Aplicacion_30)', 'Frases H usadas', 'Nivel matriz', 'Nivel motor', 'Determinantes matriz', 'Determinantes motor',
                                'Palabra matriz', 'Palabra motor', 'Pictogramas matriz', 'Pictogramas motor', 'Resultado'),
              [(f['sustancia'], f['frases'], f['matriz']['nivel'], f['motor']['nivel'], f['matriz']['determinantes'], f['motor']['determinantes'],
                f['matriz']['palabra'], f['motor']['palabra'], sorted(f['matriz']['pictogramas']), f['motor']['pictogramas'],
                'Coincide' if f['coincide'] else 'DIFIERE en ' + ', '.join(k for k, v in f['iguales'].items() if not v)) for f in reglas],
              (34, 34, 8, 8, 20, 20, 22, 22, 22, 22, 26))
    for fila in ws.iter_rows(min_row=2):
        fila[10].fill = verde if fila[10].value == 'Coincide' else rojo

    ws = hoja('Fichas_vs_Excel', ('Ficha', 'Producto según la FDS', 'Sustancia en Aplicacion_30', 'Nivel con la FDS', 'Nivel con el Excel', 'Resultado',
                                  'Frases H solo en la FDS', 'Frases H solo en el Excel', 'Fila compartida por dos sustancias'),
              [(f['archivo'], f['producto'], f['matriz']['sustancia'] if f['matriz'] else None, f"{f['nivel']} {f['nombre']}",
                f['matriz']['nivel'] if f['matriz'] else None, f['resultado'], f.get('solo_fds'), f.get('solo_excel'),
                'Sí' if f.get('compartida') else '') for f in fichas], (34, 32, 30, 14, 10, 30, 30, 30, 12))
    for fila in ws.iter_rows(min_row=2):
        v = str(fila[5].value)
        fila[5].fill = verde if v == 'Mismo nivel' else ambar if v.startswith('Sin fila') else rojo

    hoja('Resultado_32', ('Ficha', 'Producto según la FDS', 'Uso en minería (Excel)', 'Nivel', 'Frases determinantes', 'Palabra de advertencia',
                          'Pictogramas', 'Frases H en etiqueta', 'Consejos P por defecto', 'Consejos P disponibles', 'Límites para Ir (hoja TLV)',
                          'Notas del motor', 'Frases por revisar (fuera del cálculo)', 'Hallazgos de la ficha'),
         [(r['archivo'], r['producto'], r['uso_en_mineria'], f"{r['nivel']['nivel']} {r['nivel']['nombre']}", r['nivel']['determinantes'],
           r['palabra_de_advertencia']['palabra'] or 'ninguna', [x['codigo'] for x in r['pictogramas']['pictogramas']] or 'ninguno',
           [f['codigo'] for f in r['etiqueta']['H'] if f['incluida']], [f['codigo'] for f in r['etiqueta']['P'] if f['incluida']],
           [f['codigo'] for f in r['etiqueta']['P'] if not f['incluida']],
           '; '.join(f"{l['id']} {l['tipo']} {l['valor']:g} {l['unidad']} ({l['usar']})" if isinstance(l['valor'], (int, float))
                     else f"{l['id']} sin valor ({l['usar']})" for l in r['indice_ir']['limites']) or 'sin filas en la hoja TLV',
           ' | '.join(n['texto'] for n in r['notas']), ' | '.join(r['por_revisar']['frases_h'] + r['por_revisar']['consejos_p']),
           len(r['hallazgos_de_la_ficha'])) for r in resultados], (34, 30, 34, 14, 18, 12, 22, 26, 44, 16, 50, 60, 40, 10))
    wb.save(destino)


if __name__ == '__main__':
    sys.stdout.reconfigure(encoding='utf-8')
    tablas = cargar_tablas()
    for aviso in avisos_de_reglas(tablas):
        print('AVISO:', aviso)
    os.makedirs(SALIDAS, exist_ok=True)
    resultados = []
    for ficha in cargar_fichas():
        r = clasificar(ficha, tablas)
        resultados.append(r)
        with open(os.path.join(SALIDAS, os.path.splitext(r['archivo'])[0] + '.json'), 'w', encoding='utf-8') as f:
            json.dump(r, f, ensure_ascii=False, indent=1)
    aplicacion = cargar_aplicacion_30()
    reglas, fichas = prueba_de_reglas(aplicacion, tablas), prueba_de_fichas(resultados, aplicacion)
    destino = os.path.join(SALIDAS, 'validacion_motor.xlsx')
    escribir(resultados, reglas, fichas, destino)

    print(f'{len(resultados)} resultados guardados en {SALIDAS}')
    print(f"\nPRUEBA A — reglas: {sum(f['coincide'] for f in reglas)} de {len(reglas)} filas de Aplicacion_30 coinciden")
    for f in reglas:
        if not f['coincide']:
            print(f"   DIFIERE {f['sustancia']}: " + '; '.join(f"{k}: matriz {f['matriz'][k]} / motor {f['motor'][k]}" for k, v in f['iguales'].items() if not v))
    print('\nPRUEBA B — fichas frente al Excel:')
    cuenta = {}
    for f in fichas:
        clave = f['resultado'].split(':')[0]
        cuenta[clave] = cuenta.get(clave, 0) + 1
    print('   ' + ' | '.join(f'{k}: {v}' for k, v in cuenta.items()))
    for f in fichas:
        if not f['resultado'].startswith('Mismo'):
            extra = f" | solo FDS: {', '.join(f['solo_fds']) or '—'} | solo Excel: {', '.join(f['solo_excel']) or '—'}" if f['matriz'] else ''
            print(f"   {f['archivo'][:26]:<27}{f['resultado']}{extra}")
    print(f'\nValidación guardada en {destino}')
