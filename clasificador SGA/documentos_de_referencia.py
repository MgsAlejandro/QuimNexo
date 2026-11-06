# -*- coding: utf-8 -*-
"""Genera, con datos fijos de PRUEBA, los documentos que no deben cambiar cuando solo se toca la apariencia de la pantalla:
el informe de consulta de cinco productos (con y sin medición) y su etiqueta de uso en la mina a 20 L y a 500 mL.
Sirve para comparar una versión con otra: se genera antes del cambio, se genera después y se comparan las dos carpetas.
No lee ni escribe datos del usuario (mediciones, elecciones, consultas): todo va a la carpeta que se le indique.
Uso:
    py "clasificador SGA/documentos_de_referencia.py" CARPETA_DE_SALIDA
    py "clasificador SGA/documentos_de_referencia.py" --comparar CARPETA_A CARPETA_B"""
import filecmp, os, re, sys

AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, AQUI)
PRODUCTOS = ('02', '04', '06', '14', '27')   # diésel, monóxido de carbono, estireno, ácido sulfúrico y dióxido de carbono
FECHA = '2026-01-01 08:00'                   # fecha fija, para que el informe no cambie de un día a otro
FRACCION = 0.62                              # medición de PRUEBA: 62 % del valor límite


def generar(destino):
    from tablas import cargar_tablas
    from motor import clasificar, cargar_fichas, limites_de_la_sustancia, cas_de_la_ficha
    from lectura import construir_lectura
    from reporte_html import dibujar_reporte
    from etiqueta_mina import contenido_de_etiqueta_mina
    from etiqueta_mina_html import dibujar_etiqueta_mina
    os.makedirs(destino, exist_ok=True)
    tablas = cargar_tablas()
    tablas['ajustes_etiqueta'] = {'proveedor_obligatorio': False, 'tamano_minimo_de_letra_pt': 7}
    hechos = []
    for ficha in cargar_fichas():
        numero = ficha['archivo'].split('_')[0]
        if numero not in PRODUCTOS:
            continue
        documentos = {}
        limite = next((l for l in limites_de_la_sustancia(cas_de_la_ficha(ficha), tablas)
                       if l['usar'] == 'Sí' and l['tipo'] in ('TLV-TWA', 'Valor guía') and isinstance(l['valor'], (int, float))), None)
        casos = [('sin_medicion', {})]
        if limite:
            casos.append(('con_medicion', {limite['id']: {'tareas': [(limite['valor'] * FRACCION, tablas['jornada_horas'])], 'unidad': limite['unidad']}}))
        for nombre, mediciones in casos:
            r = clasificar(ficha, tablas, mediciones=mediciones)
            x = construir_lectura(ficha, r, tablas)
            if mediciones:
                x['fecha_de_medicion'] = FECHA
            documentos[f'informe_{nombre}'] = dibujar_reporte(r, tablas, x, detalle=False, fecha=FECHA)
        r = clasificar(ficha, tablas)
        for cantidad, unidad in ((20, 'L'), (500, 'mL')):
            documentos[f'etiqueta_mina_{cantidad}{unidad}'] = dibujar_etiqueta_mina(contenido_de_etiqueta_mina(ficha, r, {'capacidad': cantidad, 'unidad': unidad}, tablas), tablas)
        for nombre, html in documentos.items():
            with open(os.path.join(destino, f'{numero}_{nombre}.html'), 'w', encoding='utf-8') as f:
                f.write(html)
            hechos.append(f'{numero}_{nombre}.html')
    return hechos


def comparar(a, b):
    """Compara dos carpetas de referencia. Devuelve (iguales, distintos, [(archivo, primeras diferencias)])."""
    nombres = sorted(set(os.listdir(a)) | set(os.listdir(b)))
    nombres = [n for n in nombres if n.endswith('.html')]
    iguales, distintos, detalle = [], [], []
    for n in nombres:
        ra, rb = os.path.join(a, n), os.path.join(b, n)
        if os.path.isfile(ra) and os.path.isfile(rb) and filecmp.cmp(ra, rb, shallow=False):
            iguales.append(n)
            continue
        distintos.append(n)
        ta = open(ra, encoding='utf-8').read() if os.path.isfile(ra) else ''
        tb = open(rb, encoding='utf-8').read() if os.path.isfile(rb) else ''
        pa, pb = re.split(r'(?<=>)', ta), re.split(r'(?<=>)', tb)
        diferencias = [(x[:160], y[:160]) for x, y in zip(pa, pb) if x != y][:4]
        detalle.append((n, diferencias or [('falta en una de las dos carpetas', '')]))
    return iguales, distintos, detalle


if __name__ == '__main__':
    sys.stdout.reconfigure(encoding='utf-8')
    if len(sys.argv) >= 4 and sys.argv[1] == '--comparar':
        iguales, distintos, detalle = comparar(sys.argv[2], sys.argv[3])
        print(f'{len(iguales)} documentos idénticos; {len(distintos)} distintos.')
        for nombre, diferencias in detalle:
            print(f'\n{nombre}')
            for x, y in diferencias:
                print(f'   antes:   {x}\n   después: {y}')
        sys.exit(1 if distintos else 0)
    hechos = generar(sys.argv[1])
    print(f'{len(hechos)} documentos de referencia en {sys.argv[1]}:')
    print('  ' + '\n  '.join(hechos))
