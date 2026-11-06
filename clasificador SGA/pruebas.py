# -*- coding: utf-8 -*-
"""Pruebas de QuimNexo: semáforo del índice Ir, selección de consejos para minería, lectura por perfil,
informe de una hoja, etiqueta, mediciones guardadas y manejo de errores.
No cambian ninguna tabla ni ninguna ficha. Los casos inventados llevan la palabra PRUEBA y no se guardan:
lo que haya que escribir en disco va a una carpeta temporal.
Uso:  py "clasificador SGA/pruebas.py"
Requiere: haber corrido antes  py extraccion/extraer.py FDS"""
import copy, glob, os, re, sys, tempfile

AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, AQUI)
import almacen
from tablas import cargar_tablas
from motor import clasificar, cargar_fichas, limites_de_la_sustancia, cas_de_la_ficha, avisos_de_reglas
from lectura import construir_lectura, prioridad, glosario, ORDEN_DE_COLORES
from reporte_html import dibujar_reporte
from etiqueta import contenido_de_etiqueta, capacidad_del_empaque, sin_sueltos
from etiqueta_html import dibujar_etiqueta
from etiqueta_mina import contenido_de_etiqueta_mina
from etiqueta_mina_html import dibujar_etiqueta_mina
from fichas_nuevas import procesar_fds_nueva, ErrorDeFicha, Recursos, CARPETA

RESULTADOS = []


def caso(nombre, condicion, detalle=''):
    RESULTADOS.append((nombre, bool(condicion), detalle))
    print(('  BIEN  ' if condicion else '  FALLA ') + nombre + (f'  [{detalle}]' if detalle else ''))


def sin_originales(html):
    """Quita del informe lo que por diseño va en el idioma de la ficha: nombre del producto y textos «originales»."""
    html = re.sub(r'<style>.*?</style>', '', html, flags=re.S)
    html = re.sub(r'<p class="datos">.*?</p>', '', html, flags=re.S)
    html = re.sub(r'<(span|div) class="(qr-)?original">.*?</(span|div)>', '', html, flags=re.S)
    return re.sub(r'<[^>]+>', ' ', html)


INGLES = re.compile(r'\b(the|and|with|may|causes?|keep|avoid|wear|danger|warning|report|label|download|upload|none|if|do not)\b', re.I)


def medir(ficha, tablas, limite, fraccion):
    """Resultado del motor con una medición de PRUEBA igual a 'fraccion' veces el valor del límite."""
    horas = tablas['jornada_horas'] if limite['tipo'] in ('TLV-TWA', 'Valor guía') else 0.0
    return clasificar(ficha, tablas, mediciones={limite['id']: {'tareas': [(limite['valor'] * fraccion, horas)], 'unidad': limite['unidad']}})


if __name__ == '__main__':
    sys.stdout.reconfigure(encoding='utf-8')
    tablas = cargar_tablas()
    tablas['ajustes_etiqueta'] = {'proveedor_obligatorio': False, 'tamano_minimo_de_letra_pt': 7}
    fichas = cargar_fichas()
    por_numero = {f['archivo'].split('_')[0]: f for f in fichas}
    todo = [(f, clasificar(f, tablas)) for f in fichas]
    lecturas = [(f, r, construir_lectura(f, r, tablas)) for f, r in todo]
    sem = tablas['semaforo']
    limites_de = lambda f: limites_de_la_sustancia(cas_de_la_ficha(f), tablas)
    usable = lambda l, tipos: l['usar'] == 'Sí' and l['tipo'] in tipos and isinstance(l['valor'], (int, float))

    print('\n1. Tablas y reglas')
    caso('las reglas programadas coinciden con el texto de la hoja Reglas (incluidas C06 y C07)', not avisos_de_reglas(tablas), '; '.join(avisos_de_reglas(tablas))[:200])
    caso('la hoja Semaforo_Ir trae las tres bandas seguidas y sus mensajes', sem['verde']['hasta'] == sem['amarillo']['desde'] and sem['amarillo']['hasta'] == sem['rojo']['desde']
         and all(v['mensaje'] for v in sem.values()))
    fuente = ' '.join(open(r, encoding='utf-8').read() for r in glob.glob(os.path.join(AQUI, '*.py')) if not os.path.basename(r).startswith('pruebas'))
    caso('ningún umbral del semáforo está escrito en el código (se leen de la tabla)', not re.search(r'0[.,]85\b', fuente) and not re.search(r'[<>]=?\s*0[.,]5\b', fuente))

    print('\n2. Semáforo del índice Ir (mediciones de PRUEBA sobre el diésel, nivel de peligro 3)')
    diesel = por_numero['02']
    twa = next(l for l in limites_de(diesel) if usable(l, ('TLV-TWA',)))
    for fraccion, esperado in ((0.49, 'verde'), (0.5, 'amarillo'), (0.84, 'amarillo'), (0.85, 'rojo'), (1.5, 'rojo')):
        r = medir(diesel, tablas, twa, fraccion)
        caso(f'Ir = {str(fraccion).replace(".", ",")} da {esperado}', r['semaforo']['clave'] == esperado and abs(r['semaforo']['ir'] - fraccion) < 1e-9,
             f"salió {r['semaforo']['clave']}, Ir {r['semaforo']['ir']}")
    caso('en rojo el mensaje pide parar la operación; por encima de 1 añade que se supera el valor límite',
         'PARO INMEDIATO' in medir(diesel, tablas, twa, 0.85)['semaforo']['mensaje'] and sem['rojo']['complemento'] not in medir(diesel, tablas, twa, 0.85)['semaforo']['mensaje']
         and sem['rojo']['complemento'] in medir(diesel, tablas, twa, 1.5)['semaforo']['mensaje'])
    r = clasificar(diesel, tablas)
    caso('sin medición y nivel de peligro alto: pide medir con urgencia y se muestra en amarillo', r['semaforo']['clave'] == 'gris_urgente' and r['semaforo']['color'] == 'Amarillo'
         and r['nivel']['nivel'] >= sem['gris_urgente']['nivel_desde'])
    bajo = next(((f, r) for f, r in todo if r['nivel']['nivel'] < sem['gris_urgente']['nivel_desde'] and any(usable(l, ('TLV-TWA', 'Valor guía')) for l in limites_de(f))), None)
    caso('sin medición y nivel de peligro bajo: gris', bajo is not None and bajo[1]['semaforo']['clave'] == 'gris' and bajo[1]['semaforo']['color'] == 'Gris',
         bajo[0]['archivo'][:22] if bajo else 'no hay ficha de nivel bajo con valor límite')
    con_stel = next(((f, l) for f in fichas for l in limites_de(f) if usable(l, ('TLV-STEL', 'TLV-C'))), None)
    r = medir(con_stel[0], tablas, con_stel[1], 1.2)
    caso(f"límite de corta duración o techo superado ({con_stel[1]['tipo']} de {con_stel[0]['archivo'][:14]}): rojo", r['semaforo']['clave'] == 'rojo' and r['indice_ir']['supera'])
    con_maximo = next(((f, l) for f in fichas for l in limites_de(f) if l['uso'] == 'Alerta de seguridad' and isinstance(l['valor'], (int, float))), None)
    if con_maximo:
        arriba, abajo = medir(con_maximo[0], tablas, con_maximo[1], 1.2), medir(con_maximo[0], tablas, con_maximo[1], 0.5)
        caso(f"máximo de seguridad superado ({con_maximo[0]['archivo'][:14]}): rojo de alerta; sin superarlo: verde", arriba['semaforo']['clave'] == 'rojo_alerta'
             and arriba['semaforo']['color'] == 'Rojo' and abajo['semaforo']['clave'] == 'alerta_ok', f"{arriba['semaforo']['clave']} / {abajo['semaforo']['clave']}")
    sin_limite = [r for f, r in todo if r['semaforo']['clave'] == 'sin_limite']
    caso('un producto sin valor límite utilizable dice que el semáforo no aplica', all(r['semaforo']['ir'] is None for r in sin_limite), f'{len(sin_limite)} fichas')
    dos = next((f for f in fichas if sum(usable(l, ('TLV-TWA', 'TLV-STEL', 'TLV-C', 'Valor guía')) for l in limites_de(f)) >= 2), None)
    if dos:
        a, b = [l for l in limites_de(dos) if usable(l, ('TLV-TWA', 'TLV-STEL', 'TLV-C', 'Valor guía'))][:2]
        hora = lambda l: tablas['jornada_horas'] if l['tipo'] in ('TLV-TWA', 'Valor guía') else 0.0
        r = clasificar(dos, tablas, mediciones={a['id']: {'tareas': [(a['valor'] * 0.2, hora(a))], 'unidad': a['unidad']},
                                                b['id']: {'tareas': [(b['valor'] * 0.9, hora(b))], 'unidad': b['unidad']}})
        caso('con dos límites medidos manda el mayor índice Ir', abs(r['semaforo']['ir'] - 0.9) < 1e-9 and r['semaforo']['clave'] == 'rojo' and r['semaforo']['limite']['id'] == b['id'])
    caso('todo semáforo lleva su rastro (regla C06 y fila de la hoja)', all(r['semaforo']['regla'] == 'C06' and 'Semaforo_Ir' in r['semaforo']['fila'] for _, r in todo))

    print('\n3. Selección de consejos para minería subterránea (regla C07)')
    topes = tablas['seleccion_p']
    fallas = []
    for f, r, x in lecturas:
        elegidos = x['que_hacer'] + x['accidente']
        de_la_ficha = {fr['codigo'].replace(' ', '') for fr in f['frases']['P'] if fr['estado'] == 'encontrado'}
        if len(elegidos) > topes['Total']:
            fallas.append(f"{f['archivo'][:10]}: {len(elegidos)} consejos")
        for tipo, maximo in topes.items():
            if tipo != 'Total' and sum(z['tipo'] == tipo for z in elegidos) > maximo:
                fallas.append(f"{f['archivo'][:10]}: más de {maximo} de {tipo}")
        if not {z['codigo'] for z in elegidos} <= de_la_ficha:
            fallas.append(f"{f['archivo'][:10]}: consejo que la ficha no trae")
        if any(not z['criterio'] for z in elegidos):
            fallas.append(f"{f['archivo'][:10]}: consejo sin criterio")
        disponibles = {tipo: sum(m['tipo'] == tipo for m in r['etiqueta']['P'] if m['texto']) for tipo in topes}
        if any(sum(z['tipo'] == tipo for z in elegidos) != min(maximo, disponibles[tipo]) for tipo, maximo in topes.items() if tipo != 'Total'):
            fallas.append(f"{f['archivo'][:10]}: no usó los cupos de su tipo o los rellenó con otro")
    caso(f"nunca salen más de {topes['Total']} consejos ni más de {topes['Prevención']}/{topes['Intervención']}/{topes['Almacenamiento']} por tipo, y no se rellena un tipo con otro",
         not fallas, '; '.join(fallas)[:300])
    f, r, x = next(z for z in lecturas if z[0]['archivo'].startswith('02_'))
    caso('cada consejo elegido del diésel lleva código, página de la ficha y criterio', all(z['codigo'] and z['pagina'] and z['criterio'] for z in x['que_hacer'] + x['accidente']))
    caso('el diésel (inflamable y nocivo al respirarlo) recibe un consejo de ignición y uno de ventilación', {'P210', 'P271'} <= {z['codigo'] for z in x['que_hacer']},
         ', '.join(z['codigo'] for z in x['que_hacer']))
    caso('los demás consejos siguen disponibles para la etiqueta', len(r['etiqueta']['P']) == 26 and sum(m['incluida'] for m in r['etiqueta']['P']) > topes['Total'])

    print('\n4. Frases H, lenguaje claro y protección personal')
    caso('se muestran todas las frases H de la ficha, respetando la regla R06',
         all({p['codigo'] for p in x['visibles']} == {h['codigo'] for h in r['etiqueta']['H'] if h['incluida']} for _, r, x in lecturas))
    caso('sin validar la hoja H_lenguaje_claro, se muestra el texto oficial', all(p['claro'] is None for _, _, x in lecturas for p in x['peligros']))
    validadas = copy.deepcopy(tablas['h_claro'])
    validadas['H226']['estado'] = 'Validado'   # PRUEBA: como si el autor hubiera validado esa fila
    x2 = construir_lectura(diesel, clasificar(diesel, tablas), dict(tablas, h_claro=validadas))
    caso('una frase validada por el autor se muestra en lenguaje claro, y las demás siguen oficiales',
         [p['codigo'] for p in x2['peligros'] if p['claro']] == ['H226'])
    caso('las 32 fichas traen protección personal en español, con página y estado de la traducción',
         all(x['proteccion']['texto'] and x['proteccion']['pagina'] and not x['proteccion']['original'] and 'validar' in x['proteccion']['estado'] for _, _, x in lecturas))
    caso('la acción recomendada nunca pasa de tres puntos y cada uno dice por qué', all(0 < len(x['recomendaciones']) <= 3 and all(z['porque'] for z in x['recomendaciones'])
                                                                                    for _, _, x in lecturas))

    print('\n5. Fichas con información faltante')
    sin_h = [(f, r, x) for f, r, x in lecturas if not x['peligros']]
    caso('hay fichas sin frases H y todas lo dicen en «Por verificar»',
         sin_h and all(any('indicaciones de peligro' in p for p in x['pendientes']) for _, _, x in sin_h), f'{len(sin_h)} fichas')
    caso('ninguna ficha sin frases H recibe nivel distinto de 0 ni consejos inventados', all(r['nivel']['nivel'] == 0 and len(x['que_hacer'] + x['accidente']) <= len(r['etiqueta']['P'])
                                                                                            for _, r, x in sin_h))
    f23 = por_numero.get('23')
    caso('la ficha 23 (protegida y anterior al SGA) se procesó y deja pendientes', f23 is not None and next(x for f, r, x in lecturas if f is f23)['pendientes'])
    por_revisar = [(f, r, x) for f, r, x in lecturas if r['por_revisar']['frases_h'] or r['por_revisar']['consejos_p']]
    caso('las frases «por revisar» se cuentan y no entran a la selección ni al nivel', por_revisar and all(x['por_confirmar'] for _, _, x in por_revisar), f'{len(por_revisar)} fichas')

    print('\n6. Errores de procesamiento (archivos de PRUEBA)')
    recursos = Recursos()
    pdf_en_blanco = (b'%PDF-1.4\n1 0 obj<</Type/Catalog/Pages 2 0 R>>endobj\n2 0 obj<</Type/Pages/Kids[3 0 R]/Count 1>>endobj\n'
                     b'3 0 obj<</Type/Page/Parent 2 0 R/MediaBox[0 0 612 792]>>endobj\nxref\n0 4\n0000000000 65535 f \n0000000009 00000 n \n'
                     b'0000000052 00000 n \n0000000101 00000 n \ntrailer<</Size 4/Root 1 0 R>>\nstartxref\n164\n%%EOF\n')
    for nombre, datos, esperado in (('PRUEBA_sin_texto.pdf', pdf_en_blanco, 'PDF sin texto (como un escaneo)'),
                                    ('PRUEBA_danado.pdf', b'%PDF-1.4\nesto no es un PDF de verdad', 'PDF dañado'),
                                    ('PRUEBA_no_es_pdf.pdf', b'Hola, soy un archivo de texto', 'archivo que no es PDF')):
        try:
            procesar_fds_nueva(nombre, datos, recursos)
            caso(f'{esperado} se rechaza con mensaje', False, 'se aceptó')
        except ErrorDeFicha as error:
            caso(f'{esperado} se rechaza con mensaje', True, str(error)[:70])
        except Exception as error:
            caso(f'{esperado} se rechaza con mensaje', False, f'{type(error).__name__}: {error}')
        caso(f'   y no deja archivos de «{nombre}»', not os.path.exists(os.path.join(CARPETA, nombre)))

    print('\n7. Producto que no está en el inventario (caso inventado: PRUEBA)')
    ficha = copy.deepcopy(por_numero['14'])
    ficha['archivo'] = 'PRUEBA_producto_sin_inventario.pdf'
    ficha['cotejo_cas'] = dict(ficha['cotejo_cas'], archivo=None, ficha=['0000-00-0'])
    r = clasificar(ficha, tablas)
    x = construir_lectura(ficha, r, tablas)
    caso('sin valor límite: el semáforo dice que no aplica y no calcula Ir', r['semaforo']['clave'] == 'sin_limite' and r['indice_ir']['ir_mayor'] is None)
    caso('sin traducción en la tabla: la protección personal se muestra como texto original de la ficha', x['proteccion']['original'] and 'sin traducir' in x['proteccion']['estado'])
    caso('conserva sus peligros y su selección de consejos', x['visibles'] and x['que_hacer'] and 'Informe de consulta' in dibujar_reporte(r, tablas, x))

    print('\n8. Lista «qué revisar primero»')
    muestra = [dict(archivo=f'PRUEBA_{c}_{n}', semaforo={'clave': c}, nivel={'nivel': n}) for c in ('gris', 'verde', 'rojo', 'amarillo', 'gris_urgente', 'rojo_alerta') for n in (1, 4, 2)]
    orden = [(q['semaforo']['clave'], q['nivel']['nivel']) for q in prioridad(muestra)]
    caso('ordena por color (rojo, amarillo, verde, gris) y, dentro del color, por nivel de peligro de mayor a menor',
         orden == sorted(orden, key=lambda q: (ORDEN_DE_COLORES.index(q[0]), -q[1])) and [q[0] for q in orden[::3]] == ['rojo_alerta', 'rojo', 'amarillo', 'verde', 'gris_urgente', 'gris'],
         str(orden[:6]))
    reales = prioridad([r for _, r in todo])
    caso('las 32 fichas, sin mediciones, quedan con las de nivel alto antes que las de nivel bajo',
         [ORDEN_DE_COLORES.index(q['semaforo']['clave']) for q in reales] == sorted(ORDEN_DE_COLORES.index(q['semaforo']['clave']) for q in reales))

    print('\n9. Etiqueta de uso en la mina (regla C08)')
    capacidades = ((500, 'mL'), (4, 'L'), (5, 'galón (US)'), (20, 'kg'), (200, 'L'))
    fallas, total = [], 0
    for f, r, x in lecturas:
        numero = f['archivo'].split('_')[0]
        for cantidad, unidad in capacidades:
            total += 1
            try:
                c = contenido_de_etiqueta_mina(f, r, {'capacidad': cantidad, 'unidad': unidad}, tablas)
                pagina = dibujar_etiqueta_mina(c, tablas)
                textos = [z['texto'] for z in c['renglones']] + [c['epp']['texto'] or '', (c['peligro_principal'] or {}).get('texto') or '', c['nombre']]
                assert c['lista_para_imprimir'] and 'Imprimir etiqueta' in pagina, 'no queda lista para imprimir'
                assert not any(re.search(r'\.\.\.|…|\[|\]', z) for z in textos), 'trae «...» o corchetes'
                assert c['nombre'] == tablas['nombres_etiqueta'][numero]['nombre'], 'el nombre no es el de la hoja Nombres_etiqueta'
                assert len(c['renglones']) <= topes['Total'] and all(sum(z['tipo'] == tipo for z in c['renglones']) <= maximo for tipo, maximo in topes.items() if tipo != 'Total'), 'más renglones que el tope'
                assert not sin_sueltos([z['codigo'] for z in c['renglones']])[1] and len({z['texto'] for z in c['renglones']}) == len(c['renglones']), 'renglones repetidos'
                assert c['palabra'] or c['sin_peligros'], 'sin palabra de advertencia ni «Sin peligros clasificados»'
                if r['etiqueta']['H']:
                    assert c['origen_de_los_peligros'] == 'ficha' and c['palabra'] == r['palabra_de_advertencia']['palabra'], 'la palabra no es la del motor'
                    assert [q['codigo'] for q in c['pictogramas']] == [q['codigo'] for q in r['pictogramas']['pictogramas']], 'los pictogramas no son los del motor'
                assert abs(c['litros'] - cantidad * tablas['unidades'][unidad]['litros']) < 1e-9 and c['tamano']['litros'] == c['litros'], 'capacidad mal convertida'
            except AssertionError as error:
                fallas.append(f'{numero} con {cantidad} {unidad}: {error}')
            except Exception as error:
                fallas.append(f'{numero} con {cantidad} {unidad}: {type(error).__name__} {error}')
    caso(f'las {total} etiquetas (32 fichas × 5 capacidades) quedan listas, en español, sin «...» ni corchetes, sin repetidos y con lo del motor', not fallas, '; '.join(fallas)[:400])
    origenes = {f['archivo'].split('_')[0]: contenido_de_etiqueta_mina(f, r, {'capacidad': 20, 'unidad': 'L'}, tablas) for f, r, x in lecturas if not r['etiqueta']['H']}
    caso('las fichas sin frases H toman sus peligros del inventario y lo dicen (09, 20, 23 y 27); caliza y poliacrilamida quedan «Sin peligros clasificados»',
         {n for n, c in origenes.items() if c['origen_de_los_peligros'] == 'inventario'} == {'09', '20', '23', '27'}
         and all(any('inventario' in a for a in origenes[n]['avisos']) and origenes[n]['pictogramas'] and origenes[n]['palabra'] for n in ('09', '20', '23', '27'))
         and all(origenes[n]['sin_peligros'] and not origenes[n]['pictogramas'] for n in ('12', '30')), str({n: c['origen_de_los_peligros'] for n, c in origenes.items()}))
    caso('usar el inventario en la etiqueta no cambia el motor: esas fichas siguen en nivel 0 y sin frases H en el resultado',
         all(r['nivel']['nivel'] == 0 and not r['etiqueta']['H'] for f, r, x in lecturas if f['archivo'].split('_')[0] in origenes))
    co = contenido_de_etiqueta_mina(por_numero['04'], clasificar(por_numero['04'], tablas), {'capacidad': 20, 'unidad': 'L'}, tablas)
    caso('en el monóxido de carbono entra la frase combinada P403+P233 y no la suelta P403', 'P403+P233' in [z['codigo'] for z in co['renglones']] and 'P403' not in [z['codigo'] for z in co['renglones']])
    caso('todo renglón lleva su código, su criterio y su turno de retiro; el de almacenamiento es el primero que se quita',
         all(z['codigo'] and z['criterio'] and z['retiro'] for z in co['renglones']) and next(z for z in co['renglones'] if z['tipo'] == 'Almacenamiento')['retiro'] == 1)
    escrito = contenido_de_etiqueta_mina(diesel, clasificar(diesel, tablas), {'capacidad': 20, 'unidad': 'L', 'nombre': 'PRUEBA ACPM', 'proveedor': {'nombre': 'PRUEBA S.A.S.'}, 'lote': 'PRUEBA-1'}, tablas)
    pagina = dibujar_etiqueta_mina(escrito, tablas)
    caso('el nombre escrito a mano, el proveedor y el lote salen en la etiqueta; sin proveedor no se imprime ningún aviso',
         'PRUEBA ACPM' in pagina and 'PRUEBA S.A.S.' in pagina and 'Lote: PRUEBA-1' in pagina
         and 'roveedor' not in re.search(r'<div class="etiqueta".*?</div></div>', dibujar_etiqueta_mina(co, tablas), re.S).group(0))
    sin_capacidad = contenido_de_etiqueta_mina(diesel, clasificar(diesel, tablas), {}, tablas)
    caso('lo único que impide imprimir es que falte la capacidad del empaque', not sin_capacidad['lista_para_imprimir'] and len(sin_capacidad['faltantes']) == 1 and 'empaque' in sin_capacidad['faltantes'][0])

    print('\n9b. Unidades de capacidad del empaque (hoja Unidades_empaque)')
    litros = lambda cantidad, unidad: capacidad_del_empaque({'capacidad': cantidad, 'unidad': unidad}, tablas)
    caso('20 galones (US) son 75,7 L y les toca la etiqueta de 105 × 148 mm', abs(litros(20, 'galón (US)')[0] - 75.70823568) < 1e-6
         and contenido_de_etiqueta_mina(diesel, clasificar(diesel, tablas), {'capacidad': 20, 'unidad': 'galón (US)'}, tablas)['tamano']['etiqueta_mm']['minimo'] == (105, 148))
    caso('500 mL = 0,5 L; 250 g = 0,25 L y 20 kg = 20 L (equivalencia en borrador)', abs(litros(500, 'mL')[0] - 0.5) < 1e-9 and abs(litros(250, 'g')[0] - 0.25) < 1e-9 and litros(20, 'kg')[0] == 20)
    caso('la capacidad se muestra en la unidad que eligió el usuario', litros(5, 'galón (US)')[1] == '5 gal (US)' and litros(0.5, 'L')[1] == '0,5 L' and litros(500, 'mL')[1] == '500 mL')
    vieja = {'envase_litros': 20, 'uso_interno': False, 'borde': 'rojo', 'consejos_p': [m['codigo'] for m in clasificar(por_numero['06'], tablas)['etiqueta']['P']]}   # PRUEBA: elección guardada por una versión anterior
    caso('una elección vieja, que solo trae envase_litros, se lee como 20 L', capacidad_del_empaque(vieja, tablas) == (20.0, '20 L'))
    r6 = clasificar(por_numero['06'], tablas)
    mina6, completa6 = contenido_de_etiqueta_mina(por_numero['06'], r6, vieja, tablas), contenido_de_etiqueta(r6, vieja, tablas)
    caso('con esa elección vieja (todos los consejos marcados) la etiqueta de mina sigue con cinco renglones y la completa respeta lo marcado',
         len(mina6['renglones']) <= topes['Total'] and len(completa6['consejos_p']) > topes['Total'])
    try:
        capacidad_del_empaque({'capacidad': 1, 'unidad': 'PRUEBA'}, tablas)
        caso('una unidad que no está en la tabla se rechaza con mensaje', False)
    except ValueError as error:
        caso('una unidad que no está en la tabla se rechaza con mensaje', 'Unidades_empaque' in str(error))

    print('\n9c. Etiqueta completa (Res. 773), como opción secundaria')
    fallas = []
    for f, r, x in lecturas:
        try:
            c = contenido_de_etiqueta(r, {'capacidad': 20, 'unidad': 'L'}, tablas)
            pagina = dibujar_etiqueta(c, tablas)
            assert 'Imprimir etiqueta' in pagina and 'No sustituye la etiqueta original' in pagina and 'Empaque: 20 L' in pagina
            assert {q['codigo'] for q in c['pictogramas']} == {q['codigo'] for q in r['pictogramas']['pictogramas']}
            assert c['palabra'] == r['palabra_de_advertencia']['palabra']
            assert not sin_sueltos([m['codigo'] for m in c['consejos_p']])[1], 'consejos repetidos'
            elegidos = {z['codigo'] for z in r['seleccion_p']['elegidos']} - set(sin_sueltos([m['codigo'] for m in r['etiqueta']['P']])[1])
            assert {m['codigo'] for m in c['consejos_p']} == elegidos or not elegidos, 'no vienen marcados los de la selección para minería'
        except Exception as error:
            fallas.append(f"{f['archivo'][:20]}: {type(error).__name__} {error}")
    caso('las 32 etiquetas completas se arman con «empaque», sin consejos repetidos y con los de la selección minera marcados por defecto', not fallas, '; '.join(fallas)[:300])
    con_puntos = next((contenido_de_etiqueta(r, {'capacidad': 200, 'unidad': 'L', 'consejos_p': [m['codigo'] for m in r['etiqueta']['P'] if m['completar']]}, tablas)
                       for f, r, x in lecturas if any(m['completar'] for m in r['etiqueta']['P'])), None)
    caso('en la completa, un consejo con «...» se resalta y se avisa, sin impedir la descarga',
         con_puntos is not None and any(m['sin_completar'] for m in con_puntos['consejos_p']) and any('sin completar' in a for a in con_puntos['avisos'])
         and con_puntos['lista_para_imprimir'] and 'class=completar' in dibujar_etiqueta(con_puntos, tablas))
    fuente_visible = ' '.join(open(os.path.join(AQUI, a), encoding='utf-8').read()
                              for a in ('app.py', 'etiqueta_mina.py', 'etiqueta_mina_html.py', 'reporte_html.py', 'presentacion.py', 'lectura.py'))
    cadenas = re.findall(r"'([^'\n]{8,})'", fuente_visible) + re.findall(r'"([^"\n]{8,})"', fuente_visible)
    visibles = re.sub(r'envase_litros|llama envase', '', ' '.join(cadenas))
    caso('en los textos de la pantalla, la etiqueta de mina y el informe ya no aparece «envase» (salvo la ayuda que cita la norma)', 'envase' not in visibles.lower())

    print('\n10. Informe de consulta (una hoja) de todas las fichas')
    fallas = []
    for f, r, x in lecturas:
        try:
            corto, completo = dibujar_reporte(r, tablas, x, detalle=False), dibujar_reporte(r, tablas, x, detalle=True)
            assert all(k in corto for k in ('Peligro del producto', 'Riesgo por exposición', 'Qué puede pasar', 'Qué hacer', 'Protección personal', 'Acción recomendada', 'Fuente de los datos'))
            assert 'class="anexo"' in completo and 'class="anexo"' not in corto
            assert '>None<' not in completo and 'None</' not in sin_originales(completo)
            assert all(p['codigo'] in corto for p in x['visibles']) and all(z['codigo'] in corto for z in x['que_hacer'] + x['accidente'])
        except Exception as error:
            fallas.append(f"{f['archivo'][:20]}: {type(error).__name__} {error}")
    caso(f'los {2 * len(lecturas)} informes se generan con los dos indicadores, las frases H, los consejos elegidos, la acción y la fuente', not fallas, '; '.join(fallas)[:200])
    r = medir(diesel, tablas, twa, 0.62)
    informe = dibujar_reporte(r, tablas, construir_lectura(diesel, r, tablas))
    caso('con una medición de PRUEBA el informe muestra el color, el mensaje y el número del índice («Ir = 0,62»)',
         'Amarillo' in informe and 'Ir = 0,62' in informe and sem['amarillo']['mensaje'][:40] in informe)
    sospechosos = {}
    for f, r, x in lecturas:
        for palabra in INGLES.findall(sin_originales(dibujar_reporte(r, tablas, x, detalle=False))):
            sospechosos.setdefault(palabra.lower(), set()).add(f['archivo'].split('_')[0])
    caso('el informe no trae palabras en inglés fuera del nombre del producto', not sospechosos,
         '; '.join(f'{k}: fichas {", ".join(sorted(v))}' for k, v in sospechosos.items())[:300])
    textos = ' '.join(re.findall(r"'([^'\n]{12,})'", open(os.path.join(AQUI, 'app.py'), encoding='utf-8').read()))
    caso('los textos de la interfaz no traen palabras en inglés', not re.search(r'\b(the|with|download|upload|report|label|settings|warning|loading)\b', textos, re.I))
    caso('el glosario explica el semáforo con las bandas de la tabla', any(str(sem['rojo']['desde']).replace('.', ',') in e for _, e in glosario(tablas)))

    print('\n11. Coherencia entre extracción, clasificación y resultado')
    fallas = []
    for f, r, x in lecturas:
        h_ficha = {parte[:4] for fr in f['frases']['H'] if fr['estado'] == 'encontrado' for parte in fr['codigo'].replace(' ', '').split('+')}
        if not {p['codigo'] for p in x['peligros']} <= h_ficha:
            fallas.append(f"{f['archivo'][:12]}: peligro que la ficha no trae")
        if x['peligros'] and r['nivel']['nivel'] != max(p['nivel'] for p in x['peligros']):
            fallas.append(f"{f['archivo'][:12]}: el nivel no es el máximo de sus frases")
    caso('los peligros mostrados salen todos de la ficha, y el nivel es el máximo de sus frases', not fallas, '; '.join(fallas)[:200])

    print('\n12. Mediciones e historial de consultas (en una carpeta temporal de PRUEBA)')
    original = almacen.AQUI, almacen.CONSULTAS, almacen.REGISTRO
    with tempfile.TemporaryDirectory() as carpeta:
        almacen.AQUI, almacen.CONSULTAS, almacen.REGISTRO = carpeta, os.path.join(carpeta, 'consultas'), os.path.join(carpeta, 'consultas', 'registro.csv')
        try:
            uno = {twa['id']: {'tareas': [(40.0, 8.0)], 'unidad': twa['unidad']}}
            dos_ = {twa['id']: {'tareas': [(90.0, 8.0)], 'unidad': twa['unidad']}}
            almacen.guardar_medicion('PRUEBA.pdf', uno)
            almacen.guardar_medicion('PRUEBA.pdf', uno)   # igual a la última: no se repite
            almacen.guardar_medicion('PRUEBA.pdf', dos_)
            historial = almacen.mediciones_de('PRUEBA.pdf')
            caso('cada medición se guarda con fecha y hora, sin repetir la que ya estaba', len(historial) == 2 and all(re.fullmatch(r'\d{4}-\d\d-\d\d \d\d:\d\d', m['fecha']) for m in historial))
            r = clasificar(diesel, tablas, mediciones=almacen.ultima_medicion('PRUEBA.pdf')['mediciones'])
            caso('el resultado usa la medición más reciente', abs(r['semaforo']['ir'] - 0.9) < 1e-9 and r['semaforo']['clave'] == 'rojo')
            caso('la carpeta de mediciones está dentro de la carpeta de datos del equipo', os.path.isfile(os.path.join(carpeta, 'mediciones', 'PRUEBA.json')))
            x = construir_lectura(diesel, r, tablas)
            ruta = almacen.guardar_consulta(r, x, dibujar_reporte(r, tablas, x))
            guardadas = almacen.consultas()
            caso('guardar una consulta deja el informe y una fila en el registro, con el color y el índice',
                 os.path.isfile(ruta) and len(guardadas) == 1 and guardadas[0]['Riesgo por exposición'] == 'Rojo' and guardadas[0]['Índice Ir'] == '0,9')
        finally:
            almacen.AQUI, almacen.CONSULTAS, almacen.REGISTRO = original

    print('\n13. Las etiquetas de uso en la mina abiertas en un navegador (Chromium)')
    try:
        from verificar_etiquetas import verificar
        resumen, fallas, total = verificar(capturas=False)
        caso(f'las {total} etiquetas caben en su tamaño de norma, con letra de {tablas["ajustes_etiqueta"]["tamano_minimo_de_letra_pt"]:g} pt o más, y ninguna se bloquea',
             not fallas, '; '.join(fallas)[:300])
    except ImportError:
        print('  (no comprobado: Playwright no está instalado en este equipo)')

    bien = sum(ok for _, ok, _ in RESULTADOS)
    print(f'\nRESULTADO: {bien} de {len(RESULTADOS)} comprobaciones bien.')
    sys.exit(0 if bien == len(RESULTADOS) else 1)
