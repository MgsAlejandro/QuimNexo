# -*- coding: utf-8 -*-
"""Contenido de la etiqueta de uso en la mina (regla C08): corta, en español y siempre imprimible.
Lleva: nombre del producto en español, pictogramas y palabra de advertencia del motor, el peligro principal,
hasta cinco renglones de «qué hacer» (los consejos de la selección para minería, en su texto corto) y la
protección personal. No clasifica ni redacta: cada texto sale de una tabla
  nombre              hoja Nombres_etiqueta (o el inventario, o la ficha, o lo que escriba el usuario)
  pictogramas, palabra   motor.py, con sus reglas de siempre
  peligro principal   frase H de mayor nivel, con el texto de la hoja H_lenguaje_claro
  qué hacer           motor.seleccion_de_consejos (regla C07) y hoja P_corto
  protección personal columnas «EPP corto para etiqueta» e «Iconos de EPP» de proteccion_seccion8.xlsx
Si la ficha no trae frases H, usa las que la matriz reconoce en el inventario del proyecto (hoja Aplicacion_30) y
los consejos P clave del inventario, y lo dice. Eso solo vale para esta etiqueta: el nivel y el informe no cambian.
No reemplaza la etiqueta completa de la Resolución 773 de 2021 (etiqueta.py).
Uso de prueba:
    py "clasificador SGA/etiqueta_mina.py" 14_ --capacidad 20 --unidad L
    py "clasificador SGA/etiqueta_mina.py" 27_ --capacidad 5 --unidad "galón (US)"
Requiere: haber corrido antes  py extraccion/extraer.py FDS"""
import sys
from tablas import cargar_tablas
from motor import (clasificar, cargar_fichas, nivel_de_peligro, palabra_de_advertencia, pictogramas, frases_de_etiqueta, seleccion_de_consejos,
                   tamano_de_etiqueta)
from etiqueta import capacidad_del_empaque, sin_sueltos


def _nombre(ficha, r, elecciones, tablas):
    """Nombre en español para la etiqueta y de dónde salió."""
    escrito = str((elecciones or {}).get('nombre') or '').strip()
    if escrito:
        return escrito, 'escrito por el usuario'
    de_la_hoja = tablas['nombres_etiqueta'].get(ficha['archivo'].split('_')[0]) if ficha.get('origen') != 'nueva' else None
    if de_la_hoja:
        return de_la_hoja['nombre'], f"hoja Nombres_etiqueta ({de_la_hoja['estado']})"
    if r['sustancia_en_excel']:
        return r['sustancia_en_excel'], 'inventario del proyecto (se cruzó por CAS)'
    return r['producto'] or ficha['archivo'], 'nombre de la ficha: no hay nombre en español; puede escribirlo a mano'


def contenido_de_etiqueta_mina(ficha, r, elecciones, tablas):
    """ficha = resultado de la extracción; r = salida de motor.clasificar; elecciones = lo que dio el usuario
    (capacidad, unidad, nombre, proveedor, cantidad, lote). Devuelve el contenido de la etiqueta, con su rastro.
    Lo único que puede faltar para imprimirla es la capacidad del empaque."""
    e, m = elecciones or {}, tablas['etiqueta_mina']
    avisos = [m['texto_no_es_res_773']]
    nombre, origen_del_nombre = _nombre(ficha, r, e, tablas)

    # --- Peligros: los de la ficha; si la ficha no trae frases H, los que la matriz reconoce en el inventario ---
    h_ficha = [f['codigo'] for f in r['etiqueta']['H'] if f['incluida']]
    origen = 'ficha'
    if h_ficha:
        nivel, palabra, pictos, consejos, codigos_h = r['nivel'], r['palabra_de_advertencia']['palabra'], r['pictogramas']['pictogramas'], r['etiqueta']['P'], h_ficha
    else:
        codigos_h = next((tablas['h_del_inventario'][c] for c in r['cas'] if tablas['h_del_inventario'].get(c)), [])
        origen = 'inventario' if codigos_h else 'ninguno'
        nivel = nivel_de_peligro(codigos_h, tablas)
        palabra = palabra_de_advertencia(codigos_h, {}, tablas)['palabra'] if codigos_h else None
        pictos = pictogramas(codigos_h, tablas)['pictogramas'] if codigos_h else []
        consejos = r['etiqueta']['P']
        if codigos_h:
            avisos.append(f"{m['texto_inventario']} Frases tomadas: {', '.join(codigos_h)}.")
    if not consejos:   # la ficha tampoco trae consejos P: se usan los «consejos P clave» del inventario
        del_inventario = next((tablas['excel_fichas'][c].get('frases_p') for c in r['cas'] if tablas['excel_fichas'].get(c, {}).get('frases_p')), [])
        if del_inventario:
            consejos = frases_de_etiqueta([], [(p, None) for p in del_inventario], tablas)['P']
            avisos.append(f"La ficha no trae consejos de prudencia: se usaron los consejos clave del inventario del proyecto ({', '.join(del_inventario)}).")

    # --- Peligro principal: la frase H de mayor nivel, en lenguaje claro ---
    detalle = sorted(nivel['detalle'], key=lambda d: -d['nivel'])
    principal = None
    if detalle:
        claro = tablas['h_claro'].get(detalle[0]['codigo'])
        principal = dict(codigo=detalle[0]['codigo'], texto=(claro or {}).get('texto') or detalle[0]['texto'],
                         estado=(claro or {}).get('estado') or 'Texto oficial', fila=f"H_lenguaje_claro, {detalle[0]['codigo']}" if claro else 'Matriz_H')
    asfixiante = m['texto_asfixiante'] if any(nota['regla'] == 'C01' for nota in r['notas']) else None

    # --- Qué hacer: selección para minería (regla C07) sobre los consejos que tienen texto corto, sin sueltos repetidos ---
    quedan, repetidos = sin_sueltos([f['codigo'] for f in consejos])
    sin_texto = [c for c in quedan if c not in tablas['p_corto']]
    candidatos = [f for f in consejos if f['codigo'] in quedan and f['codigo'] in tablas['p_corto']]
    elegidos = seleccion_de_consejos(candidatos, codigos_h, tablas)['elegidos']
    # Orden en que se quitan si el contenido no cabe: primero el de almacenamiento, luego los de menor puntaje
    retiro = sorted(elegidos, key=lambda x: (x['tipo'] != 'Almacenamiento', x['puntaje'], x['tipo'] != 'Intervención', x['codigo']))
    renglones = [dict(codigo=x['codigo'], tipo=x['tipo'], texto=tablas['p_corto'][x['codigo']]['texto'], estado=tablas['p_corto'][x['codigo']]['estado'],
                      puntaje=x['puntaje'], criterio=x['criterio'], retiro=retiro.index(x) + 1) for x in elegidos]
    if repetidos:
        avisos.append('No se repiten, porque ya van dentro de un consejo combinado: ' + ', '.join(repetidos) + '.')
    if sin_texto:
        avisos.append('Consejos de la ficha sin texto corto en la hoja P_corto (no entran a esta etiqueta): ' + ', '.join(sin_texto) + '.')

    # --- Protección personal ---
    fila = tablas['proteccion'].get(ficha['archivo'])
    if fila and fila.get('corto') and fila['original'] == (ficha.get('proteccion') or {}).get('texto'):
        epp = dict(texto=str(fila['corto']).strip(), iconos=fila['iconos'], estado=fila['estado_corto'], pagina=fila['pagina'])
    else:
        epp = dict(texto=None, iconos=[], estado='Sin texto corto en proteccion_seccion8.xlsx', pagina=(ficha.get('proteccion') or {}).get('pagina'))
        avisos.append('Esta ficha no tiene línea corta de protección personal en la tabla: la etiqueta va sin ella. Consulte la sección 8 de la ficha.')

    # --- Capacidad del empaque y tamaño de la etiqueta (hoja Etiqueta_tamano, sin cambios) ---
    litros, capacidad = capacidad_del_empaque(e, tablas)
    tamano = tamano_de_etiqueta(litros, tablas) if litros else None
    faltantes = [] if tamano else ['La capacidad del empaque: de ella depende el tamaño de la etiqueta.']
    if tamano and e.get('unidad') in tablas['unidades'] and 'validar' in str(tablas['unidades'][e['unidad']]['estado']).lower():
        avisos.append(f"La capacidad se dio en {e['unidad']}: {tablas['unidades'][e['unidad']]['origen']} ({tablas['unidades'][e['unidad']]['estado']}).")
    proveedor = {c: str(((e.get('proveedor') or {}).get(c)) or '').strip() for c in ('nombre', 'direccion', 'telefono')}
    borradores = sorted({x for x in [principal and principal['estado'], epp['estado']] + [z['estado'] for z in renglones] if x and 'validar' in str(x).lower()})
    if borradores:
        avisos.append('Los textos de esta etiqueta son borradores que el autor aún no ha validado.')
    return dict(
        archivo=ficha['archivo'], nombre=nombre, origen_del_nombre=origen_del_nombre, nombre_en_la_ficha=r['producto'],
        pictogramas=[dict(codigo=p['codigo'], simbolo=p['simbolo']) for p in pictos], palabra=palabra,
        origen_de_los_peligros=origen, frases_h=codigos_h, sin_peligros=m['texto_sin_peligros'] if origen == 'ninguno' else None,
        peligro_principal=principal, asfixiante=asfixiante, renglones=renglones, epp=epp,
        proveedor=proveedor, cantidad=str(e.get('cantidad') or '').strip(), lote=str(e.get('lote') or '').strip(),
        litros=litros, capacidad=capacidad, tamano=tamano,
        tamano_mas_pequeno=bool(tamano and tamano['capacidad'] == tablas['etiqueta_tamano'][0]['capacidad']),
        orden_de_retiro=[m[f'retiro_{i}'] for i in range(1, 6)], minimo_de_renglones=int(m['minimo_de_renglones']),
        rotulos=dict(peligro=m['rotulo_peligro'], hacer=m['rotulo_hacer'], epp=m['rotulo_epp']),
        faltantes=faltantes, avisos=avisos, reglas=['C08', 'C07'], lista_para_imprimir=not faltantes)


if __name__ == '__main__':
    sys.stdout.reconfigure(encoding='utf-8')
    argumentos = sys.argv[1:]
    tablas = cargar_tablas()
    valor = lambda bandera, defecto: argumentos[argumentos.index(bandera) + 1] if bandera in argumentos else defecto
    elecciones = {'capacidad': float(str(valor('--capacidad', '20')).replace(',', '.')), 'unidad': valor('--unidad', 'L')}
    for ficha in cargar_fichas(argumentos[0] if argumentos and not argumentos[0].startswith('--') else ''):
        c = contenido_de_etiqueta_mina(ficha, clasificar(ficha, tablas), elecciones, tablas)
        print(f"\n{c['nombre']}  ({c['nombre_en_la_ficha']})  ·  {c['palabra'] or c['sin_peligros'] or 'sin palabra de advertencia'}  ·  "
              f"{', '.join(p['codigo'] for p in c['pictogramas']) or 'sin pictogramas'}  ·  peligros de: {c['origen_de_los_peligros']}")
        print(f"   Empaque de {c['capacidad']} = {c['litros']:g} L -> etiqueta {c['tamano']['etiqueta']}")
        if c['peligro_principal']:
            print(f"   {c['rotulos']['peligro']}: {c['peligro_principal']['texto']}" + (f" {c['asfixiante']}" if c['asfixiante'] else ''))
        for z in c['renglones']:
            print(f"   - {z['texto']}   [{z['codigo']}, se quita de {z['retiro']}.º]")
        print(f"   {c['rotulos']['epp']}: {c['epp']['texto']}  [{', '.join(c['epp']['iconos'])}]")
