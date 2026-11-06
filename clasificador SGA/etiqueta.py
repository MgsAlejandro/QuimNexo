# -*- coding: utf-8 -*-
"""Paso E2 del reporte y la etiqueta: contenido de la etiqueta.
Junta lo que calculó el motor con lo que decide el usuario (el "responsable del etiquetado" de la Res. 773)
y comprueba el contenido mínimo. No dibuja nada todavía y no decide nada de clasificación.
Si falta algo obligatorio lo dice en 'faltantes' y la etiqueta no queda lista para imprimir.

Elecciones del usuario (todas opcionales salvo la capacidad del empaque):
    capacidad, unidad   capacidad del empaque y su unidad (mL, L, galón (US), g o kg; hoja Unidades_empaque)
    envase_litros   capacidad en litros; es lo que traen las elecciones guardadas por versiones anteriores, y se lee como litros
    proveedor       {'nombre', 'direccion', 'telefono'}; solo entra a la etiqueta si el usuario lo da. Si no se da,
                    la etiqueta va sin proveedor (permitido por configuracion.json) y lleva un aviso
    uso_interno     True si el producto no sale del lugar de trabajo
    borde           'rojo' (por defecto) o 'negro' (solo con uso_interno, regla R09)
    consejos_p      lista de códigos P que van en la etiqueta; si no se da, los de la selección para minería (reglas C05 y C07)
    textos          {código P: texto completo} para los consejos cuyo texto oficial trae "..."
    componentes     texto del renglón de componentes, para mezclas (regla R11)
    cantidad, lote  opcionales (Res. 773, art. 9)

Uso de prueba:
    py "clasificador SGA/etiqueta.py" 14_ --envase 1
    py "clasificador SGA/etiqueta.py" 14_ --elecciones "clasificador SGA/ejemplo_elecciones.json"
Requiere: haber corrido antes  py "clasificador SGA/validacion.py"."""
import glob, json, os, re, sys
from tablas import cargar_tablas
from motor import tamano_de_etiqueta

SALIDAS = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'salidas')
CONFIGURACION = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'configuracion.json')
SIN_COMPLETAR = re.compile(r'\.\.\.|…|\[[^\]]*\]')   # «...» o un texto entre corchetes: lo completa o lo decide quien etiqueta


def numero(valor):
    """Número con coma decimal y sin ceros de sobra, como se escribe en Colombia."""
    return f'{round(valor, 2):g}'.replace('.', ',')


def capacidad_del_empaque(elecciones, tablas):
    """Capacidad del empaque que dio el usuario. Devuelve (litros, texto en la unidad elegida), o (None, None) si falta.
    Los factores a litros salen de la hoja Unidades_empaque. Una elección vieja, que solo trae 'envase_litros', son litros."""
    e = elecciones or {}
    cantidad, unidad = (e.get('capacidad'), e.get('unidad') or 'L') if e.get('capacidad') not in (None, '') else (e.get('envase_litros'), 'L')
    if cantidad in (None, ''):
        return None, None
    if unidad not in tablas['unidades']:
        raise ValueError(f"La unidad «{unidad}» no está en la hoja Unidades_empaque. Unidades que hay: {', '.join(tablas['unidades'])}.")
    return float(cantidad) * tablas['unidades'][unidad]['litros'], f"{numero(float(cantidad))} {tablas['unidades'][unidad]['abreviatura']}"


def sin_sueltos(codigos):
    """Quita los consejos sueltos que ya están dentro de uno combinado de la misma lista (P403 si está P403+P233).
    Devuelve (los que quedan, los que se quitaron)."""
    partes = {c: set(c.split('+')) for c in codigos}
    repetidos = [c for c in codigos if any(partes[c] < partes[otro] for otro in codigos)]
    return [c for c in codigos if c not in repetidos], repetidos


def ajustes(tablas=None):
    """Ajustes de operación de la etiqueta, de configuracion.json. No son reglas de la norma.
    Para las pruebas se pueden pasar dentro de 'tablas' con la clave 'ajustes_etiqueta'."""
    if tablas and 'ajustes_etiqueta' in tablas:
        return tablas['ajustes_etiqueta']
    with open(CONFIGURACION, encoding='utf-8') as f:
        return {k: v for k, v in json.load(f)['etiqueta'].items() if not k.startswith('_')}


def contenido_de_etiqueta(resultado, elecciones, tablas):
    """resultado = salida de motor.clasificar para una ficha. Devuelve el contenido de la etiqueta:
    {'tipo', 'producto', 'proveedor', 'pictogramas', 'borde', 'palabra', 'frases_h', 'consejos_p', 'componentes',
     'cantidad', 'lote', 'tamano', 'faltantes', 'avisos', 'reglas', 'lista_para_imprimir'}."""
    e = elecciones or {}
    faltantes, avisos, reglas = [], [], []

    # --- Tamaño, y si el empaque es tan pequeño que basta el contenido mínimo (R10) ---
    tamano = None
    litros, capacidad = capacidad_del_empaque(e, tablas)
    if litros is None:
        faltantes.append('La capacidad del empaque: de ella depende el tamaño de la etiqueta.')
    else:
        tamano = tamano_de_etiqueta(litros, tablas)
    envase_pequeno = bool(tamano and 'R10' in tamano['reglas'])

    # --- Frases H y tipo de etiqueta ---
    frases_h = [dict(codigo=f['codigo'], texto=f['texto']) for f in resultado['etiqueta']['H'] if f['incluida']]
    omitidas = [f"{f['codigo']} ({f['motivo']})" for f in resultado['etiqueta']['H'] if not f['incluida']]
    if omitidas:
        avisos.append('Frases H que no van en la etiqueta: ' + '; '.join(omitidas) + '.')
    peligroso = bool(frases_h)
    tipo = 'empaque pequeño (R10)' if envase_pequeno else 'producto peligroso' if peligroso else 'producto no peligroso (R07)'
    reglas.append('R10' if envase_pequeno else 'R08' if peligroso else 'R07')

    # --- Identificación del producto: igual a la de la FDS (R08) ---
    producto = resultado['producto']
    if not producto:
        faltantes.append('La identificación del producto: la ficha no la dio con claridad.')

    # --- Proveedor: solo entra si el usuario lo da. El de la sección 1 de la ficha queda como sugerencia, no se pone solo. ---
    dado = {c: (str(v).strip() or None) if v is not None else None for c, v in (e.get('proveedor') or {}).items()}
    proveedor = {c: dado.get(c) for c in ('nombre', 'direccion', 'telefono')}
    nombres = {'nombre': 'el nombre', 'direccion': 'la dirección', 'telefono': 'el teléfono'}
    faltan = [nombres[c] for c in nombres if not proveedor[c]]
    proveedor['incluido'] = any(proveedor[c] for c in nombres)
    proveedor['sugerido_por_la_ficha'] = {c: resultado['proveedor'].get(c) for c in ('nombre', 'direccion', 'telefono', 'estado')}
    if faltan and not envase_pequeno:
        norma = 'La Resolución 773 de 2021, art. 7.2 (regla R08), lista nombre, dirección y teléfono del proveedor en el contenido mínimo.'
        que = 'sin proveedor' if not proveedor['incluido'] else f"sin {' ni '.join(faltan).replace('el ', '').replace('la ', '')} del proveedor"
        if ajustes(tablas)['proveedor_obligatorio']:
            faltantes.append(f'Falta {" y ".join(faltan) if len(faltan) < 3 else "el proveedor (nombre, dirección y teléfono)"} del proveedor. {norma}'
                             if proveedor['incluido'] else f'Falta el proveedor (nombre, dirección y teléfono). {norma}')
        else:
            avisos.append(f'La etiqueta va {que}, por decisión registrada en configuracion.json. {norma}')

    # --- Pictogramas y borde (R09) ---
    pictogramas = [dict(codigo=p['codigo'], simbolo=p['simbolo']) for p in resultado['pictogramas']['pictogramas']]
    borde = e.get('borde') or 'rojo'
    if borde not in ('rojo', 'negro'):
        faltantes.append(f'El borde de los pictogramas debe ser «rojo» o «negro»; se recibió «{borde}».')
        borde = 'rojo'
    if borde == 'negro':
        reglas.append('R09')
        if not e.get('uso_interno'):
            faltantes.append('El borde negro solo se admite si el producto no sale del lugar de trabajo (regla R09). '
                             'Marque «uso interno» o use borde rojo.')
    if peligroso and not pictogramas:
        avisos.append('Con las frases H de esta ficha la matriz no asigna ningún pictograma: la etiqueta va sin pictogramas.')

    # --- Palabra de advertencia ---
    palabra = resultado['palabra_de_advertencia']['palabra']
    if peligroso and not palabra:
        avisos.append('Con las frases H de esta ficha la matriz no asigna palabra de advertencia: la etiqueta va sin ella.')

    # --- Consejos de prudencia: los marca el responsable del etiquetado (C05) ---
    disponibles = {f['codigo']: f for f in resultado['etiqueta']['P']}
    quedan, repetidos = sin_sueltos(list(disponibles))   # un consejo suelto no se repite si ya va dentro de uno combinado (regla C05)
    if repetidos:
        disponibles = {c: disponibles[c] for c in quedan}
        avisos.append('Consejos que no se repiten porque ya van dentro de uno combinado: ' + ', '.join(repetidos) + '.')
    if e.get('consejos_p') is None:
        # Por defecto, los de la selección para minería (reglas C05 y C07); los demás quedan disponibles para marcarlos
        elegidos = [x['codigo'] for x in (resultado.get('seleccion_p') or {}).get('elegidos', [])]
        marcados = [c for c in elegidos if c in disponibles] or [c for c, f in disponibles.items() if f['incluida']]
    else:
        marcados = [c for c in dict.fromkeys(e['consejos_p']) if c not in repetidos]
        ajenos = [c for c in marcados if c not in disponibles]
        if ajenos:
            faltantes.append(f"Consejos P marcados que la ficha no trae: {', '.join(ajenos)}. Solo se pueden marcar los de la ficha.")
            marcados = [c for c in marcados if c in disponibles]
    textos = e.get('textos') or {}
    consejos_p = []
    for codigo in sorted(marcados, key=lambda c: (disponibles[c]['prioridad'], c)):
        f = disponibles[codigo]
        texto = (textos.get(codigo) or '').strip() or f['texto']
        if not texto:
            faltantes.append(f'{codigo}: no tiene texto en español en las tablas; escríbalo o desmárquelo.')
            continue
        pendiente = bool(SIN_COMPLETAR.search(texto))
        if pendiente:
            avisos.append(f"{codigo} trae «...» o un texto entre corchetes sin completar. La ficha dice: «{f['texto_ficha']}».")
        consejos_p.append(dict(codigo=codigo, tipo=f['tipo'], texto=texto, texto_oficial=f['texto'], texto_ficha=f['texto_ficha'],
                               completado_por_el_usuario=codigo in textos and bool(textos[codigo].strip()), sin_completar=pendiente))
    if disponibles:
        reglas.append('C05')
    if not consejos_p and not envase_pequeno:
        if disponibles:
            faltantes.append('Ningún consejo de prudencia marcado: la etiqueta debe llevar consejos de prudencia (reglas R07 y R08).')
        else:
            # Decisión del autor: la herramienta no inventa consejos. Si la ficha no los trae, la etiqueta va sin ellos.
            avisos.append('La ficha no trae consejos de prudencia: la etiqueta va sin ellos. La herramienta no los inventa '
                          '(las reglas R07 y R08 los incluyen en el contenido de la etiqueta).')
    if resultado['por_revisar']['frases_h'] or resultado['por_revisar']['consejos_p']:
        avisos.append('Hay frases de la ficha «por revisar» que no entraron a la etiqueta: '
                      + '; '.join(resultado['por_revisar']['frases_h'] + resultado['por_revisar']['consejos_p']))

    # --- Empaque pequeño: lo obligatorio es el nombre del producto y los pictogramas (R10) ---
    if envase_pequeno:
        avisos.append(tamano['notas'][0] + ' Lo demás es opcional en este empaque.')

    return dict(archivo=resultado['archivo'], tipo=tipo, producto=producto, proveedor=proveedor, pictogramas=pictogramas, borde=borde,
                palabra=palabra, frases_h=frases_h, consejos_p=consejos_p, componentes=(e.get('componentes') or '').strip() or None,
                cantidad=e.get('cantidad'), lote=e.get('lote'), uso_interno=bool(e.get('uso_interno')), tamano=tamano, capacidad=capacidad,
                solo_contenido_minimo=envase_pequeno, faltantes=faltantes, avisos=avisos, reglas=reglas,
                lista_para_imprimir=not faltantes)


def cargar_resultado(prefijo):
    """Resultado del motor ya guardado, por el inicio del nombre de la ficha."""
    rutas = sorted(glob.glob(os.path.join(SALIDAS, prefijo + '*.json')))
    if not rutas:
        raise SystemExit(f'No hay resultado del motor que empiece por «{prefijo}». Ejecute antes: py "clasificador SGA/validacion.py"')
    with open(rutas[0], encoding='utf-8') as f:
        return json.load(f)


if __name__ == '__main__':
    sys.stdout.reconfigure(encoding='utf-8')
    argumentos = sys.argv[1:]
    elecciones = {}
    if '--elecciones' in argumentos:
        with open(argumentos[argumentos.index('--elecciones') + 1], encoding='utf-8') as f:
            elecciones = {k: v for k, v in json.load(f).items() if not k.startswith('_')}
    if '--envase' in argumentos:
        elecciones['envase_litros'] = float(argumentos[argumentos.index('--envase') + 1].replace(',', '.'))
    elecciones.update(uso_interno='--interno' in argumentos or elecciones.get('uso_interno', False))
    if '--negro' in argumentos:
        elecciones['borde'] = 'negro'
    c = contenido_de_etiqueta(cargar_resultado(argumentos[0]), elecciones, cargar_tablas())
    print(f"{c['archivo']}  ->  {c['tipo']}   (reglas: {', '.join(c['reglas'])})")
    print(f"LISTA PARA IMPRIMIR: {'sí' if c['lista_para_imprimir'] else 'NO'}")
    if c['tamano']:
        print(f"Tamaño: {c['tamano']['etiqueta']} | pictograma: {c['tamano']['pictograma']}   (empaque de {c['capacidad']})")
    print(f"Producto: {c['producto']}")
    p = c['proveedor']
    print(f"Proveedor: {p['nombre']} | {p['direccion']} | {p['telefono']}" if p['incluido']
          else f"Proveedor: no se incluye   (la ficha sugiere: {p['sugerido_por_la_ficha']['nombre'] or 'nada'})")
    print(f"Pictogramas (borde {c['borde']}): {', '.join(x['codigo'] for x in c['pictogramas']) or 'ninguno'}")
    print(f"Palabra de advertencia: {c['palabra'] or 'ninguna'}")
    print('Indicaciones de peligro:')
    for f in c['frases_h']:
        print(f"   {f['codigo']}  {f['texto']}")
    print(f"Consejos de prudencia ({len(c['consejos_p'])}):")
    for f in c['consejos_p']:
        print(f"   {f['codigo']:<17}{'*' if f['sin_completar'] else ' '} {f['texto']}" + ('   [completado por el usuario]' if f['completado_por_el_usuario'] else ''))
    for campo in ('componentes', 'cantidad', 'lote'):
        if c[campo]:
            print(f'{campo.capitalize()}: {c[campo]}')
    for x in c['faltantes']:
        print('FALTA:', x)
    for x in c['avisos']:
        print('Aviso:', x)
