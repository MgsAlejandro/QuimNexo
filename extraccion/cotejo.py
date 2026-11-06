# -*- coding: utf-8 -*-
"""Paso 9 de la extracción: cotejo de lo extraído contra las tablas del proyecto.
Compara, ficha por ficha:
  - frases H y P de la ficha  vs  hoja 'reaccio' de matriz/fichas_de_datos_de_seguridad.xlsx
  - límites ACGIH de la ficha vs  hoja 'TLV' de tlv/tlv_limites.xlsx
  - incompatibilidades de la ficha vs columna de incompatibilidades de 'reaccio' (lado a lado)
Solo LISTA las diferencias en extraccion/salidas/cotejo_fichas.xlsx. No corrige ninguna tabla
ni cambia lo extraído: decidir cuál es el dato correcto le corresponde al autor.
La ficha se empareja con su fila por número CAS.
Uso:  py extraccion/cotejo.py FDS
Requiere: pip install pdfplumber openpyxl"""
import glob, os, re, sys
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
from identificacion import CAS, FICHAS
from frases import CODIGO
from extraer import Recursos, extraer_ficha, guardar, SALIDAS

TLV = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'tlv', 'tlv_limites.xlsx')
# Cómo escribe cada ficha el tipo de límite que la tabla llama TLV-TWA, TLV-STEL y TLV-C (misma cosa, otra notación)
TIPOS = {'TLV-TWA': {'TWA'}, 'TLV-STEL': {'STEL', 'ST'}, 'TLV-C': {'CEILING', 'CEIL', 'C'}}


def _codigos(texto):
    """Códigos H o P escritos en una celda del Excel: {código base: como está escrito}."""
    salida = {}
    for m in CODIGO.finditer(str(texto or '')):
        for c in re.sub(r'\s+', '', m.group(1)).split('+'):
            salida.setdefault(c[:4], c)
    return salida


def cargar_reaccio():
    hoja = openpyxl.load_workbook(FICHAS, read_only=True, data_only=True)['reaccio']
    filas = []
    for f in hoja.iter_rows(min_row=2, values_only=True):
        if f[0] and f[1]:
            filas.append(dict(sustancia=f[0], cas=[m.group(0) for m in CAS.finditer(str(f[1]))], cas_literal=str(f[1]),
                              h_literal=str(f[3] or ''), H=_codigos(f[3]), P=_codigos(f[4]), incompatibilidades=f[7]))
    return filas


def cargar_tlv():
    filas = list(openpyxl.load_workbook(TLV, read_only=True, data_only=True)['TLV'].iter_rows(values_only=True))
    col = {nombre: i for i, nombre in enumerate(filas[0])}
    return [dict(id=f[col['ID']], sustancia=f[col['Sustancia']], cas=[m.group(0) for m in CAS.finditer(str(f[col['CAS']] or ''))],
                 agente=f[col['Agente que se mide']], cas_agente=[m.group(0) for m in CAS.finditer(str(f[col['CAS del agente']] or ''))],
                 tipo=f[col['Tipo de límite']], valor=f[col['Valor']], unidad=f[col['Unidad']], usar=f[col['Usar para Ir']])
            for f in filas[1:] if f[col['ID']]]


def _unidad(u):
    return str(u or '').replace('³', '3').replace(' ', '').replace('μ', 'µ').lower()


def _numero(v):
    try:
        return float(str(v).replace(',', '.'))
    except ValueError:
        return None


def cotejar(r, reaccio, tlv):
    """Compara el resultado de una ficha con sus filas de 'reaccio' y de 'TLV'."""
    ident = r['identificacion']
    del_archivo = r['cotejo_cas']['archivo']
    propios = [del_archivo] if del_archivo else [c for c in r['cotejo_cas']['ficha']]
    fila = next((f for f in reaccio if set(f['cas']) & set(propios)), None)
    c = {'archivo': r['archivo'], 'producto': ident['nombre']['valor'], 'fila_reaccio': fila['sustancia'] if fila else None,
         'fila_compartida': bool(fila and len(fila['cas']) > 1)}

    # --- Frases H y P ---
    ficha_h = {f['codigo'][:4] if '+' not in f['codigo'] else f['codigo']: f for f in r['frases']['H'] if f['estado'] == 'encontrado'}
    ficha_p = {p for f in r['frases']['P'] if f['estado'] == 'encontrado' for p in f['codigo'].split('+')}
    por_revisar = [f for f in r['frases']['H'] + r['frases']['P'] if f['estado'] == 'por revisar']
    excel_h, excel_p = (fila['H'], fila['P']) if fila else ({}, {})
    c['H'] = {'coinciden': sorted(set(ficha_h) & set(excel_h)), 'solo_ficha': sorted(set(ficha_h) - set(excel_h)),
              'solo_excel': sorted(set(excel_h) - set(ficha_h)), 'excel_literal': fila['h_literal'] if fila else None,
              'por_revisar': [f"«{f['texto_ficha']}»" + (f" (candidato {f['candidatos'][0][0]})" if f.get('candidatos') else f" ({f['codigo']})")
                              for f in por_revisar if f in r['frases']['H']]}
    c['P'] = {'excel': sorted(excel_p), 'faltan_en_ficha': sorted(set(excel_p) - ficha_p), 'en_ficha': len(ficha_p)}

    # --- Límites: tabla TLV frente a lo que la ficha declara como ACGIH ---
    filas_tlv = [t for t in tlv if set(t['cas']) & set(propios)]
    acgih = [l for l in r['limites']['limites'] if l['organismo'] and 'ACGIH' in l['organismo'].upper()]
    agentes = {l['agente'] for l in r['limites']['limites']}
    usados, c['limites'] = set(), []
    for t in filas_tlv:
        entrada = dict(id=t['id'], agente=t['agente'], tipo=t['tipo'], valor=t['valor'], unidad=t['unidad'], usar=t['usar'], en_ficha=None)
        if t['tipo'] not in TIPOS:
            entrada['resultado'] = 'Sin valor en la tabla' if t['valor'] is None else 'No es un TLV: no se coteja con la ficha'
        else:
            # mismo agente: por CAS si la ficha lo escribe; si no, cuando la tabla de la ficha tiene un solo agente
            iguales = [(i, l) for i, l in enumerate(acgih) if (l['tipo'] or '').upper() in TIPOS[t['tipo']] and _unidad(l['unidad']) == _unidad(t['unidad'])
                       and (l['cas'] in t['cas_agente'] if l['cas'] else len(agentes) == 1)]
            if not iguales:
                otros = [l for l in acgih if (l['tipo'] or '').upper() in TIPOS[t['tipo']] and (l['cas'] in t['cas_agente'] if l['cas'] else len(agentes) == 1)]
                entrada['en_ficha'] = '; '.join(f"{l['valor']} {l['unidad']}" for l in otros) or None
                entrada['resultado'] = 'Otra unidad en la ficha: comparar a mano' if otros else 'La ficha no declara este límite ACGIH'
            else:
                usados.update(i for i, _ in iguales)
                entrada['en_ficha'] = '; '.join(f"{l['valor']} {l['unidad']}" + (f" {l['nota']}" if l['nota'] else '') for _, l in iguales)
                entrada['resultado'] = 'Coincide' if any(_numero(l['valor']) == _numero(t['valor']) for _, l in iguales) else 'DIFIERE'
        c['limites'].append(entrada)
    c['limites_solo_ficha'] = [f"{l['agente'] or '?'} ({l['cas'] or 'sin CAS'}): {l['tipo']} {l['valor']} {l['unidad']}"
                               for i, l in enumerate(acgih) if i not in usados]
    c['sin_fila_tlv'] = not filas_tlv

    # --- Incompatibilidades, lado a lado ---
    ap = r['reacciones']['apartados']
    c['incompatibilidades'] = {'excel': fila['incompatibilidades'] if fila else None,
                               **{k: (ap[k]['texto'] if k in ap else None)
                                  for k in ('materiales_incompatibles', 'condiciones_a_evitar', 'reacciones_peligrosas')}}
    return c


def escribir(cotejos, resultados, destino):
    AR = 'Arial'
    wb = openpyxl.Workbook()
    rojo, ambar = PatternFill('solid', fgColor='FF7C80'), PatternFill('solid', fgColor='FFE699')

    def hoja(nombre, encabezado, filas, anchos, primera=False):
        ws = wb.active if primera else wb.create_sheet()
        ws.title = nombre
        ws.append(encabezado)
        for f in filas:
            ws.append([', '.join(x) if isinstance(x, (list, tuple)) else x for x in f])
        for fila in ws.iter_rows():
            for celda in fila:
                celda.font = Font(name=AR, size=10)
                celda.alignment = Alignment(wrap_text=True, vertical='top')
        for celda in ws[1]:
            celda.font = Font(name=AR, bold=True, color='FFFFFF')
            celda.fill = PatternFill('solid', fgColor='1F3864')
        for i, a in enumerate(anchos, 1):
            ws.column_dimensions[get_column_letter(i)].width = a
        ws.freeze_panes = 'B2'
        return ws

    hoja('LEEME', ('Tema', 'Detalle'), [
        ('Qué es', 'Cotejo entre lo que el módulo de extracción leyó de cada FDS y las tablas del proyecto. Lo genera extraccion/cotejo.py; se puede regenerar cuando cambien las fichas o las tablas.'),
        ('Qué NO hace', 'No corrige nada. No cambia el Excel de fichas, ni la tabla de límites, ni lo extraído. Solo pone las diferencias a la vista para que el autor decida.'),
        ('Cómo se empareja', 'Cada ficha se busca en las tablas por el número CAS de su nombre de archivo (o, si no lo trae, por los CAS que declara la ficha). Si la fila del Excel cubre dos sustancias (cemento y óxido de calcio; acetileno y oxígeno), sus frases H se comparan contra las dos fichas y habrá códigos "solo en Excel" que pertenecen a la otra sustancia.'),
        ('Hoja Resumen', 'Una fila por ficha.'),
        ('Hoja Frases_H', 'Código por código: si está en la ficha, si está en la hoja "reaccio", y el resultado.'),
        ('Hoja Frases_P', 'Los consejos P "clave" de la hoja "reaccio" que no aparecen en la ficha. La ficha suele traer más consejos que los clave; eso no es diferencia.'),
        ('Hoja Limites', 'Cada fila de tlv_limites.xlsx frente al límite ACGIH que declara la ficha. El índice Ir se calcula SIEMPRE con tlv_limites.xlsx; el valor de la ficha es informativo.'),
        ('Hoja Incompatibilidades', 'Texto del Excel junto al texto de la sección 10 de la ficha. Están en idiomas distintos, así que no se califica: es para leer lado a lado.'),
        ('Hoja Hallazgos', 'Todo lo que la extracción encontró y conserva: producto distinto al del Excel, otros componentes, frases por revisar, frases de otros sistemas, apartados que la ficha no informa. Pueden pesar en el índice de riesgo y en las reacciones con otras sustancias.'),
        ('Colores', 'Rojo: diferencia que conviene revisar. Ámbar: dato por revisar o que no se pudo comparar.'),
    ], (26, 130), primera=True)

    ws = hoja('Resumen', ('Ficha', 'Producto según la ficha', 'Fila en "reaccio"', 'H que coinciden', 'H solo en la ficha', 'H solo en el Excel',
                          'H por revisar', 'P clave del Excel que faltan en la ficha', 'Límites: coinciden', 'Límites: difieren',
                          'Límites ACGIH que la ficha no declara', 'Hallazgos'),
              [(c['archivo'], c['producto'], (c['fila_reaccio'] or 'SIN FILA (no hay CAS en común)') + (' — fila compartida por dos sustancias' if c['fila_compartida'] else ''),
                c['H']['coinciden'], c['H']['solo_ficha'], c['H']['solo_excel'], len(c['H']['por_revisar']), c['P']['faltan_en_ficha'],
                sum(l['resultado'] == 'Coincide' for l in c['limites']), sum(l['resultado'] == 'DIFIERE' for l in c['limites']),
                sum(l['resultado'].startswith('La ficha no declara') for l in c['limites']), len(r['hallazgos']))
               for c, r in zip(cotejos, resultados)], (34, 34, 34, 26, 22, 22, 9, 24, 10, 10, 14, 10))
    for fila in ws.iter_rows(min_row=2):
        for i in (4, 5):
            if fila[i].value:
                fila[i].fill = rojo
        if fila[9].value:
            fila[9].fill = rojo
        if fila[6].value or str(fila[2].value).startswith('SIN FILA'):
            fila[6].fill = ambar

    filas_h = []
    for c, r in zip(cotejos, resultados):
        en_ficha = {(f['codigo'][:4] if '+' not in f['codigo'] else f['codigo']): f for f in r['frases']['H'] if f['estado'] == 'encontrado'}
        for cod in sorted(set(en_ficha) | set(c['H']['solo_excel'])):
            f = en_ficha.get(cod)
            resultado = 'Coincide' if cod in c['H']['coinciden'] else 'Solo en la ficha' if f else 'Solo en el Excel'
            filas_h.append((c['archivo'], cod, 'Sí' if f else 'No', r['frases']['metodo']['H'] if f else None, f['texto_ficha'] if f else None,
                            f['pagina'] if f else None, 'Sí' if cod in c['H']['coinciden'] or not f else 'No', c['H']['excel_literal'], resultado))
        for texto in c['H']['por_revisar']:
            filas_h.append((c['archivo'], '?', 'Por revisar', None, texto, None, None, c['H']['excel_literal'], 'Por revisar'))
    ws = hoja('Frases_H', ('Ficha', 'Código H', 'En la ficha', 'Cómo se obtuvo', 'Texto en la ficha', 'Página', 'En "reaccio"',
                           'Celda de frases H en "reaccio"', 'Resultado'), filas_h, (34, 10, 11, 12, 50, 8, 11, 60, 18))
    for fila in ws.iter_rows(min_row=2):
        fila[8].fill = rojo if str(fila[8].value).startswith('Solo') else ambar if fila[8].value == 'Por revisar' else PatternFill()

    hoja('Frases_P', ('Ficha', 'P clave en "reaccio"', 'P clave que no están en la ficha', 'Códigos P distintos en la ficha'),
         [(c['archivo'], c['P']['excel'], c['P']['faltan_en_ficha'], c['P']['en_ficha']) for c in cotejos], (34, 40, 40, 14))

    filas_l = []
    for c in cotejos:
        for l in c['limites']:
            filas_l.append((c['archivo'], l['id'], l['agente'], l['tipo'], l['valor'], l['unidad'], l['usar'], l['en_ficha'], l['resultado']))
        for s in c['limites_solo_ficha']:
            filas_l.append((c['archivo'], None, None, None, None, None, None, s, 'Límite ACGIH de la ficha que no está en la tabla'))
        if c['sin_fila_tlv']:
            filas_l.append((c['archivo'], None, None, None, None, None, None, None, 'SIN FILA en tlv_limites.xlsx (no hay CAS en común)'))
    ws = hoja('Limites', ('Ficha', 'ID en la tabla', 'Agente que se mide', 'Tipo', 'Valor en la tabla', 'Unidad', 'Usar para Ir',
                          'ACGIH declarado en la ficha', 'Resultado'), filas_l, (34, 9, 40, 14, 10, 9, 12, 44, 44))
    for fila in ws.iter_rows(min_row=2):
        v = str(fila[8].value)
        fila[8].fill = rojo if v == 'DIFIERE' else ambar if v != 'Coincide' and not v.startswith(('Sin valor', 'No es un TLV')) else PatternFill()

    hoja('Incompatibilidades', ('Ficha', 'Incompatibilidades / reacciones en "reaccio"', 'Ficha: materiales incompatibles',
                                'Ficha: condiciones que deben evitarse', 'Ficha: posibilidad de reacciones peligrosas'),
         [(c['archivo'], c['incompatibilidades']['excel'], c['incompatibilidades']['materiales_incompatibles'],
           c['incompatibilidades']['condiciones_a_evitar'], c['incompatibilidades']['reacciones_peligrosas']) for c in cotejos],
         (34, 60, 50, 50, 50))

    hoja('Hallazgos', ('Ficha', 'Paso', 'Hallazgo'),
         [(r['archivo'], h['paso'], h['texto']) for r in resultados for h in r['hallazgos']], (34, 16, 140))
    wb.save(destino)


if __name__ == '__main__':
    sys.stdout.reconfigure(encoding='utf-8')
    recursos, reaccio, tlv = Recursos(), cargar_reaccio(), cargar_tlv()
    rutas = sorted(glob.glob(os.path.join(sys.argv[1], '*.pdf')))
    resultados = [extraer_ficha(ruta, recursos) for ruta in rutas]
    for r in resultados:
        guardar(r)
    cotejos = [cotejar(r, reaccio, tlv) for r in resultados]
    destino = os.path.join(SALIDAS, 'cotejo_fichas.xlsx')
    escribir(cotejos, resultados, destino)
    print(f"{'ficha':<24}{'H =':>4}{'H ficha':>8}{'H Excel':>8}  {'límites':<22} diferencias")
    for c in cotejos:
        lim = ', '.join(f"{n} {k}" for k, n in (('coinc.', sum(l['resultado'] == 'Coincide' for l in c['limites'])),
                                                ('DIF.', sum(l['resultado'] == 'DIFIERE' for l in c['limites'])),
                                                ('no decl.', sum(l['resultado'].startswith('La ficha no') for l in c['limites']))) if n)
        dif = []
        if not c['fila_reaccio']:
            dif.append('SIN FILA en reaccio')
        if c['H']['solo_ficha']:
            dif.append('solo ficha: ' + ', '.join(c['H']['solo_ficha']))
        if c['H']['solo_excel']:
            dif.append('solo Excel: ' + ', '.join(c['H']['solo_excel']))
        print(f"{c['archivo'][:23]:<24}{len(c['H']['coinciden']):>4}{len(c['H']['solo_ficha']):>8}{len(c['H']['solo_excel']):>8}  {lim or '—':<22} {' | '.join(dif)}")
    print(f'Cotejo guardado en {destino}')
