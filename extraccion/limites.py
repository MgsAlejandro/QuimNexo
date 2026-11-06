# -*- coding: utf-8 -*-
"""Paso 7 de la extracción: límites de exposición que la FDS declara en la sección 8.
Son INFORMATIVOS. El índice Ir se calcula siempre con tlv/tlv_limites.xlsx (TLV de ACGIH);
lo que sale de aquí solo sirve para mostrarlo y para cotejarlo contra esa tabla.
Las tablas se leen por posición en la página, porque en texto corrido se mezclan las columnas.
Cada valor se acepta solo si aparece literal en el texto de la sección 8.
Uso de prueba:
    py extraccion/limites.py FDS                          (resumen de todas)
    py extraccion/limites.py FDS/14_Acido_sulfurico_7664-93-9.pdf   (detalle)
Requiere: pip install pdfplumber openpyxl"""
import glob, os, re, sys
import pdfplumber
from pdfplumber.utils.exceptions import PdfminerException
from leer_pdf import leer_pdf
from secciones import cortar_secciones, cargar_titulos_oficiales
from identificacion import CAS, cas_valido

ENC, REV = 'encontrado', 'por revisar'
UNIDAD = r'(?:mg/m\s?[3³]|[µμu]g/m\s?[3³]|ppm|f/cc|mppcf)'
VALOR = re.compile(r'(?<![\w.,])(\d+(?:[.,]\d+)?)\s*(' + UNIDAD + r')', re.I)
# Tipos de límite tal como los escriben las fichas (no se traducen ni se convierten)
TIPO = re.compile(r'(?<![A-Za-z])(TWA|STEL|ST|Ceiling|CEIL|IDLH|REL|PEL|TLV|WEEL|DNEL|VLA-ED|VLA-EC|LMPE-PPT|LMPE-CT)(?![A-Za-z])', re.I)
ORGANISMO = re.compile(r'ACGIH|OSHA|NIOSH|AIHA|MSHA|CMRG|IDLH|Mexico|Vietnam|United Kingdom|European Union|INSST|DFG|MAK', re.I)
ENCABEZADO_AGENTE = re.compile(r'^(Components?|Chemical name|Ingredientes?|Componentes?|Substance|Sustancia)\b', re.I)
ENCABEZADO_LARGA = re.compile(r'Basis|Agencia|Value type|Tipo de l[íi]mite|CAS', re.I)
ASFIXIANTE = re.compile(r'simple asphyxiant|asfixiante simple', re.I)


def _plano(celda):
    return ' '.join((celda or '').split())


def _valores(celda, tipo_fila=None):
    """Valores de una celda: [(tipo, número literal, unidad, nota, texto literal)].
    El tipo es el último escrito antes del valor; si la celda no lo trae, el de la fila.
    La nota junta lo que acompaña a ese valor: '(Vacated)' antes del tipo, la fracción entre
    el tipo y el número, y la aclaración que le sigue ('total dust', '(TWA)', '15 min')."""
    texto = _plano(celda)
    marcas = sorted([(m.start(), m.end(), 'T', m) for m in TIPO.finditer(texto)] +
                    [(m.start(), m.end(), 'V', m) for m in VALOR.finditer(texto)])
    salida, tipo, prefijo, previo = [], tipo_fila, '', 0

    def anotar(extra):
        extra = extra.strip(' :=;,/')
        if extra and salida:
            salida[-1][3] = _plano(f"{salida[-1][3] or ''} {extra}")

    for inicio, fin, clase, m in marcas:
        hueco = texto[previo:inicio]
        if clase == 'T':
            if salida and hueco.strip() == '(' and texto[fin:fin + 1] == ')':   # "1 mg/m3 (TWA)"
                anotar(f'({m.group(1)})')
                previo = fin + 1
                continue
            antes = re.search(r'(\([^()]*\))\s*$', hueco)    # "(Vacated) TWA": va con el tipo que sigue
            anotar(hueco[:antes.start()] if antes else hueco)
            tipo, prefijo = m.group(1), (antes.group(1) if antes else '')
        else:
            if salida and not prefijo and salida[-1][5] == tipo and hueco.strip(' :=;,/'):
                anotar(hueco)                                  # aclaración del valor anterior, mismo tipo
                hueco = ''
            salida.append([tipo, m.group(1), m.group(2).replace(' ', ''), _plano(f"{prefijo} {hueco.strip(' :=;,/')}") or None,
                           None, tipo])
        previo = fin
    anotar(texto[previo:])
    return [(t, n, u, nota, f"{t or ''} {n} {u}".strip()) for t, n, u, nota, _, _ in salida]


def _agente_y_cas(texto):
    m = next((m for m in CAS.finditer(texto) if cas_valido(m.group(0))), None)
    nombre = re.sub(r'\(\s*[<>=\d\s.%-]*\)', '', texto.replace(m.group(0), '') if m else texto)
    return _plano(nombre).strip(' ()-,') or None, (m.group(0) if m else None)


def _tabla_ancha(tabla, pagina):
    """Un organismo por columna (Component | ACGIH TLV | OSHA PEL | ...)."""
    encabezado = [_plano(c) for c in tabla[0]]
    registros = []
    for fila in tabla[1:]:
        agente, cas = _agente_y_cas(_plano(fila[0]))
        for organismo, celda in zip(encabezado[1:], fila[1:]):
            tipo_columna = (TIPO.findall(organismo) or [None])[-1]
            for tipo, numero, unidad, nota, literal in _valores(celda, tipo_columna):
                registros.append(dict(agente=agente, cas=cas, organismo=organismo, tipo=tipo, valor=numero, unidad=unidad,
                                      nota=nota, cita=f'{organismo}: {_plano(celda)}', pagina=pagina))
    return registros


def _tabla_larga(tabla, pagina, actual):
    """Una fila por límite. Las celdas se reconocen por su contenido (CAS, valor, organismo, tipo), porque
    en estas tablas las columnas se corren. 'actual' recuerda el agente cuando la fila no lo repite."""
    registros = []
    for fila in tabla:
        celdas = [_plano(c) for c in fila if _plano(c)]
        if not celdas:
            continue
        con_valor = [c for c in celdas if VALOR.search(c)]
        if not con_valor:
            if ENCABEZADO_AGENTE.match(celdas[0]) or ENCABEZADO_LARGA.search(' '.join(celdas)):
                continue                                   # fila de encabezados
            if len(celdas) == 1 or ASFIXIANTE.search(' '.join(celdas)) is None and any(cas_valido(m.group(0)) for m in CAS.finditer(celdas[0])):
                actual['agente'], cas = _agente_y_cas(celdas[0])  # fila que solo nombra la sustancia
                actual['cas'] = cas
            continue
        cas = next((m.group(0) for c in celdas for m in CAS.finditer(c) if cas_valido(m.group(0)) and not VALOR.search(c)), None)
        primera = celdas[0]
        if not (VALOR.search(primera) or CAS.fullmatch(primera) or TIPO.match(primera) or ORGANISMO.match(primera)):
            actual['agente'], actual['cas'] = _agente_y_cas(primera)[0], cas
        elif cas:
            actual['cas'] = cas
        otras = [c for c in celdas if c not in con_valor and c != actual.get('agente') and not CAS.fullmatch(c)]
        organismo = next((c for c in otras if ORGANISMO.search(c)), None)
        tipo_fila = next((TIPO.search(c).group(1) for c in otras if TIPO.search(c)), None)
        for celda in con_valor:
            for tipo, numero, unidad, nota, literal in _valores(celda, tipo_fila):
                registros.append(dict(agente=actual.get('agente'), cas=actual.get('cas'), organismo=organismo, tipo=tipo, valor=numero,
                                      unidad=unidad, nota=nota, cita=' | '.join(celdas), pagina=pagina))
    return registros


def extraer_limites(doc, corte):
    """Devuelve {'limites': [{agente, cas, organismo, tipo, valor, unidad, nota, estado, pagina, cita, origen}],
                 'hallazgos': [frases de la sección 8 que se conservan], 'sin_limites': True/False}."""
    lineas = corte['secciones'][8]['lineas']
    texto_seccion = re.sub(r'\s+', '', ' '.join(l for _, l in lineas)).replace('³', '3')
    paginas = sorted({p for p, _ in lineas})
    registros = []
    try:
        with pdfplumber.open(doc['ruta']) as pdf:
            actual = {}
            for pagina in paginas:
                for tabla in pdf.pages[pagina - 1].extract_tables():
                    if not any(VALOR.search(_plano(c)) for fila in tabla for c in fila):
                        continue
                    primera = [_plano(c) for c in tabla[0]]
                    ancha = bool(ENCABEZADO_AGENTE.match(primera[0])) and not ENCABEZADO_LARGA.search(' '.join(primera[1:]))
                    registros += _tabla_ancha(tabla, pagina) if ancha else _tabla_larga(tabla, pagina, actual)
    except PdfminerException:
        pass  # PDF protegido: se sigue con el texto corrido, más abajo
    # Solo valen los valores que están, literales, en el texto de la sección 8 (descarta tablas de otras secciones)
    registros = [r for r in registros if (r['valor'] + r['unidad']).replace('³', '3').replace(' ', '') in texto_seccion]
    for r in registros:
        r['origen'] = 'tabla'
        r['estado'] = ENC if r['agente'] and r['organismo'] and r['tipo'] else REV

    if not registros:  # sin tabla legible: se toman los valores del texto corrido, todos por revisar
        for pagina, linea in lineas:
            for tipo, numero, unidad, nota, literal in _valores(linea):
                organismo = ORGANISMO.search(linea)
                registros.append(dict(agente=None, cas=None, organismo=organismo.group(0) if organismo else None, tipo=tipo, valor=numero,
                                      unidad=unidad, nota=nota, cita=linea, pagina=pagina, origen='texto corrido', estado=REV))

    hallazgos = [dict(texto_ficha=l.strip(), pagina=p) for p, l in lineas if ASFIXIANTE.search(l)]
    return {'limites': registros, 'hallazgos': hallazgos, 'sin_limites': not registros}


if __name__ == '__main__':
    sys.stdout.reconfigure(encoding='utf-8')
    oficiales = cargar_titulos_oficiales()
    destino = sys.argv[1]
    rutas = sorted(glob.glob(os.path.join(destino, '*.pdf'))) if os.path.isdir(destino) else [destino]
    if len(rutas) > 1:
        print(f"{'ficha':<24}{'límites':>8}{'revisar':>8}{'agentes':>8}  organismos | límites ACGIH declarados")
    for ruta in rutas:
        doc = leer_pdf(ruta)
        r = extraer_limites(doc, cortar_secciones(doc, oficiales))
        lim = r['limites']
        if len(rutas) > 1:
            organismos = sorted({re.split(r'\s', l['organismo'])[0] for l in lim if l['organismo']})
            acgih = [f"{l['agente'] or '?'} {l['tipo'] or '?'} {l['valor']} {l['unidad']}" for l in lim if l['organismo'] and 'ACGIH' in l['organismo']]
            print(f"{os.path.basename(ruta)[:23]:<24}{len(lim):>8}{sum(l['estado'] == REV for l in lim):>8}"
                  f"{len({l['agente'] for l in lim if l['agente']}):>8}  {', '.join(organismos) or '—'} | {'; '.join(acgih) or '—'}"
                  + (f" | hallazgo: {r['hallazgos'][0]['texto_ficha']}" if r['hallazgos'] else ''))
            continue
        print(f"--- Límites declarados en la sección 8: {len(lim)} ---")
        for l in lim:
            print(f"  p.{l['pagina']:<3}[{l['estado']}] {l['agente'] or '?'} ({l['cas'] or 'sin CAS'}) | {l['organismo'] or '?'} | "
                  f"{l['tipo'] or '?'} {l['valor']} {l['unidad']}" + (f" | nota: {l['nota']}" if l['nota'] else '') + f"   [{l['origen']}]")
        for h in r['hallazgos']:
            print(f"  hallazgo p.{h['pagina']}: {h['texto_ficha']}")
        if r['sin_limites']:
            print('  La sección 8 de esta ficha no declara valores límite.')
