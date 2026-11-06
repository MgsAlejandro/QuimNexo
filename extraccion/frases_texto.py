# -*- coding: utf-8 -*-
"""Paso 5 de la extracción: frases H y P que la FDS trae sin código.
Compara cada frase de la sección 2 con los textos oficiales del diccionario del SGA
(español e inglés) y le asigna el código solo si el parecido es alto y no hay rival.
Si hay duda, la deja 'por revisar' con los candidatos a la vista; no adivina.
Las frases que no se parecen a ninguna del SGA se conservan como hallazgo.
Uso de prueba:
    py extraccion/frases_texto.py FDS                     (resumen de las fichas sin código)
    py extraccion/frases_texto.py FDS/14_Acido_sulfurico_7664-93-9.pdf   (detalle)
    py extraccion/frases_texto.py --examen FDS            (mide el acierto con las fichas que sí traen código)
Requiere: pip install pdfplumber openpyxl"""
import glob, os, re, sys, unicodedata
from difflib import SequenceMatcher
from functools import lru_cache
from leer_pdf import leer_pdf
from secciones import cortar_secciones, cargar_titulos_oficiales
from frases import cargar_diccionario, extraer_frases_con_codigo, CODIGO, OTRO_SISTEMA, ENC, REV

SIN = 'sin código SGA'
ACEPTAR = 0.85   # parecido mínimo para asignar el código sin revisión
MARGEN = 0.10    # ...y ningún otro código puede quedar a menos de esta distancia
PROPONER = 0.60  # parecido mínimo para proponerlo como candidato
INICIO = {'H': re.compile(r'^\s*(?:\d+(?:\.\d+)*\.?\s+)?(?:Hazard statements?|Indicaci[óo]n(?:es)? de peligro)\b[^:]*:?\s*(.*)$', re.I),
          'P': re.compile(r'^\s*(?:\d+(?:\.\d+)*\.?\s+)?(?:Precautionary statements?|Consejos? de prudencia)\b[^:]*:?\s*(.*)$', re.I)}
FIN = re.compile(r'^\s*(\d+\.\d+\.?\s|Hazards not otherwise|Hazards classified under|Other hazards|Otros peligros|'
                 r'Unknown acute|Other information|Supplemental|Additional|Signal word|Palabra de advertencia|WARNING\b)', re.I)
# La ficha declara que el producto no es peligroso o que no requiere etiqueta (se guarda literal)
NO_PELIGROSO = re.compile(r'not (considered|classified as) (a )?hazardous|contains no substances which|^\s*None required|'
                          r'no (se considera|est[áa] clasificad\w+ como|es) (un[a]? )?(producto |sustancia )?peligros\w+(?! para el transporte)', re.I)
# Subtítulos dentro de los bloques: no son frases
SUBTITULO = re.compile(r'^\s*[-–]?\s*(Prevention|Response|Storage|Disposal|General|Inhalation|Skin|Eyes?|Ingestion|'
                       r'Prevenci[óo]n|Respuesta|Intervenci[óo]n|Almacenamiento|Eliminaci[óo]n|Desecho)\s*:?\s*$', re.I)


def _palabras(texto):
    """Palabras comparables de una frase: sin tildes, sin signos y sin lo que va entre paréntesis
    (en el SGA son instrucciones para el fabricante; en la ficha, órganos o vías que este agrega)."""
    t = unicodedata.normalize('NFD', re.sub(r'\([^)]*\)', ' ', texto.lower()))
    t = ''.join(ch for ch in t if unicodedata.category(ch) != 'Mn')
    return tuple(re.findall(r'[a-z0-9]+', t))


@lru_cache(maxsize=None)
def _misma(a, b):
    """Dos palabras cuentan como la misma si son iguales o casi (vapours/vapors, plural/singular)."""
    return a == b or (min(len(a), len(b)) >= 5 and SequenceMatcher(None, a, b).ratio() >= 0.84)


def _en_comun(a, b):
    """Cuántas palabras comparten dos frases respetando el orden."""
    fila = [0] * (len(b) + 1)
    for x in a:
        previa = 0
        for j, y in enumerate(b, 1):
            actual = fila[j]
            fila[j] = previa + 1 if _misma(x, y) else max(fila[j], fila[j - 1])
            previa = actual
    return fila[-1]


def preparar(diccionario):
    """Índice para comparar: {'H': [(código, idioma, palabras, admite texto extra)], 'P': [...]}.
    'Admite texto extra' = el texto oficial trae '...' para que el fabricante lo complete."""
    indice = {}
    for tipo, tabla in diccionario.items():
        indice[tipo] = [(codigo, idioma, _palabras(texto), bool(re.search(r'\.\.\.|…', texto)))
                        for codigo, textos in tabla.items() for idioma, texto in zip(('es', 'en'), textos)]
    return indice


def candidatos(texto, tipo, indice):
    """Códigos del diccionario ordenados por parecido con la frase: [(parecido, palabras comunes, código, idioma)]."""
    ficha = _palabras(texto)
    if not ficha:
        return []
    mejores = {}
    for codigo, idioma, oficial, admite_extra in indice[tipo]:
        comunes = _en_comun(oficial, ficha)
        cubre_oficial, cubre_ficha = comunes / len(oficial), comunes / len(ficha)
        parecido = cubre_oficial if admite_extra else min(cubre_oficial, cubre_ficha)
        clave = (round(parecido, 2), comunes)
        if parecido >= PROPONER and clave > mejores.get(codigo, ((0, 0), None))[0]:
            mejores[codigo] = (clave, idioma)
    return sorted(((p, n, c, i) for c, ((p, n), i) in mejores.items()), reverse=True)


def asignar(texto, tipo, indice):
    """Decide el código de una frase. Devuelve (estado, código o None, candidatos)."""
    lista = candidatos(texto, tipo, indice)
    if not lista:
        return SIN, None, []
    parecido, comunes, codigo, _ = lista[0]
    # Rival: otro código casi igual de parecido, que no sea simplemente una parte de la combinación ganadora
    partes = set(codigo.split('+'))
    piso = parecido if parecido >= 0.995 else parecido - MARGEN  # una coincidencia exacta solo empata con otra exacta
    rival = any(p >= piso and n >= comunes and not set(c.split('+')) < partes for p, n, c, _ in lista[1:])
    if parecido >= ACEPTAR and not rival:
        return ENC, codigo, lista[:3]
    return REV, None, lista[:3]


def _bloques(lineas):
    """Frases candidatas de la sección 2: [(tipo, página, texto, cita)], tomadas de los bloques de
    indicaciones de peligro (H) y consejos de prudencia (P). Las líneas con código se saltan (paso 4)."""
    modo, salida = None, []
    for pag, linea in lineas:
        nuevo = next((t for t in 'HP' if INICIO[t].match(linea)), None)
        if nuevo:
            modo, texto = nuevo, INICIO[nuevo].match(linea).group(1).strip(' -–')
        elif FIN.match(linea):
            modo = None
            continue
        else:
            texto = linea.strip()
        if modo and texto and not SUBTITULO.match(texto) and not CODIGO.search(linea) and not OTRO_SISTEMA.match(texto):
            salida.append((modo, pag, texto, linea))
    return salida


def extraer_frases_sin_codigo(corte, diccionario, indice):
    """Devuelve {'H': [...], 'P': [...], 'sin_codigo': [...]}. Cada frase: codigo, texto_ficha, texto_sga,
    estado, parecido, candidatos, pagina, cita."""
    frases = _bloques(corte['secciones'][2]['lineas'])
    resultado, i = {'H': [], 'P': [], 'sin_codigo': []}, 0

    def registrar(tipo, texto, pag, cita):
        estado, codigo, lista = asignar(texto, tipo, indice)
        if estado == SIN:
            resultado['sin_codigo'].append(dict(bloque=tipo, texto_ficha=texto, pagina=pag, cita=cita))
            return
        idioma = 0 if lista[0][3] == 'es' else 1
        resultado[tipo].append(dict(
            codigo=codigo, texto_ficha=texto, estado=estado, parecido=lista[0][0], pagina=pag, cita=cita,
            texto_sga=diccionario[tipo][codigo][idioma] if codigo else None,
            candidatos=[(c, p) for p, _, c, _ in lista]))

    while i < len(frases):
        tipo, pag, texto, cita = frases[i]
        solo = asignar(texto, tipo, indice)
        # Frase partida en dos renglones: se une con el siguiente si así se parece más a una del SGA
        if i + 1 < len(frases) and frases[i + 1][0] == tipo:
            unido = f'{texto} {frases[i + 1][2]}'
            junto = asignar(unido, tipo, indice)
            if junto[0] == ENC and (solo[0] != ENC or asignar(frases[i + 1][2], tipo, indice)[0] != ENC):
                registrar(tipo, unido, pag, f'{cita} / {frases[i + 1][3]}')
                i += 2
                continue
        # Renglón con varias frases que el SGA no combina: se intenta frase por frase
        partes = [p.strip() for p in re.split(r'(?<=[.;])\s+', texto) if p.strip()]
        if solo[0] != ENC and len(partes) > 1 and all(asignar(p, tipo, indice)[0] == ENC for p in partes):
            for p in partes:
                registrar(tipo, p, pag, cita)
        else:
            registrar(tipo, texto, pag, cita)
        i += 1
    return resultado


def extraer_frases(corte, diccionario, indice):
    """Resultado único de los pasos 4 y 5 para una ficha. Por cada tipo (H, P): si la sección 2 trae
    códigos se usan esos (método 'código'); si no, se asignan por texto (método 'texto').
    Además conserva como hallazgo: frases sin código SGA, frases de otros sistemas, códigos de
    otras secciones y la declaración de producto no peligroso, si la ficha la hace."""
    con = extraer_frases_con_codigo(corte, diccionario)
    sin = extraer_frases_sin_codigo(corte, diccionario, indice)
    declaraciones = {l.strip(): p for p, l in reversed(corte['secciones'][2]['lineas']) if NO_PELIGROSO.search(l)}
    resultado = {'metodo': {}, 'otras_frases': con['otras_frases'], 'fuera_de_seccion_2': con['fuera_de_seccion_2'],
                 'sin_codigo': [], 'declaraciones': [dict(texto_ficha=t, pagina=p) for t, p in reversed(declaraciones.items())]}
    for tipo in 'HP':
        if con[tipo]:
            resultado[tipo], resultado['metodo'][tipo] = con[tipo], 'código'
        else:
            resultado[tipo], resultado['metodo'][tipo] = sin[tipo], 'texto' if sin[tipo] else 'ninguna'
            resultado['sin_codigo'] += [s for s in sin['sin_codigo'] if s['bloque'] == tipo
                                        and not NO_PELIGROSO.search(s['texto_ficha'])]
    return resultado


def examen(rutas, oficiales, diccionario, indice):
    """Con las fichas que traen código: se compara el código que da el texto con el que declara la ficha."""
    cuenta = {'acierta': 0, 'por revisar': 0, 'sin código': 0, 'ERROR': 0}
    detalle = []
    for ruta in rutas:
        r = extraer_frases_con_codigo(cortar_secciones(leer_pdf(ruta), oficiales), diccionario)
        for tipo in 'HP':
            for f in r[tipo]:
                if not f['texto_ficha'] or not f['codigo_base']:
                    continue
                estado, codigo, lista = asignar(f['texto_ficha'], tipo, indice)
                caso = {ENC: 'acierta' if codigo == f['codigo_base'] else 'ERROR', REV: 'por revisar', SIN: 'sin código'}[estado]
                cuenta[caso] += 1
                if caso != 'acierta':
                    propuestos = ', '.join(f'{c} {p:.2f}' for p, _, c, _ in lista) or '—'
                    detalle.append(f"  {caso:<12}{os.path.basename(ruta)[:14]:<15}{f['codigo']:<16}«{f['texto_ficha'][:70]}» -> {propuestos}")
    return cuenta, detalle


if __name__ == '__main__':
    sys.stdout.reconfigure(encoding='utf-8')
    oficiales, diccionario = cargar_titulos_oficiales(), cargar_diccionario()
    indice = preparar(diccionario)
    modo_examen = sys.argv[1] == '--examen'
    destino = sys.argv[-1]
    rutas = sorted(glob.glob(os.path.join(destino, '*.pdf'))) if os.path.isdir(destino) else [destino]
    if modo_examen:
        cuenta, detalle = examen(rutas, oficiales, diccionario, indice)
        total = sum(cuenta.values())
        print(f'Frases con código puestas a prueba: {total}')
        for caso, n in cuenta.items():
            print(f'  {caso:<12}{n:>4}  ({100 * n / total:.0f} %)')
        print('\n'.join(detalle))
    elif len(rutas) > 1:
        print('Método: cód = la ficha trae códigos · txt = asignados por texto · — = la ficha no trae frases')
        print(f"{'ficha':<24}{'mét.':>5}{'H':>3}{'P':>4}{'revisar':>8}{'sin cód.':>9}{'decl.':>6}  códigos H")
        for ruta in rutas:
            r = extraer_frases(cortar_secciones(leer_pdf(ruta), oficiales), diccionario, indice)
            todas = r['H'] + r['P']
            metodo = {'código': 'cód', 'texto': 'txt', 'ninguna': '—'}[r['metodo']['H']]
            print(f"{os.path.basename(ruta)[:23]:<24}{metodo:>5}{sum(f['estado'] == ENC for f in r['H']):>3}"
                  f"{sum(f['estado'] == ENC for f in r['P']):>4}{sum(f['estado'] == REV for f in todas):>8}"
                  f"{len(r['sin_codigo']):>9}{'sí' if r['declaraciones'] else '':>6}  "
                  + ', '.join(f['codigo'] for f in r['H'] if f['estado'] == ENC))
    else:
        r = extraer_frases(cortar_secciones(leer_pdf(rutas[0]), oficiales), diccionario, indice)
        for tipo, titulo in (('H', 'Frases H'), ('P', 'Frases P')):
            print(f"--- {titulo}: {len(r[tipo])} (método: {r['metodo'][tipo]}) ---")
            for f in r[tipo]:
                parecido = f", {f['parecido']:.2f}" if 'parecido' in f else ''
                print(f"  {f['codigo'] or '?':<16} p.{f['pagina']:<3}[{f['estado']}{parecido}] {f['texto_ficha']}")
                if f.get('nota'):
                    print(f"  {'':<16}      nota: {f['nota']}")
                if f['estado'] == REV and f.get('candidatos'):
                    print(f"  {'':<16}      candidatos: " + ', '.join(f'{c} ({p:.2f})' for c, p in f['candidatos']))
        print('--- Hallazgos que se conservan ---')
        for d in r['declaraciones']:
            print(f"  declaración       p.{d['pagina']:<3}{d['texto_ficha']}")
        for s in r['sin_codigo']:
            print(f"  sin código SGA    p.{s['pagina']:<3}(bloque {s['bloque']}) {s['texto_ficha']}")
        for o in r['otras_frases']:
            print(f"  otro sistema      p.{o['pagina']:<3}{o['codigo']} {o['texto_ficha']}")
        for n, cods in r['fuera_de_seccion_2'].items():
            print(f"  no tomados        sección {n}: {', '.join(cods)}")
