# -*- coding: utf-8 -*-
"""Paso 4 de la extracción: frases H y P que la FDS trae con código.
Solo se toman de la sección 2 (la que define la etiqueta del producto). Los códigos que
aparecen en otras secciones describen componentes o son listados de referencia: se
informan aparte y no se mezclan. Cada código se valida contra el diccionario del SGA.
Aquí no se clasifica nada: solo se dice qué códigos declara la ficha.
Uso de prueba:
    py extraccion/frases.py FDS                          (resumen de todas)
    py extraccion/frases.py FDS/02_Diesel_ACPM_68334-30-5.pdf   (detalle)
Requiere: pip install pdfplumber openpyxl"""
import glob, os, re, sys
import openpyxl
from leer_pdf import leer_pdf
from secciones import cortar_secciones, cargar_titulos_oficiales, DICCIONARIO

ENC, REV = 'encontrado', 'por revisar'
UNO_H = r'(?:EUH\d{3}[A-Za-z]?|H[2-4]\d{2}[A-Za-z]{0,2})'
UNO_P = r'P[1-5]\d{2}'
# Un código o varios unidos con "+", sin letras ni cifras pegadas antes o después
CODIGO = re.compile(r'(?<![A-Za-z0-9\-])(' + UNO_H + r'(?:\s*\+\s*' + UNO_H + r')*|' +
                    UNO_P + r'(?:\s*\+\s*' + UNO_P + r')*)(?![A-Za-z0-9])')
# Frases con código de otro sistema (OSHA-H01, CGA-PG02...): se conservan como hallazgo
OTRO_SISTEMA = re.compile(r'^\s*([A-Z]{2,6}-[A-Z]{1,3}\d{2,3})\s*[-–:]\s*(.+)$')
# Líneas que cierran una frase: subtítulos de la sección 2
SUBTITULO = re.compile(r'^\s*(\d+\.\d+\.?\s|Prevention|Response|Storage|Disposal|Prevenci[óo]n|Respuesta|Intervenci[óo]n|'
                       r'Almacenamiento|Eliminaci[óo]n|Desecho|Precautionary|Hazard statements?|Signal word|'
                       r'Consejos de prudencia|Indicaciones de peligro|Palabra de advertencia|Other hazards|'
                       r'Additional|Supplemental|Otros peligros)', re.I)
MAX_CONTINUACION = 3  # líneas siguientes que pueden completar una frase


def cargar_diccionario():
    """{'H': {código: (texto es, texto en)}, 'P': {...}} de frases_HP_es_en.xlsx."""
    libro = openpyxl.load_workbook(DICCIONARIO, read_only=True)
    return {t: {c: (es, en) for c, es, en, *_ in libro['Frases_' + t].iter_rows(min_row=2, values_only=True)}
            for t in 'HP'}


def _validar(codigo, diccionario):
    """Estado del código frente al diccionario y nota explicativa."""
    tabla = diccionario[codigo[-4] if codigo[0] != 'E' else 'H'] if '+' not in codigo else diccionario[codigo[0]]
    if codigo in tabla:
        return ENC, codigo, None
    partes = codigo.split('+')
    if len(partes) > 1:
        faltan = [p for p in partes if p not in tabla and p[:4] not in tabla]
        if faltan:
            return REV, None, 'no están en el Anexo 3 del SGA rev. 6: ' + ', '.join(faltan)
        return REV, None, 'el SGA rev. 6 no lista esta combinación; sus códigos por separado sí están'
    if codigo[:4] in tabla:  # H360D, H361fd: sufijo sobre un código que sí está
        return ENC, codigo[:4], f'lleva sufijo; se resuelve con {codigo[:4]}'
    return REV, None, 'no está en el Anexo 3 del SGA rev. 6'


def _frase(lineas, i, fin_codigo):
    """Texto que sigue a un código: el resto de su línea y, si la frase no cerró, las siguientes."""
    texto = lineas[i][1][fin_codigo:].lstrip(' \t-–:)').rstrip()
    for _, siguiente in lineas[i + 1:i + 1 + MAX_CONTINUACION]:
        if texto.endswith('.') or CODIGO.search(siguiente) or SUBTITULO.match(siguiente) or OTRO_SISTEMA.match(siguiente):
            break
        texto = texto[:-1] + siguiente.strip() if re.search(r'[a-záéíóúñ]-$', texto) and siguiente.strip()[:1].islower() \
            else f'{texto} {siguiente.strip()}'
    return texto.strip()


def _codigos(lineas):
    """Apariciones de códigos en un bloque de líneas: [(código, página, posición, fin, al inicio de línea)]."""
    salida = []
    for i, (pag, linea) in enumerate(lineas):
        for m in CODIGO.finditer(linea):
            codigo = re.sub(r'\s+', '', m.group(1))
            al_inicio = not re.search(r'[a-z]{3}', linea[:m.start()].split(':')[-1], re.I)
            salida.append((codigo, pag, i, m.end(), al_inicio))
    return salida


def extraer_frases_con_codigo(corte, diccionario=None):
    """corte = salida de cortar_secciones. Devuelve:
    'H', 'P': [{'codigo', 'codigo_base', 'texto_ficha', 'estado', 'nota', 'pagina', 'cita', 'apariciones'}]
    'otras_frases': frases de la sección 2 con código de otro sistema (hallazgo)
    'fuera_de_seccion_2': {sección: códigos que no están en la sección 2} (no se toman)"""
    diccionario = diccionario or cargar_diccionario()
    lineas = corte['secciones'][2]['lineas']
    resultado = {'H': [], 'P': [], 'otras_frases': [], 'fuera_de_seccion_2': {}}

    por_codigo = {}
    for codigo, pag, i, fin, al_inicio in _codigos(lineas):
        por_codigo.setdefault(codigo, []).append((al_inicio, _frase(lineas, i, fin), pag, lineas[i][1]))
    for codigo, apariciones in por_codigo.items():
        # La frase se lee donde el código encabeza su renglón; si no, de la aparición con más texto
        _, texto, pag, cita = max(apariciones, key=lambda a: (a[0] and bool(a[1]), len(a[1])))
        estado, base, nota = _validar(codigo, diccionario)
        tipo = 'H' if codigo.lstrip('EU')[0] == 'H' else 'P'
        resultado[tipo].append(dict(codigo=codigo, codigo_base=base, texto_ficha=texto or None, estado=estado, nota=nota,
                                    pagina=pag, cita=cita, apariciones=len(apariciones)))

    for pag, linea in lineas:
        m = OTRO_SISTEMA.match(linea) or re.search(r'(?<=: )([A-Z]{2,6}-[A-Z]{1,3}\d{2,3})\s*[-–:]\s*(.+)$', linea)
        if m:
            resultado['otras_frases'].append(dict(codigo=m.group(1), texto_ficha=m.group(2).strip(), pagina=pag, cita=linea))

    en_seccion_2 = set(por_codigo)
    for n, s in corte['secciones'].items():
        ajenos = sorted({c for c, *_ in _codigos(s['lineas'])} - en_seccion_2) if n != 2 else []
        if ajenos:
            resultado['fuera_de_seccion_2'][n] = ajenos
    return resultado


if __name__ == '__main__':
    sys.stdout.reconfigure(encoding='utf-8')
    oficiales, diccionario = cargar_titulos_oficiales(), cargar_diccionario()
    destino = sys.argv[1]
    if os.path.isdir(destino):
        print(f"{'ficha':<26}{'H':>3}{'P':>4}{'por revisar':>13}  observaciones")
        for ruta in sorted(glob.glob(os.path.join(destino, '*.pdf'))):
            r = extraer_frases_con_codigo(cortar_secciones(leer_pdf(ruta), oficiales), diccionario)
            dudosos = [f['codigo'] for f in r['H'] + r['P'] if f['estado'] == REV]
            obs = []
            if not r['H'] and not r['P']:
                obs.append('sin códigos en la sección 2 (va al paso 5)')
            if dudosos:
                obs.append('revisar: ' + ', '.join(dudosos))
            if r['otras_frases']:
                obs.append('otro sistema: ' + ', '.join(o['codigo'] for o in r['otras_frases']))
            if r['fuera_de_seccion_2']:
                obs.append('no tomados: ' + '; '.join(f"sec. {n}: {', '.join(c)}" for n, c in r['fuera_de_seccion_2'].items()))
            print(f"{os.path.basename(ruta)[:25]:<26}{len(r['H']):>3}{len(r['P']):>4}{len(dudosos):>13}  {' | '.join(obs)}")
    else:
        r = extraer_frases_con_codigo(cortar_secciones(leer_pdf(destino), oficiales), diccionario)
        for tipo, titulo in (('H', 'Frases H'), ('P', 'Frases P')):
            print(f'--- {titulo} de la sección 2: {len(r[tipo])} ---')
            for f in r[tipo]:
                print(f"  {f['codigo']:<22} p.{f['pagina']:<3} [{f['estado']}] {f['texto_ficha']}")
                if f['nota']:
                    print(f"  {'':<22}       nota: {f['nota']}")
        print(f"--- Frases con código de otro sistema: {len(r['otras_frases'])} ---")
        for o in r['otras_frases']:
            print(f"  {o['codigo']:<22} p.{o['pagina']:<3} {o['texto_ficha']}")
        print('--- Códigos en otras secciones que no están en la sección 2 (no se toman) ---')
        print('\n'.join(f'  sección {n}: ' + ', '.join(c) for n, c in r['fuera_de_seccion_2'].items()) or '  ninguno')
