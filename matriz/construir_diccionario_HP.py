# -*- coding: utf-8 -*-
"""Genera frases_HP_es_en.xlsx: diccionario de frases H y P en español e inglés,
leído directamente de las tablas del Anexo 3 del SGA sexta edición revisada
(PDF de Naciones Unidas guardados en referencias/SGA_rev6). No contiene textos
escritos a mano: si cambian los PDF, se ejecuta de nuevo.
Requiere: pip install openpyxl pdfplumber"""
import os, re, unicodedata
import openpyxl, pdfplumber
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

B = os.path.dirname(os.path.abspath(__file__)) + os.sep
REF = os.path.join(B, '..', 'referencias', 'SGA_rev6') + os.sep
# Archivo y páginas del PDF donde están las tablas A3.1.1 a A3.2.5 (sección 1 y 2 del Anexo 3)
PDF = {'es': ('SGA_rev6_es_ST-SG-AC10-30-Rev6sp.pdf', 293, 337),
       'en': ('SGA_rev6_en_ST-SG-AC10-30-Rev6e.pdf', 274, 308)}
# Páginas del PDF donde está la tabla 1.5.2 (información mínima de una FDS)
TABLA_SECCIONES = {'es': (43, 45), 'en': (44, 45)}
TIPO_H = {'2': 'Físico', '3': 'Salud', '4': 'Ambiente'}
TIPO_P = {'1': 'General', '2': 'Prevención', '3': 'Intervención', '4': 'Almacenamiento', '5': 'Eliminación'}
ES_CODIGO = re.compile(r'^[HP]\d{3}')


def plano(t):
    return ' '.join((t or '').split())


def codigo(celda):
    """'P301\\n+\\nP310' -> 'P301+P310'; quita la marca de continuación de página."""
    return re.sub(r'\(cont[^)]*\)', '', celda or '', flags=re.I).replace('\n', '').replace(' ', '')


def leer(idioma):
    """Devuelve {código: {texto, clase, categoria, pagina}} en el orden del PDF.
    Se lee por tablas (posición en la página), no por líneas de texto, porque en
    estos PDF el texto plano descuadra la columna de códigos."""
    archivo, ini, fin = PDF[idioma]
    datos = {}
    with pdfplumber.open(REF + archivo) as pdf:
        for n in range(ini, fin + 1):
            for tabla in pdf.pages[n - 1].extract_tables():
                for fila in tabla:
                    c = codigo(fila[0])
                    if not ES_CODIGO.match(c):
                        continue
                    texto = plano(fila[1])
                    if c not in datos:
                        datos[c] = {'texto': texto, 'clase': plano(fila[2]),
                                    'categoria': plano(fila[3]), 'pagina': n}
                    elif texto and not datos[c]['texto']:
                        datos[c]['texto'] = texto
                    elif texto and texto != datos[c]['texto']:
                        raise SystemExit(f'{c} ({idioma}): dos textos distintos en el PDF, revisar.')
    return datos


def leer_clases_p(idioma='es'):
    """Para cada consejo P, las clases de peligro a las que el SGA lo asigna (columnas 3 a 5 de las tablas A3.2.1 a A3.2.5).
    Un consejo ocupa varias filas, una por clase; en las filas siguientes la celda del código viene vacía.
    Devuelve [(código P, clase, capítulo, categoría, condiciones de uso, página)]."""
    archivo, ini, fin = PDF[idioma]
    filas, actual = [], None
    with pdfplumber.open(REF + archivo) as pdf:
        for n in range(ini, fin + 1):
            for tabla in pdf.pages[n - 1].extract_tables():
                for fila in tabla:
                    if len(fila) < 5:
                        continue
                    c = codigo(fila[0])
                    if ES_CODIGO.match(c):
                        actual = c if c[0] == 'P' else None
                    elif c:          # encabezado u otra cosa: no es continuación de un consejo
                        continue
                    clase = plano(fila[2])
                    capitulo = re.search(r'cap[íi]tulo\s+(\d+\.\d+)', clase)
                    if actual and capitulo:
                        filas.append((actual, clase, capitulo.group(1), plano(fila[3]), plano(fila[4]), n))
    return filas


def leer_secciones(idioma):
    """Títulos de las 16 secciones de la FDS (tabla 1.5.2 del SGA): {número: (título, página)}."""
    archivo = PDF[idioma][0]
    ini, fin = TABLA_SECCIONES[idioma]
    titulos = {}
    with pdfplumber.open(REF + archivo) as pdf:
        for n in range(ini, fin + 1):
            for tabla in pdf.pages[n - 1].extract_tables():
                for fila in tabla:
                    num = plano(fila[0]).rstrip('.')
                    if num.isdigit() and int(num) not in titulos:
                        titulos[int(num)] = (plano(fila[1]), n)
    if sorted(titulos) != list(range(1, 17)):
        raise SystemExit(f'Tabla 1.5.2 ({idioma}): no se leyeron las 16 secciones, revisar.')
    return titulos


def comparable(t):
    t = unicodedata.normalize('NFD', t.lower())
    return re.sub(r'[^a-z0-9]+', ' ', ''.join(ch for ch in t if unicodedata.category(ch) != 'Mn')).strip()


def cotejar(hoja, sga):
    """Compara los textos de una hoja de matriz_decision_HP.xlsx con los del SGA."""
    filas = []
    for cod, texto in hoja.iter_rows(min_row=2, max_col=2, values_only=True):
        if not (cod and ES_CODIGO.match(str(cod))):
            continue
        oficial = sga.get(cod, {}).get('texto')
        if oficial is None:
            r = 'El código no está en el Anexo 3'
        elif plano(texto).rstrip('.') == oficial.rstrip('.'):
            r = 'Igual'
        elif comparable(texto) == comparable(oficial):
            r = 'Solo cambia puntuación, tildes o mayúsculas'
        else:
            r = 'Diferente'
        filas.append((cod, texto, oficial, r))
    return filas


es, en = leer('es'), leer('en')
if list(es) != list(en):
    raise SystemExit('Los códigos del PDF en español y en inglés no coinciden: '
                     + ', '.join(sorted(set(es) ^ set(en))))
vacios = [c for c in es if not es[c]['texto'] or not en[c]['texto']]
if vacios:
    raise SystemExit('Códigos sin texto: ' + ', '.join(vacios))

filas_h = [(c, es[c]['texto'], en[c]['texto'], TIPO_H[c[1]], es[c]['clase'], en[c]['clase'], es[c]['categoria'],
            es[c]['pagina'], en[c]['pagina']) for c in es if c[0] == 'H']
filas_p = [(c, es[c]['texto'], en[c]['texto'], TIPO_P[c[1]], es[c]['pagina'], en[c]['pagina'])
           for c in es if c[0] == 'P']

filas_pc = leer_clases_p('es')
sin_clase = [c for c in es if c[0] == 'P' and c not in {f[0] for f in filas_pc}]

sec_es, sec_en = leer_secciones('es'), leer_secciones('en')
filas_s = [(n, sec_es[n][0], sec_en[n][0], sec_es[n][1], sec_en[n][1]) for n in range(1, 17)]

matriz = openpyxl.load_workbook(B + 'matriz_decision_HP.xlsx', read_only=True, data_only=True)
cot_h, cot_p = cotejar(matriz['Matriz_H'], es), cotejar(matriz['Matriz_P'], es)

leeme = [
    ('Qué es', 'Diccionario de frases H (indicaciones de peligro) y P (consejos de prudencia) con su texto oficial en español y en inglés. El módulo de extracción lo usa para convertir a código las frases de las FDS que vienen sin código, y para reconocer códigos que no están en Matriz_P.'),
    ('De dónde sale', 'Sistema Globalmente Armonizado (SGA), sexta edición revisada, Naciones Unidas, 2015 (ST/SG/AC.10/30/Rev.6), Anexo 3, tablas A3.1.1 a A3.1.3 (frases H) y A3.2.1 a A3.2.5 (frases P). Los dos PDF están en referencias/SGA_rev6.'),
    ('Cómo se obtuvieron los PDF', 'unece.org bloquea las descargas automáticas. Se descargaron las copias de esos mismos archivos de unece.org que conserva Internet Archive (web.archive.org): captura del 12 de septiembre de 2025 (inglés) y del 6 de febrero de 2024 (español). Descarga: 10 de octubre de 2026.'),
    ('Cómo se generó', 'Lo genera construir_diccionario_HP.py leyendo las tablas de los PDF. Ningún texto se escribió a mano ni de memoria. Las columnas de página indican la página del archivo PDF donde aparece cada código por primera vez.'),
    ('Contenido', f'Frases_H: {len(filas_h)} códigos ({sum("+" not in f[0] for f in filas_h)} simples y {sum("+" in f[0] for f in filas_h)} combinados). Frases_P: {len(filas_p)} códigos ({sum("+" not in f[0] for f in filas_p)} simples y {sum("+" in f[0] for f in filas_p)} combinados).'),
    ('Textos con puntos suspensivos, barras o corchetes', 'Son del SGA: "..." es un dato que completa el fabricante; "/" separa alternativas entre las que el fabricante escoge; "[ ]" es texto opcional. Por eso la frase de una FDS real puede no ser idéntica a la del diccionario.'),
    ('Hoja Secciones_FDS', 'Títulos oficiales de las 16 secciones de una FDS, de la tabla 1.5.2 del SGA (capítulo 1.5). El módulo de extracción los usa para comprobar que el título de cada sección de una ficha corresponde a su número.'),
    ('Hoja P_por_clase', f'Para cada consejo P, las clases de peligro a las que el SGA lo asigna, con la categoría y las condiciones de uso (columnas 3 a 5 de las tablas A3.2.1 a A3.2.5). Una fila por consejo y clase: {len(filas_pc)} filas. La herramienta la usa para mostrar, junto a cada peligro de una ficha, los consejos de esa misma ficha que el SGA asigna a esa clase de peligro. El enlace se hace por el número de capítulo del SGA. Consejos sin clase específica en el SGA (generales o de uso general): {", ".join(sin_clase) or "ninguno"}.'),
    ('Hojas Cotejo_H y Cotejo_P', 'Comparan el texto de Matriz_H y Matriz_P (matriz_decision_HP.xlsx) con el del SGA. Son un insumo para el pendiente "cotejar textos contra el Anexo 3"; aquí no se corrige nada de la matriz.'),
    ('Qué no es', 'No clasifica ni asigna niveles. La clase y la categoría de Frases_H se copian del SGA solo como referencia; las reglas del motor siguen en matriz_decision_HP.xlsx.'),
    ('Pendiente', 'Verificación del autor contra el PDF, al menos por muestreo, usando las columnas de página.'),
]

AR = 'Arial'
hf = Font(name=AR, bold=True, color='FFFFFF'); hfill = PatternFill('solid', fgColor='1F3864')
wb = openpyxl.Workbook()


def hoja(nombre, encabezado, filas, anchos, primera=False):
    ws = wb.active if primera else wb.create_sheet()
    ws.title = nombre
    ws.append(encabezado)
    for f in filas:
        ws.append(f)
    for fila in ws.iter_rows():
        for c in fila:
            c.font = Font(name=AR, size=10)
            c.alignment = Alignment(wrap_text=True, vertical='top')
    for c in ws[1]:
        c.font = hf; c.fill = hfill
    for i, a in enumerate(anchos, 1):
        ws.column_dimensions[get_column_letter(i)].width = a
    ws.freeze_panes = 'A2'
    return ws


hoja('LEEME', ('Tema', 'Detalle'), leeme, (34, 120), primera=True)
hoja('Frases_H', ('Código H', 'Texto en español (SGA)', 'Texto en inglés (SGA)', 'Tipo de peligro',
                  'Clase de peligro (SGA)', 'Clase de peligro en inglés (SGA)', 'Categoría (SGA)',
                  'Pág. PDF español', 'Pág. PDF inglés'),
     filas_h, (18, 60, 60, 12, 50, 50, 16, 10, 10))
hoja('Frases_P', ('Código P', 'Texto en español (SGA)', 'Texto en inglés (SGA)', 'Tipo',
                  'Pág. PDF español', 'Pág. PDF inglés'), filas_p, (24, 75, 75, 16, 10, 10))
hoja('Secciones_FDS', ('N°', 'Título en español (SGA)', 'Título en inglés (SGA)',
                       'Pág. PDF español', 'Pág. PDF inglés'), filas_s, (6, 70, 70, 10, 10))
hoja('P_por_clase', ('Código P', 'Clase de peligro (SGA)', 'Capítulo del SGA', 'Categoría de peligro', 'Condiciones de uso',
                     'Pág. PDF español'), filas_pc, (22, 60, 12, 26, 80, 10))
enc = ('Código', 'Texto en matriz_decision_HP.xlsx', 'Texto del SGA rev. 6 (español)', 'Resultado')
for nombre, filas in (('Cotejo_H', cot_h), ('Cotejo_P', cot_p)):
    ws = hoja(nombre, enc, filas, (20, 70, 70, 36))
    for fila in ws.iter_rows(min_row=2):
        if fila[3].value != 'Igual':
            fila[3].fill = PatternFill('solid', fgColor='FFE699' if fila[3].value.startswith('Solo') else 'FF7C80')

wb.save(B + 'frases_HP_es_en.xlsx')
print(f'Frases_H: {len(filas_h)} | Frases_P: {len(filas_p)} | P_por_clase: {len(filas_pc)} filas; consejos sin clase: {", ".join(sin_clase) or "ninguno"}')
for nombre, filas in (('Cotejo_H', cot_h), ('Cotejo_P', cot_p)):
    cuenta = {}
    for f in filas:
        cuenta[f[3]] = cuenta.get(f[3], 0) + 1
    print(nombre, cuenta)
