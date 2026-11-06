# -*- coding: utf-8 -*-
"""Paso 8b de la extracción: protección personal (sección 8 de la FDS).
Ubica dentro de la sección 8 el apartado de protección personal y guarda su texto literal, con la página.
No interpreta, no resume y no traduce: la traducción al español la valida el autor en matriz/proteccion_seccion8.xlsx.
Si la sección 8 no trae un rótulo de protección personal reconocible, el dato queda 'no encontrado'.
Uso de prueba:
    py extraccion/proteccion.py FDS                         (resumen de todas)
    py extraccion/proteccion.py FDS/15_Acetileno_74-86-2.pdf     (detalle)
    py extraccion/proteccion.py FDS --guardar              (agrega la clave 'proteccion' a los resultados ya guardados, sin tocar lo demás)
Requiere: pip install pdfplumber openpyxl"""
import glob, json, os, re, sys
from leer_pdf import leer_pdf
from secciones import cortar_secciones, cargar_titulos_oficiales

ENC, NO_ENC = 'encontrado', 'no encontrado'
MAXIMO = 1800   # tope del texto guardado; si la ficha trae más, se corta y se marca
# Rótulos con que las fichas abren el apartado. Son palabras para ubicar el texto; el contenido se copia literal.
INICIO = re.compile(r'^\s*(?:\d+(?:\.\d+)*\.?\s+|[a-z]\)\s*)?(?:'
                    r'individual protection measures|personal protecti(?:ve|on)(?: equipment| measures)?|protective equipment|'
                    r'eye\s*(?:/|and|or)?\s*face protection|eye protection|respiratory protection|skin(?: and body)? protection|hand protection|'
                    r'medidas de protecci[óo]n (?:individual|personal)|protecci[óo]n (?:personal|individual)|equipos? de protecci[óo]n(?: personal| individual)?|'
                    r'protecci[óo]n (?:de (?:los )?ojos|ocular|respiratoria|de (?:la )?piel|de (?:las )?manos|cut[áa]nea)|'
                    r'controles? de (?:la )?exposici[óo]n|exposure controls|appropriate engineering controls|engineering (?:controls|measures)|controles t[ée]cnicos)\b', re.I)


def extraer_proteccion(corte):
    """corte = salida de cortar_secciones. Devuelve {'texto', 'pagina', 'estado', 'cortado', 'cita'}."""
    lineas = (corte['secciones'].get(8) or {}).get('lineas') or []
    desde = next((i for i, (_, linea) in enumerate(lineas) if INICIO.match(linea)), None)
    if desde is None:
        return dict(texto=None, pagina=None, estado=NO_ENC, cortado=False, cita=None)
    texto = ' '.join(' '.join(linea.split()) for _, linea in lineas[desde:])
    return dict(texto=texto[:MAXIMO], pagina=lineas[desde][0], estado=ENC, cortado=len(texto) > MAXIMO, cita=lineas[desde][1].strip())


if __name__ == '__main__':
    sys.stdout.reconfigure(encoding='utf-8')
    oficiales = cargar_titulos_oficiales()
    destino = sys.argv[1]
    rutas = sorted(glob.glob(os.path.join(destino, '*.pdf'))) if os.path.isdir(destino) else [destino]
    salidas = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'salidas')
    for ruta in rutas:
        p = extraer_proteccion(cortar_secciones(leer_pdf(ruta), oficiales))
        if '--guardar' in sys.argv:
            archivo = os.path.join(salidas, os.path.splitext(os.path.basename(ruta))[0] + '.json')
            with open(archivo, encoding='utf-8') as f:
                resultado = json.load(f)
            resultado['proteccion'] = p
            with open(archivo, 'w', encoding='utf-8') as f:
                json.dump(resultado, f, ensure_ascii=False, indent=1)
        if len(rutas) > 1:
            print(f"{os.path.basename(ruta)[:22]:<23}{p['estado']:<14} p.{p['pagina'] or '—':<3} {len(p['texto'] or ''):>5}{'+' if p['cortado'] else ' '} | {(p['cita'] or '')[:70]}")
        else:
            print(f"[{p['estado']}] página {p['pagina']}{' (cortado)' if p['cortado'] else ''}\n{p['texto']}")
