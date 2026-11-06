# -*- coding: utf-8 -*-
"""Paso E3 del reporte y la etiqueta: etiqueta imprimible, a tamaño real.
Recibe el contenido que armó etiqueta.py y lo dibuja en una página HTML con medidas en milímetros.
Se abre en el navegador y se imprime (o se guarda como PDF) en hoja tamaño carta, al 100 %, sin "ajustar a la página".
  - Etiqueta y pictogramas: del tamaño que da la hoja Etiqueta_tamano para el empaque. Ese tamaño no se cambia
    solo: si el texto no cabe, la etiqueta se bloquea (decisión del autor: se mantienen las medidas de la norma).
  - Letra: la más grande que quepa, nunca menor que el mínimo de configuracion.json.
  - Si el texto no cabe ni con la letra mínima, o falta algo obligatorio, la etiqueta NO se imprime:
    en pantalla y en papel sale el motivo. No se recorta ni se encoge en silencio.
No decide nada: solo presenta.
Uso de prueba:
    py "clasificador SGA/etiqueta_html.py" 14_ --envase 20
    py "clasificador SGA/etiqueta_html.py" 14_ --elecciones "clasificador SGA/ejemplo_elecciones.json"
Requiere: haber corrido antes  py "clasificador SGA/validacion.py"."""
import json, os, sys
from html import escape
from tablas import cargar_tablas
from etiqueta import contenido_de_etiqueta, cargar_resultado, ajustes, SALIDAS
from pictogramas_svg import svg_de_pictograma

LETRA_MAXIMA_PT = 14   # tope de la búsqueda de tamaño de letra; es de presentación, no de la norma

ESTILO = """
*{box-sizing:border-box}
body{font-family:Arial,Helvetica,sans-serif;margin:0;padding:8mm;background:#f1f3f5;color:#000}
.hoja{display:flex;gap:8mm;align-items:flex-start;flex-wrap:wrap}
.etiqueta{flex:none;background:#fff;border:0.35mm solid #000;padding:2mm;overflow:hidden;line-height:1.12}
.interior{display:flex;flex-direction:column;gap:0.3em}
.producto{font-size:1.3em;font-weight:bold;text-align:center;overflow-wrap:anywhere}
.pictogramas{display:flex;flex-wrap:wrap;justify-content:center;gap:1mm}
.pictogramas div{flex:none}
.pictogramas svg{width:100%;height:100%;display:block}
.palabra{font-size:1.4em;font-weight:bold;text-align:center;text-transform:uppercase;letter-spacing:.04em}
.frases{margin:0;text-align:left}
.frases.h{font-weight:bold}
.pie{border-top:0.2mm solid #000;padding-top:.25em}
.pie p{margin:0}
.completar{background:#ffe08a}
.panel{background:#fff;border:1px solid #d5dbe0;border-radius:3mm;padding:5mm;max-width:95mm;font-size:10pt;font-family:-apple-system,'Segoe UI',Helvetica,Arial,sans-serif;line-height:1.35}
.boton{font:inherit;font-weight:600;padding:2mm 5mm;border-radius:2.5mm;border:0;background:#1f5fa8;color:#fff;cursor:pointer;width:100%}
.boton[disabled]{background:#b9c2c9;cursor:not-allowed}
.nota{color:#5f6368;font-size:8.5pt}
.panel h1{font-size:12pt;margin:0 0 3mm 0}.panel h2{font-size:10pt;margin:4mm 0 1mm 0}
.panel ul{margin:0;padding-left:5mm}.panel li{margin-bottom:1.2mm}
.bloqueo{border:0.6mm solid #b00020;color:#8b1a1a;background:#fff;padding:5mm;font-size:11pt;max-width:120mm;border-radius:3mm;margin-bottom:4mm}
.bloqueo h1{font-size:13pt;margin:0 0 2mm 0}
.oculto{display:none}
@page{size:letter portrait;margin:10mm}
@media print{
 body{background:#fff;padding:0}
 .panel{display:none}
 .completar{background:none}
 html.no-imprimir .etiqueta{display:none}
}
@media screen{html.no-imprimir .etiqueta{outline:1.2mm solid #b00020}}
"""

# Busca la letra más grande que quepa, de mayor a menor; si ni la mínima cabe, bloquea la impresión.
AJUSTE = """
(function(){
  var e=document.querySelector('.etiqueta'); if(!e) return;
  var i=e.querySelector('.interior'), min=parseFloat(e.dataset.letraMinima), max=parseFloat(e.dataset.letraMaxima);
  var alto=function(){var s=getComputedStyle(e);return e.clientHeight-parseFloat(s.paddingTop)-parseFloat(s.paddingBottom);};
  var ancho=function(){var s=getComputedStyle(e);return e.clientWidth-parseFloat(s.paddingLeft)-parseFloat(s.paddingRight);};
  var cabe=function(pt){e.style.fontSize=pt+'pt';return i.scrollHeight<=alto()+0.5 && i.scrollWidth<=ancho()+0.5;};
  var elegida=null;
  for(var pt=max; pt>=min-0.001; pt-=0.25){ if(cabe(pt)){elegida=pt;break;} }
  var raiz=document.documentElement;
  if(elegida===null){ e.style.fontSize=min+'pt'; raiz.dataset.ajuste='no cabe'; raiz.classList.add('no-imprimir');
    document.getElementById('no-cabe').classList.remove('oculto'); }
  else { raiz.dataset.ajuste='cabe'; }
  raiz.dataset.letra=(elegida===null?min:elegida);
  var d=document.getElementById('letra'); if(d) d.textContent=(elegida===null?'no cabe ni a '+min+' pt':elegida+' pt');
  var b=document.getElementById('imprimir'); if(b && raiz.classList.contains('no-imprimir')){ b.disabled=true; b.textContent='No se puede imprimir todavía'; }
})();
"""


def _medida(mm, cual):
    """Medida a usar: el mínimo exigible si la tabla lo da; si no, el recomendado."""
    par = mm['minimo'] or mm['recomendado']
    if not par:
        raise ValueError(f'La hoja Etiqueta_tamano no da medidas de {cual} para este empaque.')
    return par


def dibujar_etiqueta(contenido, tablas):
    """Página HTML de la etiqueta. Su tamaño es siempre el de la hoja Etiqueta_tamano para el empaque; no hay forma
    de pedir otro (decisión del autor: la etiqueta cumple el tamaño de la norma; la hoja carta es solo el papel)."""
    c = contenido
    if not c['tamano']:
        return _pagina(c, None, '<p>No se puede dibujar sin la capacidad del empaque.</p>', None, None)
    ancho, alto = _medida(c['tamano']['etiqueta_mm'], 'la etiqueta')
    avisos_medida = []
    if not c['tamano']['etiqueta_mm']['minimo']:
        avisos_medida.append(f'Para este empaque la tabla no exige un mínimo de etiqueta; se usó el recomendado ({ancho:g} × {alto:g} mm).')
    lado, _ = _medida(c['tamano']['pictograma_mm'], 'el pictograma')
    minimo_pt = ajustes(tablas)['tamano_minimo_de_letra_pt']

    partes = [f'<div class="producto">{escape(c["producto"] or "")}</div>']
    if c['pictogramas']:
        partes.append('<div class="pictogramas">' + ''.join(
            f'<div style="width:{lado:g}mm;height:{lado:g}mm">{svg_de_pictograma(p["codigo"], c["borde"], p["simbolo"])}</div>'
            for p in c['pictogramas']) + '</div>')
    if not c['solo_contenido_minimo']:   # en empaques de menos de 30 ml basta el nombre y los pictogramas (R10)
        if c['palabra']:
            partes.append(f'<div class="palabra">{escape(c["palabra"])}</div>')
        # Las frases van seguidas, en párrafo, como en las etiquetas comerciales: así ocupan mucho menos alto.
        punto = lambda t: t if t.rstrip().endswith(('.', '…')) else t.rstrip() + '.'
        texto = lambda t: escape(punto(t)).replace('/', '/<wbr>')   # permite cortar el renglón después de una barra
        if c['frases_h']:
            partes.append('<p class="frases h">' + ' '.join(f'<span>{texto(f["texto"])}</span>' for f in c['frases_h']) + '</p>')
        if c['consejos_p']:
            partes.append('<p class="frases p">' + ' '.join(
                f'<span{" class=completar" if f["sin_completar"] else ""}>{texto(f["texto"])}</span>' for f in c['consejos_p']) + '</p>')
        pie = []
        if c['componentes']:
            pie.append(f'<p><b>Componentes:</b> {escape(c["componentes"])}</p>')
        p = c['proveedor']
        if p['incluido']:
            pie.append('<p><b>Proveedor:</b> ' + escape(' · '.join(x for x in (p['nombre'], p['direccion'], p['telefono'] and 'Tel. ' + p['telefono']) if x)) + '</p>')
        extras = ' · '.join(x for x in (c['cantidad'] and f'Cantidad: {c["cantidad"]}', c['lote'] and f'Lote: {c["lote"]}') if x)
        if extras:
            pie.append(f'<p>{escape(extras)}</p>')
        if pie:
            partes.append('<div class="pie">' + ''.join(pie) + '</div>')
    etiqueta = (f'<div class="etiqueta" style="width:{ancho:g}mm;height:{alto:g}mm" data-letra-minima="{minimo_pt:g}" '
                f'data-letra-maxima="{LETRA_MAXIMA_PT}" data-ancho-mm="{ancho:g}" data-alto-mm="{alto:g}" data-pictograma-mm="{lado:g}">'
                f'<div class="interior">{"".join(partes)}</div></div>')
    return _pagina(c, etiqueta, None, (ancho, alto, lado, minimo_pt), avisos_medida)


def _pagina(c, etiqueta, mensaje, medidas, avisos_medida):
    lista = lambda xs: '<ul>' + ''.join(f'<li>{escape(x)}</li>' for x in xs) + '</ul>'
    bloqueos = ''
    if c['faltantes']:   # falta algo obligatorio: no se imprime
        bloqueos += f'<div class="bloqueo"><h1>Esta etiqueta no está lista para imprimir</h1>{lista(c["faltantes"])}</div>'
    bloqueos += ('<div class="bloqueo oculto" id="no-cabe"><h1>El texto no cabe en la etiqueta</h1><p>Ni con la letra mínima configurada cabe todo el '
                 'contenido en el tamaño que exige la norma para este empaque. Deje menos consejos de prudencia (los demás se consultan en la ficha). No se imprime recortada.</p></div>')
    panel = '<div class="panel"><h1>Datos de esta etiqueta</h1>'
    if etiqueta:
        panel += '<p><button class="boton" id="imprimir" onclick="window.print()">Imprimir etiqueta</button></p>'
    panel += f'<p>Ficha: {escape(c["archivo"])}<br>Tipo: {escape(c["tipo"])}<br>Reglas aplicadas: {escape(", ".join(c["reglas"]))}</p>'
    if medidas:
        ancho, alto, lado, minimo_pt = medidas
        panel += (f'<p>Empaque: {escape(c.get("capacidad") or "")} ({escape(c["tamano"]["capacidad"].replace("litros", "L"))})<br>Etiqueta: {ancho:g} × {alto:g} mm<br>'
                  f'Pictogramas: {lado:g} × {lado:g} mm, borde {c["borde"]}<br>Letra: <span id="letra">—</span> (mínimo configurado: {minimo_pt:g} pt)</p>')
        panel += '<p><b>Para imprimir a tamaño real:</b> hoja tamaño carta, escala 100 %, sin «ajustar a la página». Solo sale la etiqueta; este panel no se imprime.</p>'
    avisos = (avisos_medida or []) + c['avisos']
    if avisos:
        panel += '<h2>Avisos</h2>' + lista(avisos)
    if any(f['sin_completar'] for f in c['consejos_p']):
        panel += '<p>Los renglones resaltados en amarillo traen «...» o un texto entre corchetes sin completar: complételos o quítelos. El resaltado no se imprime.</p>'
    panel += ('<p class="nota">Esta etiqueta se arma con los datos leídos de la ficha y las tablas del proyecto. No sustituye la etiqueta original '
              'del fabricante y, por sí sola, no garantiza el cumplimiento de la norma: el responsable del etiquetado debe revisarla.</p>')
    panel += '</div>'
    clase = ' class="no-imprimir"' if c['faltantes'] or not etiqueta else ''
    return (f'<!doctype html><html lang="es"{clase}><head><meta charset="utf-8"><title>Etiqueta — {escape(c["producto"] or c["archivo"])}</title>'
            f'<style>{ESTILO}</style></head><body><div class="hoja">{etiqueta or ""}{mensaje or ""}<div>{bloqueos}</div>{panel}</div>'
            f'<script>{AJUSTE}</script></body></html>')


def guardar_etiqueta(html, contenido, sufijo=''):
    destino = os.path.join(SALIDAS, 'etiquetas')
    os.makedirs(destino, exist_ok=True)
    litros = f"{contenido['tamano']['litros']:g}L" if contenido['tamano'] else 'sin_envase'
    ruta = os.path.join(destino, f"{os.path.splitext(contenido['archivo'])[0]}_{litros}{sufijo}.html")
    with open(ruta, 'w', encoding='utf-8') as f:
        f.write(html)
    return ruta


if __name__ == '__main__':
    sys.stdout.reconfigure(encoding='utf-8')
    argumentos = sys.argv[1:]
    elecciones, sufijo = {}, ''
    if '--elecciones' in argumentos:
        with open(argumentos[argumentos.index('--elecciones') + 1], encoding='utf-8') as f:
            elecciones = {k: v for k, v in json.load(f).items() if not k.startswith('_')}
        sufijo = '_con_elecciones'
    if '--envase' in argumentos:
        elecciones['envase_litros'] = float(argumentos[argumentos.index('--envase') + 1].replace(',', '.'))
    elecciones.update(uso_interno='--interno' in argumentos or elecciones.get('uso_interno', False))
    if '--negro' in argumentos:
        elecciones['borde'] = 'negro'
        sufijo += '_negro' + ('_interno' if elecciones['uso_interno'] else '')
    tablas = cargar_tablas()
    contenido = contenido_de_etiqueta(cargar_resultado(argumentos[0]), elecciones, tablas)
    ruta = guardar_etiqueta(dibujar_etiqueta(contenido, tablas), contenido, sufijo)
    print(f"{contenido['archivo']} -> {contenido['tipo']} | lista para imprimir: {'sí' if contenido['lista_para_imprimir'] else 'NO'}")
    print(f'Etiqueta guardada en {ruta}')
