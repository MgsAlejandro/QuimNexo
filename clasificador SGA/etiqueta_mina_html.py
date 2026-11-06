# -*- coding: utf-8 -*-
"""Dibujo de la etiqueta de uso en la mina, a tamaño real (regla C08). Recibe el contenido que armó etiqueta_mina.py.
Se abre en el navegador y se imprime en hoja carta, al 100 %, sin «ajustar a la página».
  - Etiqueta y pictogramas: del tamaño que da la hoja Etiqueta_tamano para la capacidad del empaque. Eso no se cambia.
  - Letra: la más grande que quepa, nunca menor que el mínimo de configuracion.json.
  - SIEMPRE se puede imprimir. Si el contenido no cabe con la letra mínima, la página quita contenido en el orden de
    la hoja Etiqueta_mina (almacenamiento; renglones de menor prioridad, dejando al menos uno; iconos de protección; y,
    solo en el tamaño más pequeño, el texto de protección y el peligro principal) hasta que quepa. Nada se recorta en
    silencio: el panel de pantalla dice qué se quitó.
No decide nada: solo presenta.
Uso de prueba:
    py "clasificador SGA/etiqueta_mina_html.py" 14_ --capacidad 20 --unidad L
Requiere: haber corrido antes  py extraccion/extraer.py FDS"""
import os, re, sys
from html import escape
from tablas import cargar_tablas
from motor import clasificar, cargar_fichas
from etiqueta import ajustes
from etiqueta_html import _medida
from etiqueta_mina import contenido_de_etiqueta_mina
from pictogramas_svg import svg_de_pictograma

AQUI = os.path.dirname(os.path.abspath(__file__))
ICONOS = os.path.join(AQUI, '..', 'plantillas', 'epp')   # dibujos propios y genéricos: gafas, careta, guantes, respirador, botas, ropa
SALIDAS = os.path.join(AQUI, 'salidas')
LETRA_MAXIMA_PT = 16   # tope de la búsqueda de tamaño de letra; es de presentación, no de la norma

ESTILO = """
*{box-sizing:border-box}
body{font-family:Arial,Helvetica,sans-serif;margin:0;padding:8mm;background:#f1f3f5;color:#000}
.hoja{display:flex;gap:8mm;align-items:flex-start;flex-wrap:wrap}
.etiqueta{flex:none;background:#fff;border:0.5mm solid #000;padding:2mm;overflow:hidden;line-height:1.14}
.interior{display:flex;flex-direction:column;gap:0.32em}
.nombre{font-size:1.5em;font-weight:bold;text-align:center;overflow-wrap:anywhere;line-height:1.05}
.nombre-fds{font-size:.68em;text-align:center;color:#333;overflow-wrap:anywhere}
.pictogramas{display:flex;flex-wrap:wrap;justify-content:center;gap:1mm}
.pictogramas div{flex:none}.pictogramas svg{width:100%;height:100%;display:block}
.palabra{font-size:1.5em;font-weight:bold;text-align:center;text-transform:uppercase;letter-spacing:.04em}
.sin-peligros{font-size:1.15em;font-weight:bold;text-align:center}
.peligro{margin:0;font-weight:bold}
.rotulo{font-weight:bold;text-transform:uppercase;font-size:.78em;letter-spacing:.03em;margin:0}
.hacer{margin:0;padding-left:1.1em}.hacer li{margin:0 0 .12em 0}
.epp{display:flex;gap:.5em;align-items:center;border-top:0.2mm solid #000;padding-top:.3em}
.iconos{display:flex;gap:.25em;flex:none}.iconos svg{width:2em;height:2em;display:block}
.epp p{margin:0}
.pie{border-top:0.2mm solid #000;padding-top:.2em;font-size:.8em}.pie p{margin:0}
.panel{background:#fff;border:1px solid #d5dbe0;border-radius:3mm;padding:5mm;max-width:95mm;font-size:10pt;font-family:-apple-system,'Segoe UI',Helvetica,Arial,sans-serif;line-height:1.35}
.panel h1{font-size:12pt;margin:0 0 3mm 0}.panel h2{font-size:10pt;margin:4mm 0 1mm 0}
.panel ul{margin:0;padding-left:5mm}.panel li{margin-bottom:1.2mm}
.boton{font:inherit;font-weight:600;padding:2mm 5mm;border-radius:2.5mm;border:0;background:#1f5fa8;color:#fff;cursor:pointer;width:100%}
.nota{color:#5f6368;font-size:8.5pt}.quitado{color:#8a4b00}
.falta{border:0.6mm solid #b00020;color:#8b1a1a;background:#fff;padding:5mm;font-size:11pt;max-width:120mm;border-radius:3mm}
@page{size:letter portrait;margin:10mm}
@media print{body{background:#fff;padding:0}.panel{display:none}}
"""

# Busca la letra más grande que quepa. Si no cabe con la mínima, quita contenido en el orden acordado y vuelve a buscar.
AJUSTE = """
(function(){
  var e=document.querySelector('.etiqueta'); if(!e) return;
  var i=e.querySelector('.interior'), min=parseFloat(e.dataset.letraMinima), max=parseFloat(e.dataset.letraMaxima);
  var alto=function(){var s=getComputedStyle(e);return e.clientHeight-parseFloat(s.paddingTop)-parseFloat(s.paddingBottom);};
  var ancho=function(){var s=getComputedStyle(e);return e.clientWidth-parseFloat(s.paddingLeft)-parseFloat(s.paddingRight);};
  var cabe=function(pt){e.style.fontSize=pt+'pt';return i.scrollHeight<=alto()+0.5 && i.scrollWidth<=ancho()+0.5;};
  var buscar=function(){for(var pt=max; pt>=min-0.001; pt-=0.25){ if(cabe(pt)) return pt; } return null;};
  var renglones=function(){return [].slice.call(e.querySelectorAll('.hacer li'));};
  var minimo=parseInt(e.dataset.minimoDeRenglones,10), pequeno=e.dataset.masPequeno==='si', quitado=[];
  var quitar=function(el,que){ if(el){ el.remove(); quitado.push(que); return true; } return false; };
  var limpiar=function(){ if(!renglones().length){ var h=e.querySelector('.bloque-hacer'); if(h) h.remove(); } };
  // Cada paso devuelve true si quitó algo; se prueba de nuevo después de cada uno
  var pasos={
    almacenamiento:function(){ var l=renglones(); if(l.length<=minimo) return false;
      var a=l.filter(function(x){return x.dataset.tipo==='Almacenamiento';})[0]; return a?quitar(a,'el renglón de almacenamiento ('+a.dataset.codigo+')'):false; },
    que_hacer:function(){ var l=renglones(); if(l.length<=minimo) return false;
      l.sort(function(a,b){return a.dataset.retiro-b.dataset.retiro;}); return quitar(l[0],'el renglón «'+l[0].textContent.trim()+'» ('+l[0].dataset.codigo+')'); },
    iconos_epp:function(){ return quitar(e.querySelector('.iconos'),'los iconos de protección personal (queda el texto)'); },
    texto_epp:function(){ return pequeno && quitar(e.querySelector('.epp'),'la línea de protección personal'); },
    peligro_principal:function(){ return pequeno && quitar(e.querySelector('.peligro'),'la línea de peligro principal'); }
  };
  var orden=e.dataset.ordenDeRetiro.split(','), elegida=buscar(), k=0;
  while(elegida===null && k<orden.length){
    var paso=pasos[orden[k]];
    if(paso && paso()){ limpiar(); elegida=buscar(); } else { k++; }
  }
  var raiz=document.documentElement, forzado=false;
  if(elegida===null){   // último recurso: ya no hay qué quitar. Se baja la letra por debajo del mínimo y se avisa; no se recorta.
    forzado=true;
    for(var pt=min; pt>=3; pt-=0.25){ if(cabe(pt)){ elegida=pt; break; } }
    if(elegida===null){ elegida=3; e.style.fontSize='3pt'; }
  }
  raiz.dataset.ajuste=forzado?'letra menor que la mínima':(quitado.length?'cabe quitando contenido':'cabe');
  raiz.dataset.letra=elegida; raiz.dataset.quitado=quitado.join(' | ');
  var d=document.getElementById('letra'); if(d) d.textContent=elegida+' pt'+(forzado?' (menor que la mínima configurada: revise la etiqueta)':'');
  var q=document.getElementById('quitado');
  if(q){ q.innerHTML=quitado.length?'<h2>Para que cupiera se quitó</h2><ul class="quitado">'+quitado.map(function(x){return '<li>'+x.replace(/</g,'&lt;')+'</li>';}).join('')+'</ul>'
         +'<p class="nota">Lo que se quitó sigue en la ficha y en la etiqueta completa.</p>':'<p class="nota">Cupo todo el contenido.</p>'; }
})();
"""

_iconos = {}


def icono_de_epp(nombre):
    """SVG de un icono de protección personal, o None si no existe ese dibujo."""
    if nombre not in _iconos:
        ruta = os.path.join(ICONOS, re.sub(r'[^a-z]', '', nombre) + '.svg')
        _iconos[nombre] = open(ruta, encoding='utf-8').read().strip() if os.path.isfile(ruta) else None
    return _iconos[nombre]


def dibujar_etiqueta_mina(contenido, tablas):
    """Página HTML de la etiqueta de uso en la mina. Su tamaño es el de la hoja Etiqueta_tamano para la capacidad del empaque."""
    c = contenido
    if not c['tamano']:
        return _pagina(c, None, '<div class="falta"><b>Falta la capacidad del empaque.</b> De ella depende el tamaño de la etiqueta.</div>', None)
    ancho, alto = _medida(c['tamano']['etiqueta_mm'], 'la etiqueta')
    lado, _ = _medida(c['tamano']['pictograma_mm'], 'el pictograma')
    minimo_pt = ajustes(tablas)['tamano_minimo_de_letra_pt']
    r = c['rotulos']

    partes = [f'<div class="nombre">{escape(c["nombre"])}</div>']
    if c['nombre_en_la_ficha'] and c['nombre_en_la_ficha'].strip().lower() != c['nombre'].strip().lower():
        partes.append(f'<div class="nombre-fds">{escape(c["nombre_en_la_ficha"])}</div>')
    if c['pictogramas']:
        partes.append('<div class="pictogramas">' + ''.join(
            f'<div style="width:{lado:g}mm;height:{lado:g}mm">{svg_de_pictograma(p["codigo"], "rojo", p["simbolo"])}</div>' for p in c['pictogramas']) + '</div>')
    if c['palabra']:
        partes.append(f'<div class="palabra">{escape(c["palabra"])}</div>')
    elif c['sin_peligros']:
        partes.append(f'<div class="sin-peligros">{escape(c["sin_peligros"])}</div>')
    if c['peligro_principal'] or c['asfixiante']:
        texto = ' '.join(x for x in ((c['peligro_principal'] or {}).get('texto'), c['asfixiante']) if x)
        partes.append(f'<p class="peligro">{escape(r["peligro"])}: {escape(texto)}</p>')
    if c['renglones']:
        partes.append(f'<div class="bloque-hacer"><p class="rotulo">{escape(r["hacer"])}</p><ul class="hacer">' + ''.join(
            f'<li data-tipo="{escape(z["tipo"])}" data-codigo="{escape(z["codigo"])}" data-retiro="{z["retiro"]}">{escape(z["texto"])}</li>' for z in c['renglones'])
            + '</ul></div>')
    if c['epp']['texto']:
        dibujos = [icono_de_epp(n) for n in c['epp']['iconos']]
        partes.append('<div class="epp">' + ('<div class="iconos">' + ''.join(d for d in dibujos if d) + '</div>' if any(dibujos) else '')
                      + f'<p>{escape(c["epp"]["texto"])}</p></div>')
    p = c['proveedor']
    pie = [escape(' · '.join(x for x in (p['nombre'], p['direccion'], p['telefono'] and 'Tel. ' + p['telefono']) if x))] if any(p.values()) else []
    pie.append(escape(' · '.join(x for x in (f'Empaque de {c["capacidad"]}', c['cantidad'] and f'Cantidad: {c["cantidad"]}', c['lote'] and f'Lote: {c["lote"]}') if x)))
    partes.append('<div class="pie">' + ''.join(f'<p>{x}</p>' for x in pie if x) + '</div>')

    etiqueta = (f'<div class="etiqueta" style="width:{ancho:g}mm;height:{alto:g}mm" data-letra-minima="{minimo_pt:g}" data-letra-maxima="{LETRA_MAXIMA_PT}" '
                f'data-ancho-mm="{ancho:g}" data-alto-mm="{alto:g}" data-pictograma-mm="{lado:g}" data-minimo-de-renglones="{c["minimo_de_renglones"]}" '
                f'data-mas-pequeno="{"si" if c["tamano_mas_pequeno"] else "no"}" data-orden-de-retiro="{escape(",".join(c["orden_de_retiro"]))}">'
                f'<div class="interior">{"".join(partes)}</div></div>')
    return _pagina(c, etiqueta, None, (ancho, alto, lado, minimo_pt))


def _pagina(c, etiqueta, mensaje, medidas):
    panel = '<div class="panel"><h1>Etiqueta de uso en la mina</h1>'
    if etiqueta:
        ancho, alto, lado, minimo_pt = medidas
        panel += '<p><button class="boton" id="imprimir" onclick="window.print()">Imprimir etiqueta</button></p>'
        panel += (f'<p>Ficha: {escape(c["archivo"])}<br>Empaque: {escape(c["capacidad"])}' + (f' = {c["litros"]:.3g} L'.replace('.', ',') if not c['capacidad'].endswith(' L') else '')
                  + f' ({escape(c["tamano"]["capacidad"].lower().replace("litros", "L"))})<br>Etiqueta: {ancho:g} × {alto:g} mm<br>Pictogramas: {lado:g} × {lado:g} mm<br>'
                  f'Letra: <span id="letra">—</span> (mínimo configurado: {minimo_pt:g} pt)<br>Nombre: {escape(c["origen_del_nombre"])}</p>')
        panel += '<div id="quitado"></div>'
        panel += '<p><b>Para imprimir a tamaño real:</b> hoja tamaño carta, escala 100 %, sin «ajustar a la página». Solo sale la etiqueta; este panel no se imprime.</p>'
    panel += '<h2>Avisos</h2><ul>' + ''.join(f'<li>{escape(a)}</li>' for a in c['avisos']) + '</ul>'
    if c['renglones']:
        panel += '<h2>De dónde sale cada renglón</h2><ul>' + ''.join(f'<li>{escape(z["codigo"])}: {escape(z["criterio"])}</li>' for z in c['renglones']) + '</ul>'
    panel += '</div>'
    return (f'<!doctype html><html lang="es"><head><meta charset="utf-8"><title>Etiqueta · {escape(c["nombre"])}</title><style>{ESTILO}</style></head>'
            f'<body><div class="hoja">{etiqueta or ""}{mensaje or ""}{panel}</div><script>{AJUSTE}</script></body></html>')


def guardar_etiqueta_mina(html, contenido, sufijo=''):
    destino = os.path.join(SALIDAS, 'etiquetas_mina')
    os.makedirs(destino, exist_ok=True)
    capacidad = re.sub(r'[^\w]+', '_', contenido['capacidad'] or 'sin_capacidad').strip('_')
    ruta = os.path.join(destino, f"{os.path.splitext(contenido['archivo'])[0]}_{capacidad}{sufijo}.html")
    with open(ruta, 'w', encoding='utf-8') as f:
        f.write(html)
    return ruta


if __name__ == '__main__':
    sys.stdout.reconfigure(encoding='utf-8')
    argumentos = sys.argv[1:]
    tablas = cargar_tablas()
    valor = lambda bandera, defecto: argumentos[argumentos.index(bandera) + 1] if bandera in argumentos else defecto
    elecciones = {'capacidad': float(str(valor('--capacidad', '20')).replace(',', '.')), 'unidad': valor('--unidad', 'L')}
    for ficha in cargar_fichas(argumentos[0] if argumentos and not argumentos[0].startswith('--') else ''):
        c = contenido_de_etiqueta_mina(ficha, clasificar(ficha, tablas), elecciones, tablas)
        print('Etiqueta guardada en', guardar_etiqueta_mina(dibujar_etiqueta_mina(c, tablas), c))
