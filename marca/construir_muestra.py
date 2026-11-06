# -*- coding: utf-8 -*-
"""Fase 2 de QuimNexo: genera la página de muestra de la identidad visual (marca/muestra_identidad.html) con las tres
propuestas de símbolo, la paleta con su contraste medido y la escala tipográfica. También guarda cada símbolo como SVG
en marca/propuestas/. Todo es dibujo propio, hecho a mano con geometría simple.
Uso:  py marca/construir_muestra.py"""
import os

AQUI = os.path.dirname(os.path.abspath(__file__))
PETROLEO, ACCION, ACENTO = '#0F3D4C', '#0B6E82', '#17A2B8'
PALETA = [
    ('--qn-marca', PETROLEO, 'Azul petróleo oscuro: marca, títulos y cabecera'),
    ('--qn-marca-profundo', '#0A2A35', 'Fondo oscuro del logotipo y del paso activo'),
    ('--qn-accion', ACCION, 'Turquesa oscuro: botones, enlaces y navegación'),
    ('--qn-acento', ACENTO, 'Turquesa: detalle del símbolo y marcas grandes; no se usa para texto pequeño'),
    ('--qn-accion-suave', '#E3F1F4', 'Fondo de elementos seleccionados'),
    ('--qn-texto', '#17232B', 'Texto principal'),
    ('--qn-texto-suave', '#44525C', 'Texto secundario y metadatos'),
    ('--qn-borde', '#C9D3DA', 'Bordes de casillas y separadores'),
    ('--qn-fondo', '#F4F6F8', 'Fondo de la página'),
    ('--qn-superficie', '#FFFFFF', 'Tarjetas y formularios'),
]
RESERVADOS = [('Verde', '#1A7F37', 'Semáforo: Ir menor que 0,5'), ('Ámbar', '#E3B100', 'Semáforo: Ir de 0,5 a menos de 0,85, y avisos'),
              ('Rojo', '#B3261E', 'Semáforo: Ir de 0,85 o más, y alertas'), ('Gris', '#8A94A0', 'Semáforo: sin medición')]
NIVELES = [('0', '#EEF1F5', '#3C4A5C'), ('1', '#DBE9F9', '#123A63'), ('2', '#B7D3F2', '#0B2A4F'), ('3', '#1F5FA8', '#FFFFFF'), ('4', '#0B2A4F', '#FFFFFF')]


def luminancia(color):
    canales = [int(color[i:i + 2], 16) / 255 for i in (1, 3, 5)]
    lineal = [c / 12.92 if c <= 0.03928 else ((c + 0.055) / 1.055) ** 2.4 for c in canales]
    return 0.2126 * lineal[0] + 0.7152 * lineal[1] + 0.0722 * lineal[2]


def contraste(a, b):
    """Razón de contraste WCAG entre dos colores."""
    la, lb = sorted((luminancia(a), luminancia(b)), reverse=True)
    return (la + 0.05) / (lb + 0.05)


def simbolos(principal=PETROLEO, acento=ACENTO, fondo='#FFFFFF'):
    """Las tres propuestas, en un lienzo de 64 × 64. 'principal' y 'acento' cambian para las versiones monocroma y de fondo oscuro."""
    return {
        'A': ('Eslabón', 'Dos piezas enlazadas: lo que dice la ficha y lo que se decide en la mina quedan unidos.',
              f'<rect x="7" y="7" width="34" height="34" rx="9" fill="none" stroke="{principal}" stroke-width="7"/>'
              f'<rect x="23" y="23" width="34" height="34" rx="9" fill="none" stroke="{acento}" stroke-width="7"/>'
              f'<path d="M14 41 H30" fill="none" stroke="{principal}" stroke-width="7"/>'),
        'B': ('Monograma N', 'La N de «nexo» como un puente: dos apoyos firmes (el dato y la acción) unidos por un trazo.',
              f'<rect x="2" y="2" width="60" height="60" rx="15" fill="{principal}"/>'
              f'<path d="M20 46 V18 L44 46 V18" fill="none" stroke="{fondo}" stroke-width="7" stroke-linecap="round" stroke-linejoin="round"/>'
              f'<circle cx="20" cy="18" r="5.500" fill="{acento}"/><circle cx="44" cy="46" r="5.500" fill="{acento}"/>'),
        'C': ('Convergencia', 'Tres líneas que llegan a un solo punto: peligros, consejos y límites de la ficha convertidos en una decisión.',
              f'<path d="M7 13 H20 C34 13 31 32 45 32 M7 32 H45 M7 51 H20 C34 51 31 32 45 32" fill="none" stroke="{principal}" stroke-width="6.500" stroke-linecap="round"/>'
              f'<circle cx="50" cy="32" r="9" fill="{acento}"/>'),
    }


def svg(interior, lado=64):
    return f'<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 64 64" width="{lado}" height="{lado}" role="img">{interior}</svg>'


def construir():
    os.makedirs(os.path.join(AQUI, 'propuestas'), exist_ok=True)
    claro, mono, oscuro = simbolos(), simbolos('#17232B', '#17232B'), simbolos('#FFFFFF', '#5FD3E3', '#0A2A35')
    for letra, (nombre, _, interior) in claro.items():
        with open(os.path.join(AQUI, 'propuestas', f'simbolo_{letra}_{nombre.lower().replace(" ", "_").replace("ó", "o")}.svg'), 'w', encoding='utf-8') as f:
            f.write(svg(interior) + '\n')

    def tarjeta(letra):
        nombre, idea, interior = claro[letra]
        marca = lambda color_quim, color_nexo: (f'<span class="palabra" style="color:{color_quim}">Quim<b style="color:{color_nexo}">Nexo</b></span>')
        return f'''<section class="propuesta"><h3>Propuesta {letra} · {nombre}</h3><p class="idea">{idea}</p>
 <div class="fila"><div class="grande">{svg(interior, 132)}</div>
  <div class="usos">
   <div class="uso claro">{svg(interior, 44)}{marca(PETROLEO, ACCION)}</div>
   <div class="uso oscuro">{svg(oscuro[letra][2], 44)}{marca('#FFFFFF', '#5FD3E3')}</div>
   <div class="uso claro">{svg(mono[letra][2], 44)}{marca('#17232B', '#17232B')}<small>una tinta</small></div>
   <div class="uso claro tamanos">{svg(interior, 32)}<small>32 px</small>{svg(interior, 16)}<small>16 px</small>{svg(mono[letra][2], 16)}<small>16 px, una tinta</small></div>
  </div></div></section>'''

    def muestra(variable, color, uso):
        sobre_blanco, con_blanco = contraste(color, '#FFFFFF'), contraste(color, '#17232B')
        claro_ = luminancia(color) > 0.4
        return (f'<tr><td><span class="chip" style="background:{color};color:{"#17232B" if claro_ else "#FFFFFF"}">{color}</span></td><td><code>{variable}</code></td><td>{uso}</td>'
                f'<td>{sobre_blanco:.1f} : 1</td><td>{con_blanco:.1f} : 1</td></tr>')

    pares = [('Texto principal sobre el fondo', '#17232B', '#F4F6F8', 4.5), ('Texto secundario sobre el fondo', '#44525C', '#F4F6F8', 4.5),
             ('Texto secundario sobre tarjeta', '#44525C', '#FFFFFF', 4.5), ('Texto blanco sobre el botón (acción)', '#FFFFFF', ACCION, 4.5),
             ('Texto blanco sobre la marca', '#FFFFFF', PETROLEO, 4.5), ('Enlace o texto de acción sobre tarjeta', ACCION, '#FFFFFF', 4.5),
             ('Borde de casilla sobre tarjeta', '#7C8B96', '#FFFFFF', 3.0), ('Acento del símbolo sobre blanco (gráfico)', ACENTO, '#FFFFFF', 3.0),
             ('Texto sobre verde del semáforo', '#FFFFFF', '#1A7F37', 4.5), ('Texto oscuro sobre ámbar del semáforo', '#17232B', '#E3B100', 4.5),
             ('Texto blanco sobre rojo del semáforo', '#FFFFFF', '#B3261E', 4.5)]
    filas_de_contraste = ''.join(
        f'<tr><td>{nombre}</td><td><span class="chip" style="background:{fondo};color:{texto};border:1px solid #C9D3DA">Aa 18 px</span></td><td>{contraste(texto, fondo):.1f} : 1</td>'
        f'<td>{minimo:g} : 1</td><td class="{"bien" if contraste(texto, fondo) >= minimo else "mal"}">{"Cumple" if contraste(texto, fondo) >= minimo else "NO cumple"}</td></tr>'
        for nombre, texto, fondo, minimo in pares)
    escala = [('Título de pantalla', 36, 700), ('Sección', 26, 700), ('Tarjeta', 21, 700), ('Texto, formularios e instrucciones', 18, 400), ('Botones', 18, 600),
              ('Tablas', 17, 400), ('Metadatos', 16, 400)]
    html = f'''<!doctype html><html lang="es"><head><meta charset="utf-8"><title>QuimNexo · identidad visual (Fase 2)</title><style>
@font-face{{font-family:"Inter";src:url("../clasificador%20SGA/static/fuentes/InterVariable.woff2") format("woff2");font-weight:100 900}}
*{{box-sizing:border-box}}body{{font-family:"Inter","Segoe UI",Arial,sans-serif;background:#F4F6F8;color:#17232B;margin:0;padding:32px 40px;font-size:18px;line-height:1.5}}
h1{{font-size:36px;margin:0 0 4px;color:{PETROLEO};letter-spacing:-.02em}}h2{{font-size:26px;margin:40px 0 10px;color:{PETROLEO}}}h3{{font-size:21px;margin:0 0 4px;color:{PETROLEO}}}
p{{margin:0 0 10px}}.idea{{color:#44525C;max-width:760px}}small{{font-size:15px;color:#44525C;margin:0 14px 0 6px}}
.propuesta{{background:#fff;border:1px solid #C9D3DA;border-radius:16px;padding:22px 26px;margin:0 0 18px}}
.fila{{display:flex;gap:28px;align-items:center;flex-wrap:wrap}}.grande{{padding:14px;border:1px dashed #C9D3DA;border-radius:12px;line-height:0}}
.usos{{display:grid;grid-template-columns:repeat(2,minmax(300px,1fr));gap:10px;flex:1}}.uso{{display:flex;align-items:center;gap:12px;padding:12px 16px;border-radius:12px;border:1px solid #C9D3DA;min-height:70px}}
.uso.oscuro{{background:#0A2A35;border-color:#0A2A35}}.palabra{{font-size:30px;font-weight:500;letter-spacing:-.02em}}.palabra b{{font-weight:800}}.tamanos svg{{flex:none}}
table{{border-collapse:collapse;background:#fff;border-radius:12px;overflow:hidden;width:100%;max-width:1100px}}th,td{{text-align:left;padding:9px 14px;border-bottom:1px solid #E3E8EC;font-size:17px;vertical-align:middle}}
th{{font-size:15px;color:#44525C;font-weight:600}}.chip{{display:inline-block;padding:6px 12px;border-radius:8px;font-weight:600;font-size:16px;min-width:96px;text-align:center}}
code{{font-size:16px}}.bien{{color:#1A7F37;font-weight:700}}.mal{{color:#B3261E;font-weight:700}}.nivel{{display:inline-block;width:54px;padding:8px 0;text-align:center;border-radius:8px;font-weight:700;margin-right:6px}}
.boton{{display:inline-block;background:{ACCION};color:#fff;font-weight:600;font-size:18px;padding:12px 26px;border-radius:999px;margin-right:10px}}.boton.dos{{background:#fff;color:{ACCION};border:2px solid {ACCION}}}
</style></head><body>
<h1>QuimNexo · identidad visual</h1><p class="idea">Fase 2. Tres propuestas de símbolo para elegir una, la paleta con su contraste medido y la escala de letra. Nada de esto está aplicado todavía a la aplicación.</p>
<h2>1. Símbolo: tres propuestas</h2>{''.join(tarjeta(l) for l in 'ABC')}
<h2>2. Paleta</h2><p class="idea">La marca usa azul petróleo y turquesa. El verde, el ámbar y el rojo quedan reservados para el semáforo y los avisos; los niveles de peligro siguen en tonos de azul, como hoy.</p>
<table><tr><th>Color</th><th>Variable</th><th>Uso</th><th>Contraste con blanco</th><th>Contraste con el texto</th></tr>{''.join(muestra(*p) for p in PALETA)}</table>
<h3 style="margin-top:22px">Reservados (no cambian; son los que ya usan el informe y la pantalla)</h3>
<p>{''.join(f'<span class="chip" style="background:{c};color:{"#17232B" if n == "Ámbar" else "#fff"};margin-right:8px">{n}</span><small>{u}</small>' for n, c, u in RESERVADOS)}</p>
<p>Niveles de peligro: {''.join(f'<span class="nivel" style="background:{f};color:{t}">{n}</span>' for n, f, t in NIVELES)}</p>
<h3 style="margin-top:22px">Contraste medido (WCAG AA)</h3>
<table><tr><th>Combinación</th><th>Muestra</th><th>Contraste</th><th>Mínimo</th><th>Resultado</th></tr>{filas_de_contraste}</table>
<h2>3. Tipografía: Inter, guardada en el proyecto</h2>
<table><tr><th>Uso</th><th>Tamaño</th><th>Muestra</th></tr>{''.join(f'<tr><td>{u}</td><td>{t} px</td><td style="font-size:{t}px;font-weight:{p};line-height:1.3">Riesgo por exposición · Ir = 0,62</td></tr>' for u, t, p in escala)}</table>
<p style="margin-top:16px">Interlineado de 1,5 en el texto. <span class="boton">Continuar: Resultado →</span><span class="boton dos">← Atrás</span></p>
</body></html>'''
    with open(os.path.join(AQUI, 'muestra_identidad.html'), 'w', encoding='utf-8') as f:
        f.write(html)
    return [(nombre, f'{contraste(texto, fondo):.2f}', minimo, contraste(texto, fondo) >= minimo) for nombre, texto, fondo, minimo in pares]


if __name__ == '__main__':
    import sys
    sys.stdout.reconfigure(encoding='utf-8')
    for nombre, valor, minimo, cumple in construir():
        print(f"{'cumple   ' if cumple else 'NO CUMPLE'} {valor:>6} (mínimo {minimo:g})  {nombre}")
    print('Página de muestra:', os.path.join(AQUI, 'muestra_identidad.html'))
