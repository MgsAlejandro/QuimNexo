# -*- coding: utf-8 -*-
"""Prueba de la interfaz sin abrir el navegador. Recorre los tres pasos con todos los productos y busca errores;
luego, con el diésel y datos de PRUEBA, comprueba la navegación, las validaciones, el semáforo, los dos perfiles,
el detalle técnico, los documentos, las mediciones guardadas y la lista «qué revisar primero».
Todo lo que escribe va a una carpeta temporal: no toca los datos del usuario.
Uso (desde la carpeta del proyecto):  py "clasificador SGA/pruebas_interfaz.py" """
import os, sys, tempfile
from streamlit.testing.v1 import AppTest

AQUI = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, AQUI)
import almacen
from lectura import PERFILES

RESULTADOS = []


def caso(nombre, condicion, detalle=''):
    RESULTADOS.append(bool(condicion))
    print(('  BIEN  ' if condicion else '  FALLA ') + nombre + (f'  [{detalle}]' if detalle and not condicion else ''))


def texto(at):
    return ' '.join(m.value for m in at.markdown)


def paso_actual(at):
    return at.session_state['paso']


def oprimir(at, inicio_del_rotulo):
    next(b for b in at.button if b.label.startswith(inicio_del_rotulo)).click().run()


def sano(at):
    return not at.exception and not at.error


def casilla(at, tipo, inicio):
    return next(c for c in getattr(at, tipo) if c.label.startswith(inicio))


def abrir(at, producto, paso=1):
    at.session_state['sustancia'], at.session_state['paso'], at.session_state['ver_detalle'] = producto, paso, False
    at.run()


if __name__ == '__main__':
    sys.stdout.reconfigure(encoding='utf-8')
    with tempfile.TemporaryDirectory() as temporal:
        # Lo que la aplicación guarde durante la prueba (elecciones, mediciones, consultas) va a la carpeta temporal
        almacen.AQUI, almacen.CONSULTAS, almacen.REGISTRO = temporal, os.path.join(temporal, 'consultas'), os.path.join(temporal, 'consultas', 'registro.csv')
        # PRUEBA: una elección guardada por una versión anterior, que solo trae los litros. Se deja antes de abrir la aplicación, como le pasaría al usuario
        almacen.guardar('elecciones', '06_Estireno_100-42-5.pdf', {'envase_litros': 20, 'uso_interno': False, 'borde': 'rojo', 'consejos_p': ['P210', 'P233', 'P240', 'P241', 'P242', 'P243', 'P260'],
                                                               'textos': {}, 'componentes': '', 'proveedor': {'nombre': '', 'direccion': '', 'telefono': ''}, 'cantidad': '', 'lote': ''})
        at = AppTest.from_file(os.path.join(AQUI, 'app.py'), default_timeout=180).run()

        print('\nPaso 1')
        caso('la aplicación abre en el paso 1 de 3, con el selector de perfil', sano(at) and paso_actual(at) == 1 and 'Paso 1 de 3' in texto(at)
             and list(casilla(at, 'radio', '¿Quién consulta?').options) == list(PERFILES))
        oprimir(at, 'Continuar')
        caso('sin producto no deja continuar y dice qué falta, dónde, por qué y qué hacer',
             paso_actual(at) == 1 and all(k in texto(at) for k in ('No se puede continuar', 'Qué falta', 'Dónde', 'Por qué se necesita', 'Qué hacer')))
        caso('el detalle técnico está en el panel izquierdo y espera a que haya producto', any(i.label == 'Ver el detalle técnico' and i.disabled for i in at.sidebar.toggle))
        caso('el pie explica el conteo de fichas', 'la 08 tiene dos' in texto(at))

        print('\nLos tres pasos con todos los productos, en los dos perfiles')
        opciones = list(casilla(at, 'selectbox', 'Producto').options)
        malos = []
        for i in range(len(opciones) - 1):
            abrir(at, i)
            at.session_state['perfil'] = PERFILES[i % 2]
            for paso in (2, 3):
                sin = [c for c in at.checkbox if c.label == 'No tengo medición']
                if paso == 3 and sin:
                    sin[0].set_value(True).run()
                oprimir(at, 'Continuar')
                if not sano(at) or paso_actual(at) != paso:
                    malos.append(f'{opciones[i + 1]} al llegar al paso {paso}: ' + '; '.join([str(e.value)[:120] for e in at.exception] + [e.value[:120] for e in at.error]) + texto(at)[-200:])
                    break
            else:
                at.sidebar.toggle[0].set_value(True).run()
                if not sano(at) or 'Detalle técnico' not in texto(at):
                    malos.append(f'{opciones[i + 1]}: falla el detalle técnico')
        caso(f'los {len(opciones) - 1} productos recorren los tres pasos y el detalle técnico sin errores', not malos, ' | '.join(malos)[:500])

        print('\nExposición y resultado con el diésel (datos de PRUEBA)')
        diesel = next(i for i, o in enumerate(opciones[1:]) if o.startswith('02 '))
        at.session_state['perfil'] = PERFILES[0]
        abrir(at, diesel, 2)
        if casilla(at, 'checkbox', 'No tengo medición').value:
            casilla(at, 'checkbox', 'No tengo medición').set_value(False).run()
        caso('paso 2: solo pide la medición (ya no pide actividad, tarea, cantidad ni controles)',
             paso_actual(at) == 2 and not any(s.label in ('Actividad', 'Frecuencia de uso') for s in at.selectbox) and not any(i.label.startswith(('Tarea', 'Cantidad', 'Controles')) for i in at.text_input))
        oprimir(at, 'Continuar')
        caso('sin medición y sin marcar «No tengo medición» no deja continuar', paso_actual(at) == 2 and 'No se puede continuar' in texto(at) and 'No tengo medición' in texto(at))
        casilla(at, 'number_input', 'Concentración medida').set_value(62.0)   # PRUEBA: 62 mg/m3 durante 8 h, contra un límite de 100 mg/m3
        casilla(at, 'number_input', 'Horas de exposición').set_value(12.0)    # PRUEBA: más horas que la jornada, para que la medición no sea válida
        at.run()
        oprimir(at, 'Continuar')
        caso('una medición mal escrita impide continuar y explica cómo corregirla', paso_actual(at) == 2 and 'No se puede continuar' in texto(at) and 'Horas de exposición' in texto(at))
        casilla(at, 'number_input', 'Horas de exposición').set_value(8.0)
        at.run()
        oprimir(at, 'Continuar')
        pantalla = texto(at)
        caso('con la medición corregida se llega al resultado, y la medición NO se guarda sola al continuar',
             paso_actual(at) == 3 and len(almacen.mediciones_de('02_Diesel_ACPM_68334-30-5.pdf')) == 0)
        caso('resultado: los dos indicadores lado a lado, semáforo amarillo con su mensaje y el número del índice',
             all(k in pantalla for k in ('Peligro del producto', 'Riesgo por exposición', 'Amarillo', 'Ir = 0,62')) and pantalla.index('Peligro del producto') < pantalla.index('Acción recomendada'))
        caso('perfil profesional de SST: acción recomendada, por verificar y acceso al detalle técnico',
             'Acción recomendada' in pantalla and any(e.label.startswith('Por verificar') for e in at.expander) and any(b.label == 'Abrir el detalle técnico' for b in at.button))
        caso('resultado: qué puede pasar, qué hacer, protección personal y qué hacer en un accidente',
             all(k in pantalla for k in ('Qué puede pasar', 'Qué hacer', 'Protección personal', 'Si hay un accidente')))
        caso('el método y los factores no están en el resultado', 'Cómo se obtuvo' not in pantalla and '¿Qué representa?' not in pantalla)

        at.sidebar.radio[0].set_value(PERFILES[1]).run()
        pantalla = texto(at)
        caso('perfil trabajador o supervisor: mismos indicadores y frases, sin acción recomendada, pendientes ni detalle',
             sano(at) and all(k in pantalla for k in ('Peligro del producto', 'Riesgo por exposición', 'Qué puede pasar', 'Qué hacer', 'Si hay un accidente'))
             and 'Acción recomendada' not in pantalla and not any(e.label.startswith('Por verificar') for e in at.expander)
             and not any(b.label == 'Abrir el detalle técnico' for b in at.button))
        at.sidebar.radio[0].set_value(PERFILES[0]).run()

        print('\nSemáforo rojo y lista «qué revisar primero»')
        oprimir(at, '← Atrás')
        casilla(at, 'number_input', 'Concentración medida').set_value(85.0)   # PRUEBA: Ir = 0,85, nivel de intervención
        at.run()
        oprimir(at, 'Continuar')
        pantalla = texto(at)
        caso('Ir = 0,85 da rojo con la alerta de paro inmediato', 'Rojo' in pantalla and 'ALERTA DE PARO INMEDIATO' in pantalla and 'Ir = 0,85' in pantalla)
        caso('cambiar la medición y continuar tampoco la guarda', len(almacen.mediciones_de('02_Diesel_ACPM_68334-30-5.pdf')) == 0)
        oprimir(at, 'Guardar en el historial')
        guardadas = almacen.consultas()
        caso('«Guardar en el historial» deja la consulta en el registro, con el color', len(guardadas) == 1 and guardadas[0]['Riesgo por exposición'] == 'Rojo')
        medidas = almacen.mediciones_de('02_Diesel_ACPM_68334-30-5.pdf')
        caso('«Guardar en el historial» guarda también la medición de esta consulta, con fecha',
             len(medidas) == 1 and list(medidas[0]['mediciones'].values())[0]['tareas'] == [[85.0, 8.0]])
        caso('al final del resultado está la descarga del informe', 'Descargar informe' in [b.label for b in at.get('download_button')])
        etiqueta = lambda: next(b for b in at.get('download_button') if b.label == 'Descargar etiqueta para imprimir')
        caso('el botón de la etiqueta está a la vista junto al del informe; sin capacidad del empaque está apagado y dice qué falta',
             etiqueta().disabled and 'Para activar la descarga, escriba la capacidad del empaque' in texto(at) and 'Falta un dato para ver la etiqueta' in texto(at))
        caso('la unidad se elige entre las de la hoja Unidades_empaque', list(casilla(at, 'selectbox', 'Unidad').options) == ['mL', 'L', 'galón (US)', 'g', 'kg'])
        casilla(at, 'number_input', 'Capacidad del empaque').set_value(20.0)
        casilla(at, 'selectbox', 'Unidad').set_value('galón (US)').run()
        notas = ' '.join(c.value for c in at.caption)
        caso('con la capacidad el botón se activa y la pantalla muestra la conversión («20 gal (US) = 75,71 L → etiqueta de 105 × 148 mm»)',
             sano(at) and not etiqueta().disabled and '20 gal (US) = 75,71 L' in notas and 'etiqueta de 105 × 148 mm' in notas, notas[-300:])
        guardado = almacen.leer('elecciones', '02_Diesel_ACPM_68334-30-5.pdf')
        caso('la capacidad y la unidad quedan guardadas (y los litros, para las versiones anteriores)',
             guardado.get('capacidad') == 20.0 and guardado.get('unidad') == 'galón (US)' and abs(guardado.get('envase_litros') - 75.70823568) < 1e-6)
        caso('la etiqueta completa (Res. 773) no aparece hasta que se pide', 'Descargar etiqueta completa (Res. 773)' not in [b.label for b in at.get('download_button')])
        casilla(at, 'toggle', 'Etiqueta completa (Res. 773)').set_value(True).run()
        marcados = [c for c in at.checkbox if c.label.startswith('**P') and c.value]
        caso('al pedirla aparece como opción secundaria, con los cinco consejos de la selección minera marcados',
             sano(at) and 'Descargar etiqueta completa (Res. 773)' in [b.label for b in at.get('download_button')] and len(marcados) == 5, f'{len(marcados)} marcados')
        visible = ' '.join([texto(at)] + [c.value for c in at.caption] + [w.label for tipo in ('number_input', 'selectbox', 'checkbox', 'button', 'text_input', 'toggle')
                                                                         for w in getattr(at, tipo)] + [e.label for e in at.expander])
        caso('en la pantalla del resultado ya no aparece la palabra «envase»', 'envase' not in visible.lower())

        estireno = next(i for i, o in enumerate(opciones[1:]) if o.startswith('06 '))
        abrir(at, estireno, 2)
        casilla(at, 'checkbox', 'No tengo medición').set_value(True).run()
        oprimir(at, 'Continuar')
        caso('una elección vieja sin unidad se lee como 20 L y la etiqueta queda lista',
             sano(at) and casilla(at, 'number_input', 'Capacidad del empaque').value == 20.0 and casilla(at, 'selectbox', 'Unidad').value == 'L' and not etiqueta().disabled)
        casilla(at, 'toggle', 'Etiqueta completa (Res. 773)').set_value(True).run()
        caso('y en la etiqueta completa se respetan los consejos que traía marcados', len([c for c in at.checkbox if c.label.startswith('**P') and c.value]) == 7)
        abrir(at, diesel, 3)
        oprimir(at, 'Empezar otra consulta')
        pantalla = texto(at)
        primera_fila = pantalla[pantalla.index('<th>N.º</th>'):].split('<tr>')[1] if '<th>N.º</th>' in pantalla else ''
        caso('la lista «qué revisar primero» pone de primero al diésel en rojo, con el semáforo y el nivel lado a lado',
             paso_actual(at) == 1 and 'Combustible diésel' in primera_fila and 'Rojo' in primera_fila and 'Nivel 3' in primera_fila, primera_fila[:200])
        caso('y aparece la consulta guardada', any('Consultas guardadas · 1' in e.label for e in at.expander))

        print('\nDetalle técnico y conservación de datos')
        abrir(at, diesel, 2)
        caso('una consulta nueva del mismo producto empieza en blanco (no trae la medición guardada)',
             casilla(at, 'number_input', 'Concentración medida').value is None and not casilla(at, 'checkbox', 'No tengo medición').value)
        oprimir(at, 'Continuar')
        caso('y sin escribir nada no deja continuar: no usa la medición guardada por su cuenta', paso_actual(at) == 2 and 'No se puede continuar' in texto(at))
        casilla(at, 'number_input', 'Concentración medida').set_value(30.0)   # PRUEBA: Ir = 0,30
        at.run()
        oprimir(at, 'Continuar')
        caso('el resultado usa solo lo escrito en esta consulta (Ir = 0,3, verde)', paso_actual(at) == 3 and 'Verde' in texto(at) and 'Ir = 0,3' in texto(at))
        oprimir(at, '← Atrás')
        caso('al volver al paso 2 dentro de la misma consulta las casillas conservan lo escrito', casilla(at, 'number_input', 'Concentración medida').value == 30.0)
        at.sidebar.toggle[0].set_value(True).run()
        rotulos = ' '.join(e.label for e in at.expander)
        caso('el detalle técnico conserva todo: frases por confirmar, clasificación, consejos y criterios, semáforo y método, historial, protección, sección 10, reglas',
             sano(at) and all(k in rotulos for k in ('Frases de la ficha por confirmar', 'Clasificación de peligros', 'Consejos de prudencia', 'Riesgo por exposición',
                                                     'Historial de mediciones · 1', 'Protección personal', 'Estabilidad y reactividad', 'Reglas aplicadas')))
        oprimir(at, '← Volver al paso')
        caso('«Volver» regresa al mismo paso de la consulta', paso_actual(at) == 2 and not at.session_state['ver_detalle'])

        print('\nSin medición')
        sin_medir = next(i for i, o in enumerate(opciones[1:]) if o.startswith('14 '))
        abrir(at, sin_medir, 2)
        casilla(at, 'checkbox', 'No tengo medición').set_value(True).run()
        oprimir(at, 'Continuar')
        pantalla = texto(at)
        caso('con «No tengo medición» se continúa y el semáforo dice «Sin medición» y pide medir con urgencia (nivel de peligro alto)',
             paso_actual(at) == 3 and 'Sin medición' in pantalla and 'urgencia' in pantalla)
    print(f'\nRESULTADO: {sum(RESULTADOS)} de {len(RESULTADOS)} comprobaciones bien.')
    sys.exit(0 if all(RESULTADOS) else 1)
