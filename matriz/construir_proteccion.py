# -*- coding: utf-8 -*-
"""Genera o actualiza proteccion_seccion8.xlsx: la protección personal que cada FDS da en su sección 8, con el
texto original (tal como lo leyó la extracción), su página y una versión corta en español de Colombia.
La herramienta no traduce sola. La versión en español de las 32 fichas del proyecto es un borrador de este
guion, para que el autor lo valide o lo corrija EN EL EXCEL: al volver a ejecutar, lo que el autor haya escrito
en las columnas «Protección personal en español» y «Estado» se conserva mientras el texto original no cambie.
Uso:  py matriz/construir_proteccion.py
Requiere: pip install openpyxl, y haber corrido antes  py extraccion/proteccion.py FDS --guardar"""
import glob, json, os, sys
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

B = os.path.dirname(os.path.abspath(__file__)) + os.sep
SALIDAS = os.path.join(B, '..', 'extraccion', 'salidas')
DESTINO = B + 'proteccion_seccion8.xlsx'
BORRADOR = 'Traducción resumida por validar del autor'
BORRADOR_ES = 'Resumen del texto en español por validar del autor'
COMUN = 'Piel y manos: guantes y ropa de protección adecuados para evitar el contacto. '
OJOS = 'Ojos: gafas de seguridad o monogafas para productos químicos. '
RESP = 'Respiratoria: respirador aprobado{} si se superan los valores límite o si hay irritación. '
VENT = 'Ventilación adecuada, sobre todo en espacios confinados'
LAVA = 'lavaojos y ducha de seguridad cerca del puesto de trabajo'
# Borradores por número de ficha. Resumen fiel del apartado: no agrega equipos que la ficha no nombra.
ESPANOL = {
 '01': 'Ojos: monogafas de seguridad (EN 166). Manos: guantes de protección (caucho natural, nitrilo, neopreno o PVC). Piel: ropa de protección adecuada. '
       f'Respiratoria: respirador certificado si la concentración supera el valor límite. {VENT}; {LAVA}.',
 '02': 'Respiratoria: si los controles de ingeniería no mantienen la concentración en un nivel seguro, usar protección respiratoria adecuada a las condiciones de uso. '
       'Ventilación a prueba de explosión y extracción localizada; lavaojos y duchas de emergencia. (El apartado de la ficha continúa: consulte allí ojos y manos.)',
 '03': 'Manos: guantes de trabajo al manipular cilindros. Ojos: gafas de seguridad con protección lateral. Piel: ropa de protección adecuada, por ejemplo overol o ropa '
       'resistente a la llama. Respiratoria: no se necesita en operaciones normales. Calzado de seguridad al manipular cilindros. Usar detectores de oxígeno donde '
       'puedan liberarse gases asfixiantes; ventilación general y localizada.',
 '04': 'Manos: guantes de trabajo al manipular cilindros. Ojos: gafas de seguridad con protección lateral. Piel: ropa de protección adecuada, por ejemplo overol o ropa '
       'resistente a la llama. Respiratoria: usar respirador en tareas no rutinarias (como abrir líneas o tomar muestras), y en las rutinarias si una revisión del '
       'proceso lo indica. Calzado de seguridad al manipular cilindros. Usar detectores con alarma donde puedan liberarse gases tóxicos; ventilación general y localizada.',
 '05': OJOS + COMUN + RESP.format(' con filtro de partículas') + VENT + '.',
 '06': 'Ojos y cara: monogafas de ajuste hermético y careta. ' + COMUN + 'Respiratoria: no se necesita en condiciones normales de uso; filtro recomendado para gases y '
       f'vapores orgánicos, tipo A (marrón). {VENT}; equipos eléctricos, de ventilación y de iluminación a prueba de explosión; {LAVA}.',
 '07': 'Manos: guantes de protección. Ojos: gafas de seguridad. Piel: ropa de protección adecuada. Respiratoria: equipo respiratorio adecuado si la ventilación es '
       'insuficiente. Buena ventilación del puesto de trabajo.',
 '08a': 'Ojos: monogafas de ajuste hermético. Manos: guantes de caucho de nitrilo (espesor mínimo 0,11 mm). Cuerpo: ropa de protección. Respiratoria: necesaria cuando '
        'se genera polvo; filtro tipo P2.',
 '08b': OJOS + COMUN + RESP.format(' con filtro de partículas') + f'Usar solo bajo campana de extracción; {VENT.lower()}; {LAVA}.',
 '09': OJOS + COMUN + 'Respiratoria: no se necesita en condiciones normales de uso; filtro recomendado de partículas.',
 '10': OJOS + COMUN + RESP.format(' con filtro de partículas') + f'{VENT}; {LAVA}.',
 '11': OJOS + COMUN + RESP.format(' con filtro de partículas y de gases y vapores inorgánicos, tipo B (gris),') + VENT + '.',
 '12': 'Ojos: gafas de seguridad con protección lateral o monogafas. Manos: guantes de protección (norma EN 374). Piel: ropa de protección adecuada. Respiratoria: no se '
       'necesita en condiciones normales de uso; si se superan los valores límite o hay irritación, puede requerirse ventilación y evacuación. Duchas, lavaojos y ventilación.',
 '13': 'Ojos: gafas de seguridad con protección lateral o monogafas. Manos: guantes de caucho butilo (norma EN 374). Piel: ropa de protección adecuada. Respiratoria: no se '
       'necesita en condiciones normales de uso; si se superan los valores límite o hay irritación, puede requerirse ventilación y evacuación. Duchas, lavaojos y ventilación.',
 '14': OJOS + COMUN + RESP.format(' con filtro de partículas y de gases ácidos, tipo E (amarillo),') + f'Manipular en un lugar con extracción localizada; {LAVA}.',
 '15': 'Respiratoria: equipo autónomo de respiración (SCBA) o máscara con línea de aire si el oxígeno está por debajo de 19,5 % o en emergencias por escape de gas; los '
       'purificadores de aire no protegen lo suficiente. Guantes industriales libres de aceite y grasa, gafas de seguridad, calzado con puntera de acero y ropa de algodón.',
 '16': OJOS + COMUN + RESP.format('') + VENT + '.',
 '17': 'Ojos: gafas de seguridad con protección lateral o monogafas (EN 166). Manos: guantes de protección (Viton). Piel: ropa de manga larga. Respiratoria: respirador '
       f'certificado si la concentración supera el valor límite. {VENT}; {LAVA}.',
 '18': OJOS + COMUN + RESP.format(' con filtro de partículas') + f'Usar solo bajo campana de extracción; {LAVA}.',
 '19': 'Respiratoria: no se requiere en condiciones normales. Manos: no se requieren guantes, salvo para evitar heridas por abrasión y reducir el contacto con la piel '
       '(neopreno o goma natural). Ojos: gafas de seguridad. Cuerpo: la indumentaria correspondiente. Reducir al mínimo el número de personas en el área; solo personal autorizado.',
 '20': OJOS + COMUN + RESP.format(' con filtro de partículas') + f'{VENT}; {LAVA}.',
 '21': 'Ojos: gafas de seguridad. Manos: guantes de caucho de nitrilo para salpicaduras (0,4 mm; resisten 10 minutos) o de Viton para contacto pleno (0,7 mm; 480 minutos). '
       'Piel: ropa de protección antiestática y retardante de llama. Respiratoria: necesaria cuando se generan vapores o aerosoles; filtro tipo AX.',
 '22': OJOS + COMUN + RESP.format(' con filtro para gases y vapores orgánicos, tipo A (marrón),') + VENT + '.',
 '23': 'Respiratoria: necesaria cuando se generan aerosoles. Ojos: monogafas de seguridad. Manos y cuerpo: ropa de protección. Cambiar de inmediato la ropa contaminada; '
       'lavarse manos y cara después de trabajar con la sustancia.',
 '24': OJOS + COMUN + 'Respiratoria: respirador certificado si la concentración supera el valor límite; filtro para gases ácidos, tipo E (amarillo), o de partículas. '
       + LAVA.capitalize() + '.',
 '25': 'Ojos y cara: monogafas de ajuste hermético y careta. ' + COMUN + RESP.format(' con filtro de partículas') + f'Usar solo bajo campana de extracción; {LAVA}.',
 '26': OJOS + COMUN + RESP.format(' con filtro de partículas') + f'{VENT}; {LAVA}.',
 '27': OJOS + COMUN + 'Respiratoria: no se necesita en condiciones normales de uso; filtro recomendado de partículas.',
 '28': OJOS + COMUN + 'Respiratoria: no se necesita en condiciones normales de uso; filtro recomendado de partículas.',
 '29': 'Ojos: gafas de seguridad. Manos: guantes de caucho de nitrilo (0,11 mm; 480 minutos). Piel: ropa de protección. Respiratoria: necesaria cuando se generan vapores '
       'o aerosoles; filtro tipo ABEK-P.',
 '30': OJOS + COMUN + 'Respiratoria: no se necesita en condiciones normales de uso; filtro recomendado de partículas.',
 '31': 'Ojos y cara: antiparras con ventilación indirecta. Manos y piel: guantes de fluoroelastómero o de caucho de nitrilo y ropa de protección, según la evaluación de la '
       'exposición. Respiratoria: según esa evaluación puede necesitarse respirador (el apartado de la ficha continúa). No permanecer donde pueda faltar oxígeno; '
       'ventilación general o extracción localizada.',
}
EN_ESPANOL = {'15', '19', '31'}   # fichas cuyo apartado ya viene en español: aquí solo se resume

# Borradores para la etiqueta de uso en la mina: una línea corta y los iconos que la acompañan (gafas, careta, guantes, respirador, botas, ropa).
# Salen de la columna «Protección personal en español»; no agregan equipos que la ficha no nombra.
BORRADOR_CORTO = 'Borrador por validar del autor'
_SI_SUPERA = 'si se supera el valor límite o hay irritación'
_BASE = 'Use: gafas de seguridad o monogafas, guantes y ropa de protección'
EPP_CORTO = {
 '01': ('Use: monogafas, guantes y ropa de protección; respirador si se supera el valor límite.', 'gafas, guantes, ropa, respirador'),
 '02': ('Use: protección respiratoria si la ventilación no basta. Para ojos y manos, consulte la ficha.', 'respirador'),
 '03': ('Use: guantes de trabajo, gafas con protección lateral, ropa resistente a la llama y calzado de seguridad.', 'guantes, gafas, ropa, botas'),
 '04': ('Use: guantes de trabajo, gafas con protección lateral, ropa resistente a la llama y calzado de seguridad; respirador en tareas no rutinarias.',
        'guantes, gafas, ropa, botas, respirador'),
 '05': (f'{_BASE}; respirador con filtro de partículas {_SI_SUPERA}.', 'gafas, guantes, ropa, respirador'),
 '06': ('Use: monogafas herméticas y careta, guantes y ropa de protección.', 'gafas, careta, guantes, ropa'),
 '07': ('Use: guantes, gafas de seguridad y ropa de protección; respirador si la ventilación no basta.', 'guantes, gafas, ropa, respirador'),
 '08a': ('Use: monogafas herméticas, guantes de nitrilo y ropa de protección; respirador con filtro P2 si hay polvo.', 'gafas, guantes, ropa, respirador'),
 '08b': (f'{_BASE}; respirador con filtro de partículas {_SI_SUPERA}.', 'gafas, guantes, ropa, respirador'),
 '09': (f'{_BASE}.', 'gafas, guantes, ropa'),
 '10': (f'{_BASE}; respirador con filtro de partículas {_SI_SUPERA}.', 'gafas, guantes, ropa, respirador'),
 '11': (f'{_BASE}; respirador con filtro de partículas y de gases {_SI_SUPERA}.', 'gafas, guantes, ropa, respirador'),
 '12': ('Use: gafas con protección lateral, guantes y ropa de protección.', 'gafas, guantes, ropa'),
 '13': ('Use: gafas con protección lateral, guantes de caucho butilo y ropa de protección.', 'gafas, guantes, ropa'),
 '14': (f'{_BASE}; respirador para gases ácidos {_SI_SUPERA}.', 'gafas, guantes, ropa, respirador'),
 '15': ('Use: guantes sin aceite ni grasa, gafas de seguridad, calzado con puntera y ropa de algodón; equipo autónomo de respiración si falta oxígeno o hay escape de gas.',
        'guantes, gafas, botas, ropa, respirador'),
 '16': (f'{_BASE}; respirador {_SI_SUPERA}.', 'gafas, guantes, ropa, respirador'),
 '17': ('Use: gafas con protección lateral, guantes de Viton y ropa de manga larga; respirador si se supera el valor límite.', 'gafas, guantes, ropa, respirador'),
 '18': (f'{_BASE}; respirador con filtro de partículas {_SI_SUPERA}.', 'gafas, guantes, ropa, respirador'),
 '19': ('Use: gafas de seguridad y la indumentaria de trabajo; guantes solo para evitar abrasión.', 'gafas, ropa'),
 '20': (f'{_BASE}; respirador con filtro de partículas {_SI_SUPERA}.', 'gafas, guantes, ropa, respirador'),
 '21': ('Use: gafas de seguridad, guantes de nitrilo o de Viton y ropa antiestática; respirador con filtro AX si hay vapores.', 'gafas, guantes, ropa, respirador'),
 '22': (f'{_BASE}; respirador para vapores orgánicos {_SI_SUPERA}.', 'gafas, guantes, ropa, respirador'),
 '23': ('Use: monogafas y ropa de protección; respirador si se generan aerosoles.', 'gafas, ropa, respirador'),
 '24': (f'{_BASE}; respirador para gases ácidos si se supera el valor límite.', 'gafas, guantes, ropa, respirador'),
 '25': ('Use: monogafas herméticas y careta, guantes y ropa de protección; respirador con filtro de partículas ' + _SI_SUPERA + '.', 'gafas, careta, guantes, ropa, respirador'),
 '26': (f'{_BASE}; respirador con filtro de partículas {_SI_SUPERA}.', 'gafas, guantes, ropa, respirador'),
 '27': (f'{_BASE}.', 'gafas, guantes, ropa'),
 '28': (f'{_BASE}.', 'gafas, guantes, ropa'),
 '29': ('Use: gafas de seguridad, guantes de nitrilo y ropa de protección; respirador con filtro ABEK-P si hay vapores.', 'gafas, guantes, ropa, respirador'),
 '30': (f'{_BASE}.', 'gafas, guantes, ropa'),
 '31': ('Use: antiparras, guantes de nitrilo o de fluoroelastómero y ropa de protección; respirador según la evaluación de la exposición.', 'gafas, guantes, ropa, respirador'),
}

if __name__ == '__main__':
    sys.stdout.reconfigure(encoding='utf-8')
    anteriores = {}
    if os.path.isfile(DESTINO):   # lo que el autor ya revisó se conserva
        for fila in openpyxl.load_workbook(DESTINO, read_only=True, data_only=True)['Proteccion'].iter_rows(min_row=2, values_only=True):
            if fila and fila[0]:
                fila = tuple(fila) + (None,) * 9
                anteriores[fila[0]] = dict(original=fila[3], espanol=fila[4], estado=fila[5], corto=fila[6], iconos=fila[7], estado_corto=fila[8])
    filas, conservadas = [], 0
    for ruta in sorted(glob.glob(os.path.join(SALIDAS, '*.json'))):
        with open(ruta, encoding='utf-8') as f:
            ficha = json.load(f)
        p = ficha.get('proteccion') or {}
        numero = ficha['archivo'].split('_')[0]
        previo = anteriores.get(ficha['archivo'])
        corto, iconos = EPP_CORTO.get(numero, (None, None))
        estado_corto = BORRADOR_CORTO if corto else 'Sin texto corto'
        if previo and previo['original'] == p.get('texto'):
            espanol, estado = previo['espanol'], previo['estado']
            if previo['estado_corto']:   # las columnas del texto corto ya existían: se respeta lo que tenga el autor
                corto, iconos, estado_corto = previo['corto'], previo['iconos'], previo['estado_corto']
            conservadas += 1
        else:
            espanol = ESPANOL.get(numero)
            estado = (BORRADOR_ES if numero in EN_ESPANOL else BORRADOR) if espanol else 'Sin traducción'
        filas.append((ficha['archivo'], p.get('pagina'), 'español' if numero in EN_ESPANOL else 'inglés', p.get('texto'), espanol, estado, corto, iconos, estado_corto))

    wb = openpyxl.Workbook()
    leeme = wb.active
    leeme.title = 'LEEME'
    for fila in (('Tema', 'Detalle'),
                 ('Qué es', 'Protección personal que cada Ficha de Datos de Seguridad da en su sección 8: texto original, página y una versión corta en español de Colombia.'),
                 ('De dónde sale', 'El texto original lo copia la extracción (extraccion/proteccion.py), literal y con tope de 1800 caracteres. La versión en español es un borrador '
                                   'hecho por el asistente al construir la herramienta; la herramienta no traduce sola.'),
                 ('Qué debe hacer el autor', 'Revisar la columna «Protección personal en español» contra el texto original, corregirla aquí mismo si hace falta y cambiar el Estado a '
                                             '«Validado». Lo que escriba se conserva al regenerar el archivo, mientras el texto original de esa ficha no cambie.'),
                 ('Cómo la usa la herramienta', 'En el resultado muestra la versión en español con su Estado y la página de la ficha. Para una ficha que no esté en esta tabla '
                                                '(por ejemplo, una cargada por el usuario) muestra el texto original, rotulado como tal.'),
                 ('Columnas para la etiqueta de uso en la mina', '«EPP corto para etiqueta» es la línea que se imprime en la etiqueta («Use: …»). «Iconos de EPP» dice qué dibujos '
                                                              'la acompañan; valores posibles, separados por comas: gafas, careta, guantes, respirador, botas, ropa. '
                                                              'Las dos son borrador hasta que «Estado del EPP corto» diga «Validado».'),
                 ('Advertencia', 'Varias fichas son de reactivos de laboratorio y recomiendan «campana de extracción»: se tradujo tal cual. Si eso no aplica en la mina, es una '
                                 'razón para conseguir la ficha del proveedor real.')):
        leeme.append(fila)
    hoja = wb.create_sheet('Proteccion')
    hoja.append(('Ficha', 'Página', 'Idioma del original', 'Texto original de la ficha (sección 8)', 'Protección personal en español', 'Estado',
                 'EPP corto para etiqueta', 'Iconos de EPP', 'Estado del EPP corto'))
    for fila in filas:
        hoja.append(fila)
    for ws, anchos in ((leeme, (28, 130)), (hoja, (36, 8, 12, 90, 90, 30, 70, 30, 30))):
        for fila in ws.iter_rows():
            for c in fila:
                c.font = Font(name='Arial', size=10)
                c.alignment = Alignment(wrap_text=True, vertical='top')
        for c in ws[1]:
            c.font = Font(name='Arial', bold=True, color='FFFFFF')
            c.fill = PatternFill('solid', fgColor='1F3864')
        for i, ancho in enumerate(anchos, 1):
            ws.column_dimensions[get_column_letter(i)].width = ancho
        ws.freeze_panes = 'A2'
    wb.save(DESTINO)
    print(f'{len(filas)} fichas | con versión en español: {sum(bool(f[4]) for f in filas)} | filas del autor conservadas: {conservadas}')
