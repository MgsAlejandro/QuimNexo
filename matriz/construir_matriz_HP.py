# -*- coding: utf-8 -*-
"""Genera matriz_decision_HP.xlsx a partir de las reglas definidas aquí y de la
hoja "reaccio" del Excel de fichas. Ejecutar de nuevo si se cambia un nivel.
Requiere: pip install openpyxl"""
import openpyxl, glob, re, os
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

B = os.path.dirname(os.path.abspath(__file__)) + os.sep
src = glob.glob(B + 'fichas*.xlsx')[0]
NIV = {0: 'No clasificado', 1: 'Bajo', 2: 'Medio', 3: 'Alto', 4: 'Muy alto'}
F = 'Físico'; S = 'Salud'; A = 'Ambiente'
# (código, texto, tipo, clase, categoría, palabra de advertencia, pictograma, nivel propuesto)
H = [
('H200','Explosivo inestable',F,'Explosivos','Inestable','Peligro','GHS01',4),
('H201','Explosivo; peligro de explosión en masa',F,'Explosivos','División 1.1','Peligro','GHS01',4),
('H202','Explosivo; grave peligro de proyección',F,'Explosivos','División 1.2','Peligro','GHS01',4),
('H203','Explosivo; peligro de incendio, de onda expansiva o de proyección',F,'Explosivos','División 1.3','Peligro','GHS01',4),
('H204','Peligro de incendio o de proyección',F,'Explosivos','División 1.4','Atención','GHS01',3),
('H205','Peligro de explosión en masa en caso de incendio',F,'Explosivos','División 1.5','Peligro','Sin pictograma',4),
('H206','Peligro de incendio, de onda expansiva o de proyección; mayor riesgo de explosión si se reduce el agente insensibilizante',F,'Explosivos insensibilizados','1','Peligro','GHS02',3),
('H207','Peligro de incendio o de proyección; mayor riesgo de explosión si se reduce el agente insensibilizante',F,'Explosivos insensibilizados','2, 3','Peligro (cat. 2) / Atención (cat. 3)','GHS02',3),
('H208','Peligro de incendio; mayor riesgo de explosión si se reduce el agente insensibilizante',F,'Explosivos insensibilizados','4','Atención','GHS02',2),
('H220','Gas extremadamente inflamable',F,'Gases inflamables','1','Peligro','GHS02',4),
('H221','Gas inflamable',F,'Gases inflamables','2','Atención','Sin pictograma',3),
('H222','Aerosol extremadamente inflamable',F,'Aerosoles','1','Peligro','GHS02',3),
('H223','Aerosol inflamable',F,'Aerosoles','2','Atención','GHS02',2),
('H224','Líquido y vapores extremadamente inflamables',F,'Líquidos inflamables','1','Peligro','GHS02',4),
('H225','Líquido y vapores muy inflamables',F,'Líquidos inflamables','2','Peligro','GHS02',3),
('H226','Líquido y vapores inflamables',F,'Líquidos inflamables','3','Atención','GHS02',2),
('H227','Líquido combustible',F,'Líquidos inflamables','4','Atención','Sin pictograma',1),
('H228','Sólido inflamable',F,'Sólidos inflamables','1, 2','Peligro (cat. 1) / Atención (cat. 2)','GHS02',3),
('H229','Recipiente a presión: puede reventar si se calienta',F,'Aerosoles','1, 2, 3','(acompaña a H222/H223)','Sin pictograma propio',2),
('H230','Puede explotar incluso en ausencia de aire',F,'Gases químicamente inestables','A','(sin palabra propia)','Sin pictograma propio',4),
('H231','Puede explotar incluso en ausencia de aire, a presión y/o temperatura elevadas',F,'Gases químicamente inestables','B','(sin palabra propia)','Sin pictograma propio',3),
('H232','Puede inflamarse espontáneamente en contacto con el aire',F,'Gases pirofóricos','1','Peligro','GHS02',4),
('H240','Puede explotar al calentarse',F,'Sustancias autorreactivas / peróxidos orgánicos','Tipo A','Peligro','GHS01',4),
('H241','Puede incendiarse o explotar al calentarse',F,'Sustancias autorreactivas / peróxidos orgánicos','Tipo B','Peligro','GHS01, GHS02',4),
('H242','Puede incendiarse al calentarse',F,'Sustancias autorreactivas / peróxidos orgánicos','Tipos C a F','Peligro (C, D) / Atención (E, F)','GHS02',3),
('H250','Se inflama espontáneamente en contacto con el aire',F,'Líquidos y sólidos pirofóricos','1','Peligro','GHS02',4),
('H251','Se calienta espontáneamente; puede inflamarse',F,'Sustancias que experimentan calentamiento espontáneo','1','Peligro','GHS02',3),
('H252','Se calienta espontáneamente en grandes cantidades; puede inflamarse',F,'Sustancias que experimentan calentamiento espontáneo','2','Atención','GHS02',2),
('H260','En contacto con el agua desprende gases inflamables que pueden inflamarse espontáneamente',F,'Sustancias que con el agua desprenden gases inflamables','1','Peligro','GHS02',4),
('H261','En contacto con el agua desprende gases inflamables',F,'Sustancias que con el agua desprenden gases inflamables','2, 3','Peligro (cat. 2) / Atención (cat. 3)','GHS02',3),
('H270','Puede provocar o agravar un incendio; comburente',F,'Gases comburentes','1','Peligro','GHS03',3),
('H271','Puede provocar un incendio o una explosión; muy comburente',F,'Líquidos y sólidos comburentes','1','Peligro','GHS03',4),
('H272','Puede agravar un incendio; comburente',F,'Líquidos y sólidos comburentes','2, 3','Peligro (cat. 2) / Atención (cat. 3)','GHS03',3),
('H280','Contiene gas a presión; puede explotar si se calienta',F,'Gases a presión','Comprimido, licuado, disuelto','Atención','GHS04',2),
('H281','Contiene gas refrigerado; puede provocar quemaduras o lesiones criogénicas',F,'Gases a presión','Licuado refrigerado','Atención','GHS04',2),
('H290','Puede ser corrosiva para los metales',F,'Corrosivos para los metales','1','Atención','GHS05',2),
('H300','Mortal en caso de ingestión',S,'Toxicidad aguda (oral)','1, 2','Peligro','GHS06',4),
('H301','Tóxico en caso de ingestión',S,'Toxicidad aguda (oral)','3','Peligro','GHS06',3),
('H302','Nocivo en caso de ingestión',S,'Toxicidad aguda (oral)','4','Atención','GHS07',2),
('H303','Puede ser nocivo en caso de ingestión',S,'Toxicidad aguda (oral)','5','Atención','Sin pictograma',1),
('H304','Puede ser mortal en caso de ingestión y de penetración en las vías respiratorias',S,'Peligro por aspiración','1','Peligro','GHS08',3),
('H305','Puede ser nocivo en caso de ingestión y de penetración en las vías respiratorias',S,'Peligro por aspiración','2','Atención','GHS08',2),
('H310','Mortal en contacto con la piel',S,'Toxicidad aguda (cutánea)','1, 2','Peligro','GHS06',4),
('H311','Tóxico en contacto con la piel',S,'Toxicidad aguda (cutánea)','3','Peligro','GHS06',3),
('H312','Nocivo en contacto con la piel',S,'Toxicidad aguda (cutánea)','4','Atención','GHS07',2),
('H313','Puede ser nocivo en contacto con la piel',S,'Toxicidad aguda (cutánea)','5','Atención','Sin pictograma',1),
('H314','Provoca graves quemaduras en la piel y lesiones oculares',S,'Corrosión cutánea','1A, 1B, 1C','Peligro','GHS05',3),
('H315','Provoca irritación cutánea',S,'Irritación cutánea','2','Atención','GHS07',2),
('H316','Provoca una leve irritación cutánea',S,'Irritación cutánea','3','Atención','Sin pictograma',1),
('H317','Puede provocar una reacción cutánea alérgica',S,'Sensibilización cutánea','1, 1A, 1B','Atención','GHS07',3),
('H318','Provoca lesiones oculares graves',S,'Lesiones oculares graves','1','Peligro','GHS05',3),
('H319','Provoca irritación ocular grave',S,'Irritación ocular','2 / 2A','Atención','GHS07',2),
('H320','Provoca irritación ocular',S,'Irritación ocular','2B','Atención','Sin pictograma',1),
('H330','Mortal si se inhala',S,'Toxicidad aguda (inhalación)','1, 2','Peligro','GHS06',4),
('H331','Tóxico si se inhala',S,'Toxicidad aguda (inhalación)','3','Peligro','GHS06',3),
('H332','Nocivo si se inhala',S,'Toxicidad aguda (inhalación)','4','Atención','GHS07',2),
('H333','Puede ser nocivo si se inhala',S,'Toxicidad aguda (inhalación)','5','Atención','Sin pictograma',1),
('H334','Puede provocar síntomas de alergia o asma o dificultades respiratorias si se inhala',S,'Sensibilización respiratoria','1, 1A, 1B','Peligro','GHS08',4),
('H335','Puede irritar las vías respiratorias',S,'Toxicidad específica de órganos diana (exposición única)','3','Atención','GHS07',2),
('H336','Puede provocar somnolencia o vértigo',S,'Toxicidad específica de órganos diana (exposición única)','3','Atención','GHS07',2),
('H340','Puede provocar defectos genéticos',S,'Mutagenicidad en células germinales','1A, 1B','Peligro','GHS08',4),
('H341','Susceptible de provocar defectos genéticos',S,'Mutagenicidad en células germinales','2','Atención','GHS08',3),
('H350','Puede provocar cáncer',S,'Carcinogenicidad','1A, 1B','Peligro','GHS08',4),
('H351','Susceptible de provocar cáncer',S,'Carcinogenicidad','2','Atención','GHS08',3),
('H360','Puede perjudicar la fertilidad o dañar al feto',S,'Toxicidad para la reproducción','1A, 1B','Peligro','GHS08',4),
('H361','Susceptible de perjudicar la fertilidad o dañar al feto',S,'Toxicidad para la reproducción','2','Atención','GHS08',3),
('H362','Puede ser nocivo para los lactantes',S,'Toxicidad para la reproducción (lactancia)','Categoría adicional','(sin palabra de advertencia)','Sin pictograma',3),
('H370','Provoca daños en los órganos',S,'Toxicidad específica de órganos diana (exposición única)','1','Peligro','GHS08',4),
('H371','Puede provocar daños en los órganos',S,'Toxicidad específica de órganos diana (exposición única)','2','Atención','GHS08',3),
('H372','Provoca daños en los órganos tras exposiciones prolongadas o repetidas',S,'Toxicidad específica de órganos diana (exposiciones repetidas)','1','Peligro','GHS08',4),
('H373','Puede provocar daños en los órganos tras exposiciones prolongadas o repetidas',S,'Toxicidad específica de órganos diana (exposiciones repetidas)','2','Atención','GHS08',3),
('H400','Muy tóxico para los organismos acuáticos',A,'Peligro para el medio acuático (agudo)','1','Atención','GHS09',2),
('H401','Tóxico para los organismos acuáticos',A,'Peligro para el medio acuático (agudo)','2','(sin palabra de advertencia)','Sin pictograma',1),
('H402','Nocivo para los organismos acuáticos',A,'Peligro para el medio acuático (agudo)','3','(sin palabra de advertencia)','Sin pictograma',1),
('H410','Muy tóxico para los organismos acuáticos, con efectos nocivos duraderos',A,'Peligro para el medio acuático (crónico)','1','Atención','GHS09',2),
('H411','Tóxico para los organismos acuáticos, con efectos nocivos duraderos',A,'Peligro para el medio acuático (crónico)','2','(sin palabra de advertencia)','GHS09',2),
('H412','Nocivo para los organismos acuáticos, con efectos nocivos duraderos',A,'Peligro para el medio acuático (crónico)','3','(sin palabra de advertencia)','Sin pictograma',1),
('H413','Puede ser nocivo para los organismos acuáticos, con efectos nocivos duraderos',A,'Peligro para el medio acuático (crónico)','4','(sin palabra de advertencia)','Sin pictograma',1),
('H420','Causa daños a la salud pública y el medio ambiente al destruir el ozono en la atmósfera superior',A,'Peligro para la capa de ozono','1','Atención','GHS07',2),
]
CRIT = {4: 'Efecto mortal o irreversible grave, o evento mayor inmediato (categoría 1)',
        3: 'Efecto grave o irreversible, o sospecha de efecto CMR (palabra Peligro o categoría 2 CMR/STOT)',
        2: 'Efecto nocivo o reversible (palabra Atención)',
        1: 'Categoría de menor peligro, sin pictograma'}
HD = {h[0]: h for h in H}
P = [
('P201','Procurarse las instrucciones antes del uso.'),
('P202','No manipular antes de haber leído y comprendido todas las precauciones de seguridad.'),
('P210','Mantener alejado del calor, superficies calientes, chispas, llamas al descubierto y otras fuentes de ignición. No fumar.'),
('P220','Mantener alejado de la ropa y otros materiales combustibles.'),
('P230','Mantener humidificado con...'),
('P234','Conservar únicamente en el embalaje original.'),
('P240','Toma de tierra y enlace equipotencial del recipiente y del equipo receptor.'),
('P250','Evitar abrasiones/choques/fricciones/...'),
('P260','No respirar polvos/humos/gases/nieblas/vapores/aerosoles.'),
('P261','Evitar respirar polvos/humos/gases/nieblas/vapores/aerosoles.'),
('P262','Evitar todo contacto con los ojos, la piel o la ropa.'),
('P264','Lavarse ... cuidadosamente después de la manipulación.'),
('P270','No comer, beber o fumar mientras se manipula este producto.'),
('P271','Utilizar sólo al aire libre o en un lugar bien ventilado.'),
('P272','La ropa de trabajo contaminada no debe salir del lugar de trabajo.'),
('P273','No dispersar en el medio ambiente.'),
('P280','Usar guantes/ropa de protección/equipo de protección para los ojos/la cara.'),
('P281','(Código eliminado en revisiones recientes del SGA; sustituir por P280. Verificar.)'),
('P282','Usar guantes aislantes contra el frío y equipo de protección para la cara o los ojos.'),
('P284','[En caso de ventilación insuficiente] llevar equipo de protección respiratoria.'),
('P301+P310','EN CASO DE INGESTIÓN: Llamar inmediatamente a un CENTRO DE TOXICOLOGÍA/médico.'),
('P301+P312','EN CASO DE INGESTIÓN: Llamar a un CENTRO DE TOXICOLOGÍA/médico si la persona se encuentra mal.'),
('P302+P352','EN CASO DE CONTACTO CON LA PIEL: Lavar con abundante agua.'),
('P303+P361+P353','EN CASO DE CONTACTO CON LA PIEL (o el pelo): Quitar inmediatamente toda la ropa contaminada. Enjuagar la piel con agua [o ducharse].'),
('P304+P340','EN CASO DE INHALACIÓN: Transportar a la persona al aire libre y mantenerla en una posición que le facilite la respiración.'),
('P305+P351+P338','EN CASO DE CONTACTO CON LOS OJOS: Enjuagar con agua cuidadosamente durante varios minutos. Quitar las lentes de contacto cuando estén presentes y pueda hacerse con facilidad. Proseguir con el lavado.'),
('P308+P313','EN CASO DE exposición demostrada o supuesta: Consultar a un médico.'),
('P310','Llamar inmediatamente a un CENTRO DE TOXICOLOGÍA/médico.'),
('P331','NO provocar el vómito.'),
('P336+P315','Descongelar las partes congeladas con agua tibia. No frotar la parte afectada. Buscar inmediatamente atención médica.'),
('P342+P311','En caso de síntomas respiratorios: Llamar a un CENTRO DE TOXICOLOGÍA/médico.'),
('P370+P378','En caso de incendio: Utilizar ... para la extinción.'),
('P377','Fuga de gas en llamas: No apagar las llamas del gas inflamado si no puede hacerse sin riesgo.'),
('P381','En caso de fuga, eliminar todas las fuentes de ignición.'),
('P403','Almacenar en un lugar bien ventilado.'),
('P403+P235','Almacenar en un lugar bien ventilado. Mantener fresco.'),
('P405','Guardar bajo llave.'),
('P410+P403','Proteger de la luz solar. Almacenar en un lugar bien ventilado.'),
('P501','Eliminar el contenido/recipiente en ...'),
]
PT = {'1': ('General', 5), '2': ('Prevención', 1), '3': ('Intervención', 2), '4': ('Almacenamiento', 3), '5': ('Eliminación', 4)}
PD = dict(P)

wb = openpyxl.Workbook()
AR = 'Arial'
hf = Font(name=AR, bold=True, color='FFFFFF'); hfill = PatternFill('solid', fgColor='1F3864')
thin = Side(style='thin', color='BFBFBF'); bd = Border(left=thin, right=thin, top=thin, bottom=thin)
NF = {0: 'D9D9D9', 1: 'C6E0B4', 2: 'FFE699', 3: 'F4B183', 4: 'FF7C80'}
YEL = PatternFill('solid', fgColor='FFFF00')

def sheet(name, head, rows, widths, first=False):
    ws = wb.active if first else wb.create_sheet()
    ws.title = name
    ws.append(head)
    for r in rows:
        ws.append(list(r))
    for c in ws[1]:
        c.font = hf; c.fill = hfill
        c.alignment = Alignment(wrap_text=True, vertical='center', horizontal='center')
    for row in ws.iter_rows(min_row=2):
        for c in row:
            c.font = Font(name=AR, size=10)
            c.alignment = Alignment(wrap_text=True, vertical='top'); c.border = bd
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = 'A2'
    return ws

def colorlvl(ws, ncol, lcol=None):
    for row in ws.iter_rows(min_row=2):
        v = row[ncol - 1].value
        if isinstance(v, int):
            for k in [ncol] + ([lcol] if lcol else []):
                row[k - 1].fill = PatternFill('solid', fgColor=NF[v])
                row[k - 1].alignment = Alignment(horizontal='center', vertical='top', wrap_text=True)

lee = [
('Qué es','Matriz de decisión que convierte las frases H y P extraídas de una FDS en: clase y categoría SGA, palabra de advertencia, pictogramas, nivel de peligro y consejos de prudencia priorizados. Es la tabla de reglas del motor de clasificación estático (sin IA en la decisión).'),
('Versión','v0.4 — borrador para validación experta. 13 de octubre de 2026. Cambio frente a la v0.3: se agregan las hojas Unidades_empaque, P_corto, Etiqueta_mina y Nombres_etiqueta y la regla C08, y cambia el texto de la regla C05 (qué consejos vienen marcados en la etiqueta completa). Niveles, frases y demás reglas no cambian. (v0.3: hojas Semaforo_Ir, Seleccion_P, Pertinencia_mineria y H_lenguaje_claro, y reglas C06 y C07. v0.2: hoja Prioridad_P y regla C05.)'),
('Qué es normativo','Hoja Matriz_H, columnas A a G (código, texto, clase, categoría, palabra de advertencia, pictograma): corresponden al SGA sexta edición revisada (2015), adoptado por el Decreto 1496 de 2018 (art. 1 y 4) y exigido en lugares de trabajo por la Resolución 773 de 2021 (art. 4, 5 y 7). Los textos se transcribieron sin tener el SGA a la vista: cotejar contra el Anexo 3 del SGA rev. 6 antes de la versión final.'),
('Qué es propuesta','Hoja Matriz_H, columnas H a J (nivel de peligro) y hoja Niveles. Ni el Decreto 1496 ni la Resolución 773 definen niveles de severidad; el SGA clasifica, no pondera. La escala de 0 a 4 es una propuesta metodológica basada en la categoría SGA y la palabra de advertencia, en la línea de los esquemas de bandas de peligro (p. ej. COSHH Essentials del HSE). Debe validarla el experto y sustentarse en la tesis.'),
('Cómo se calcula el nivel de una sustancia','1) Se toman todas las frases H de la FDS. 2) Cada una se busca en Matriz_H. 3) El nivel de la sustancia es el MÁXIMO de los niveles de sus frases (principio de precaución). 4) Si no hay frases H, nivel 0 y se aplica la etiqueta de producto no peligroso (Res. 773, art. 8). 5) Sufijos como H360D o H361d se resuelven con el código base de tres dígitos.'),
('Relación con el índice Ir','Este nivel mide el peligro intrínseco. El índice de riesgo Ir = Cexp/TLV mide la exposición. La regla que combina ambos en la valoración final está PENDIENTE de definir.'),
('Hojas','Unidades_empaque: unidades de capacidad del empaque y su paso a litros. P_corto, Etiqueta_mina y Nombres_etiqueta: textos y reglas de la etiqueta de uso en la mina. Semaforo_Ir: bandas y mensajes del índice Ir. Seleccion_P y Pertinencia_mineria: cuántos y cuáles consejos P se muestran en el resultado. H_lenguaje_claro: redacción sencilla de las frases H, por validar. Matriz_H: reglas por frase H. Niveles: definición de la escala. Matriz_P: consejos de prudencia usados en las 30 sustancias y su prioridad en la etiqueta. Prioridad_P: prioridad de cada tipo de consejo P y si va en la etiqueta por defecto. Reglas: precedencias del SGA y reglas complementarias. Pictogramas: catálogo GHS01–GHS09. Etiqueta_tamano: dimensiones de etiqueta. Aplicacion_30: resultado de aplicar la matriz a las 30 sustancias y comparación con la severidad asignada en el Excel de fichas.'),
('Qué puede editar','Celdas en amarillo: nivel propuesto por frase H (Matriz_H, columna H), reglas complementarias (Reglas) y dimensiones de etiqueta. Este archivo contiene valores y no fórmulas, porque el motor lo leerá con pandas. Por eso, si cambia un nivel aquí, la hoja Aplicacion_30 no se actualiza sola: cambie el nivel en construir_matriz_HP.py y ejecútelo de nuevo para regenerar todo el archivo.'),
('Fuentes','Decreto 1496 de 2018; Resolución 773 de 2021 (Ministerios del Trabajo y de Salud); SGA de Naciones Unidas, sexta edición revisada (2015); Reglamento (CE) 1272/2008 (CLP), anexo I, sección 1.2.1, para el tamaño de etiqueta; archivo de fichas de datos de seguridad, hoja "reaccio".'),
('Pendientes','a) Cotejar textos H y P con el SGA rev. 6. b) Validar los niveles propuestos. c) Confirmar la Tabla 4 de la Res. 773 (en el PDF es una imagen y no se pudo leer). d) Resolver las discrepancias de Aplicacion_30. e) Definir la regla de combinación con Ir. f) Completar Matriz_P con el resto de códigos P del SGA si se van a analizar FDS distintas a las 30.'),
]
ws = sheet('LEEME', ['Tema', 'Detalle'], lee, [30, 130], first=True)
for row in ws.iter_rows(min_row=2, max_col=1):
    row[0].font = Font(name=AR, size=10, bold=True)

rows = [h + (NIV[h[7]], CRIT[h[7]]) for h in H]
ws = sheet('Matriz_H', ['Código H', 'Indicación de peligro', 'Tipo de peligro', 'Clase de peligro SGA', 'Categoría',
                        'Palabra de advertencia', 'Pictograma', 'Nivel (0-4) — PROPUESTO', 'Nivel (texto)', 'Criterio del nivel'],
           rows, [10, 62, 12, 40, 22, 26, 18, 14, 14, 52])
colorlvl(ws, 8, 9)
ws['H1'].fill = YEL; ws['H1'].font = Font(name=AR, bold=True)
ws.auto_filter.ref = ws.dimensions

nv = [
(0,'No clasificado','Sin frases H: no cumple criterios de peligro del SGA.','Etiqueta de producto no peligroso (Res. 773, art. 8): identificación, proveedor y consejos de prudencia.'),
(1,'Bajo','Categorías de menor peligro, sin pictograma (p. ej. toxicidad aguda cat. 5, irritación leve).','Buenas prácticas de higiene; etiqueta y FDS disponibles.'),
(2,'Medio','Palabra de advertencia "Atención": efecto nocivo o reversible (irritación, toxicidad aguda cat. 4, líquido inflamable cat. 3, gas a presión).','EPP según sección 8 de la FDS; ventilación; capacitación anual (Res. 773, art. 21.7).'),
(3,'Alto','Palabra "Peligro" o efecto grave: corrosión, lesión ocular grave, tóxico cat. 3, sensibilización cutánea, sospecha CMR (cat. 2), STOT cat. 2, muy inflamable, comburente.','Controles de ingeniería y procedimiento escrito; medir exposición y calcular Ir; vigilancia de la salud.'),
(4,'Muy alto','Efecto mortal o irreversible grave (mortal cat. 1-2, CMR cat. 1, sensibilización respiratoria, STOT cat. 1) o evento mayor inmediato (explosivos, gas o líquido extremadamente inflamable, pirofóricos).','Sustitución si es viable; control estricto de fuentes de ignición y atmósfera; medición obligatoria; acceso restringido; plan de emergencia específico.'),
]
ws = sheet('Niveles', ['Nivel', 'Nombre', 'Criterio (propuesto)', 'Acción genérica recomendada (propuesta, a validar)'], nv, [8, 16, 70, 75])
colorlvl(ws, 1, 2)

pr = [(c, t) + PT[c[1]] for c, t in P]
ws = sheet('Matriz_P', ['Código P', 'Consejo de prudencia', 'Tipo', 'Prioridad en etiqueta (1 = primero)'], pr, [18, 100, 16, 16])
ws.append([])
ws.append(['Regla', 'El tipo se deduce del primer dígito del (primer) código: P1xx general, P2xx prevención, P3xx intervención, P4xx almacenamiento, P5xx eliminación. La prioridad sigue la Res. 773, art. 7, parágrafo 2: en la etiqueta van principalmente los de prevención, intervención y almacenamiento; los demás se consultan en la FDS. Esta hoja solo contiene los códigos presentes en las 30 sustancias.'])
for c in ws[ws.max_row]:
    c.font = Font(name=AR, size=10, italic=True); c.alignment = Alignment(wrap_text=True, vertical='top')
ws.row_dimensions[ws.max_row].height = 55

# Escala de prioridad por tipo de consejo P (la misma de PT), a la vista para que el motor la lea, y si el tipo
# va en la etiqueta por defecto (regla C05). El usuario puede marcar o desmarcar cualquiera antes de imprimir.
PD = {'Prevención': ('Sí', 'Res. 773 de 2021, art. 7, parágrafo 2: se atiende principalmente a los de prevención, intervención y almacenamiento.', 'Normativo (el orden entre los tres es propuesto)'),
      'Intervención': ('Sí', 'Res. 773 de 2021, art. 7, parágrafo 2.', 'Normativo (el orden entre los tres es propuesto)'),
      'Almacenamiento': ('Sí', 'Res. 773 de 2021, art. 7, parágrafo 2.', 'Normativo (el orden entre los tres es propuesto)'),
      'Eliminación': ('No', 'Res. 773 de 2021, art. 7, parágrafo 2: "los demás consejos podrán ser consultados en la FDS".', 'Propuesto — confirmar'),
      'General': ('No', 'SGA rev. 6, anexo 3, tabla A3.2.1: P101, P102 y P103 tienen como condición de uso "Productos de consumo".', 'Propuesto — confirmar')}
pp = sorted(((d, 'P' + d + 'xx') + PT[d] + PD[PT[d][0]] for d in PT), key=lambda f: f[3])
ws = sheet('Prioridad_P', ['Primer dígito', 'Códigos', 'Tipo', 'Prioridad en etiqueta (1 = primero)', 'En la etiqueta por defecto',
                           'Sustento', 'Carácter'], pp, [10, 10, 16, 16, 14, 80, 30])

rg = [
('R01','Nivel de la sustancia','Nivel = máximo de los niveles de todas sus frases H.','Propuesta (principio de precaución)','Propuesto'),
('R02','Palabra de advertencia','Si alguna frase lleva "Peligro", la etiqueta lleva "Peligro" y no se muestra "Atención".','SGA rev. 6, 1.4.10.5.3.2; Res. 773 art. 7.3.2','Normativo'),
('R03','Precedencia de pictogramas','Si aplica GHS06 (calavera), no se muestra GHS07 (exclamación).','SGA rev. 6, 1.4.10.5.3.1','Normativo'),
('R04','Precedencia de pictogramas','Si aplica GHS05 (corrosión), no se muestra GHS07 por irritación cutánea u ocular (H315, H319).','SGA rev. 6, 1.4.10.5.3.1','Normativo'),
('R05','Precedencia de pictogramas','Si aplica GHS08 por sensibilización respiratoria (H334), no se muestra GHS07 por sensibilización cutánea ni por irritación (H317, H315, H319).','SGA rev. 6, 1.4.10.5.3.1','Normativo'),
('R06','Precedencia de frases H','Si aplica H314, se puede omitir H318. Si aplica H410, se puede omitir H400.','SGA rev. 6, 1.4.10.5.3.3','Normativo'),
('R07','Producto no peligroso','Sin frases H: etiqueta con identificación del producto, proveedor y consejos de prudencia.','Res. 773 art. 8','Normativo'),
('R08','Contenido mínimo de la etiqueta','Identificación del producto (igual a la FDS), proveedor (nombre, dirección, teléfono), pictogramas, palabra de advertencia, indicaciones de peligro y consejos de prudencia, juntos y en español.','Res. 773 art. 7 y parágrafo 1','Normativo'),
('R09','Color del pictograma','Borde rojo, fondo blanco, símbolo negro. Si el producto no sale del lugar de trabajo se admite borde negro.','Res. 773 art. 12','Normativo'),
('R10','Envases de menos de 30 ml','Mínimo: nombre del producto y pictogramas.','Res. 773 art. 13','Normativo'),
('R11','Mezclas','La etiqueta indica los componentes que aportan toxicidad aguda, corrosión, lesión ocular grave, CMR, sensibilización o STOT. Los valores de corte de las tablas 1 a 3 son imágenes en el PDF: transcribir.','Res. 773 art. 10','Normativo (pendiente transcribir)'),
('R12','TLV de referencia','Los límites de exposición son los TLV de la ACGIH vigentes a la fecha de la FDS.','Res. 773 art. 16.9','Normativo'),
('C01','Asfixiante simple sin frase H propia (metano, CO2, argón)','El SGA no asigna frase H a la asfixia por desplazamiento de oxígeno. Propuesta: nivel mínimo 3 (Alto) para trabajo subterráneo. En esta versión NO se aplica al cálculo; solo se anota.','Propuesta; Decreto 1886 de 2015 (atmósfera minera)','Propuesto — confirmar'),
('C02','Frases H condicionadas en la ficha ("en solución concentrada", "si hay aspiración")','Se toma el peor caso salvo que la FDS del producto real indique la concentración.','Propuesta','Propuesto — confirmar'),
('C03','Peligro atribuible a un componente residual (p. ej. acrilamida en poliacrilamida)','Clasificar la forma comercial; marcar advertencia de nivel condicional por el componente residual.','Propuesta','Propuesto — confirmar'),
('C04','Peligros ambientales (H400–H420)','Se informan en la etiqueta. Decidir si entran al nivel de salud ocupacional. En esta versión SÍ entran al máximo.','Propuesta','Propuesto — decidir'),
('C05','Consejos de prudencia en la etiqueta','Se ordenan por la prioridad de la hoja Prioridad_P. En la etiqueta completa vienen marcados por defecto los consejos de la selección para minería (regla C07); los demás consejos de la ficha quedan disponibles para que el responsable del etiquetado los marque. Un consejo suelto no se repite si ya está dentro de uno combinado de la misma ficha. No se fija un número máximo: la selección final es del responsable del etiquetado.','Res. 773 art. 7, parágrafo 2; SGA rev. 6, anexo 3, tabla A3.2.1 (P101 a P103: productos de consumo)','Propuesto — confirmar'),
]
rg += [
('C06','Semáforo del índice Ir','El color del riesgo por exposición sale del mayor índice Ir medido y de las bandas de la hoja Semaforo_Ir. Una alerta de seguridad superada da rojo. Sin medición el semáforo es gris, y pide medir con urgencia desde el nivel de peligro indicado en esa hoja.','Decisión del autor, 13 de octubre de 2026 (hoja Semaforo_Ir)','Validado por el autor, 26 de octubre de 2026 (bandas y mensajes)'),
('C07','Selección de consejos de prudencia para minería subterránea','Para el resultado y el informe se eligen los consejos de la ficha con mayor puntaje de pertinencia (hoja Pertinencia_mineria), hasta el máximo por tipo de la hoja Seleccion_P. No se rellena un tipo con otro. La etiqueta no usa esta selección: sigue la regla C05.','Decisión del autor, 13 de octubre de 2026 (hojas Seleccion_P y Pertinencia_mineria)','Propuesto — criterios por validar'),
]
rg += [
('C08','Etiqueta de uso en la mina','Etiqueta corta para el puesto de trabajo: nombre en español (hoja Nombres_etiqueta), pictogramas y palabra de advertencia del motor, peligro principal (frase H de mayor nivel, en lenguaje claro), los consejos de la selección para minería en su texto corto (hoja P_corto) y la protección personal. Si la ficha no trae frases H, usa las del inventario del proyecto y lo dice. Si el contenido no cabe en el tamaño de la norma, se quita en el orden de la hoja Etiqueta_mina; nunca se bloquea. No reemplaza la etiqueta completa de la Res. 773.','Decisión del autor, 13 de octubre de 2026 (hojas Etiqueta_mina, P_corto, Nombres_etiqueta y Unidades_empaque)','Propuesto — textos por validar'),
]
ws = sheet('Reglas', ['ID', 'Tema', 'Regla', 'Fuente', 'Carácter'], rg, [7, 34, 80, 40, 22])
for row in ws.iter_rows(min_row=2):
    if str(row[4].value).startswith('Propuesto'):
        row[2].fill = YEL

pc = [('GHS01','Bomba explotando','Explosivos, autorreactivas y peróxidos tipo A y B'),
('GHS02','Llama','Inflamables, pirofóricos, calentamiento espontáneo, desprenden gases inflamables con agua'),
('GHS03','Llama sobre círculo','Comburentes'),('GHS04','Botella de gas','Gases a presión'),
('GHS05','Corrosión','Corrosión cutánea, lesiones oculares graves, corrosivos para metales'),
('GHS06','Calavera y tibias cruzadas','Toxicidad aguda categorías 1 a 3'),
('GHS07','Signo de exclamación','Toxicidad aguda cat. 4, irritación, sensibilización cutánea, STOT única cat. 3'),
('GHS08','Peligro para la salud','CMR, sensibilización respiratoria, STOT, peligro por aspiración'),
('GHS09','Medio ambiente','Peligro para el medio acuático')]
sheet('Pictogramas', ['Código', 'Símbolo', 'Clases de peligro'], pc, [10, 30, 90])

et = [('Hasta 3 litros','Si es posible, al menos 52 × 74','No menor de 10 × 10; si es posible, al menos 16 × 16'),
('Más de 3 y hasta 50 litros','Al menos 74 × 105','Al menos 23 × 23'),
('Más de 50 y hasta 500 litros','Al menos 105 × 148','Al menos 32 × 32'),
('Más de 500 litros','Al menos 148 × 210','Al menos 46 × 46')]
ws = sheet('Etiqueta_tamano', ['Capacidad del envase', 'Dimensiones de la etiqueta (mm)', 'Dimensiones de cada pictograma (mm)'], et, [32, 36, 80])
for row in ws.iter_rows(min_row=2, min_col=2):
    for c in row:
        c.fill = YEL
ws.append([])
ws.append(['Fuente', '', 'Res. 773 de 2021, art. 11, que remite al Reglamento (CE) 1272/2008 (CLP), anexo I, sección 1.2.1. La Tabla 4 de la resolución es una imagen en el PDF y no se pudo leer; estos valores son los del CLP y deben cotejarse con la tabla impresa.'])
for c in ws[ws.max_row]:
    c.font = Font(name=AR, size=10, italic=True); c.alignment = Alignment(wrap_text=True, vertical='top')
ws.row_dimensions[ws.max_row].height = 55

s = openpyxl.load_workbook(src)['reaccio']

def usr(t):
    t = (t or '').split('(')[0].strip().lower()
    m = {'bajo': {0, 1}, 'medio': {2}, 'alto': {3, 4}, 'bajo-medio': {0, 1, 2}, 'medio-alto': {2, 3, 4}}
    return m.get(t)

ap = []; unk = set()
for r in s.iter_rows(min_row=2, values_only=True):
    if not r[0] or not r[3]:
        continue
    name, cas, _, ht, ptx, sev = r[:6]
    ptx = ptx or ''; obs = []
    codes = []
    for c in re.findall(r'H\d{3}', ht):
        if c not in codes:
            codes.append(c)
    for p in re.findall(r'P\d{3}(?:\+P\d{3})*', ptx):
        if p not in PD:
            unk.add(p)
    cond = 'monómero' in ht
    if cond:
        obs.append('Regla C03: las frases ' + ', '.join(codes) + ' aplican al monómero residual (acrilamida), no al polímero. Nivel condicional 4 si hay monómero libre.')
        codes = []
    for c in codes:
        if c not in HD:
            unk.add(c)
    codes = [c for c in codes if c in HD]
    if re.search(r'\((si |en sol)', ht) or 'concentraciones altas' in ht:
        obs.append('Regla C02: contiene frase H condicionada; se tomó el peor caso.')
    if 'asfixiante' in ht.lower() or name.startswith('Metano'):
        obs.append('Regla C01: asfixiante simple; nivel mínimo propuesto 3 (no aplicado).')
    if 'P281' in ptx:
        obs.append('P281 fue eliminado del SGA; sustituir por P280.')
    if 'P282' in ptx and 'H281' not in codes:
        obs.append('P282 y P336+P315 corresponden a gas licuado refrigerado (H281), no a H280: revisar.')
    lv = max([HD[c][7] for c in codes], default=0)
    det = ', '.join(c for c in codes if HD[c][7] == lv) if codes else '—'
    words = [HD[c][5] for c in codes]
    if any(w == 'Peligro' for w in words):
        pal = 'Peligro'
    elif any(w.startswith('Peligro (') for w in words):
        pal = 'Peligro o Atención según categoría (ver FDS)'
    elif any('Atención' in w for w in words):
        pal = 'Atención'
    else:
        pal = 'Sin palabra de advertencia'
    pic = []
    for c in codes:
        for g in HD[c][6].split(', '):
            if not g.startswith('GHS'):
                continue
            if g == 'GHS07':
                if any(HD[x][6] == 'GHS06' for x in codes):
                    continue
                if c in ('H315', 'H319') and any('GHS05' in HD[x][6] for x in codes):
                    continue
                if c in ('H317', 'H315', 'H319') and 'H334' in codes:
                    continue
            if g not in pic:
                pic.append(g)
    u = usr('bajo' if cond else sev)
    co = 'Sí' if (u is not None and lv in u) else 'No'
    ap.append((name, cas, ', '.join(codes) or 'Sin frases H', lv, NIV[lv], det, pal,
               ', '.join(sorted(pic)) or 'Sin pictograma', sev, co, ' '.join(obs)))

ws = sheet('Aplicacion_30', ['Sustancia', 'CAS', 'Frases H reconocidas', 'Nivel matriz', 'Nivel (texto)', 'Frase(s) determinante(s)',
                             'Palabra de advertencia', 'Pictogramas (con precedencia)', 'Severidad en el Excel de fichas', '¿Coincide?', 'Observaciones'],
           ap, [34, 16, 30, 9, 13, 18, 14, 20, 34, 10, 60])
colorlvl(ws, 4, 5)
for row in ws.iter_rows(min_row=2):
    if row[9].value == 'No':
        row[9].fill = PatternFill('solid', fgColor='FF7C80')
ws.auto_filter.ref = ws.dimensions

# ---------- v0.3: semáforo del índice Ir, selección de consejos P para minería y lenguaje claro ----------
BORRADOR = 'Borrador por validar del autor'
DECISION = 'Decisión del autor, 13 de octubre de 2026'
VALIDADO_SEM = 'Validado por el autor, 26 de octubre de 2026'
# Clave, color en pantalla, condición, Ir desde (incluido), Ir hasta (sin incluir), nivel de peligro desde, mensaje, complemento si Ir > 1, fuente, estado
sem = [
('verde', 'Verde', 'Índice Ir menor que 0,5', 0, 0.5, None,
 'Exposición por debajo de la mitad del valor límite. Mantenga los controles actuales y repita la medición según su programa de higiene.', None, DECISION, VALIDADO_SEM),
('amarillo', 'Amarillo', 'Índice Ir de 0,5 a menos de 0,85', 0.5, 0.85, None,
 'Exposición entre el 50 % y el 85 % del valor límite. Revise y mejore los controles (ventilación, tiempo de exposición, protección respiratoria) y programe pronto una nueva medición.', None, DECISION, VALIDADO_SEM),
('rojo', 'Rojo', 'Índice Ir de 0,85 o más (nivel de intervención)', 0.85, None, None,
 'ALERTA DE PARO INMEDIATO DE LA OPERACIÓN. La exposición llegó al nivel de intervención (85 % del valor límite o más). Detenga la tarea, retire al personal expuesto y corrija los controles antes de reanudar; después vuelva a medir.',
 'Además, la exposición supera el valor límite permitido (Ir mayor que 1).', DECISION, VALIDADO_SEM),
('rojo_alerta', 'Rojo', 'Alerta de seguridad superada (máximo permitido)', None, None, None,
 'ALERTA DE PARO INMEDIATO DE LA OPERACIÓN. Se superó un máximo de seguridad. Detenga la labor y evacúe según el plan de emergencia; no reanude hasta que la medición vuelva a estar por debajo del máximo.', None, DECISION, VALIDADO_SEM),
('alerta_ok', 'Verde', 'Solo hay medición de un máximo de seguridad y no se supera', None, None, None,
 'La medición no supera el máximo de seguridad. Mantenga el monitoreo continuo de la atmósfera.', None, 'Propuesta', VALIDADO_SEM),
('gris', 'Gris', 'Sin medición', None, None, None,
 'Sin medición. No se puede saber si la exposición es aceptable. Programe la medición del agente en el puesto de trabajo.', None, DECISION, VALIDADO_SEM),
('gris_urgente', 'Amarillo', 'Sin medición y nivel de peligro alto', None, None, 3,
 'Sin medición y con nivel de peligro alto. Mida con urgencia; mientras tanto, aplique los controles de la ficha como si la exposición fuera alta.', None, DECISION + ' (umbral propuesto: nivel 3)', VALIDADO_SEM),
('sin_limite', 'Gris', 'El producto no tiene valor límite utilizable', None, None, None,
 'Este producto no tiene valor límite en la tabla: el semáforo no aplica. Guíese por el nivel de peligro y por las medidas de la ficha.', None, 'Propuesta', VALIDADO_SEM),
]
ws = sheet('Semaforo_Ir', ['Clave', 'Color en pantalla', 'Condición', 'Ir desde (incluido)', 'Ir hasta (sin incluir)', 'Nivel de peligro desde',
                           'Mensaje para el profesional de SST', 'Complemento si Ir > 1', 'Fuente', 'Estado'], sem, [14, 12, 34, 10, 10, 10, 90, 40, 34, 28])
ws.append([])
ws.append(['Nota', '', 'Las bandas no vienen de la norma: la norma solo fija que Ir mayor que 1 supera el valor límite (hoja Calculo de tlv_limites.xlsx). '
           'El rojo empieza en 0,85 por decisión del autor, como nivel de intervención. Con varios agentes manda el mayor índice Ir. '
           'El motor lee de aquí las bandas, el umbral de nivel y los mensajes: para cambiarlos se edita construir_matriz_HP.py y se vuelve a ejecutar.'])

# Tipo de consejo, máximo en la selección, estado. La fila Total es el tope general.
selp = [('Prevención', 2, DECISION), ('Intervención', 2, DECISION), ('Almacenamiento', 1, DECISION), ('Eliminación', 0, DECISION), ('General', 0, DECISION),
        ('Total', 5, DECISION)]
sheet('Seleccion_P', ['Tipo de consejo', 'Máximo en la selección', 'Estado'], selp, [22, 18, 44])

# ID, vía o peligro bajo tierra, peso, capítulos del SGA que lo activan, códigos P que responden, por qué importa, estado
pert = [
('M1', 'Inhalación', 3, '3.1, 3.4, 3.6, 3.8, 3.9',
 'P260, P261, P271, P284, P304+P340, P403+P233, P342+P311, P314, P308+P313, P311',
 'Bajo tierra el aire se renueva solo con la ventilación de la mina: polvos, gases y vapores se acumulan en el frente de trabajo.', BORRADOR),
('M2', 'Inflamabilidad o explosión', 3, '2.1, 2.2, 2.3, 2.4, 2.6, 2.7, 2.13, 2.14',
 'P210, P211, P220, P233, P240, P241, P242, P243, P250, P251, P370+P378, P370+P376, P377, P381, P403, P403+P235, P410+P412, P410+P403, P372, P373',
 'En una mina de carbón hay metano y polvo de carbón: cualquier fuente de ignición puede iniciar un incendio o una explosión.', BORRADOR),
('M3', 'Atmósfera confinada o deficiente en oxígeno', 2, '2.2, 2.4, 2.5',
 'P271, P403, P410+P403, P377, P381, P304+P340, P403+P233',
 'Un gas que escapa en una labor sin salida desplaza el oxígeno o enriquece la atmósfera.', BORRADOR),
('M4', 'Ventilación', 2, '2.2, 2.3, 2.6, 3.1, 3.8, 3.9',
 'P271, P403+P233, P403+P235, P403',
 'La ventilación es el control principal de la atmósfera minera (Decreto 1886 de 2015).', BORRADOR),
('M5', 'Contacto con piel u ojos', 2, '3.1, 3.2, 3.3, 3.4',
 'P280, P264, P302+P352, P303+P361+P353, P305+P351+P338, P332+P313, P337+P313, P362+P364, P362, P301+P330+P331, P310, P363, P333+P313, P272',
 'El lavado y la atención médica están lejos del frente: la protección y el primer auxilio inmediato pesan más.', BORRADOR),
]
ws = sheet('Pertinencia_mineria', ['ID', 'Vía o peligro bajo tierra', 'Peso', 'Capítulos del SGA (clases de peligro) que lo activan',
                                   'Códigos P que responden', 'Por qué importa en minería subterránea', 'Estado'], pert, [6, 34, 7, 30, 70, 70, 28])
ws.append([])
ws.append(['Nota', '', '', 'Un criterio se activa si la ficha trae alguna frase H de esos capítulos del SGA (columna «Clase de peligro» del diccionario frases_HP_es_en.xlsx). '
           'El puntaje de un consejo P es la suma de los pesos de los criterios activos que lo nombran. Se eligen los de mayor puntaje, hasta el máximo por tipo de '
           'la hoja Seleccion_P; si empatan, primero el de código menor. Un consejo con puntaje 0 solo entra si sobran cupos de su tipo, y queda marcado.'])

# Código H, en lenguaje claro (borrador), estado. La pantalla usa este texto solo cuando el Estado es «Validado»; mientras tanto muestra el oficial.
claro = [
('H201', 'Puede explotar todo de una vez.'), ('H220', 'Gas que se enciende con mucha facilidad.'), ('H222', 'Aerosol que se enciende con mucha facilidad.'),
('H225', 'Líquido que se enciende muy fácilmente, y sus vapores también.'), ('H226', 'Líquido que puede encenderse, y sus vapores también.'),
('H227', 'Líquido que puede arder si se calienta.'), ('H229', 'El envase está a presión y puede reventar si se calienta.'),
('H270', 'Puede iniciar un fuego o avivarlo: hace que otros materiales ardan con más fuerza.'), ('H272', 'Aviva el fuego: hace que otros materiales ardan con más fuerza.'),
('H280', 'El cilindro tiene gas a presión y puede explotar si se calienta.'), ('H290', 'Puede corroer los metales.'), ('H302', 'Hace daño si se traga.'),
('H304', 'Si se traga y pasa a los pulmones, puede matar.'), ('H314', 'Quema gravemente la piel y los ojos.'), ('H315', 'Irrita la piel.'),
('H316', 'Irrita levemente la piel.'), ('H317', 'Puede causar alergia en la piel.'), ('H318', 'Causa daño grave en los ojos.'),
('H319', 'Irrita fuertemente los ojos.'), ('H320', 'Irrita los ojos.'), ('H330', 'Respirarlo puede matar.'), ('H331', 'Respirarlo es tóxico.'),
('H332', 'Respirarlo hace daño.'), ('H335', 'Puede irritar la nariz, la garganta y los pulmones.'), ('H336', 'Puede dar sueño o mareo.'),
('H340', 'Puede dañar el material genético (los genes).'), ('H350', 'Puede causar cáncer.'), ('H351', 'Se sospecha que causa cáncer.'),
('H360', 'Puede afectar la fertilidad o dañar al bebé durante el embarazo.'), ('H361', 'Se sospecha que afecta la fertilidad o daña al bebé durante el embarazo.'),
('H370', 'Daña órganos del cuerpo.'), ('H372', 'Daña órganos del cuerpo si la exposición es prolongada o repetida.'),
('H373', 'Puede dañar órganos del cuerpo si la exposición es prolongada o repetida.'), ('H401', 'Es tóxico para la vida acuática.'),
('H410', 'Es muy tóxico para la vida acuática y su efecto dura mucho tiempo.'), ('H411', 'Es tóxico para la vida acuática y su efecto dura mucho tiempo.'),
('H412', 'Hace daño a la vida acuática y su efecto dura mucho tiempo.'),
]
oficial = {h[0]: h[1] for h in H}
ws = sheet('H_lenguaje_claro', ['Código H', 'Texto oficial (Matriz_H)', 'En lenguaje claro', 'Estado'],
           [(c, oficial.get(c), texto, BORRADOR) for c, texto in claro], [10, 70, 70, 30])
ws.append([])
ws.append(['Nota', '', 'Son los 37 códigos H que aparecen en las 32 fichas. La herramienta muestra el texto en lenguaje claro solo cuando el Estado dice «Validado»; '
           'mientras diga otra cosa, muestra el texto oficial. La etiqueta siempre lleva el texto oficial.'])


# ---------- v0.4: etiqueta de uso en la mina y unidades de capacidad del empaque ----------
# Unidad, abreviatura con que se muestra, litros por unidad, cómo se obtuvo, estado
uni = [
('mL', 'mL', 0.001, 'Conversión exacta: 1 L = 1000 mL.', 'Conversión exacta'),
('L', 'L', 1, 'La hoja Etiqueta_tamano está en litros.', 'Conversión exacta'),
('galón (US)', 'gal (US)', 3.785411784, 'Conversión exacta: el galón líquido de Estados Unidos mide 3,785411784 L.', 'Conversión exacta'),
('g', 'g', 0.001, 'Equivalencia propuesta: 1 kg = 1 L. La norma da los tamaños de etiqueta por capacidad en litros; para sólidos se toma el mismo número.', BORRADOR),
('kg', 'kg', 1, 'Equivalencia propuesta: 1 kg = 1 L. La norma da los tamaños de etiqueta por capacidad en litros; para sólidos se toma el mismo número.', BORRADOR),
]
sheet('Unidades_empaque', ['Unidad', 'Abreviatura', 'Litros por unidad', 'De dónde sale', 'Estado'], uni, [14, 12, 16, 90, 30])

# Código P -> texto corto para la etiqueta de uso en la mina. Sin «...» ni corchetes. Son los consejos que aparecen en las fichas y en el inventario.
pc_corto = [
('P101', 'Si necesita atención médica, lleve el empaque o la etiqueta.'), ('P102', 'Mantenga fuera del alcance de los niños.'),
('P201', 'Pida las instrucciones antes de usarlo.'), ('P202', 'No lo manipule sin haber leído las precauciones de seguridad.'),
('P210', 'Lejos de calor, chispas y llamas. No fumar.'), ('P211', 'No rocíe sobre llamas ni fuentes de ignición.'),
('P220', 'Mantenga lejos de ropa y materiales combustibles.'), ('P230', 'Manténgalo humedecido.'), ('P233', 'Mantenga el recipiente bien cerrado.'),
('P234', 'Consérvelo solo en su empaque original.'), ('P235', 'Manténgalo fresco.'), ('P240', 'Conecte a tierra el recipiente y el equipo.'),
('P241', 'Use equipo eléctrico, de ventilación e iluminación a prueba de explosión.'), ('P242', 'No use herramientas que produzcan chispas.'),
('P243', 'Evite las descargas de electricidad estática.'), ('P250', 'Evite golpes, fricción y abrasión.'),
('P251', 'No perfore ni queme el recipiente, ni siquiera vacío.'), ('P260', 'No respire polvos, gases ni vapores.'),
('P261', 'Evite respirar polvos, gases o vapores.'), ('P264', 'Lávese bien después de manipularlo.'), ('P270', 'No coma, beba ni fume mientras lo usa.'),
('P271', 'Use solo al aire libre o con buena ventilación.'), ('P273', 'Evite que llegue al suelo, al agua o a los desagües.'),
('P280', 'Use guantes, ropa de protección y protección para ojos y cara.'), ('P282', 'Use guantes aislantes contra el frío y protección para la cara o los ojos.'),
('P284', 'Si la ventilación no basta, use protección respiratoria.'),
('P301+P310', 'Si se traga: llame de inmediato a un médico.'), ('P301+P312', 'Si se traga y se siente mal: llame a un médico.'),
('P301+P330+P331', 'Si se traga: enjuague la boca. No provoque el vómito.'), ('P302+P352', 'Si cae en la piel: lave con abundante agua.'),
('P303+P361+P353', 'Si cae en la piel o el pelo: quítese la ropa contaminada y lave con agua.'), ('P304+P340', 'Si lo respira: salga al aire libre y repose.'),
('P305+P351+P338', 'Si cae en los ojos: lave con agua varios minutos y quite los lentes de contacto.'),
('P308+P311', 'Si hubo exposición: llame a un médico.'), ('P308+P313', 'Si hubo exposición o la sospecha: consulte a un médico.'),
('P310', 'Llame de inmediato a un médico.'), ('P311', 'Llame a un médico.'), ('P312', 'Si se siente mal, llame a un médico.'),
('P314', 'Si se siente mal, consulte a un médico.'), ('P321', 'Siga el tratamiento que indica la ficha de seguridad.'), ('P330', 'Enjuague la boca.'),
('P331', 'No provoque el vómito.'), ('P332+P313', 'Si la piel se irrita: consulte a un médico.'),
('P336+P315', 'Si hay congelación: descongele con agua tibia, sin frotar, y busque atención médica.'),
('P337+P313', 'Si la irritación de los ojos sigue: consulte a un médico.'), ('P362', 'Quítese la ropa contaminada.'),
('P362+P364', 'Quítese la ropa contaminada y lávela antes de volver a usarla.'), ('P370+P376', 'En caso de incendio: detenga la fuga si puede hacerlo sin riesgo.'),
('P370+P378', 'En caso de incendio: use el medio de extinción que indica la ficha de seguridad.'),
('P377', 'Fuga de gas en llamas: no apague la llama si no puede detener la fuga sin riesgo.'), ('P381', 'En caso de fuga, elimine todas las fuentes de ignición.'),
('P390', 'Absorba el derrame para que no dañe otros materiales.'), ('P391', 'Recoja el derrame.'), ('P402', 'Guarde en un lugar seco.'),
('P403', 'Guarde en un lugar bien ventilado.'), ('P403+P233', 'Guarde bien cerrado, en un lugar ventilado.'),
('P403+P235', 'Guarde en un lugar fresco y bien ventilado.'), ('P405', 'Guarde bajo llave.'), ('P410+P403', 'Proteja del sol y guarde en un lugar ventilado.'),
('P410+P412', 'Proteja del sol y de temperaturas mayores de 50 °C.'), ('P501', 'Deseche el contenido y el empaque según las normas locales.'),
]
oficial_p = {p[0]: p[1] for p in P}
ws = sheet('P_corto', ['Código P', 'Texto corto para la etiqueta de uso en la mina', 'Texto oficial (Matriz_P, si lo trae)', 'Estado'],
           [(c, corto, oficial_p.get(c), BORRADOR) for c, corto in pc_corto], [18, 80, 80, 30])
ws.append([])
ws.append(['Nota', 'Un consejo de la ficha que no esté en esta hoja no sale en la etiqueta de uso en la mina (sí en la etiqueta completa). El texto corto no lleva «...» ni corchetes: '
           'donde el texto oficial pide un dato del fabricante, remite a la ficha de seguridad.'])

# Clave, valor, para qué sirve, estado
etm = [
('retiro_1', 'almacenamiento', 'Si el contenido no cabe, lo primero que se quita: el renglón de almacenamiento.', DECISION),
('retiro_2', 'que_hacer', 'Después, los renglones de «Qué hacer» de menor prioridad, uno por uno.', DECISION),
('minimo_de_renglones', 1, 'Renglones de «Qué hacer» que siempre quedan.', DECISION),
('retiro_3', 'iconos_epp', 'Después, los iconos de protección personal (queda el texto).', DECISION),
('retiro_4', 'texto_epp', 'Solo en el tamaño de etiqueta más pequeño (primera fila de Etiqueta_tamano): el texto de protección personal.', DECISION),
('retiro_5', 'peligro_principal', 'Solo en el tamaño más pequeño: la línea de peligro principal. Queda el mínimo: nombre, pictogramas, palabra de advertencia y un renglón.', DECISION),
('rotulo_peligro', 'Peligro principal', 'Rótulo de la línea de peligro.', BORRADOR),
('rotulo_hacer', 'Qué hacer', 'Rótulo de los renglones de acción.', BORRADOR),
('rotulo_epp', 'Protección personal', 'Rótulo de la línea de protección personal.', BORRADOR),
('texto_sin_peligros', 'Sin peligros clasificados', 'Lo que dice la etiqueta cuando no hay frases H ni en la ficha ni en el inventario.', BORRADOR),
('texto_asfixiante', 'Asfixiante simple: desplaza el oxígeno del aire.', 'Se agrega al peligro principal cuando aplica la regla C01.', BORRADOR),
('texto_inventario', 'Peligros tomados del inventario del proyecto: la FDS no los declara en la sección 2.', 'Aviso del panel de pantalla; no se imprime.', BORRADOR),
('texto_no_es_res_773', 'Etiqueta de uso en la mina. No trae todas las indicaciones de peligro ni los consejos oficiales: no reemplaza la etiqueta completa de la Resolución 773 de 2021 ni la del fabricante.',
 'Aviso del panel de pantalla; no se imprime.', BORRADOR),
]
sheet('Etiqueta_mina', ['Clave', 'Valor', 'Para qué sirve', 'Estado'], etm, [22, 70, 80, 34])

# Número de la ficha (inicio del nombre del archivo), nombre en español para la etiqueta
nom = [('01', 'Nitrato de amonio'), ('02', 'Combustible diésel (ACPM)'), ('03', 'Metano'), ('04', 'Monóxido de carbono'), ('05', 'Sílice cristalina (cuarzo)'),
       ('06', 'Estireno'), ('07', 'MDI polimérico (isocianato)'), ('08a', 'Cemento Portland'), ('08b', 'Óxido de calcio (cal viva)'), ('09', 'Silicato de sodio'),
       ('10', 'Grasas y aceites lubricantes'), ('11', 'Sulfuro de hidrógeno'), ('12', 'Polvo de roca caliza'), ('13', 'Fluido hidráulico ignífugo'),
       ('14', 'Ácido sulfúrico'), ('15', 'Acetileno'), ('16', 'Oxígeno comprimido'), ('17', 'Etilenglicol'), ('18', 'Cal hidratada'), ('19', 'PETN (pentrita)'),
       ('20', 'Polvo químico seco ABC'), ('21', 'Gasolina'), ('22', 'Queroseno'), ('23', 'Nitrato de calcio'), ('24', 'Ácido clorhídrico'),
       ('25', 'Hidróxido de sodio (soda cáustica)'), ('26', 'Dodecilbencenosulfonato de sodio'), ('27', 'Dióxido de carbono'), ('28', 'Argón'),
       ('29', 'Resina epóxica (bisfenol A)'), ('30', 'Poliacrilamida'), ('31', 'Grasa de litio en aerosol')]
ws = sheet('Nombres_etiqueta', ['N.º de ficha', 'Nombre en español para la etiqueta', 'Estado'], [(n, nombre, BORRADOR) for n, nombre in nom], [14, 60, 34])
ws.append([])
ws.append(['Nota', 'El inventario agrupa algunos productos en una sola fila (acetileno y oxígeno; cemento y óxido de calcio) y trae el uso entre paréntesis; por eso el '
           'nombre de la etiqueta va aquí, uno por ficha. Una ficha cargada por el usuario usa el nombre del inventario si su CAS coincide; si no, el de la ficha, '
           'y el usuario puede escribirlo a mano.'])

wb.save(B + 'matriz_decision_HP.xlsx')
print('filas', len(ap), '| códigos no reconocidos:', unk or 'ninguno')
for a in ap:
    print(a[3], a[9], '|', a[0][:28], '|', a[2], '|', a[7], '|', a[6], '|', (a[8] or '')[:24])
