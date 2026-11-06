# QuimNexo

Herramienta local para leer fichas de datos de seguridad (FDS) y clasificar el riesgo químico según el Sistema Globalmente Armonizado (SGA), pensada para la pequeña y mediana minería subterránea de carbón en Sogamoso, Boyacá.

Hace parte de la tesis de maestría en Seguridad y Salud en el Trabajo *Clasificación automatizada de riesgos químicos (SGA) mediante el procesamiento de lenguaje natural en la pequeña y mediana minería del carbón en Sogamoso, Boyacá* (Universidad Pedagógica y Tecnológica de Colombia, seccional Sogamoso).

**Autor:** Javier Alejandro Castro Nuván  
**Director:** M.Sc. Fabián Alfredo Torres  
**Versión:** 2.03

## Qué hace

1. Lee la FDS en PDF y extrae las frases H y P, los pictogramas y los datos de identificación.
2. Clasifica el nivel de peligro con reglas fijas tomadas de la norma (Decreto 1496 de 2018 y Resolución 773 de 2021). Ningún modelo de lenguaje interviene en esta decisión.
3. Calcula el índice de riesgo (I.R. = concentración de exposición / TLV) para una jornada de 8 horas y lo muestra con semáforo: verde < 0,5; amarillo 0,5 a < 0,85; rojo ≥ 0,85.
4. Entrega un informe de consulta corto y la etiqueta del envase (de uso en la mina y completa).

El corpus incluye 32 fichas de 30 sustancias usadas en la operación minera (carpeta `FDS/`), ya procesadas en `extraccion/salidas/`.

## Cómo abrirla (Windows)

Requiere Python 3.12 o superior.

```
py -m pip install -r requirements.txt
```

Luego, doble clic en **Abrir QuimNexo.bat**. Se abre en el navegador del mismo equipo (http://localhost:8501) y los datos no salen de él.

Para volver a procesar las fichas del corpus:

```
py extraccion/extraer.py FDS
```

## Estructura

| Carpeta | Contenido |
|---|---|
| `clasificador SGA/` | Aplicación (Streamlit), motor de clasificación, informe y etiquetas |
| `extraccion/` | Lectura del PDF y extracción por secciones |
| `matriz/` | Matriz de decisión de frases H y P, diccionario ES-EN y protección (sección 8) |
| `tlv/` | Valores límite (TLV) usados en el cálculo |
| `plantillas/` | Pictogramas SGA y elementos de protección personal |
| `marca/` | Identidad visual de QuimNexo |
| `FDS/` | Fichas de datos de seguridad del corpus |

## Clave de API (opcional)

La segunda pasada de la extracción puede usar un modelo de lenguaje. La clave nunca se escribe en los archivos: va en la variable de entorno `ANTHROPIC_API_KEY` u `OPENAI_API_KEY`. Sin clave, la herramienta funciona con la extracción por reglas.

## Aviso

QuimNexo apoya la lectura de las FDS. No reemplaza el criterio del profesional de SST ni la medición de las concentraciones en el lugar de trabajo.

Las fichas de datos de seguridad de `FDS/` pertenecen a sus fabricantes o proveedores y se incluyen solo como corpus de la investigación.

© 2026 Javier Alejandro Castro Nuván. Todos los derechos reservados. La letra Inter se distribuye con su propia licencia (SIL OFL), en `clasificador SGA/static/fuentes/`.
