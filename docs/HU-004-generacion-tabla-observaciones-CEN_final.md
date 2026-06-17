# HU-004 - Generación de tabla de observaciones CEN

## Contexto

A partir de los hallazgos detectados en las validaciones de mecanismos de recuperación, alimentadores y comunas, el aplicativo debe generar una tabla resumen tipo anexo, lista para ser copiada y pegada en la carta o documento de publicación.

Esta HU no reemplaza las validaciones anteriores. Su objetivo es consolidar los casos con hallazgos relevantes y entregar una salida útil para revisión de ingeniería.

---

## Input

- Archivo Excel borrador.
- Hoja: `BD Consolidado`.
- Hoja: `BD interrupciones`.
- Base de referencia: `BD Alimentadores 2025`.
- Columnas de clasificación generadas por las validaciones previas:
  - `REV_RECUPERACION_CLASIFICACION`
  - `REV_ALIMENTADORES_CLASIFICACION`
  - `REV_COMUNAS_CLASIFICACION`

---

## Precondición

Antes de generar la tabla de observaciones, deben existir las columnas de clasificación en la hoja `BD Consolidado`.

Si alguna de estas columnas no existe, el aplicativo debe reportar que falta ejecutar o completar las validaciones previas:

```text
REV_RECUPERACION_CLASIFICACION
REV_ALIMENTADORES_CLASIFICACION
REV_COMUNAS_CLASIFICACION
```

---

## Alcance

La tabla debe generarse solo para el último documento `EAF XXX-2026` detectado en la columna `DOCUMENTO` [B] de la hoja `BD Consolidado`.

Regla:

- Considerar solo valores con formato `EAF XXX-2026`.
- Extraer el número `XXX`.
- Seleccionar el documento con el número `XXX` más alto.
- Aplicar la generación de tabla solo a las filas cuyo `DOCUMENTO` [B] sea igual a ese documento.

Si no existen documentos con formato `EAF XXX-2026`, detener la generación y reportar que no existe documento objetivo 2026.

---

## Criterio para incluir filas

La tabla de observaciones se debe generar solo a partir de las filas del documento objetivo que presenten al menos un hallazgo en las columnas de clasificación.

Una fila se considera con hallazgo si al menos una de las siguientes columnas tiene un valor distinto de vacío o `OK`:

```text
REV_RECUPERACION_CLASIFICACION
REV_ALIMENTADORES_CLASIFICACION
REV_COMUNAS_CLASIFICACION
```

Las filas sin hallazgos no deben incorporarse en la tabla de observaciones.

---

## Regla cuando no existan hallazgos

Si el documento objetivo `EAF XXX-2026` no presenta filas con hallazgos en las columnas `REV_*_CLASIFICACION`, el aplicativo debe crear igualmente la hoja `ANEXO_OBSERVACIONES` con los encabezados definidos, pero sin filas de datos.

Opcionalmente, puede registrar una nota fuera de la tabla indicando:

```text
Sin hallazgos para el documento objetivo.
```

No se deben generar comentarios artificiales si no existen hallazgos.

---


## Unidad de consolidación

La tabla debe consolidarse por:

```text
DOCUMENTO + Coordinado
```

Es decir, para un mismo EAF y una misma distribuidora/coordinado, debe generarse una sola fila en la tabla de observaciones.

---

## Regla de consolidación de comentarios en la tabla final

La hoja `ANEXO_OBSERVACIONES` debe generar una sola fila por combinación:

```text
DOCUMENTO + Coordinado
```

Si para esa combinación existen hallazgos en más de un eje de revisión, los comentarios no deben sobrescribirse ni generar filas duplicadas.

Los comentarios de mecanismos de recuperación, alimentadores y comunas deben consolidarse en la misma celda `Comentario CEN`, separados por una línea en blanco.

Solo se debe crear una nueva fila cuando cambie el `DOCUMENTO` o el `Coordinado`.

Ejemplo:

| EAF | Distribuidora Involucrada | Comentario CEN |
|---|---|---|
| EAF 103-2026 | Compañía General de Electricidad S.A. | Comentario mecanismo de recuperación<br><br>Comentario alimentadores<br><br>Comentario comunas |

---


## Columnas de origen

### Hoja `BD Consolidado`

| Campo | Columna |
|---|---:|
| DOCUMENTO | B |
| NOMBRE ALIMENTADOR / CONSUMO | E |
| Barra | AA |
| Coordinado | AB |
| SUBESTACIÓN | AG |
| REV_RECUPERACION_CLASIFICACION | Al final de la tabla |
| REV_ALIMENTADORES_CLASIFICACION | Al final de la tabla |
| REV_COMUNAS_CLASIFICACION | Al final de la tabla |

### Hoja `BD interrupciones`

| Campo | Columna |
|---|---:|
| FECHA Y HORA DE INICIO | I |

### Base `BD Alimentadores 2025`

| Campo | Columna |
|---|---:|
| Subestación | C |
| Nombre Alimentador_25F | J |
| Barra_25F | K |
| Comunas_25F | L |

---

## Columnas de la tabla final

| Columna tabla | Origen / regla |
|---|---|
| `EAF` | Valor del `DOCUMENTO` analizado, correspondiente al último `EAF XXX-2026`. |
| `Fecha falla` | Hoja `BD interrupciones`, columna `FECHA Y HORA DE INICIO` [I]. Usar solo la fecha en formato `dd-mm-aaaa`. |
| `Distribuidora Involucrada` | Hoja `BD Consolidado`, columna `Coordinado` [AB]. |
| `S/E con consumo afectado` | Hoja `BD Consolidado`, columna `SUBESTACIÓN` [AG]. Agrupar todas las S/E distintas asociadas al documento y coordinado con hallazgo. |
| `¿Envía información?` | Siempre `Parcial`. |
| `Presenta formato solicitado` | Pendiente de definición. |
| `Comentario CEN` | Se genera según los comentarios tipo asociados a las clasificaciones detectadas. Por ahora se consideran comentarios tipo para mecanismos de recuperación, alimentadores y comunas. |

---

## Tratamiento de fecha de falla

Para la columna `Fecha falla`, tomar el valor de `FECHA Y HORA DE INICIO` [I] desde la hoja `BD interrupciones`.

Si el valor contiene fecha y hora, se debe conservar solo la fecha y presentarla en formato:

```text
dd-mm-aaaa
```

Ejemplo:

```text
18-01-2026 08:43
```

Debe quedar como:

```text
18-01-2026
```

Si existen múltiples valores de `FECHA Y HORA DE INICIO` para el documento objetivo, usar la fecha más temprana y registrar advertencia para revisión.

---

## Tratamiento de S/E con consumo afectado

Para cada combinación `DOCUMENTO + Coordinado`, se deben listar todas las `SUBESTACIÓN` [AG] distintas que aparezcan en filas con hallazgo.

Reglas:

- Eliminar duplicados.
- Ignorar valores vacíos.
- Mantener el formato original del nombre de la S/E cuando sea posible.
- Si existe más de una S/E, listarlas en la misma celda separadas por salto de línea.
- La consolidación solo debe considerar filas del documento objetivo y con hallazgo.

---

## Regla de regeneración de la tabla

La hoja `ANEXO_OBSERVACIONES` debe representar únicamente el resultado de la ejecución actual.

Si la hoja `ANEXO_OBSERVACIONES` ya existe en el archivo Excel, el aplicativo debe eliminarla o limpiar completamente su contenido antes de generar la nueva tabla.

No se deben conservar filas de ejecuciones anteriores.

La nueva tabla debe generarse solo con los hallazgos asociados al último documento `EAF XXX-2026` detectado en la hoja `BD Consolidado`.

---

## Generación de Comentario CEN

La columna `Comentario CEN` debe construirse a partir de las clasificaciones detectadas en las columnas `REV_*_CLASIFICACION`.

Para esta versión, se consideran los comentarios tipo asociados a mecanismos de recuperación, alimentadores y comunas.

---

## Orden de consolidación del Comentario CEN

Cuando existan hallazgos de más de un eje para una misma combinación `DOCUMENTO + Coordinado`, los comentarios deben escribirse en el siguiente orden:

1. Mecanismos de recuperación.
2. Alimentadores.
3. Comunas.

Cada bloque de comentario debe separarse por una línea en blanco.

Este orden busca mantener consistencia entre ejecuciones y facilitar la revisión del equipo de ingeniería.

---


## Comentario tipo: mecanismos de recuperación

Aplica cuando la columna `REV_RECUPERACION_CLASIFICACION` tenga una clasificación distinta de vacío o `OK`.

En particular, aplica para el caso:

```text
NORMALIZACION_ANTES_DISPONIBILIDAD_SIN_RESPALDO
```

También puede aplicarse si se decide observar casos con respaldo informado, según criterio del usuario o del equipo revisor.

---

## Variables del comentario de mecanismos

| Variable | Origen |
|---|---|
| `ALIMENTADOR` | Hoja `BD Consolidado`, columna `NOMBRE ALIMENTADOR / CONSUMO` [E] |
| `SUBESTACIÓN` | Hoja `BD Consolidado`, columna `SUBESTACIÓN` [AG] |

---

## Regla de no repetición para mecanismos

Para mecanismos de recuperación, el comentario no debe repetirse por cada fila inconsistente.

Se debe identificar cada combinación única de:

```text
DOCUMENTO + Coordinado + NOMBRE ALIMENTADOR / CONSUMO + SUBESTACIÓN
```

Si un mismo alimentador aparece varias veces en las filas con hallazgo, debe aparecer una sola vez en el `Comentario CEN`.

Ejemplo: si el mismo alimentador aparece en 6 filas por comunas, zonas de tarificación u otros desgloses, el comentario se genera una sola vez para ese alimentador y subestación.

---

## Texto del comentario tipo de mecanismos

Estructura:

```text
El alimentador ALIMENTADOR (SUBESTACIÓN) tiene hora de normalización equivalente previa a la hora de disponibilidad de la barra.

Se solicita detallar el mecanismo de recuperación de estas cargas para este alimentador donde esto se presente, y con el desglose requerido por la SEC, esto es comunal y por zona de tarificación.
```

Ejemplo:

```text
El alimentador CARAMPANGUE (S/E Isla de Maipo) tiene hora de normalización equivalente previa a la hora de disponibilidad de la barra.

Se solicita detallar el mecanismo de recuperación de estas cargas para este alimentador donde esto se presente, y con el desglose requerido por la SEC, esto es comunal y por zona de tarificación.
```

---

## Consolidación de múltiples comentarios

Si para un mismo `DOCUMENTO + Coordinado` existen varios alimentadores distintos con hallazgo de mecanismos de recuperación, los comentarios deben consolidarse en la misma celda `Comentario CEN`.

Reglas:

- Generar un comentario por cada combinación única `NOMBRE ALIMENTADOR / CONSUMO + SUBESTACIÓN`.
- No repetir comentarios idénticos.
- Separar cada comentario por una línea en blanco.
- Mantener el texto listo para copiar y pegar en el documento oficial.

---


## Comentario tipo: alimentadores

Los comentarios tipo de alimentadores aplican cuando la columna `REV_ALIMENTADORES_CLASIFICACION` tenga una de las siguientes clasificaciones:

```text
SOBRAN_ALIMENTADORES
FALTAN_ALIMENTADORES
FALTAN_Y_SOBRAN_ALIMENTADORES
```

Para estos casos, el comentario debe consolidarse en la celda `Comentario CEN` correspondiente a la combinación `DOCUMENTO + Coordinado`.

---

## Comentario tipo: sobran alimentadores

Aplica cuando la clasificación sea:

```text
SOBRAN_ALIMENTADORES
```

También aplica cuando la clasificación sea:

```text
FALTAN_Y_SOBRAN_ALIMENTADORES
```

En este caso, los alimentadores a listar se obtienen desde la hoja `BD Consolidado`, ya que corresponden a alimentadores que aparecen en el borrador, pero que conforme a las bases de datos revisadas se encontrarían incorporados en forma excedente.

### Variables del comentario de alimentadores sobrantes

| Variable | Origen |
|---|---|
| `NOMBRE_ALIMENTADOR` | Hoja `BD Consolidado`, columna `NOMBRE ALIMENTADOR / CONSUMO` [E] |
| `NOMBRE_BARRA` | Hoja `BD Consolidado`, columna `Barra` [AA] |
| `NOMBRE_SUBESTACION` | Hoja `BD Consolidado`, columna `SUBESTACIÓN` [AG] |

### Regla de no repetición para alimentadores sobrantes

No se debe repetir el mismo alimentador por cada fila inconsistente.

Se debe identificar cada combinación única de:

```text
DOCUMENTO + Coordinado + NOMBRE ALIMENTADOR / CONSUMO + Barra + SUBESTACIÓN
```

Si el mismo alimentador aparece varias veces en las filas con hallazgo, debe aparecer una sola vez en el bloque de alimentadores sobrantes.

### Texto del comentario tipo para alimentadores sobrantes

Estructura:

```text
Se identifican alimentadores que, conforme a las bases de datos revisadas, se encontrarían incorporados en forma excedente:

Alimentador NOMBRE_ALIMENTADOR — Barra NOMBRE_BARRA — S/E NOMBRE_SUBESTACION

Se solicita revisar, corregir o justificar su inclusión, incorporando el desglose comunal y por zona de tarificación requerido por la SEC.
```

---

## Comentario tipo: faltan alimentadores

Aplica cuando la clasificación sea:

```text
FALTAN_ALIMENTADORES
```

También aplica cuando la clasificación sea:

```text
FALTAN_Y_SOBRAN_ALIMENTADORES
```

En este caso, los alimentadores a listar se obtienen desde `BD Alimentadores 2025`, ya que corresponden a alimentadores que pertenecerían a las barras y/o S/E informadas, pero que no se encontrarían incorporados en los antecedentes remitidos.

### Variables del comentario de alimentadores faltantes

| Variable | Origen |
|---|---|
| `NOMBRE_ALIMENTADOR` | `BD Alimentadores 2025`, columna `Nombre Alimentador_25F` [J] |
| `NOMBRE_BARRA` | `BD Alimentadores 2025`, columna `Barra_25F` [K] |
| `NOMBRE_SUBESTACION` | `BD Alimentadores 2025`, columna `Subestación` [C] |

### Regla de no repetición para alimentadores faltantes

No se debe repetir el mismo alimentador faltante por cada fila o grupo inconsistente.

Se debe identificar cada combinación única de:

```text
Nombre Alimentador_25F + Barra_25F + Subestación
```

Si el mismo alimentador faltante aparece más de una vez dentro del análisis, debe aparecer una sola vez en el bloque de alimentadores faltantes.

### Texto del comentario tipo para alimentadores faltantes

Estructura:

```text
Se identifican alimentadores que, conforme a las bases de datos revisadas, pertenecerían a las barras y/o S/E informadas, pero no se encontrarían incorporados en los antecedentes remitidos:

Alimentador NOMBRE_ALIMENTADOR — Barra NOMBRE_BARRA — S/E NOMBRE_SUBESTACION

Se solicita revisar, corregir o justificar su omisión, incorporando el desglose comunal y por zona de tarificación requerido por la SEC.
```

---

## Comentario tipo: faltan y sobran alimentadores

Aplica cuando la clasificación sea:

```text
FALTAN_Y_SOBRAN_ALIMENTADORES
```

En este caso, el `Comentario CEN` debe incluir ambos bloques:

1. Bloque de alimentadores sobrantes, usando información desde `BD Consolidado`.
2. Bloque de alimentadores faltantes, usando información desde `BD Alimentadores 2025`.

La estructura debe ser:

```text
Se identifican alimentadores que, conforme a las bases de datos revisadas, se encontrarían incorporados en forma excedente:

Alimentador NOMBRE_ALIMENTADOR_1 — Barra NOMBRE_BARRA_1 — S/E NOMBRE_SUBESTACION_1
Alimentador NOMBRE_ALIMENTADOR_2 — Barra NOMBRE_BARRA_2 — S/E NOMBRE_SUBESTACION_2

Se identifican alimentadores que, conforme a las bases de datos revisadas, pertenecerían a las barras y/o S/E informadas, pero no se encontrarían incorporados en los antecedentes remitidos:

Alimentador NOMBRE_ALIMENTADOR_1 — Barra NOMBRE_BARRA_1 — S/E NOMBRE_SUBESTACION_1
Alimentador NOMBRE_ALIMENTADOR_2 — Barra NOMBRE_BARRA_2 — S/E NOMBRE_SUBESTACION_2

Se solicita revisar, corregir o justificar su inclusión u omisión, incorporando el desglose comunal y por zona de tarificación requerido por la SEC.
```

Reglas:

- No repetir alimentadores dentro del bloque de sobrantes.
- No repetir alimentadores dentro del bloque de faltantes.
- Separar ambos bloques por una línea en blanco.
- En este caso, la frase final debe indicar `inclusión u omisión`, ya que existen alimentadores excedentes y omitidos.

---

## Consolidación de comentarios de alimentadores

Si para un mismo `DOCUMENTO + Coordinado` existen varios hallazgos de alimentadores, estos deben consolidarse en la misma celda `Comentario CEN`.

Reglas:

- Si solo sobran alimentadores, usar el comentario tipo de alimentadores sobrantes.
- Si solo faltan alimentadores, usar el comentario tipo de alimentadores faltantes.
- Si faltan y sobran alimentadores, usar el comentario combinado.
- No repetir líneas idénticas de alimentadores.
- Mantener el texto listo para copiar y pegar en el documento oficial.

---


## Comentario tipo: comunas

Los comentarios tipo de comunas aplican cuando la columna `REV_COMUNAS_CLASIFICACION` tenga una de las siguientes clasificaciones:

```text
SOBRAN_COMUNAS
FALTAN_COMUNAS
FALTAN_Y_SOBRAN_COMUNAS
```

Para estos casos, el comentario debe consolidarse en la celda `Comentario CEN` correspondiente a la combinación `DOCUMENTO + Coordinado`.

El comentario debe generarse a partir del resultado de la validación definida en la HU-003. Es decir, se deben listar solo las comunas efectivamente identificadas como sobrantes o faltantes, no todas las comunas informadas en la fila o en la base.

---

## Comentario tipo: sobran comunas

Aplica cuando la clasificación sea:

```text
SOBRAN_COMUNAS
```

También aplica cuando la clasificación sea:

```text
FALTAN_Y_SOBRAN_COMUNAS
```

En este caso, las comunas a listar se obtienen desde la hoja `BD Consolidado`, ya que corresponden a comunas que aparecen en el borrador, pero que conforme a las bases de datos revisadas se encontrarían incorporadas en forma excedente para los alimentadores y/o barras informadas.

### Variables del comentario de comunas sobrantes

| Variable | Origen |
|---|---|
| `COMUNA` | Hoja `BD Consolidado`, columna `Comunas` [X] |
| `NOMBRE_ALIMENTADOR` | Hoja `BD Consolidado`, columna `NOMBRE ALIMENTADOR / CONSUMO` [E] |
| `NOMBRE_BARRA` | Hoja `BD Consolidado`, columna `Barra` [AA] |

### Regla de no repetición para comunas sobrantes

No se debe repetir la misma comuna por cada fila inconsistente.

Se debe identificar cada combinación única de:

```text
DOCUMENTO + Coordinado + COMUNA + NOMBRE ALIMENTADOR / CONSUMO + Barra
```

Si una misma comuna aparece varias veces para el mismo alimentador y barra, debe aparecer una sola vez en el bloque de comunas sobrantes.

Si `Comunas` [X] contiene varias comunas en una misma celda, se deben separar como tokens independientes y listar solo aquellas que efectivamente sobran según la validación de la HU-003.

### Texto del comentario tipo para comunas sobrantes

Estructura:

```text
Se identifican comunas que, conforme a las bases de datos revisadas, se encontrarían incorporadas en forma excedente para los alimentadores y/o barras informadas:

Comuna COMUNA — Alimentador NOMBRE_ALIMENTADOR — Barra NOMBRE_BARRA

Se solicita revisar, corregir o justificar su inclusión, incorporando el desglose comunal y por zona de tarificación requerido por la SEC.
```

---

## Comentario tipo: faltan comunas

Aplica cuando la clasificación sea:

```text
FALTAN_COMUNAS
```

También aplica cuando la clasificación sea:

```text
FALTAN_Y_SOBRAN_COMUNAS
```

En este caso, las comunas a listar se obtienen desde `BD Alimentadores 2025`, ya que corresponden a comunas que pertenecerían a los alimentadores y/o barras informadas, pero que no se encontrarían incorporadas en los antecedentes remitidos.

### Variables del comentario de comunas faltantes

| Variable | Origen |
|---|---|
| `COMUNA` | `BD Alimentadores 2025`, columna `Comunas_25F` [L] |
| `NOMBRE_ALIMENTADOR` | `BD Alimentadores 2025`, columna `Nombre Alimentador_25F` [J] |
| `NOMBRE_BARRA` | `BD Alimentadores 2025`, columna `Barra_25F` [K] |

### Regla de no repetición para comunas faltantes

No se debe repetir la misma comuna faltante por cada fila o grupo inconsistente.

Se debe identificar cada combinación única de:

```text
COMUNA + Nombre Alimentador_25F + Barra_25F
```

Si una misma comuna faltante aparece varias veces para el mismo alimentador y barra, debe aparecer una sola vez en el bloque de comunas faltantes.

Si `Comunas_25F` [L] contiene varias comunas en una misma celda, se deben separar como tokens independientes y listar solo aquellas que efectivamente faltan según la validación de la HU-003.

### Texto del comentario tipo para comunas faltantes

Estructura:

```text
Se identifican comunas que, conforme a las bases de datos revisadas, pertenecerían a los alimentadores y/o barras informadas, pero no se encontrarían incorporadas en los antecedentes remitidos:

Comuna COMUNA — Alimentador NOMBRE_ALIMENTADOR — Barra NOMBRE_BARRA

Se solicita revisar, corregir o justificar su omisión, incorporando el desglose comunal y por zona de tarificación requerido por la SEC.
```

---

## Comentario tipo: faltan y sobran comunas

Aplica cuando la clasificación sea:

```text
FALTAN_Y_SOBRAN_COMUNAS
```

En este caso, el `Comentario CEN` debe incluir ambos bloques:

1. Bloque de comunas sobrantes, usando información desde `BD Consolidado`.
2. Bloque de comunas faltantes, usando información desde `BD Alimentadores 2025`.

La estructura debe ser:

```text
Se identifican comunas que, conforme a las bases de datos revisadas, se encontrarían incorporadas en forma excedente para los alimentadores y/o barras informadas:

Comuna COMUNA_1 — Alimentador NOMBRE_ALIMENTADOR_1 — Barra NOMBRE_BARRA_1
Comuna COMUNA_2 — Alimentador NOMBRE_ALIMENTADOR_2 — Barra NOMBRE_BARRA_2

Se identifican comunas que, conforme a las bases de datos revisadas, pertenecerían a los alimentadores y/o barras informadas, pero no se encontrarían incorporadas en los antecedentes remitidos:

Comuna COMUNA_1 — Alimentador NOMBRE_ALIMENTADOR_1 — Barra NOMBRE_BARRA_1
Comuna COMUNA_2 — Alimentador NOMBRE_ALIMENTADOR_2 — Barra NOMBRE_BARRA_2

Se solicita revisar, corregir o justificar su inclusión u omisión, incorporando el desglose comunal y por zona de tarificación requerido por la SEC.
```

Reglas:

- No repetir comunas dentro del bloque de sobrantes para el mismo alimentador y barra.
- No repetir comunas dentro del bloque de faltantes para el mismo alimentador y barra.
- Separar ambos bloques por una línea en blanco.
- En este caso, la frase final debe indicar `inclusión u omisión`, ya que existen comunas excedentes y omitidas.

---

## Consolidación de comentarios de comunas

Si para un mismo `DOCUMENTO + Coordinado` existen varios hallazgos de comunas, estos deben consolidarse en la misma celda `Comentario CEN`.

Reglas:

- Si solo sobran comunas, usar el comentario tipo de comunas sobrantes.
- Si solo faltan comunas, usar el comentario tipo de comunas faltantes.
- Si faltan y sobran comunas, usar el comentario combinado.
- No repetir líneas idénticas de comunas.
- Si una misma comuna se repite varias veces por el mismo consumo, alimentador o barra, debe consolidarse y aparecer una sola vez.
- Mantener el texto listo para copiar y pegar en el documento oficial.

---

## Output esperado

- Crear una hoja nueva llamada `ANEXO_OBSERVACIONES`.
- Si la hoja ya existe, eliminarla o limpiarla completamente antes de escribir la nueva tabla.
- La hoja debe contener una tabla lista para copiar y pegar en el documento oficial.
- La tabla debe incluir solo filas asociadas a hallazgos.
- La tabla debe consolidarse por `DOCUMENTO + Coordinado`.
- Si existen hallazgos en más de un eje para el mismo `DOCUMENTO + Coordinado`, los comentarios deben consolidarse en una sola celda `Comentario CEN`.
- No deben incluirse filas sin hallazgos.
- Si no existen hallazgos, la hoja debe quedar creada solo con encabezados y sin filas de datos.
- No deben modificarse los datos originales de `BD Consolidado`.
- Si un coordinado tiene más de una S/E con hallazgo, listarlas en la misma celda separadas por salto de línea.
- Si un coordinado tiene más de una clasificación de hallazgo, el `Comentario CEN` debe consolidar los comentarios tipo que correspondan.
- Los comentarios deben escribirse en el orden: mecanismos de recuperación, alimentadores y comunas.
- Para mecanismos de recuperación, el comentario debe aparecer una sola vez por alimentador y subestación, aunque existan múltiples filas con la misma clasificación.
- Para alimentadores, el comentario debe consolidar alimentadores faltantes y/o sobrantes sin repetir líneas idénticas.
- Para comunas, el comentario debe consolidar comunas faltantes y/o sobrantes sin repetir líneas idénticas.

---

## Estructura de salida

La hoja `ANEXO_OBSERVACIONES` debe contener las siguientes columnas, en este orden:

| EAF | Fecha falla | Distribuidora Involucrada | S/E con consumo afectado | ¿Envía información? | Presenta formato solicitado | Comentario CEN |
|---|---|---|---|---|---|---|

---

## Comentarios tipo definidos

En esta versión quedan definidos los comentarios tipo para:

| Eje | Clasificaciones consideradas |
|---|---|
| Mecanismos de recuperación | `NORMALIZACION_ANTES_DISPONIBILIDAD_SIN_RESPALDO` |
| Alimentadores | `SOBRAN_ALIMENTADORES`, `FALTAN_ALIMENTADORES`, `FALTAN_Y_SOBRAN_ALIMENTADORES` |
| Comunas | `SOBRAN_COMUNAS`, `FALTAN_COMUNAS`, `FALTAN_Y_SOBRAN_COMUNAS` |

---

## Consideraciones

- La tabla debe priorizar facilidad de uso para el equipo de ingeniería.
- El resultado debe ser copiable directamente desde Excel hacia Word o el documento oficial.
- No se debe requerir que el usuario filtre, ordene o consolide manualmente información adicional.
- Las reglas de comentarios tipo podrán agregarse posteriormente sin modificar la estructura base de la tabla.
