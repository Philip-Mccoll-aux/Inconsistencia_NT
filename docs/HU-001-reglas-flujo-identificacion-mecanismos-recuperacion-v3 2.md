# HU-001 - Validación de mecanismos de recuperación

## Contexto

La `FECHA Y HORA NORMALIZACIÓN DE CONSUMO` no debe ser menor que la `FECHA Y HORA DISPONIBILIDAD DE LA BARRA`, salvo que exista observación o mecanismo de recuperación informado.

## Input

* Archivo Excel borrador.
* Hoja: `BD Consolidado`.

## Alcance de revisión

La validación no debe ejecutarse sobre todo el borrador.

Primero se debe identificar el **último documento EAF 2026** en la columna `DOCUMENTO` [B].

Regla:

* Considerar solo valores con formato `EAF XXX-2026`.
* Extraer el número `XXX`.
* Seleccionar el documento con el número `XXX` más alto.
* Aplicar la revisión solo a las filas cuyo `DOCUMENTO` [B] sea igual a ese documento.
* Dentro del documento objetivo, considerar solo filas donde `TIPO CLIENTE` [I] sea `RE` o `LD`.
* Excluir filas con `TIPO CLIENTE` [I] igual a `LT`.

Ejemplo:

```text
EAF 096-2026
EAF 101-2026
EAF 103-2026
```

El documento objetivo es:

```text
EAF 103-2026
```

Si no existen documentos con formato `EAF XXX-2026`, detener la validación y reportar que no existe documento objetivo 2026.

## Columnas

| Campo                                   | Columna |
| --------------------------------------- | ------: |
| DOCUMENTO                               |       B |
| ID PAÑO                                 |       D |
| ID BARRA PUNTO DE CONTROL               |       G |
| ID COORDINADO AFECTADO                  |       H |
| TIPO CLIENTE                            |       I |
| FECHA Y HORA DISPONIBILIDAD DE LA BARRA |       U |
| FECHA Y HORA NORMALIZACIÓN DE CONSUMO   |       V |
| OBSERVACIONES                           |       W |
| Comunas                                 |       X |
| Mecanismo de Recuperación               |      AO |
| Comentarios mecanismo de recuperación   |      AQ |

## Pre-revisiones mínimas

Antes de aplicar la validación principal:

1. Verificar que exista la hoja `BD Consolidado`.
2. Verificar que existan las columnas requeridas.
3. Identificar el último documento `EAF XXX-2026` desde `DOCUMENTO` [B].
4. Filtrar el universo revisable solo a las filas de ese documento objetivo.
5. Dentro del documento objetivo, considerar solo filas donde `TIPO CLIENTE` [I] sea `RE` o `LD`.
6. Excluir filas con `TIPO CLIENTE` [I] igual a `LT`.
7. Validar que cada fila revisable tenga informada la triada:

   * `ID PAÑO` [D]
   * `ID BARRA PUNTO DE CONTROL` [G]
   * `ID COORDINADO AFECTADO` [H]
8. Validar que `Comunas` [X] no esté vacío.
9. Si `Comunas` [X] tiene varias comunas separadas por coma, tratarlas como tokens independientes.
10. Para cada combinación `DOCUMENTO` [B] + `ID BARRA PUNTO DE CONTROL` [G], debe existir una sola `FECHA Y HORA DISPONIBILIDAD DE LA BARRA` [U].
11. Para cada combinación `DOCUMENTO` [B] + `ID PAÑO` [D] + `ID BARRA PUNTO DE CONTROL` [G] + `ID COORDINADO AFECTADO` [H] + `Comuna` [X], debe existir una sola `FECHA Y HORA NORMALIZACIÓN DE CONSUMO` [V].

Si falla la estructura básica, detener.
Si fallan las demás pre-revisiones, registrar hallazgo.

## Regla principal

Registrar hallazgo en la fila revisable si se cumple:

```text
V < U
AND W vacío
AND AO vacío
AND AQ vacío
```

Donde:

* `U` = `FECHA Y HORA DISPONIBILIDAD DE LA BARRA`
* `V` = `FECHA Y HORA NORMALIZACIÓN DE CONSUMO`
* `W` = `OBSERVACIONES`
* `AO` = `Mecanismo de Recuperación`
* `AQ` = `Comentarios mecanismo de recuperación`

## Tratamiento de datos

* Comparar `U` y `V` como fecha/hora, no como texto.
* Considerar vacío: `null`, `NaN`, `""` o texto con solo espacios.
* Aplicar `trim` antes de evaluar campos vacíos.
* Normalizar `TIPO CLIENTE` [I] aplicando `trim` y mayúsculas antes de comparar.
* Considerar como filas revisables solo aquellas con `TIPO CLIENTE` [I] igual a `RE` o `LD`.
* No modificar valores originales.

## Output esperado

* Mismo Excel de entrada.
* Solo se revisan filas del último documento `EAF XXX-2026` cuyo `TIPO CLIENTE` [I] sea `RE` o `LD`.
* No se deben marcar filas con color.
* Se debe agregar una columna al final de la hoja `BD Consolidado` llamada:

```text
REV_RECUPERACION_CLASIFICACION
```

* La columna debe completarse solo para filas pertenecientes al último documento `EAF XXX-2026` cuyo `TIPO CLIENTE` [I] sea `RE` o `LD`.
* Si la fila presenta inconsistencia, registrar la clasificación correspondiente.
* Si la fila no presenta inconsistencia, dejar la celda vacía o registrar `OK`, según política del aplicativo.
* Como la validación de recuperación es por fila, la clasificación debe registrarse solo en la fila observada.
* Se debe generar reporte de hallazgos del documento objetivo.

---

## Clasificación del hallazgo

| Condición                                        | Clasificación                                   |
| ------------------------------------------------ | ----------------------------------------------- |
| `V < U` y `W`, `AO`, `AQ` están vacías           | NORMALIZACION_ANTES_DISPONIBILIDAD_SIN_RESPALDO |
| `V < U` y existe información en `W`, `AO` o `AQ` | NORMALIZACION_ANTES_DISPONIBILIDAD_CON_RESPALDO |
| `U` o `V` no son fechas válidas                  | DATOS_INSUFICIENTES_FECHAS                      |
| No existe inconsistencia                         | OK                                              |
