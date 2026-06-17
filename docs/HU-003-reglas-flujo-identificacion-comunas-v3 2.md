# HU-003 - Validación de comunas por alimentador

## Contexto

Esta validación corresponde al eje de **comunas**.

El objetivo es revisar si las comunas informadas en el borrador para cada alimentador coinciden con las comunas esperadas según `Comunas_25F` de `BD Alimentadores`.

La comparación se realiza por documento y alimentador, pero solo para el último documento `EAF XXX-2026` detectado en `DOCUMENTO` [B].

---

## Input

* Archivo Excel borrador.
* Hoja: `BD Consolidado`.
* Base de referencia: `BD Alimentadores`.
* Base oficial de comunas, si aplica.

---

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

---

## Columnas del borrador

| Campo                                                     |               Columna |
| --------------------------------------------------------- | --------------------: |
| DOCUMENTO                                                 |                     B |
| ID COORDINADO AFECTADO                                    |                     H |
| TIPO CLIENTE                                              |                     I |
| CLAVE UNICA / ID UNICO DE ALIMENTADOR                     | Por nombre de columna |
| Comunas                                                   |                     X |
| COMENTARIOS DISTRIBUIDORA SOBRE INCONSISTENCIA DE COMUNAS | Por nombre de columna |

---

## Columnas de BD Alimentadores

| Campo                  |
| ---------------------- |
| ID UNICO               |
| Comunas_25F            |
| ID COORDINADO          |
| Nombre Alimentador_25F |

---

## Unidad de análisis

```text
DOCUMENTO + CLAVE UNICA
```

Solo considerar unidades pertenecientes al documento objetivo `EAF XXX-2026` y cuyas filas tengan `TIPO CLIENTE` [I] igual a `RE` o `LD`.

---

## Regla principal

Para cada unidad de análisis:

```text
observado = conjunto de comunas informadas en el borrador [Comunas]

esperado = conjunto de comunas de Comunas_25F en BD Alimentadores
           donde ID UNICO = CLAVE UNICA
```

Luego comparar ambos conjuntos.

---

## Tratamiento de comunas

* Separar comunas si vienen en una misma celda usando:

  * coma `,`
  * punto y coma `;`
  * slash `/`
* Cada comuna debe tratarse como token independiente.
* Normalizar comunas antes de comparar:

  * mayúsculas;
  * sin tildes;
  * sin espacios extra;
  * conjuntos sin orden.

---

## Clasificación de hallazgos

| Condición                                     | Tipo                    |
| --------------------------------------------- | ----------------------- |
| `CLAVE UNICA` no existe en `BD Alimentadores` | ID_NO_EN_BASE           |
| Faltan datos para evaluar                     | DATOS_INSUFICIENTES     |
| `faltan ≠ vacío` y `sobran = vacío`           | FALTAN_COMUNAS          |
| `faltan = vacío` y `sobran ≠ vacío`           | SOBRAN_COMUNAS          |
| `faltan ≠ vacío` y `sobran ≠ vacío`           | FALTAN_Y_SOBRAN_COMUNAS |
| `faltan = vacío` y `sobran = vacío`           | OK                      |

---

## Flujo

1. Validar estructura del borrador.
2. Validar existencia de `BD Alimentadores`.
3. Identificar el último documento `EAF XXX-2026` desde `DOCUMENTO` [B].
4. Filtrar el universo revisable solo a las filas de ese documento objetivo.
5. Dentro del documento objetivo, considerar solo filas donde `TIPO CLIENTE` [I] sea `RE` o `LD`.
6. Excluir filas con `TIPO CLIENTE` [I] igual a `LT`.
7. Agrupar filas por `DOCUMENTO + CLAVE UNICA`.
8. Obtener comunas observadas desde `Comunas` [X].
9. Obtener comunas esperadas desde `Comunas_25F`.
10. Normalizar ambos conjuntos de comunas.
11. Comparar conjuntos.
12. Registrar hallazgo si hay diferencias.
13. Registrar la clasificación en la columna `REV_COMUNAS_CLASIFICACION`.
14. Repetir la clasificación en todas las filas revisables del grupo observado.
15. Reportar hallazgo del documento objetivo.

---

## Reglas de validación

* Comparar solo filas del documento objetivo `EAF XXX-2026` cuyo `TIPO CLIENTE` [I] sea `RE` o `LD`.
* Excluir filas con `TIPO CLIENTE` [I] igual a `LT`.
* Comparar `Comunas` del borrador contra `Comunas_25F` de `BD Alimentadores`.
* Para obtener comunas esperadas, buscar `CLAVE UNICA` del borrador como `ID UNICO` en `BD Alimentadores`.
* Si falta `DOCUMENTO`, `CLAVE UNICA` o `Comunas`, registrar `DATOS_INSUFICIENTES`.
* Si `CLAVE UNICA` no existe en `BD Alimentadores`, registrar `ID_NO_EN_BASE`.
* Si existe base oficial de comunas, validar que cada comuna informada exista en dicha base.
* Si existe base empresa-comuna, validar que la comuna esté asociada al `ID COORDINADO AFECTADO` [H].
* El comentario de la distribuidora puede respaldar el hallazgo, pero no lo elimina automáticamente.

---

## Tratamiento de datos

* Normalizar IDs antes de comparar.
* Eliminar sufijos `.0` en IDs numéricos.
* Comparar IDs como texto normalizado.
* Considerar vacío: `null`, `NaN`, texto vacío o texto compuesto solo por espacios.
* Normalizar `TIPO CLIENTE` [I] aplicando `trim` y mayúsculas antes de comparar.
* Considerar como filas revisables solo aquellas con `TIPO CLIENTE` [I] igual a `RE` o `LD`.
* No modificar valores originales del borrador.

---

## Output esperado

* Mismo Excel de entrada.
* Solo se revisan filas del último documento `EAF XXX-2026` cuyo `TIPO CLIENTE` [I] sea `RE` o `LD`.
* No se deben marcar filas con color.
* Se debe agregar una columna al final de la hoja `BD Consolidado` llamada:

```text
REV_COMUNAS_CLASIFICACION
```

* La columna debe completarse solo para filas pertenecientes al último documento `EAF XXX-2026` cuyo `TIPO CLIENTE` [I] sea `RE` o `LD`.
* Si el grupo presenta inconsistencia, registrar la clasificación correspondiente.
* Si el grupo no presenta inconsistencia, dejar la celda vacía o registrar `OK`, según política del aplicativo.
* Como la validación de comunas es grupal, la clasificación debe repetirse en todas las filas revisables del grupo `DOCUMENTO + CLAVE UNICA`.
* Se debe generar reporte de hallazgos del documento objetivo.

---

## Clasificación de hallazgos

| Condición                                     | Tipo                    |
| --------------------------------------------- | ----------------------- |
| `CLAVE UNICA` no existe en `BD Alimentadores` | ID_NO_EN_BASE           |
| Faltan datos para evaluar                     | DATOS_INSUFICIENTES     |
| `faltan ≠ vacío` y `sobran = vacío`           | FALTAN_COMUNAS          |
| `faltan = vacío` y `sobran ≠ vacío`           | SOBRAN_COMUNAS          |
| `faltan ≠ vacío` y `sobran ≠ vacío`           | FALTAN_Y_SOBRAN_COMUNAS |
| `faltan = vacío` y `sobran = vacío`           | OK                      |
