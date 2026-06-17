# HU-002 - Validación de alimentadores por barra

## Contexto

Esta validación corresponde al eje de **alimentadores**.

El objetivo es revisar si los alimentadores informados en el borrador para una barra coinciden con los alimentadores esperados según `BD Alimentadores`.

La comparación se realiza por documento y barra, pero solo para el último documento `EAF XXX-2026` detectado en `DOCUMENTO` [B].

---

## Input

* Archivo Excel borrador.
* Hoja: `BD Consolidado`.
* Base de referencia: `BD Alimentadores`.

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

| Campo                                        |               Columna |
| -------------------------------------------- | --------------------: |
| DOCUMENTO                                    |                     B |
| ID PAÑO                                      |                     D |
| ID BARRA PUNTO DE CONTROL                    |                     G |
| ID COORDINADO AFECTADO                       |                     H |
| TIPO CLIENTE                                 |                     I |
| CLAVE UNICA / ID UNICO DE ALIMENTADOR        | Por nombre de columna |
| NOMBRE ALIMENTADOR / CONSUMO                 | Por nombre de columna |
| COMENTARIO DISTRIBUIDORA SOBRE ALIMENTADORES | Por nombre de columna |

---

## Unidad de análisis

```text
DOCUMENTO + ID BARRA PUNTO DE CONTROL
```

Solo considerar unidades pertenecientes al documento objetivo `EAF XXX-2026` y cuyas filas tengan `TIPO CLIENTE` [I] igual a `RE` o `LD`.

---

## Regla principal

Para cada unidad de análisis:

```text
observado = conjunto de CLAVE UNICA informado en el borrador

esperado = conjunto de ID UNICO en BD Alimentadores
           donde ID Punto Control = ID BARRA PUNTO DE CONTROL
```

Luego comparar ambos conjuntos.

---

## Clasificación de hallazgos

| Condición                             | Tipo                          |
| ------------------------------------- | ----------------------------- |
| Barra no existe en `BD Alimentadores` | BARRA_NO_EN_BASE              |
| Faltan datos para evaluar             | DATOS_INSUFICIENTES           |
| `faltan ≠ vacío` y `sobran = vacío`   | FALTAN_ALIMENTADORES          |
| `faltan = vacío` y `sobran ≠ vacío`   | SOBRAN_ALIMENTADORES          |
| `faltan ≠ vacío` y `sobran ≠ vacío`   | FALTAN_Y_SOBRAN_ALIMENTADORES |
| `faltan = vacío` y `sobran = vacío`   | OK                            |

---

## Flujo

1. Validar estructura del borrador.
2. Validar existencia de `BD Alimentadores`.
3. Identificar el último documento `EAF XXX-2026` desde `DOCUMENTO` [B].
4. Filtrar el universo revisable solo a las filas de ese documento objetivo.
5. Dentro del documento objetivo, considerar solo filas donde `TIPO CLIENTE` [I] sea `RE` o `LD`.
6. Excluir filas con `TIPO CLIENTE` [I] igual a `LT`.
7. Agrupar filas por `DOCUMENTO + ID BARRA PUNTO DE CONTROL`.
8. Obtener alimentadores observados desde `CLAVE UNICA`.
9. Obtener alimentadores esperados desde `BD Alimentadores`.
10. Comparar conjuntos.
11. Registrar hallazgo si hay diferencias.
12. Registrar la clasificación en la columna `REV_ALIMENTADORES_CLASIFICACION`.
13. Repetir la clasificación en todas las filas revisables del grupo observado.
14. Reportar hallazgo del documento objetivo.

---

## Reglas de validación

* Comparar solo filas del documento objetivo `EAF XXX-2026` cuyo `TIPO CLIENTE` [I] sea `RE` o `LD`.
* Excluir filas con `TIPO CLIENTE` [I] igual a `LT`.
* Comparar `CLAVE UNICA` del borrador contra `ID UNICO` de `BD Alimentadores`.
* Para obtener esperados, filtrar `BD Alimentadores` por `ID Punto Control`.
* `ID Punto Control` se compara contra `ID BARRA PUNTO DE CONTROL`.
* Si falta `DOCUMENTO`, `ID BARRA PUNTO DE CONTROL` o `CLAVE UNICA`, registrar `DATOS_INSUFICIENTES`.
* Si la barra no existe en `BD Alimentadores`, registrar `BARRA_NO_EN_BASE`.
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
REV_ALIMENTADORES_CLASIFICACION
```

* La columna debe completarse solo para filas pertenecientes al último documento `EAF XXX-2026` cuyo `TIPO CLIENTE` [I] sea `RE` o `LD`.
* Si el grupo presenta inconsistencia, registrar la clasificación correspondiente.
* Si el grupo no presenta inconsistencia, dejar la celda vacía o registrar `OK`, según política del aplicativo.
* Como la validación de alimentadores es grupal, la clasificación debe repetirse en todas las filas revisables del grupo `DOCUMENTO + ID BARRA PUNTO DE CONTROL`.
* Se debe generar reporte de hallazgos del documento objetivo.

---

## Clasificación de hallazgos

| Condición                             | Tipo                          |
| ------------------------------------- | ----------------------------- |
| Barra no existe en `BD Alimentadores` | BARRA_NO_EN_BASE              |
| Faltan datos para evaluar             | DATOS_INSUFICIENTES           |
| `faltan ≠ vacío` y `sobran = vacío`   | FALTAN_ALIMENTADORES          |
| `faltan = vacío` y `sobran ≠ vacío`   | SOBRAN_ALIMENTADORES          |
| `faltan ≠ vacío` y `sobran ≠ vacío`   | FALTAN_Y_SOBRAN_ALIMENTADORES |
| `faltan = vacío` y `sobran = vacío`   | OK                            |

