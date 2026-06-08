# HU-003 - Validación de comunas por alimentador

## Contexto

Esta validación corresponde al eje de **comunas**.

El objetivo es revisar si las comunas informadas en el borrador para cada alimentador coinciden con las comunas esperadas según `Comunas_25F` de `BD Alimentadores`.

La comparación se realiza por documento y alimentador.

---

## Input

- Archivo Excel borrador.
- Hoja: `BD Consolidado`.
- Base de referencia: `BD Alimentadores`.
- Base oficial de comunas, si aplica.

---

## Columnas del borrador

| Campo | Columna |
|---|---:|
| DOCUMENTO | B |
| ID COORDINADO AFECTADO | H |
| CLAVE UNICA / ID UNICO DE ALIMENTADOR | Por nombre de columna |
| Comunas | X |
| COMENTARIOS DISTRIBUIDORA SOBRE INCONSISTENCIA DE COMUNAS | Por nombre de columna |

---

## Columnas de BD Alimentadores

| Campo |
|---|
| ID UNICO |
| Comunas_25F |
| ID COORDINADO |
| Nombre Alimentador_25F |

---

## Unidad de análisis

```text
DOCUMENTO + CLAVE UNICA
```

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

- Separar comunas si vienen en una misma celda usando:
  - coma `,`
  - punto y coma `;`
  - slash `/`
- Cada comuna debe tratarse como token independiente.
- Normalizar comunas antes de comparar:
  - mayúsculas;
  - sin tildes;
  - sin espacios extra;
  - conjuntos sin orden.

---

## Clasificación de hallazgos

| Condición | Tipo |
|---|---|
| `CLAVE UNICA` no existe en `BD Alimentadores` | ID NO EN BASE |
| Faltan datos para evaluar | DATOS INSUFICIENTES |
| `faltan ≠ vacío` y `sobran = vacío` | FALTAN COMUNAS |
| `faltan = vacío` y `sobran ≠ vacío` | SOBRAN COMUNAS |
| `faltan ≠ vacío` y `sobran ≠ vacío` | FALTAN Y SOBRAN COMUNAS |
| `faltan = vacío` y `sobran = vacío` | OK |

---

## Flujo

1. Validar estructura del borrador.
2. Validar existencia de `BD Alimentadores`.
3. Agrupar filas por `DOCUMENTO + CLAVE UNICA`.
4. Obtener comunas observadas desde `Comunas` [X].
5. Obtener comunas esperadas desde `Comunas_25F`.
6. Normalizar ambos conjuntos de comunas.
7. Comparar conjuntos.
8. Registrar hallazgo si hay diferencias.
9. Marcar en celeste las filas del grupo observado.
10. Reportar hallazgo por documento.

---

## Reglas de validación

- Comparar `Comunas` del borrador contra `Comunas_25F` de `BD Alimentadores`.
- Para obtener comunas esperadas, buscar `CLAVE UNICA` del borrador como `ID UNICO` en `BD Alimentadores`.
- Si falta `DOCUMENTO`, `CLAVE UNICA` o `Comunas`, registrar `DATOS INSUFICIENTES`.
- Si `CLAVE UNICA` no existe en `BD Alimentadores`, registrar `ID NO EN BASE`.
- Si existe base oficial de comunas, validar que cada comuna informada exista en dicha base.
- Si existe base empresa-comuna, validar que la comuna esté asociada al `ID COORDINADO AFECTADO` [H].
- El comentario de la distribuidora puede respaldar el hallazgo, pero no lo elimina automáticamente.

---

## Tratamiento de datos

- Normalizar IDs antes de comparar.
- Eliminar sufijos `.0` en IDs numéricos.
- Comparar IDs como texto normalizado.
- Considerar vacío: `null`, `NaN`, texto vacío o texto compuesto solo por espacios.
- No modificar valores originales del borrador.

---

## Output esperado

- Mismo Excel de entrada.
- Filas del grupo con inconsistencia marcadas en celeste.
- Reporte de hallazgos por documento.

---

## Color

| Caso | Color | HEX |
|---|---|---|
| Grupo con inconsistencia de comunas | Celeste | `#CCFFFF` |
