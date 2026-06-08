# HU-002 - Validación de alimentadores por barra

## Contexto

Esta validación corresponde al eje de **alimentadores**.

El objetivo es revisar si los alimentadores informados en el borrador para una barra coinciden con los alimentadores esperados según `BD Alimentadores`.

La comparación se realiza por documento y barra.

---

## Input

- Archivo Excel borrador.
- Hoja: `BD Consolidado`.
- Base de referencia: `BD Alimentadores`.

---

## Columnas del borrador

| Campo | Columna |
|---|---:|
| DOCUMENTO | B |
| ID PAÑO | D |
| ID BARRA PUNTO DE CONTROL | G |
| ID COORDINADO AFECTADO | H |
| CLAVE UNICA / ID UNICO DE ALIMENTADOR | Por nombre de columna |
| NOMBRE ALIMENTADOR / CONSUMO | Por nombre de columna |
| COMENTARIO DISTRIBUIDORA SOBRE ALIMENTADORES | Por nombre de columna |

---

## Unidad de análisis

```text
DOCUMENTO + ID BARRA PUNTO DE CONTROL
```

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

| Condición | Tipo |
|---|---|
| Barra no existe en `BD Alimentadores` | BARRA NO EN BASE |
| Faltan datos para evaluar | DATOS INSUFICIENTES |
| `faltan ≠ vacío` y `sobran = vacío` | FALTAN ALIMENTADORES |
| `faltan = vacío` y `sobran ≠ vacío` | SOBRAN ALIMENTADORES |
| `faltan ≠ vacío` y `sobran ≠ vacío` | FALTAN Y SOBRAN ALIMENTADORES |
| `faltan = vacío` y `sobran = vacío` | OK |

---

## Flujo

1. Validar estructura del borrador.
2. Validar existencia de `BD Alimentadores`.
3. Agrupar filas por `DOCUMENTO + ID BARRA PUNTO DE CONTROL`.
4. Obtener alimentadores observados desde `CLAVE UNICA`.
5. Obtener alimentadores esperados desde `BD Alimentadores`.
6. Comparar conjuntos.
7. Registrar hallazgo si hay diferencias.
8. Marcar en celeste las filas del grupo observado.
9. Reportar hallazgo por documento.

---

## Reglas de validación

- Comparar `CLAVE UNICA` del borrador contra `ID UNICO` de `BD Alimentadores`.
- Para obtener esperados, filtrar `BD Alimentadores` por `ID Punto Control`.
- `ID Punto Control` se compara contra `ID BARRA PUNTO DE CONTROL`.
- Si falta `DOCUMENTO`, `ID BARRA PUNTO DE CONTROL` o `CLAVE UNICA`, registrar `DATOS INSUFICIENTES`.
- Si la barra no existe en `BD Alimentadores`, registrar `BARRA NO EN BASE`.
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
| Grupo con inconsistencia de alimentadores | Celeste | `#CCFFFF` |
