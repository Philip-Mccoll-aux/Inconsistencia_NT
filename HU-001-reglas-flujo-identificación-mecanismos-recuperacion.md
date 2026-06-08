# HU-001 - Validación de mecanismos de recuperación

## Contexto

La `FECHA Y HORA NORMALIZACIÓN DE CONSUMO` no debe ser menor que la `FECHA Y HORA DISPONIBILIDAD DE LA BARRA`, salvo que exista observación o mecanismo de recuperación informado.

## Input

- Archivo Excel borrador.
- Hoja: `BD Consolidado`.

## Columnas

| Campo | Columna |
|---|---:|
| DOCUMENTO | B |
| ID PAÑO | D |
| ID BARRA PUNTO DE CONTROL | G |
| ID COORDINADO AFECTADO | H |
| FECHA Y HORA DISPONIBILIDAD DE LA BARRA | U |
| FECHA Y HORA NORMALIZACIÓN DE CONSUMO | V |
| OBSERVACIONES | W |
| Comunas | X |
| Mecanismo de Recuperación | AO |
| Comentarios mecanismo de recuperación | AQ |

## Pre-revisiones mínimas

Antes de aplicar la validación principal:

1. Verificar que exista la hoja `BD Consolidado`.
2. Verificar que existan las columnas requeridas.
3. Validar que cada fila tenga informada la triada:
   - `ID PAÑO` [D]
   - `ID BARRA PUNTO DE CONTROL` [G]
   - `ID COORDINADO AFECTADO` [H]
4. Validar que `Comunas` [X] no esté vacío.
5. Si `Comunas` [X] tiene varias comunas separadas por coma, tratarlas como tokens independientes.
6. Para cada combinación `DOCUMENTO` [B] + `ID BARRA PUNTO DE CONTROL` [G], debe existir una sola `FECHA Y HORA DISPONIBILIDAD DE LA BARRA` [U].
7. Para cada combinación `DOCUMENTO` [B] + `ID PAÑO` [D] + `ID BARRA PUNTO DE CONTROL` [G] + `ID COORDINADO AFECTADO` [H] + `Comuna` [X], debe existir una sola `FECHA Y HORA NORMALIZACIÓN DE CONSUMO` [V].

Si falla la estructura básica, detener.  
Si fallan las demás pre-revisiones, registrar hallazgo.

## Regla principal

Marcar la fila si se cumple:

```text
V < U
AND W vacío
AND AO vacío
AND AQ vacío
```

Donde:

- `U` = `FECHA Y HORA DISPONIBILIDAD DE LA BARRA`
- `V` = `FECHA Y HORA NORMALIZACIÓN DE CONSUMO`
- `W` = `OBSERVACIONES`
- `AO` = `Mecanismo de Recuperación`
- `AQ` = `Comentarios mecanismo de recuperación`

## Pseudocódigo

```pseudo
abrir Excel
usar hoja "BD Consolidado"

validar estructura mínima
si falla:
    detener

ejecutar pre-revisiones
registrar hallazgos

para cada fila:
    U = parse_fecha_hora(columna U)
    V = parse_fecha_hora(columna V)
    W = columna W
    AO = columna AO
    AQ = columna AQ

    si U o V no son fechas válidas:
        continuar

    si V < U y vacio(W) y vacio(AO) y vacio(AQ):
        marcar fila completa color celeste
        registrar hallazgo HU-001
```

## Tratamiento de datos

- Comparar `U` y `V` como fecha/hora, no como texto.
- Considerar vacío: `null`, `NaN`, `""` o texto con solo espacios.
- Aplicar `trim` antes de evaluar campos vacíos.
- No modificar valores originales.

## Output

- Mismo Excel de entrada.
- Filas observadas marcadas en celeste.
- Reporte de hallazgos por documento.

## Color

| Caso | Color | HEX |
|---|---|---|
| Fila observada | Celeste | `#CCFFFF` |
