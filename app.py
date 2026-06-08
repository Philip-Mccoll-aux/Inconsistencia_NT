import io
import unicodedata
import re
from copy import copy
from datetime import datetime

import openpyxl
from openpyxl.styles import PatternFill
import streamlit as st

# ── constantes ──────────────────────────────────────────────────────────────
CELESTE = PatternFill(start_color="CCFFFF", end_color="CCFFFF", fill_type="solid")
HEADER_SEARCH_ROWS = 10  # filas donde buscar el encabezado


# ── utilidades ───────────────────────────────────────────────────────────────

def normalize_id(val):
    """Convierte a str, quita .0 final y espacios."""
    if val is None:
        return ""
    s = str(val).strip()
    if s.endswith(".0") and s[:-2].lstrip("-").isdigit():
        s = s[:-2]
    return s


def is_empty(val):
    if val is None:
        return True
    s = str(val).strip()
    return s == "" or s.lower() == "nan"


def normalize_comuna(c):
    """Mayúsculas, sin tildes, sin espacios extra."""
    c = str(c).strip().upper()
    c = "".join(
        ch for ch in unicodedata.normalize("NFD", c)
        if unicodedata.category(ch) != "Mn"
    )
    return re.sub(r"\s+", " ", c)


def split_comunas(val):
    if is_empty(val):
        return set()
    tokens = re.split(r"[,;/]", str(val))
    return {normalize_comuna(t) for t in tokens if t.strip()}


def col_letter(idx):
    return openpyxl.utils.get_column_letter(idx + 1)


def find_header_row(ws):
    """Devuelve (header_row_index_0based, col_map {nombre: idx_0based})."""
    for i, row in enumerate(ws.iter_rows(values_only=True)):
        if i >= HEADER_SEARCH_ROWS:
            break
        names = [str(v).strip() if v is not None else "" for v in row]
        if "DOCUMENTO" in names and "ID PAÑO" in names:
            col_map = {name: j for j, name in enumerate(names) if name}
            return i, col_map
    return None, {}


def mark_row(ws, row_num, n_cols):
    for c in range(1, n_cols + 2):
        ws.cell(row=row_num, column=c).fill = CELESTE


def get_col(col_map, *candidates):
    """Devuelve el índice del primer candidato encontrado."""
    for name in candidates:
        if name in col_map:
            return col_map[name]
    return None


# ── carga BD Alimentadores ───────────────────────────────────────────────────

def load_bd_alimentadores(file):
    wb = openpyxl.load_workbook(file, read_only=True, data_only=True)
    if "Alimentadores" not in wb.sheetnames:
        return None, "La hoja 'Alimentadores' no existe en BD Alimentadores."
    ws = wb["Alimentadores"]
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return None, "BD Alimentadores está vacía."
    header = [str(v).strip() if v is not None else "" for v in rows[0]]
    col = {n: i for i, n in enumerate(header) if n}

    required = {"ID UNICO", "ID Punto Control", "ID COORDINADO", "Comunas_25F"}
    missing = required - set(col)
    if missing:
        return None, f"Faltan columnas en BD Alimentadores: {missing}"

    records = []
    for row in rows[1:]:
        records.append({
            "id_unico": normalize_id(row[col["ID UNICO"]]),
            "id_punto_control": normalize_id(row[col["ID Punto Control"]]),
            "id_coordinado": normalize_id(row[col["ID COORDINADO"]]),
            "comunas_25f": str(row[col["Comunas_25F"]] or "").strip(),
        })
    wb.close()

    # índices rápidos
    by_punto_ctrl = {}  # id_punto_control → lista de id_unico
    by_id_unico = {}    # id_unico → set de comunas normalizadas

    for r in records:
        by_punto_ctrl.setdefault(r["id_punto_control"], set()).add(r["id_unico"])
        by_id_unico[r["id_unico"]] = split_comunas(r["comunas_25f"])

    return {"by_punto_ctrl": by_punto_ctrl, "by_id_unico": by_id_unico}, None


# ── HU-001 ───────────────────────────────────────────────────────────────────

def run_hu001(ws, header_row, col_map, hallazgos):
    required = {
        "DOCUMENTO": "B",
        "ID PAÑO": "D",
        "ID BARRA PUNTO DE CONTROL": "G",
        "ID COORDINADO AFECTADO": "H",
        "FECHA Y HORA DISPONIBILIDAD DE LA BARRA": "U",
        "FECHA Y HORA NORMALIZACIÓN DE CONSUMO": "V",
        "OBSERVACIONES": "W",
        "Comunas": "X",
        "Mecanismo de Recuperación": "AO",
    }
    missing = [k for k in required if k not in col_map]
    if missing:
        return [f"[HU-001] Faltan columnas: {missing}"]

    c_doc = col_map["DOCUMENTO"]
    c_pano = col_map["ID PAÑO"]
    c_barra = col_map["ID BARRA PUNTO DE CONTROL"]
    c_coord = col_map["ID COORDINADO AFECTADO"]
    c_u = col_map["FECHA Y HORA DISPONIBILIDAD DE LA BARRA"]
    c_v = col_map["FECHA Y HORA NORMALIZACIÓN DE CONSUMO"]
    c_w = col_map["OBSERVACIONES"]
    c_x = col_map["Comunas"]
    c_ao = col_map["Mecanismo de Recuperación"]
    c_aq = get_col(col_map,
                   "Comentarios mecanismo de recuperación",
                   "Comentario mecanismo de recuperación")

    errors = []
    n_cols = ws.max_column

    # pre-revisiones acumuladas por tipo
    triada_vacias = []
    comunas_vacias = []
    multi_disp = {}     # (doc,barra) → set(U)
    multi_norm = {}     # (doc,pano,barra,coord,comuna) → set(V)

    data_rows = []
    for i, row in enumerate(ws.iter_rows(values_only=True)):
        if i <= header_row:
            continue
        data_rows.append((i, row))

    for i, row in data_rows:
        doc = row[c_doc]
        pano = row[c_pano]
        barra = row[c_barra]
        coord = row[c_coord]
        comunas_val = row[c_x]

        if is_empty(pano) or is_empty(barra) or is_empty(coord):
            triada_vacias.append(i + 1)

        if is_empty(comunas_val):
            comunas_vacias.append(i + 1)

        key_disp = (normalize_id(doc), normalize_id(barra))
        u_val = row[c_u]
        multi_disp.setdefault(key_disp, set()).add(u_val)

        v_val = row[c_v]
        for com in (split_comunas(comunas_val) or {None}):
            key_norm = (normalize_id(doc), normalize_id(pano),
                        normalize_id(barra), normalize_id(coord), com)
            multi_norm.setdefault(key_norm, set()).add(v_val)

    # registrar pre-revisiones
    if triada_vacias:
        hallazgos.append({
            "HU": "HU-001", "Tipo": "PRE-REVISION",
            "Documento": "-",
            "Detalle": f"Triada vacía en filas: {triada_vacias[:20]}{'...' if len(triada_vacias)>20 else ''}",
        })
    if comunas_vacias:
        hallazgos.append({
            "HU": "HU-001", "Tipo": "PRE-REVISION",
            "Documento": "-",
            "Detalle": f"Comunas vacías en filas: {comunas_vacias[:20]}{'...' if len(comunas_vacias)>20 else ''}",
        })
    for key, vals in multi_disp.items():
        if len(vals) > 1:
            hallazgos.append({
                "HU": "HU-001", "Tipo": "MÚLTIPLES FECHAS DISPONIBILIDAD",
                "Documento": key[0],
                "Detalle": f"Barra {key[1]} tiene {len(vals)} fechas distintas de disponibilidad.",
            })
    for key, vals in multi_norm.items():
        if len(vals) > 1:
            hallazgos.append({
                "HU": "HU-001", "Tipo": "MÚLTIPLES FECHAS NORMALIZACIÓN",
                "Documento": key[0],
                "Detalle": f"Combinación {key[1:]}: {len(vals)} fechas distintas de normalización.",
            })

    # regla principal
    marked = 0
    for i, row in data_rows:
        u_val = row[c_u]
        v_val = row[c_v]
        w_val = row[c_w]
        ao_val = row[c_ao]
        aq_val = row[c_aq] if c_aq is not None else None

        if not isinstance(u_val, datetime) or not isinstance(v_val, datetime):
            continue

        if v_val < u_val and is_empty(w_val) and is_empty(ao_val) and is_empty(aq_val):
            mark_row(ws, i + 1, n_cols)
            marked += 1
            hallazgos.append({
                "HU": "HU-001",
                "Tipo": "V < U SIN JUSTIFICACIÓN",
                "Documento": str(row[c_doc] or ""),
                "Detalle": (
                    f"Fila {i+1} | Paño={row[c_pano]} | Barra={row[c_barra]} | "
                    f"U={u_val} | V={v_val}"
                ),
            })

    return errors


# ── HU-002 ───────────────────────────────────────────────────────────────────

def run_hu002(ws, header_row, col_map, bd, hallazgos):
    required = ["DOCUMENTO", "ID PAÑO", "ID BARRA PUNTO DE CONTROL",
                "ID COORDINADO AFECTADO"]
    missing = [k for k in required if k not in col_map]
    if missing:
        return [f"[HU-002] Faltan columnas: {missing}"]

    c_doc = col_map["DOCUMENTO"]
    c_pano = col_map["ID PAÑO"]
    c_barra = col_map["ID BARRA PUNTO DE CONTROL"]
    c_coord = col_map["ID COORDINADO AFECTADO"]
    c_id_unico = get_col(col_map, "ID UNICO", "CLAVE UNICA", "CLAVE ALIMENTADOR")
    c_nombre = get_col(col_map, "NOMBRE ALIMENTADOR / CONSUMO")
    c_comentario = get_col(col_map,
                           "Comentario Distribuidora Inconsistencia Alimentadores",
                           "COMENTARIO DISTRIBUIDORA SOBRE ALIMENTADORES")

    n_cols = ws.max_column
    grupos = {}  # (doc, barra) → lista de (row_idx, row)

    for i, row in enumerate(ws.iter_rows(values_only=True)):
        if i <= header_row:
            continue
        doc = normalize_id(row[c_doc])
        barra = normalize_id(row[c_barra])

        if is_empty(row[c_doc]) or is_empty(row[c_barra]):
            hallazgos.append({
                "HU": "HU-002", "Tipo": "DATOS INSUFICIENTES",
                "Documento": doc,
                "Detalle": f"Fila {i+1}: falta DOCUMENTO o ID BARRA.",
            })
            continue

        key = (doc, barra)
        grupos.setdefault(key, []).append((i, row))

    errors = []
    for (doc, barra), filas in grupos.items():
        # alimentadores observados
        if c_id_unico is None:
            observados = set()
            tipo = "DATOS INSUFICIENTES"
            comentario_dx = ""
        else:
            observados = set()
            comentario_dx = ""
            for _, row in filas:
                val = normalize_id(row[c_id_unico])
                if not is_empty(val):
                    observados.add(val)
                if c_comentario is not None and not is_empty(row[c_comentario]):
                    comentario_dx = str(row[c_comentario]).strip()

        # alimentadores esperados
        esperados = bd["by_punto_ctrl"].get(barra, None)

        if esperados is None:
            tipo = "BARRA NO EN BASE"
            faltan = set()
            sobran = observados
        elif not observados:
            tipo = "DATOS INSUFICIENTES"
            faltan = set()
            sobran = set()
        else:
            faltan = esperados - observados
            sobran = observados - esperados
            if faltan and sobran:
                tipo = "FALTAN Y SOBRAN ALIMENTADORES"
            elif faltan:
                tipo = "FALTAN ALIMENTADORES"
            elif sobran:
                tipo = "SOBRAN ALIMENTADORES"
            else:
                tipo = "OK"

        if tipo != "OK":
            for i, _ in filas:
                mark_row(ws, i + 1, n_cols)
            detalle = f"Barra {barra}"
            if faltan:
                detalle += f" | Faltan: {sorted(faltan)}"
            if sobran:
                detalle += f" | Sobran: {sorted(sobran)}"
            if comentario_dx:
                detalle += f" | Comentario Dx: {comentario_dx}"
            hallazgos.append({
                "HU": "HU-002",
                "Tipo": tipo,
                "Documento": doc,
                "Detalle": detalle,
            })

    return errors


# ── HU-003 ───────────────────────────────────────────────────────────────────

def run_hu003(ws, header_row, col_map, bd, hallazgos):
    required = ["DOCUMENTO", "ID COORDINADO AFECTADO", "Comunas"]
    missing = [k for k in required if k not in col_map]
    if missing:
        return [f"[HU-003] Faltan columnas: {missing}"]

    c_doc = col_map["DOCUMENTO"]
    c_coord = col_map["ID COORDINADO AFECTADO"]
    c_comunas = col_map["Comunas"]
    c_id_unico = get_col(col_map, "ID UNICO", "CLAVE UNICA", "CLAVE ALIMENTADOR")
    c_comentario = get_col(col_map,
                           "Comentario Distribuidora sobre Inconsistencia Comunas",
                           "COMENTARIOS DISTRIBUIDORA SOBRE INCONSISTENCIA DE COMUNAS")

    n_cols = ws.max_column
    grupos = {}  # (doc, id_unico) → lista de (row_idx, row)

    for i, row in enumerate(ws.iter_rows(values_only=True)):
        if i <= header_row:
            continue
        doc = normalize_id(row[c_doc])

        if c_id_unico is None or is_empty(row[c_id_unico]):
            if is_empty(row[c_doc]) or is_empty(row[c_comunas]):
                hallazgos.append({
                    "HU": "HU-003", "Tipo": "DATOS INSUFICIENTES",
                    "Documento": doc,
                    "Detalle": f"Fila {i+1}: falta DOCUMENTO, ID UNICO o Comunas.",
                })
            continue

        id_unico = normalize_id(row[c_id_unico])
        key = (doc, id_unico)
        grupos.setdefault(key, []).append((i, row))

    errors = []
    for (doc, id_unico), filas in grupos.items():
        # comunas observadas
        comunas_obs = set()
        comentario_dx = ""
        for _, row in filas:
            comunas_obs |= split_comunas(row[c_comunas])
            if c_comentario is not None and not is_empty(row[c_comentario]):
                comentario_dx = str(row[c_comentario]).strip()

        # comunas esperadas
        esperadas = bd["by_id_unico"].get(id_unico, None)

        if esperadas is None:
            tipo = "ID NO EN BASE"
            faltan = set()
            sobran = comunas_obs
        elif not comunas_obs:
            tipo = "DATOS INSUFICIENTES"
            faltan = set()
            sobran = set()
        else:
            faltan = esperadas - comunas_obs
            sobran = comunas_obs - esperadas
            if faltan and sobran:
                tipo = "FALTAN Y SOBRAN COMUNAS"
            elif faltan:
                tipo = "FALTAN COMUNAS"
            elif sobran:
                tipo = "SOBRAN COMUNAS"
            else:
                tipo = "OK"

        if tipo != "OK":
            for i, _ in filas:
                mark_row(ws, i + 1, n_cols)
            detalle = f"ID UNICO {id_unico}"
            if faltan:
                detalle += f" | Faltan: {sorted(faltan)}"
            if sobran:
                detalle += f" | Sobran: {sorted(sobran)}"
            if comentario_dx:
                detalle += f" | Comentario Dx: {comentario_dx}"
            hallazgos.append({
                "HU": "HU-003",
                "Tipo": tipo,
                "Documento": doc,
                "Detalle": detalle,
            })

    return errors


# ── reporte ──────────────────────────────────────────────────────────────────

def build_report(wb, hallazgos):
    sheet_name = "Reporte Hallazgos"
    if sheet_name in wb.sheetnames:
        del wb[sheet_name]
    ws = wb.create_sheet(sheet_name)

    headers = ["HU", "Tipo", "Documento", "Detalle"]
    header_fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    from openpyxl.styles import Font, Alignment
    for c, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=c, value=h)
        cell.fill = header_fill
        cell.font = Font(color="FFFFFF", bold=True)
        cell.alignment = Alignment(horizontal="center")

    for r, h in enumerate(hallazgos, 2):
        ws.cell(row=r, column=1, value=h["HU"])
        ws.cell(row=r, column=2, value=h["Tipo"])
        ws.cell(row=r, column=3, value=h["Documento"])
        ws.cell(row=r, column=4, value=h["Detalle"])

    ws.column_dimensions["A"].width = 10
    ws.column_dimensions["B"].width = 32
    ws.column_dimensions["C"].width = 20
    ws.column_dimensions["D"].width = 90


# ── UI Streamlit ─────────────────────────────────────────────────────────────

st.set_page_config(page_title="Validador NT - Inconsistencias", layout="wide")
st.title("Validador de Inconsistencias NT")
st.markdown(
    "Carga los dos archivos Excel y ejecuta las validaciones **HU-001**, **HU-002** y **HU-003**."
)

col1, col2 = st.columns(2)
with col1:
    borrador_file = st.file_uploader(
        "📄 Archivo borrador EAF (Excel)",
        type=["xlsx"],
        key="borrador",
    )
with col2:
    bd_file = st.file_uploader(
        "📋 BD Alimentadores (Excel)",
        type=["xlsx"],
        key="bd",
    )

hu_opts = st.multiselect(
    "Validaciones a ejecutar",
    ["HU-001 – Mecanismos de recuperación",
     "HU-002 – Alimentadores por barra",
     "HU-003 – Comunas por alimentador"],
    default=[
        "HU-001 – Mecanismos de recuperación",
        "HU-002 – Alimentadores por barra",
        "HU-003 – Comunas por alimentador",
    ],
)

run_btn = st.button("▶ Ejecutar validaciones", type="primary",
                    disabled=(borrador_file is None or bd_file is None))

if run_btn:
    with st.spinner("Procesando…"):
        # cargar BD Alimentadores
        bd, err = load_bd_alimentadores(bd_file)
        if err:
            st.error(err)
            st.stop()

        # cargar borrador
        wb = openpyxl.load_workbook(borrador_file, data_only=True)
        if "BD Consolidado" not in wb.sheetnames:
            st.error("El archivo borrador no contiene la hoja 'BD Consolidado'.")
            st.stop()

        ws = wb["BD Consolidado"]
        header_row, col_map = find_header_row(ws)
        if header_row is None:
            st.error("No se encontró fila de encabezado con 'DOCUMENTO' e 'ID PAÑO' en 'BD Consolidado'.")
            st.stop()

        hallazgos = []
        all_errors = []

        if "HU-001" in " ".join(hu_opts):
            errs = run_hu001(ws, header_row, col_map, hallazgos)
            all_errors.extend(errs)

        if "HU-002" in " ".join(hu_opts):
            errs = run_hu002(ws, header_row, col_map, bd, hallazgos)
            all_errors.extend(errs)

        if "HU-003" in " ".join(hu_opts):
            errs = run_hu003(ws, header_row, col_map, bd, hallazgos)
            all_errors.extend(errs)

        build_report(wb, hallazgos)

        # guardar en buffer
        buf = io.BytesIO()
        wb.save(buf)
        buf.seek(0)

    # resultados
    if all_errors:
        for e in all_errors:
            st.error(e)

    hu1 = [h for h in hallazgos if h["HU"] == "HU-001"]
    hu2 = [h for h in hallazgos if h["HU"] == "HU-002"]
    hu3 = [h for h in hallazgos if h["HU"] == "HU-003"]

    st.success(
        f"Validación completa. Hallazgos: "
        f"HU-001={len(hu1)} | HU-002={len(hu2)} | HU-003={len(hu3)}"
    )

    tab1, tab2, tab3 = st.tabs(
        ["HU-001 Mecanismos", "HU-002 Alimentadores", "HU-003 Comunas"]
    )

    def show_table(tab, items, hu_label):
        with tab:
            if not items:
                st.info(f"Sin hallazgos para {hu_label}.")
            else:
                import pandas as pd
                df = pd.DataFrame(items)[["Tipo", "Documento", "Detalle"]]
                st.dataframe(df, use_container_width=True, hide_index=True)

    show_table(tab1, hu1, "HU-001")
    show_table(tab2, hu2, "HU-002")
    show_table(tab3, hu3, "HU-003")

    st.download_button(
        label="⬇ Descargar Excel con hallazgos marcados",
        data=buf,
        file_name="EAF_validado.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
