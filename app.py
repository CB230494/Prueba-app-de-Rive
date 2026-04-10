# app.py
# -*- coding: utf-8 -*-

import re
import unicodedata
from io import BytesIO

import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.cell.cell import MergedCell

from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER
from reportlab.platypus import (
    SimpleDocTemplate,
    Paragraph,
    Spacer,
    Table,
    TableStyle,
)

# =========================================================
# CONFIGURACIÓN GENERAL
# =========================================================
st.set_page_config(
    page_title="Seguimiento de líneas de acción",
    layout="wide"
)

# =========================================================
# SESSION STATE
# =========================================================
if "lineas_guardadas" not in st.session_state:
    st.session_state["lineas_guardadas"] = {}

if "archivo_nombre" not in st.session_state:
    st.session_state["archivo_nombre"] = ""

if "pdf_final" not in st.session_state:
    st.session_state["pdf_final"] = None

if "ultimo_resumen" not in st.session_state:
    st.session_state["ultimo_resumen"] = pd.DataFrame()

# =========================================================
# ESTILOS CSS
# Ajustados para modo oscuro y también funcionales en claro
# =========================================================
st.markdown("""
<style>
.block-container {
    max-width: 1550px;
    padding-top: 1rem;
    padding-bottom: 2rem;
}

.main-title {
    font-size: 2rem;
    font-weight: 700;
    margin-bottom: 0.2rem;
    color: inherit;
}

.sub-note {
    font-size: 0.95rem;
    color: rgba(250,250,250,0.78);
    margin-bottom: 1rem;
}

.info-card {
    border: 1px solid rgba(255,255,255,0.08);
    background: rgba(255,255,255,0.03);
    border-radius: 14px;
    padding: 16px 18px;
    margin-bottom: 16px;
}

.line-card {
    border: 1px solid rgba(255,255,255,0.08);
    border-radius: 14px;
    padding: 18px;
    margin-bottom: 22px;
    background: rgba(255,255,255,0.02);
    box-shadow: 0 1px 2px rgba(0,0,0,0.18);
}

.section-title {
    font-size: 1.12rem;
    font-weight: 700;
    margin-bottom: 0.6rem;
    color: inherit;
}

.kpi-box {
    border: 1px solid rgba(255,255,255,0.08);
    border-radius: 12px;
    padding: 12px;
    background: rgba(255,255,255,0.03);
    text-align: center;
    color: inherit;
    min-height: 82px;
    display: flex;
    flex-direction: column;
    justify-content: center;
}

.small-note {
    font-size: 0.9rem;
    color: rgba(250,250,250,0.72);
}

hr.soft-line {
    border: none;
    height: 1px;
    background: rgba(255,255,255,0.12);
    margin-top: 10px;
    margin-bottom: 14px;
}

div[data-testid="stDataFrame"] {
    border-radius: 12px;
    overflow: hidden;
}

div[data-testid="stExpander"] {
    border: 1px solid rgba(255,255,255,0.08);
    border-radius: 12px;
}

[data-testid="stSidebar"] .stMarkdown,
[data-testid="stSidebar"] p,
[data-testid="stSidebar"] li,
[data-testid="stSidebar"] label {
    color: inherit !important;
}
</style>
""", unsafe_allow_html=True)


# =========================================================
# UTILIDADES DE TEXTO
# =========================================================
def strip_accents(text: str) -> str:
    if text is None:
        return ""
    text = str(text)
    return "".join(
        ch for ch in unicodedata.normalize("NFD", text)
        if unicodedata.category(ch) != "Mn"
    )


def normalize_text(value) -> str:
    """
    Normaliza texto de forma robusta:
    - minúsculas
    - sin tildes
    - sin saltos de línea
    - espacios compactados
    - puntuación reducida
    """
    if value is None:
        return ""

    s = str(value)
    s = strip_accents(s).lower()
    s = s.replace("\n", " ").replace("\r", " ")
    s = s.replace("°", " ")
    s = s.replace("№", " ")
    s = re.sub(r"[|]+", " ", s)
    s = re.sub(r"[_]+", " ", s)
    s = re.sub(r"[-]+", " ", s)
    s = re.sub(r"[#]+", " # ", s)
    s = re.sub(r"[:;,]+", " ", s)
    s = re.sub(r"[()]+", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s


def clean_text(value) -> str:
    if value is None:
        return ""
    return str(value).strip()


def safe_str(value) -> str:
    if value is None:
        return ""
    return str(value).strip()


def is_nonempty(value) -> bool:
    if value is None:
        return False
    if isinstance(value, str):
        return value.strip() != ""
    return True


# =========================================================
# PALABRAS CLAVE ROBUSTAS
# =========================================================
KEY_DELEGACION = [
    "delegacion",
]

KEY_LINEA = [
    "linea de accion",
    "linea accion",
    "linea de coordinacion",
    "linea de accion #",
    "linea accion #",
]

KEY_PROBLEMATICA = [
    "problematica",
    "problematica de linea de accion",
    "problematica de linea",
    "problematica de la linea",
]

KEY_LIDER = [
    "lider estrategico",
    "lider",
]

KEY_HEADER_INDICADOR = ["indicador", "indicadores"]
KEY_HEADER_META = ["meta"]
KEY_HEADER_AVANCE = ["avance"]
KEY_HEADER_DESCRIPCION = ["descripcion", "descripcio"]
KEY_HEADER_CANTIDAD = ["cantidad"]
KEY_HEADER_OBSERVACIONES = ["observaciones", "observacion", "obs"]

BAD_LINE_TEXT = [
    "problematica",
    "lider",
    "delegacion",
    "trimestre",
    "indicador",
    "meta",
    "avance",
    "descripcion",
    "cantidad",
    "observacion",
    "linea de accion",
]


# =========================================================
# UTILIDADES DE EXCEL
# =========================================================
def get_effective_cell_value(ws, row, col):
    if row < 1 or col < 1:
        return None

    cell = ws.cell(row=row, column=col)

    if not isinstance(cell, MergedCell):
        return cell.value

    for merged_range in ws.merged_cells.ranges:
        if (
            merged_range.min_row <= row <= merged_range.max_row
            and merged_range.min_col <= col <= merged_range.max_col
        ):
            return ws.cell(merged_range.min_row, merged_range.min_col).value

    return None


def row_values(ws, row, max_col=None):
    if max_col is None:
        max_col = ws.max_column
    return [get_effective_cell_value(ws, row, c) for c in range(1, max_col + 1)]


def row_text(ws, row, max_col=None):
    vals = row_values(ws, row, max_col)
    return " | ".join("" if v is None else str(v) for v in vals)


def get_right_nonempty(ws, row, col, max_steps=12):
    for c in range(col + 1, min(ws.max_column, col + max_steps) + 1):
        val = get_effective_cell_value(ws, row, c)
        if is_nonempty(val):
            return val
    return ""


def get_left_nonempty(ws, row, col, max_steps=8):
    for c in range(col - 1, max(1, col - max_steps), -1):
        val = get_effective_cell_value(ws, row, c)
        if is_nonempty(val):
            return val
    return ""


def get_down_nonempty(ws, row, col, max_steps=6):
    for r in range(row + 1, min(ws.max_row, row + max_steps) + 1):
        val = get_effective_cell_value(ws, r, col)
        if is_nonempty(val):
            return val
    return ""


def get_up_nonempty(ws, row, col, max_steps=6):
    for r in range(row - 1, max(1, row - max_steps), -1):
        val = get_effective_cell_value(ws, r, col)
        if is_nonempty(val):
            return val
    return ""


def get_near_nonempty(ws, row, col):
    """
    Busca un valor cercano de forma flexible.
    Prioriza derecha, luego abajo, luego izquierda, luego arriba.
    """
    for func, steps in [
        (get_right_nonempty, 12),
        (get_down_nonempty, 6),
        (get_left_nonempty, 8),
        (get_up_nonempty, 6),
    ]:
        val = func(ws, row, col, max_steps=steps)
        if is_nonempty(val):
            return val
    return ""


def sheet_density_score(ws, max_row=220, max_col=40):
    score = 0
    mr = min(ws.max_row, max_row)
    mc = min(ws.max_column, max_col)

    for r in range(1, mr + 1):
        txt = normalize_text(row_text(ws, r, mc))
        score += 8 if "linea de accion" in txt else 0
        score += 6 if "problematica" in txt else 0
        score += 6 if "lider estrategico" in txt or "lider" in txt else 0
        score += 4 if "indicador" in txt else 0
        score += 4 if "meta" in txt else 0
        score += 4 if "avance" in txt else 0
        score += 4 if "descripcion" in txt else 0
        score += 4 if "cantidad" in txt else 0
        score += 4 if "observaciones" in txt else 0
        score += 3 if "delegacion" in txt else 0
        score += 2 if "trimestre" in txt else 0

    return score


# =========================================================
# DETECCIÓN DE HOJA PRINCIPAL
# =========================================================
def find_best_main_sheet(wb):
    candidates = []

    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        score = sheet_density_score(ws)

        name_norm = normalize_text(sheet_name)

        if "informe de avance" in name_norm:
            score += 50
        if "informe" in name_norm:
            score += 20
        if "avance" in name_norm:
            score += 20
        if "dashboard" in name_norm or "datos" in name_norm or "sumatoria" in name_norm:
            score -= 20

        candidates.append((sheet_name, score))

    candidates.sort(key=lambda x: x[1], reverse=True)
    return candidates[0][0] if candidates else wb.sheetnames[0]


# =========================================================
# DATOS GENERALES
# =========================================================
def get_delegacion(ws):
    max_row = min(ws.max_row, 80)
    max_col = min(ws.max_column, 25)

    best_candidate = ""

    for r in range(1, max_row + 1):
        for c in range(1, max_col + 1):
            val = get_effective_cell_value(ws, r, c)
            txt = normalize_text(val)

            if "delegacion" in txt:
                near = get_near_nonempty(ws, r, c)
                if is_nonempty(near):
                    cleaned = clean_text(near)
                    if normalize_text(cleaned) != "delegacion":
                        return cleaned

    for r in range(1, max_row + 1):
        row_vals = row_values(ws, r, max_col)
        joined = normalize_text(" | ".join("" if v is None else str(v) for v in row_vals))
        if "delegacion" in joined:
            for v in row_vals:
                vt = normalize_text(v)
                if is_nonempty(v) and "delegacion" not in vt and len(vt) > 2:
                    best_candidate = clean_text(v)
                    break
            if best_candidate:
                return best_candidate

    return ""


def get_fecha_actualizacion(ws):
    max_row = min(ws.max_row, 30)
    max_col = min(ws.max_column, 25)

    for r in range(1, max_row + 1):
        for c in range(1, max_col + 1):
            val = get_effective_cell_value(ws, r, c)
            txt = normalize_text(val)

            if "fecha de actualizacion" in txt:
                near = get_near_nonempty(ws, r, c)
                if is_nonempty(near):
                    return clean_text(near)

    return ""


# =========================================================
# BÚSQUEDA ROBUSTA DE BLOQUES
# =========================================================
def looks_like_bad_line_value(text: str) -> bool:
    t = normalize_text(text)
    if not t:
        return True
    return any(p in t for p in BAD_LINE_TEXT)


def extract_line_number_from_area(ws, start_row, start_col):
    candidates = []

    for c in range(start_col, min(ws.max_column, start_col + 10) + 1):
        val = get_effective_cell_value(ws, start_row, c)
        if is_nonempty(val):
            txt = clean_text(val)
            if not looks_like_bad_line_value(txt):
                candidates.append(txt)

    for r in range(start_row, min(ws.max_row, start_row + 3) + 1):
        for c in range(max(1, start_col - 1), min(ws.max_column, start_col + 8) + 1):
            val = get_effective_cell_value(ws, r, c)
            if is_nonempty(val):
                txt = clean_text(val)
                if not looks_like_bad_line_value(txt):
                    candidates.append(txt)

    for candidate in candidates:
        m = re.search(r"(\d+(?:[.\-]\d+)?)", str(candidate))
        if m:
            return m.group(1)

    short_candidates = [x for x in candidates if len(str(x).strip()) <= 40]
    if short_candidates:
        return str(short_candidates[0]).strip()

    return ""


def find_line_action_starts(ws):
    starts = []
    mr = ws.max_row
    mc = min(ws.max_column, 25)

    for r in range(1, mr + 1):
        for c in range(1, mc + 1):
            val = get_effective_cell_value(ws, r, c)
            txt = normalize_text(val)

            if any(k in txt for k in KEY_LINEA):
                line_value = extract_line_number_from_area(ws, r, c)
                starts.append({
                    "row": r,
                    "col": c,
                    "line_value": line_value
                })
                break

    cleaned = []
    last_row = -999

    for item in starts:
        if item["row"] - last_row > 2:
            cleaned.append(item)
            last_row = item["row"]

    return cleaned


def search_value_near_keywords(ws, start_row, end_row, keywords, value_blacklist=None):
    if value_blacklist is None:
        value_blacklist = []

    best_value = ""
    best_score = -1

    for r in range(start_row, min(end_row, ws.max_row) + 1):
        for c in range(1, ws.max_column + 1):
            val = get_effective_cell_value(ws, r, c)
            txt = normalize_text(val)

            local_score = sum(1 for kw in keywords if kw in txt)
            if local_score > 0:
                candidates = [
                    get_right_nonempty(ws, r, c, 12),
                    get_down_nonempty(ws, r, c, 5),
                    get_left_nonempty(ws, r, c, 4),
                    get_up_nonempty(ws, r, c, 2),
                ]

                for cand in candidates:
                    cand_clean = clean_text(cand)
                    cand_norm = normalize_text(cand_clean)

                    if not cand_norm:
                        continue
                    if cand_norm == txt:
                        continue
                    if any(b in cand_norm for b in value_blacklist):
                        continue

                    bonus = 0
                    if len(cand_clean) > 2:
                        bonus += 1
                    if len(cand_clean) > 8:
                        bonus += 1
                    if cand_norm not in keywords:
                        bonus += 1

                    total_score = local_score + bonus
                    if total_score > best_score:
                        best_score = total_score
                        best_value = cand_clean

    return best_value


def detect_trimester(ws, start_row, end_row):
    roman_map = {
        "i": "I",
        "ii": "II",
        "iii": "III",
        "iv": "IV",
        "1": "I",
        "2": "II",
        "3": "III",
        "4": "IV",
        "primer": "I",
        "segundo": "II",
        "tercer": "III",
        "cuarto": "IV",
    }

    for r in range(start_row, min(end_row, ws.max_row) + 1):
        for c in range(1, ws.max_column + 1):
            val = get_effective_cell_value(ws, r, c)
            txt = normalize_text(val)

            if "trimestre" in txt:
                candidates = [
                    get_right_nonempty(ws, r, c, 5),
                    get_down_nonempty(ws, r, c, 2),
                    get_left_nonempty(ws, r, c, 2),
                ]

                for cand in candidates:
                    ct = normalize_text(cand)
                    for k, out in roman_map.items():
                        if re.fullmatch(rf"{re.escape(k)}", ct) or f"{k} trimestre" in ct:
                            return out

                row_txt = normalize_text(row_text(ws, r))
                if "iv trimestre" in row_txt or "4 trimestre" in row_txt or "cuarto trimestre" in row_txt:
                    return "IV"
                if "iii trimestre" in row_txt or "3 trimestre" in row_txt or "tercer trimestre" in row_txt:
                    return "III"
                if "ii trimestre" in row_txt or "2 trimestre" in row_txt or "segundo trimestre" in row_txt:
                    return "II"
                if "i trimestre" in row_txt or "1 trimestre" in row_txt or "primer trimestre" in row_txt:
                    return "I"

    return ""


# =========================================================
# DETECCIÓN DE TABLA
# =========================================================
def detect_header_row(ws, start_row, end_row):
    best_row = None
    best_score = -1

    for r in range(start_row, min(end_row, ws.max_row) + 1):
        vals = row_values(ws, r)

        found = {
            "indicador": False,
            "meta": False,
            "avance": False,
            "descripcion": False,
            "cantidad": False,
            "observaciones": False,
        }

        for v in vals:
            t = normalize_text(v)

            if any(k in t for k in KEY_HEADER_INDICADOR):
                found["indicador"] = True
            if any(k in t for k in KEY_HEADER_META):
                found["meta"] = True
            if any(k in t for k in KEY_HEADER_AVANCE):
                found["avance"] = True
            if any(k in t for k in KEY_HEADER_DESCRIPCION):
                found["descripcion"] = True
            if any(k in t for k in KEY_HEADER_CANTIDAD):
                found["cantidad"] = True
            if any(k in t for k in KEY_HEADER_OBSERVACIONES):
                found["observaciones"] = True

        score = sum(found.values())

        if found["indicador"] and found["meta"] and found["avance"] and found["descripcion"]:
            score += 4

        if score > best_score:
            best_score = score
            best_row = r

    if best_score < 3:
        return None

    return best_row


def map_headers(ws, header_row):
    header_map = {}

    for c in range(1, ws.max_column + 1):
        val = get_effective_cell_value(ws, header_row, c)
        t = normalize_text(val)

        if any(k in t for k in KEY_HEADER_INDICADOR) and "Indicador" not in header_map:
            header_map["Indicador"] = c

        elif any(k in t for k in KEY_HEADER_META) and "Meta (editable)" not in header_map:
            header_map["Meta (editable)"] = c

        elif any(k in t for k in KEY_HEADER_AVANCE) and "Avance (Editable)" not in header_map:
            header_map["Avance (Editable)"] = c

        elif any(k in t for k in KEY_HEADER_DESCRIPCION) and "Descripción (editable)" not in header_map:
            header_map["Descripción (editable)"] = c

        elif any(k in t for k in KEY_HEADER_CANTIDAD) and "Cantidad (editable)" not in header_map:
            header_map["Cantidad (editable)"] = c

        elif any(k in t for k in KEY_HEADER_OBSERVACIONES) and "Observaciones (Editable)" not in header_map:
            header_map["Observaciones (Editable)"] = c

    return header_map


def normalize_status_value(value):
    t = normalize_text(value)
    if not t:
        return ""

    if "completo" in t:
        return "Completado"
    if "con actividades" in t:
        return "Con Actividades"
    if "sin actividades" in t:
        return "Sin Actividades"
    return clean_text(value)


def extract_table(ws, header_row, block_end_row):
    columns = [
        "Indicador",
        "Meta (editable)",
        "Avance (Editable)",
        "Descripción (editable)",
        "Cantidad (editable)",
        "Observaciones (Editable)"
    ]

    header_map = map_headers(ws, header_row)

    if "Indicador" not in header_map:
        return pd.DataFrame([{
            "Indicador": "",
            "Meta (editable)": "",
            "Avance (Editable)": "",
            "Descripción (editable)": "",
            "Cantidad (editable)": "",
            "Observaciones (Editable)": ""
        }])

    data = []
    empty_count = 0

    for r in range(header_row + 1, block_end_row + 1):
        row_data = {}

        indicador = get_effective_cell_value(ws, r, header_map.get("Indicador")) if header_map.get("Indicador") else ""
        meta = get_effective_cell_value(ws, r, header_map.get("Meta (editable)")) if header_map.get("Meta (editable)") else ""
        avance = get_effective_cell_value(ws, r, header_map.get("Avance (Editable)")) if header_map.get("Avance (Editable)") else ""
        descripcion = get_effective_cell_value(ws, r, header_map.get("Descripción (editable)")) if header_map.get("Descripción (editable)") else ""
        cantidad = get_effective_cell_value(ws, r, header_map.get("Cantidad (editable)")) if header_map.get("Cantidad (editable)") else ""
        observaciones = get_effective_cell_value(ws, r, header_map.get("Observaciones (Editable)")) if header_map.get("Observaciones (Editable)") else ""

        row_data["Indicador"] = "" if indicador is None else indicador
        row_data["Meta (editable)"] = "" if meta is None else meta
        row_data["Avance (Editable)"] = "" if avance is None else normalize_status_value(avance)
        row_data["Descripción (editable)"] = "" if descripcion is None else descripcion
        row_data["Cantidad (editable)"] = "" if cantidad is None else cantidad
        row_data["Observaciones (Editable)"] = "" if observaciones is None else observaciones

        has_content = any(str(v).strip() != "" for v in row_data.values())

        if not has_content:
            empty_count += 1
            if empty_count >= 5:
                break
            continue

        if normalize_text(row_data["Indicador"]) in ["indicador", "indicadores"]:
            continue

        empty_count = 0
        data.append(row_data)

    if not data:
        return pd.DataFrame([{
            "Indicador": "",
            "Meta (editable)": "",
            "Avance (Editable)": "",
            "Descripción (editable)": "",
            "Cantidad (editable)": "",
            "Observaciones (Editable)": ""
        }])

    return pd.DataFrame(data, columns=columns)


# =========================================================
# EXTRACCIÓN COMPLETA
# =========================================================
def extract_blocks_from_sheet(ws):
    starts = find_line_action_starts(ws)
    delegacion = get_delegacion(ws)
    fecha_actualizacion = get_fecha_actualizacion(ws)

    blocks = []

    if not starts:
        return {
            "delegacion": delegacion,
            "fecha_actualizacion": fecha_actualizacion,
            "blocks": []
        }

    for i, start in enumerate(starts):
        start_row = start["row"]
        end_row = starts[i + 1]["row"] - 1 if i + 1 < len(starts) else ws.max_row

        linea_numero = start["line_value"]
        if not linea_numero:
            linea_numero = str(i + 1)

        problematica = search_value_near_keywords(
            ws=ws,
            start_row=start_row,
            end_row=min(start_row + 12, end_row),
            keywords=KEY_PROBLEMATICA,
            value_blacklist=["problematica", "lider", "trimestre", "indicador", "meta"]
        )

        lider = search_value_near_keywords(
            ws=ws,
            start_row=start_row,
            end_row=min(start_row + 12, end_row),
            keywords=KEY_LIDER,
            value_blacklist=["lider", "trimestre", "indicador", "meta", "problematica"]
        )

        trimestre = detect_trimester(
            ws=ws,
            start_row=start_row,
            end_row=min(start_row + 12, end_row)
        )

        header_row = detect_header_row(
            ws=ws,
            start_row=start_row,
            end_row=min(start_row + 35, end_row)
        )

        tabla = extract_table(ws, header_row, end_row) if header_row else pd.DataFrame([{
            "Indicador": "",
            "Meta (editable)": "",
            "Avance (Editable)": "",
            "Descripción (editable)": "",
            "Cantidad (editable)": "",
            "Observaciones (Editable)": ""
        }])

        blocks.append({
            "delegacion": delegacion,
            "fecha_actualizacion": fecha_actualizacion,
            "linea_accion": str(linea_numero).strip(),
            "problematica": clean_text(problematica),
            "lider": clean_text(lider),
            "trimestre": clean_text(trimestre),
            "tabla": tabla,
            "rango_inicio": start_row,
            "rango_fin": end_row,
            "header_row": header_row
        })

    return {
        "delegacion": delegacion,
        "fecha_actualizacion": fecha_actualizacion,
        "blocks": blocks
    }


# =========================================================
# PDF
# =========================================================
def build_pdf_all_lines(data_lineas, delegacion_general, fecha_actualizacion=""):
    buffer = BytesIO()

    doc = SimpleDocTemplate(
        buffer,
        pagesize=letter,
        rightMargin=25,
        leftMargin=25,
        topMargin=28,
        bottomMargin=25
    )

    styles = getSampleStyleSheet()

    title_style = ParagraphStyle(
        "title_custom",
        parent=styles["Title"],
        alignment=TA_CENTER,
        fontSize=15,
        leading=18,
        spaceAfter=10
    )

    normal_style = ParagraphStyle(
        "normal_custom",
        parent=styles["Normal"],
        fontSize=9,
        leading=11,
        spaceAfter=4
    )

    small_style = ParagraphStyle(
        "small_custom",
        parent=styles["Normal"],
        fontSize=7.4,
        leading=9
    )

    elements = []

    elements.append(Paragraph("REPORTE TRIMESTRAL DE LÍNEAS DE ACCIÓN", title_style))
    elements.append(Paragraph(f"<b>Delegación:</b> {safe_str(delegacion_general)}", normal_style))

    if safe_str(fecha_actualizacion):
        elements.append(Paragraph(f"<b>Fecha de actualización:</b> {safe_str(fecha_actualizacion)}", normal_style))

    elements.append(Spacer(1, 8))

    ordered_keys = list(data_lineas.keys())

    for key in ordered_keys:
        item = data_lineas[key]
        info = item["info"]
        df = item["tabla"]
        trimestre = item["trimestre"]
        display_linea = item.get("display_linea", key)

        elements.append(Paragraph(f"<b>Línea de acción #:</b> {safe_str(display_linea)}", normal_style))
        elements.append(Paragraph(f"<b>Problemática:</b> {safe_str(info.get('problematica', ''))}", normal_style))
        elements.append(Paragraph(f"<b>Líder Estratégico:</b> {safe_str(info.get('lider', ''))}", normal_style))
        elements.append(Paragraph(f"<b>Trimestre:</b> {safe_str(trimestre)}", normal_style))
        elements.append(Spacer(1, 6))

        table_data = [[
            "Indicador",
            "Meta",
            "Avance",
            "Descripción",
            "Cantidad",
            "Observaciones"
        ]]

        for _, row in df.iterrows():
            table_data.append([
                Paragraph(safe_str(row.get("Indicador", "")), small_style),
                Paragraph(safe_str(row.get("Meta (editable)", "")), small_style),
                Paragraph(safe_str(row.get("Avance (Editable)", "")), small_style),
                Paragraph(safe_str(row.get("Descripción (editable)", "")), small_style),
                Paragraph(safe_str(row.get("Cantidad (editable)", "")), small_style),
                Paragraph(safe_str(row.get("Observaciones (Editable)", "")), small_style),
            ])

        table = Table(
            table_data,
            repeatRows=1,
            colWidths=[112, 76, 70, 95, 65, 120]
        )

        table.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#D9E2F3")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.black),
            ("GRID", (0, 0), (-1, -1), 0.45, colors.black),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("FONTSIZE", (0, 0), (-1, -1), 7.2),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("LEFTPADDING", (0, 0), (-1, -1), 4),
            ("RIGHTPADDING", (0, 0), (-1, -1), 4),
            ("TOPPADDING", (0, 0), (-1, -1), 4),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F7F9FC")]),
        ]))

        elements.append(table)
        elements.append(Spacer(1, 12))

    doc.build(elements)
    pdf_bytes = buffer.getvalue()
    buffer.close()
    return pdf_bytes


# =========================================================
# RESUMEN
# =========================================================
def build_summary_dataframe(data_lineas):
    rows = []

    for _, item in data_lineas.items():
        df = item["tabla"].copy()
        total_indicadores = len(df)

        counts = {
            "Completado": 0,
            "Con Actividades": 0,
            "Sin Actividades": 0
        }

        if "Avance (Editable)" in df.columns:
            for value in df["Avance (Editable)"].fillna("").astype(str):
                norm = normalize_status_value(value)
                if norm in counts:
                    counts[norm] += 1

        rows.append({
            "Línea": item.get("display_linea", ""),
            "Problemática": item["info"].get("problematica", ""),
            "Líder Estratégico": item["info"].get("lider", ""),
            "Trimestre": item.get("trimestre", ""),
            "Indicadores": total_indicadores,
            "Completado": counts["Completado"],
            "Con Actividades": counts["Con Actividades"],
            "Sin Actividades": counts["Sin Actividades"],
        })

    return pd.DataFrame(rows)


# =========================================================
# INTERFAZ
# =========================================================
st.markdown('<div class="main-title">Seguimiento de líneas de acción</div>', unsafe_allow_html=True)
st.markdown(
    '<div class="sub-note">Sube un archivo Excel (.xlsm o .xlsx). La app localizará de forma robusta la delegación, línea de acción, problemática, líder estratégico, trimestre y el detalle editable de indicadores.</div>',
    unsafe_allow_html=True
)

uploaded_file = st.file_uploader(
    "Arrastra y suelta el archivo Excel",
    type=["xlsm", "xlsx"]
)

if uploaded_file is not None:
    try:
        if st.session_state["archivo_nombre"] != uploaded_file.name:
            st.session_state["archivo_nombre"] = uploaded_file.name
            st.session_state["lineas_guardadas"] = {}
            st.session_state["pdf_final"] = None
            st.session_state["ultimo_resumen"] = pd.DataFrame()

        file_bytes = uploaded_file.read()

        wb = load_workbook(
            BytesIO(file_bytes),
            data_only=False,
            keep_vba=True
        )

        main_sheet = find_best_main_sheet(wb)
        ws = wb[main_sheet]

        extraction = extract_blocks_from_sheet(ws)

        delegacion = extraction["delegacion"]
        fecha_actualizacion = extraction["fecha_actualizacion"]
        blocks = extraction["blocks"]

        if not blocks:
            st.warning("No se encontraron bloques de líneas de acción en la hoja cargada.")
            st.stop()

        # =========================================
        # RESUMEN SUPERIOR
        # =========================================
        col_k1, col_k2, col_k3, col_k4 = st.columns(4)

        with col_k1:
            st.markdown(
                f'<div class="kpi-box"><b>Hoja detectada</b><br>{safe_str(main_sheet)}</div>',
                unsafe_allow_html=True
            )
        with col_k2:
            st.markdown(
                f'<div class="kpi-box"><b>Líneas detectadas</b><br>{len(blocks)}</div>',
                unsafe_allow_html=True
            )
        with col_k3:
            st.markdown(
                f'<div class="kpi-box"><b>Delegación</b><br>{safe_str(delegacion) or "-"}</div>',
                unsafe_allow_html=True
            )
        with col_k4:
            st.markdown(
                f'<div class="kpi-box"><b>Fecha actualización</b><br>{safe_str(fecha_actualizacion) or "-"}</div>',
                unsafe_allow_html=True
            )

        st.markdown("")

        st.markdown('<div class="info-card">', unsafe_allow_html=True)
        g1, g2 = st.columns([1.2, 4])

        with g1:
            st.text_input(
                "Delegación",
                value=delegacion,
                disabled=True,
                key="delegacion_general"
            )

        with g2:
            st.text_input(
                "Fecha de actualización",
                value=fecha_actualizacion,
                disabled=True,
                key="fecha_actualizacion_general"
            )

        st.markdown("</div>", unsafe_allow_html=True)

        # =========================================
        # SIDEBAR
        # =========================================
        with st.sidebar:
            st.header("Resumen")
            st.write(f"**Archivo:** {uploaded_file.name}")
            st.write(f"**Hoja detectada:** {main_sheet}")
            st.write(f"**Delegación:** {delegacion or '-'}")
            st.write(f"**Fecha de actualización:** {fecha_actualizacion or '-'}")
            st.write(f"**Líneas encontradas:** {len(blocks)}")
            st.divider()

            lineas_detectadas = [f"Línea {b['linea_accion']}" for b in blocks]
            st.write("**Líneas detectadas:**")
            for item in lineas_detectadas:
                st.write(f"• {item}")

        # =========================================
        # BLOQUES
        # =========================================
        for idx, bloque in enumerate(blocks):
            linea_id = str(bloque["linea_accion"]).strip() if bloque["linea_accion"] else str(idx + 1)

            block_key = f"bloque_{idx}_{linea_id}"
            save_key = f"{idx}_{linea_id}"

            if save_key in st.session_state["lineas_guardadas"]:
                df_base = st.session_state["lineas_guardadas"][save_key]["tabla"].copy()
                trim_base = st.session_state["lineas_guardadas"][save_key]["trimestre"]
            else:
                df_base = bloque["tabla"].copy()
                trim_base = bloque["trimestre"] if bloque["trimestre"] in ["", "I", "II", "III", "IV"] else ""

            st.markdown('<div class="line-card">', unsafe_allow_html=True)

            st.markdown(f'<div class="section-title">Línea {safe_str(linea_id)}</div>', unsafe_allow_html=True)
            st.markdown('<hr class="soft-line">', unsafe_allow_html=True)

            c1, c2, c3 = st.columns([1.2, 3.8, 2])

            with c1:
                st.text_input(
                    "Línea de acción #",
                    value=linea_id,
                    disabled=True,
                    key=f"linea_{block_key}"
                )

            with c2:
                st.text_input(
                    "Problemática",
                    value=bloque["problematica"],
                    disabled=True,
                    key=f"problematica_{block_key}"
                )

            with c3:
                st.text_input(
                    "Líder Estratégico",
                    value=bloque["lider"],
                    disabled=True,
                    key=f"lider_{block_key}"
                )

            c4, c5 = st.columns([1.2, 5])

            with c4:
                trim_options = ["", "I", "II", "III", "IV"]
                selected_trim = st.selectbox(
                    "Trimestre",
                    trim_options,
                    index=trim_options.index(trim_base if trim_base in trim_options else ""),
                    key=f"trim_{block_key}"
                )

            with c5:
                st.caption(
                    f"Filas detectadas del bloque: {bloque['rango_inicio']} a {bloque['rango_fin']} | "
                    f"Encabezado detectado: {bloque['header_row'] if bloque['header_row'] else '-'}"
                )

            st.markdown("#### Detalle editable")

            df_editado = st.data_editor(
                df_base,
                use_container_width=True,
                hide_index=True,
                num_rows="dynamic",
                key=f"tabla_{block_key}",
                column_config={
                    "Indicador": st.column_config.TextColumn("Indicador", width="medium"),
                    "Meta (editable)": st.column_config.TextColumn("Meta (editable)", width="medium"),
                    "Avance (Editable)": st.column_config.SelectboxColumn(
                        "Avance (Editable)",
                        options=["", "Completado", "Con Actividades", "Sin Actividades"],
                        width="small"
                    ),
                    "Descripción (editable)": st.column_config.TextColumn("Descripción (editable)", width="large"),
                    "Cantidad (editable)": st.column_config.TextColumn("Cantidad (editable)", width="small"),
                    "Observaciones (Editable)": st.column_config.TextColumn("Observaciones (Editable)", width="large"),
                }
            )

            btn1, btn2, btn3 = st.columns([2.3, 2.4, 5])

            with btn1:
                if st.button(f"Guardar / Actualizar línea {linea_id}", key=f"guardar_{block_key}"):
                    st.session_state["lineas_guardadas"][save_key] = {
                        "display_linea": linea_id,
                        "info": {
                            "delegacion": delegacion,
                            "linea_accion": linea_id,
                            "problematica": bloque["problematica"],
                            "lider": bloque["lider"],
                            "rango_inicio": bloque["rango_inicio"],
                            "rango_fin": bloque["rango_fin"],
                        },
                        "tabla": df_editado.copy(),
                        "trimestre": selected_trim
                    }
                    st.session_state["pdf_final"] = None
                    st.success(f"Línea {linea_id} guardada correctamente.")

            with btn2:
                if st.button(f"Restaurar línea {linea_id}", key=f"restaurar_{block_key}"):
                    if save_key in st.session_state["lineas_guardadas"]:
                        del st.session_state["lineas_guardadas"][save_key]
                    st.session_state["pdf_final"] = None
                    st.warning(f"Línea {linea_id} restaurada a los datos detectados del Excel.")
                    st.rerun()

            st.markdown("</div>", unsafe_allow_html=True)

        # =========================================
        # GUARDADO MASIVO
        # =========================================
        st.markdown("## Gestión general")

        gbtn1, gbtn2 = st.columns([2.5, 5])

        with gbtn1:
            if st.button("Guardar todas las líneas detectadas"):
                for idx, bloque in enumerate(blocks):
                    linea_id = str(bloque["linea_accion"]).strip() if bloque["linea_accion"] else str(idx + 1)
                    block_key = f"bloque_{idx}_{linea_id}"
                    save_key = f"{idx}_{linea_id}"

                    trim_value = st.session_state.get(f"trim_{block_key}", bloque["trimestre"])
                    tabla_value = st.session_state.get(f"tabla_{block_key}", bloque["tabla"])

                    if isinstance(tabla_value, pd.DataFrame):
                        tabla_guardar = tabla_value.copy()
                    else:
                        tabla_guardar = bloque["tabla"].copy()

                    st.session_state["lineas_guardadas"][save_key] = {
                        "display_linea": linea_id,
                        "info": {
                            "delegacion": delegacion,
                            "linea_accion": linea_id,
                            "problematica": bloque["problematica"],
                            "lider": bloque["lider"],
                            "rango_inicio": bloque["rango_inicio"],
                            "rango_fin": bloque["rango_fin"],
                        },
                        "tabla": tabla_guardar,
                        "trimestre": trim_value if trim_value in ["", "I", "II", "III", "IV"] else ""
                    }

                st.session_state["pdf_final"] = None
                st.success("Todas las líneas fueron guardadas correctamente.")

        with gbtn2:
            st.caption("Usa esta opción si deseas guardar de una sola vez todo lo detectado y editado en pantalla.")

        # =========================================
        # LÍNEAS GUARDADAS
        # =========================================
        st.markdown("## Líneas guardadas")

        if st.session_state["lineas_guardadas"]:
            cols_saved = st.columns(4)
            ordered_items = list(st.session_state["lineas_guardadas"].items())

            for i, (_, item) in enumerate(ordered_items):
                with cols_saved[i % 4]:
                    st.success(f"Línea {item.get('display_linea', '')}")
        else:
            st.info("Todavía no has guardado ninguna línea.")

        # =========================================
        # RESUMEN TABULAR
        # =========================================
        st.markdown("## Resumen consolidado")

        if st.session_state["lineas_guardadas"]:
            df_summary = build_summary_dataframe(st.session_state["lineas_guardadas"])
            st.session_state["ultimo_resumen"] = df_summary.copy()
            st.dataframe(df_summary, use_container_width=True, hide_index=True)
        else:
            st.info("Cuando guardes líneas, aquí verás el resumen consolidado.")

        # =========================================
        # PDF
        # =========================================
        st.markdown("## Reporte final")

        p1, p2 = st.columns([2.5, 5])

        with p1:
            if st.button("Preparar PDF con todas las líneas guardadas"):
                if not st.session_state["lineas_guardadas"]:
                    st.warning("Primero debes guardar al menos una línea.")
                else:
                    pdf_bytes = build_pdf_all_lines(
                        st.session_state["lineas_guardadas"],
                        delegacion_general=delegacion,
                        fecha_actualizacion=fecha_actualizacion
                    )
                    st.session_state["pdf_final"] = pdf_bytes
                    st.success("PDF generado correctamente.")

        if st.session_state["pdf_final"]:
            st.download_button(
                "Descargar PDF completo",
                data=st.session_state["pdf_final"],
                file_name=f"reporte_trimestral_{delegacion or 'delegacion'}.pdf",
                mime="application/pdf"
            )

        # =========================================
        # TÉCNICO
        # =========================================
        with st.expander("Resumen técnico de detección"):
            debug_rows = []
            for b in blocks:
                debug_rows.append({
                    "Línea": b["linea_accion"],
                    "Problemática": b["problematica"],
                    "Líder": b["lider"],
                    "Trimestre detectado": b["trimestre"],
                    "Filas detalle": len(b["tabla"]),
                    "Fila inicio": b["rango_inicio"],
                    "Fila fin": b["rango_fin"],
                    "Fila encabezado": b["header_row"],
                })

            st.dataframe(pd.DataFrame(debug_rows), use_container_width=True, hide_index=True)

    except Exception as e:
        st.error(f"Error al procesar el archivo: {e}")

else:
    st.info("Sube un archivo para comenzar.")
