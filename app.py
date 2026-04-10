# ============================================================================
# ============================== PARTE 1/10 =================================
# ============ Encabezado, imports, config y estilos Matplotlib =============
# ============================================================================

# App — Pareto 80/20 + Portafolio + Unificado + Sheets + Informe PDF
# Ejecuta: streamlit run app.py

import io
from textwrap import wrap
from typing import List, Dict, Tuple
from datetime import datetime

import numpy as np
import pandas as pd
import streamlit as st
import matplotlib.pyplot as plt

# ====== Google Sheets (DB) ======
import gspread
from google.oauth2.service_account import Credentials

# ====== PDF (ReportLab/Platypus) ======
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm
from reportlab.platypus import (
    BaseDocTemplate, PageTemplate, Frame,
    Paragraph, Spacer, Image as RLImage, Table, TableStyle,
    PageBreak, NextPageTemplate
)
from reportlab.platypus.flowables import KeepTogether

# ----------------- CONFIG (tu Sheets y pestaña) -----------------
SPREADSHEET_URL = "https://docs.google.com/spreadsheets/d/1XZjXQfLb5Jiptp_BXuCfg9QZNEz6ZWh9hbtp0rRAGpM/edit?usp=sharing"
WS_PARETOS = "paretos"  # cambia si tu pestaña se llama diferente

st.set_page_config(page_title="Pareto de Descriptores", layout="wide")

# Paleta
VERDE = "#1B9E77"
AZUL  = "#2C7FB8"
TEXTO = "#124559"
GRIS  = "#6B7280"

# Matplotlib
plt.rcParams.update({
    "figure.dpi": 180,
    "savefig.dpi": 180,
    "axes.titlesize": 18,
    "axes.labelsize": 13,
    "xtick.labelsize": 10,
    "ytick.labelsize": 11,
    "axes.grid": True,
    "grid.alpha": 0.25,
})
# ============================================================================
# ============================== PARTE 2/10 =================================
# ========================= Catálogo embebido (CSV) =========================
# ============================================================================

CATALOGO: List[Dict[str, str]] = [
    {"categoria": "Delito", "descriptor": "Abandono de personas (menor de edad, adulto mayor o con capacidades diferentes)"},
    {"categoria": "Delito", "descriptor": "Abigeato (robo y destace de ganado)"},
    {"categoria": "Delito", "descriptor": "Aborto"},
    {"categoria": "Delito", "descriptor": "Abuso de autoridad"},
    {"categoria": "Riesgo social", "descriptor": "Accidentes de tránsito"},
    {"categoria": "Delito", "descriptor": "Accionamiento de arma de fuego (balaceras)"},
    {"categoria": "Riesgo social", "descriptor": "Acoso escolar (bullying)"},
    {"categoria": "Riesgo social", "descriptor": "Acoso laboral (mobbing)"},
    {"categoria": "Riesgo social", "descriptor": "Acoso sexual callejero"},
    {"categoria": "Riesgo social", "descriptor": "Actos obscenos en vía pública"},
    {"categoria": "Delito", "descriptor": "Administración fraudulenta, apropiaciones indebidas o enriquecimiento ilícito"},
    {"categoria": "Delito", "descriptor": "Agresión con armas"},
    {"categoria": "Riesgo social", "descriptor": "Agrupaciones delincuenciales no organizadas"},
    {"categoria": "Delito", "descriptor": "Alteración de datos y sabotaje informático"},
    {"categoria": "Otros factores", "descriptor": "Ambiente laboral inadecuado"},
    {"categoria": "Delito", "descriptor": "Amenazas"},
    {"categoria": "Riesgo social", "descriptor": "Analfabetismo"},
    {"categoria": "Riesgo social", "descriptor": "Bajos salarios"},
    {"categoria": "Riesgo social", "descriptor": "Barras de fútbol"},
    {"categoria": "Riesgo social", "descriptor": "Búnker (eje de expendio de drogas)"},
    {"categoria": "Delito", "descriptor": "Calumnia"},
    {"categoria": "Delito", "descriptor": "Caza ilegal"},
    {"categoria": "Delito", "descriptor": "Conducción temeraria"},
    {"categoria": "Riesgo social", "descriptor": "Consumo de alcohol en vía pública"},
    {"categoria": "Riesgo social", "descriptor": "Consumo de drogas"},
    {"categoria": "Riesgo social", "descriptor": "Contaminación sónica"},
    {"categoria": "Delito", "descriptor": "Contrabando"},
    {"categoria": "Delito", "descriptor": "Corrupción"},
    {"categoria": "Delito", "descriptor": "Corrupción policial"},
    {"categoria": "Delito", "descriptor": "Cultivo de droga (marihuana)"},
    {"categoria": "Delito", "descriptor": "Daño ambiental"},
    {"categoria": "Delito", "descriptor": "Daños/vandalismo"},
    {"categoria": "Riesgo social", "descriptor": "Deficiencia en la infraestructura vial"},
    {"categoria": "Otros factores", "descriptor": "Deficiencia en la línea 9-1-1"},
    {"categoria": "Riesgo social", "descriptor": "Deficiencias en el alumbrado público"},
    {"categoria": "Delito", "descriptor": "Delincuencia organizada"},
    {"categoria": "Delito", "descriptor": "Delitos contra el ámbito de intimidad (violación de secretos, correspondencia y comunicaciones electrónicas)"},
    {"categoria": "Delito", "descriptor": "Delitos sexuales"},
    {"categoria": "Riesgo social", "descriptor": "Desaparición de personas"},
    {"categoria": "Riesgo social", "descriptor": "Desarticulación interinstitucional"},
    {"categoria": "Riesgo social", "descriptor": "Desempleo"},
    {"categoria": "Riesgo social", "descriptor": "Desvinculación estudiantil"},
    {"categoria": "Delito", "descriptor": "Desobediencia"},
    {"categoria": "Delito", "descriptor": "Desórdenes en vía pública"},
    {"categoria": "Delito", "descriptor": "Disturbios (riñas)"},
    {"categoria": "Riesgo social", "descriptor": "Enfrentamientos estudiantiles"},
    {"categoria": "Delito", "descriptor": "Estafa o defraudación"},
    {"categoria": "Delito", "descriptor": "Estupro (delitos sexuales contra menor de edad)"},
    {"categoria": "Delito", "descriptor": "Evasión y quebrantamiento de pena"},
    {"categoria": "Delito", "descriptor": "Explosivos"},
    {"categoria": "Delito", "descriptor": "Extorsión"},
    {"categoria": "Delito", "descriptor": "Fabricación, producción o reproducción de pornografía"},
    {"categoria": "Riesgo social", "descriptor": "Facilismo económico"},
    {"categoria": "Delito", "descriptor": "Falsificación de moneda y otros valores"},
    {"categoria": "Riesgo social", "descriptor": "Falta de cámaras de seguridad"},
    {"categoria": "Otros factores", "descriptor": "Falta de capacitación policial"},
    {"categoria": "Riesgo social", "descriptor": "Falta de control a patentes"},
    {"categoria": "Riesgo social", "descriptor": "Falta de control fronterizo"},
    {"categoria": "Riesgo social", "descriptor": "Falta de corresponsabilidad en seguridad"},
    {"categoria": "Riesgo social", "descriptor": "Falta de cultura vial"},
    {"categoria": "Riesgo social", "descriptor": "Falta de cultura y compromiso ciudadano"},
    {"categoria": "Riesgo social", "descriptor": "Falta de educación familiar"},
    {"categoria": "Otros factores", "descriptor": "Falta de incentivos"},
    {"categoria": "Riesgo social", "descriptor": "Falta de inversión social"},
    {"categoria": "Riesgo social", "descriptor": "Falta de legislación de extinción de dominio"},
    {"categoria": "Otros factores", "descriptor": "Falta de personal administrativo"},
    {"categoria": "Otros factores", "descriptor": "Falta de personal policial"},
    {"categoria": "Otros factores", "descriptor": "Falta de policías de tránsito"},
    {"categoria": "Riesgo social", "descriptor": "Falta de políticas públicas en seguridad"},
    {"categoria": "Riesgo social", "descriptor": "Falta de presencia policial"},
    {"categoria": "Riesgo social", "descriptor": "Falta de salubridad pública"},
    {"categoria": "Riesgo social", "descriptor": "Familias disfuncionales"},
    {"categoria": "Delito", "descriptor": "Fraude informático"},
    {"categoria": "Delito", "descriptor": "Grooming"},
    {"categoria": "Riesgo social", "descriptor": "Hacinamiento carcelario"},
    {"categoria": "Riesgo social", "descriptor": "Hacinamiento policial"},
    {"categoria": "Delito", "descriptor": "Homicidio"},
    {"categoria": "Riesgo social", "descriptor": "Hospedajes ilegales (cuarterías)"},
    {"categoria": "Delito", "descriptor": "Hurto"},
    {"categoria": "Otros factores", "descriptor": "Inadecuado uso del recurso policial"},
    {"categoria": "Riesgo social", "descriptor": "Incumplimiento al plan regulador de la municipalidad"},
    {"categoria": "Delito", "descriptor": "Incumplimiento del deber alimentario"},
    {"categoria": "Riesgo social", "descriptor": "Indiferencia social"},
    {"categoria": "Otros factores", "descriptor": "Inefectividad en el servicio de policía"},
    {"categoria": "Riesgo social", "descriptor": "Ineficiencia en la administración de justicia"},
    {"categoria": "Otros factores", "descriptor": "Infraestructura inadecuada"},
    {"categoria": "Riesgo social", "descriptor": "Intolerancia social"},
    {"categoria": "Otros factores", "descriptor": "Irrespeto a la jefatura"},
    {"categoria": "Otros factores", "descriptor": "Irrespeto al subalterno"},
    {"categoria": "Otros factores", "descriptor": "Jornadas laborales extensas"},
    {"categoria": "Delito", "descriptor": "Lavado de activos"},
    {"categoria": "Delito", "descriptor": "Lesiones"},
    {"categoria": "Delito", "descriptor": "Ley de armas y explosivos N° 7530"},
    {"categoria": "Riesgo social", "descriptor": "Ley de control de tabaco (Ley 9028)"},
    {"categoria": "Riesgo social", "descriptor": "Lotes baldíos"},
    {"categoria": "Delito", "descriptor": "Maltrato animal"},
    {"categoria": "Delito", "descriptor": "Narcotráfico"},
    {"categoria": "Riesgo social", "descriptor": "Necesidades básicas insatisfechas"},
    {"categoria": "Riesgo social", "descriptor": "Percepción de inseguridad"},
    {"categoria": "Riesgo social", "descriptor": "Pérdida de espacios públicos"},
    {"categoria": "Riesgo social", "descriptor": "Personas con exceso de tiempo de ocio"},
    {"categoria": "Riesgo social", "descriptor": "Personas en estado migratorio irregular"},
    {"categoria": "Riesgo social", "descriptor": "Personas en situación de calle"},
    {"categoria": "Delito", "descriptor": "Menores en vulnerabilidad"},
    {"categoria": "Delito", "descriptor": "Pesca ilegal"},
    {"categoria": "Delito", "descriptor": "Portación ilegal de armas"},
    {"categoria": "Riesgo social", "descriptor": "Presencia multicultural"},
    {"categoria": "Otros factores", "descriptor": "Presión por resultados operativos"},
    {"categoria": "Delito", "descriptor": "Privación de libertad sin ánimo de lucro"},
    {"categoria": "Riesgo social", "descriptor": "Problemas vecinales"},
    {"categoria": "Delito", "descriptor": "Receptación"},
    {"categoria": "Delito", "descriptor": "Relaciones impropias"},
    {"categoria": "Delito", "descriptor": "Resistencia (irrespeto a la autoridad)"},
    {"categoria": "Delito", "descriptor": "Robo a comercio (intimidación)"},
    {"categoria": "Delito", "descriptor": "Robo a comercio (tacha)"},
    {"categoria": "Delito", "descriptor": "Robo a edificación (tacha)"},
    {"categoria": "Delito", "descriptor": "Robo a personas"},
    {"categoria": "Delito", "descriptor": "Robo a transporte comercial"},
    {"categoria": "Delito", "descriptor": "Robo a vehículos (tacha)"},
    {"categoria": "Delito", "descriptor": "Robo a vivienda (intimidación)"},
    {"categoria": "Delito", "descriptor": "Robo a vivienda (tacha)"},
    {"categoria": "Delito", "descriptor": "Robo de bicicleta"},
    {"categoria": "Delito", "descriptor": "Robo de cultivos"},
    {"categoria": "Delito", "descriptor": "Robo de motocicletas/vehículos (bajonazo)"},
    {"categoria": "Delito", "descriptor": "Robo de vehículos"},
    {"categoria": "Delito", "descriptor": "Secuestro"},
    {"categoria": "Delito", "descriptor": "Simulación de delito"},
    {"categoria": "Riesgo social", "descriptor": "Sistema jurídico desactualizado"},
    {"categoria": "Riesgo social", "descriptor": "Suicidio"},
    {"categoria": "Delito", "descriptor": "Sustracción de una persona menor de edad o incapaz"},
    {"categoria": "Delito", "descriptor": "Tala ilegal"},
    {"categoria": "Riesgo social", "descriptor": "Tendencia social hacia el delito (pautas de crianza violenta)"},
    {"categoria": "Riesgo social", "descriptor": "Tenencia de droga"},
    {"categoria": "Delito", "descriptor": "Tentativa de homicidio"},
    {"categoria": "Delito", "descriptor": "Terrorismo"},
    {"categoria": "Riesgo social", "descriptor": "Trabajo informal"},
    {"categoria": "Delito", "descriptor": "Tráfico de armas"},
    {"categoria": "Delito", "descriptor": "Tráfico de influencias"},
    {"categoria": "Riesgo social", "descriptor": "Transporte informal (Uber, porteadores, piratas)"},
    {"categoria": "Delito", "descriptor": "Trata de personas"},
    {"categoria": "Delito", "descriptor": "Turbación de actos religiosos y profanaciones"},
    {"categoria": "Delito", "descriptor": "Uso ilegal de uniformes, insignias o dispositivos policiales"},
    {"categoria": "Delito", "descriptor": "Usurpación de terrenos (precarios)"},
    {"categoria": "Delito", "descriptor": "Venta de drogas"},
    {"categoria": "Riesgo social", "descriptor": "Ventas informales (ambulantes)"},
    {"categoria": "Riesgo social", "descriptor": "Vigilancia informal"},
    {"categoria": "Delito", "descriptor": "Violación de domicilio"},
    {"categoria": "Delito", "descriptor": "Violación de la custodia de las cosas"},
    {"categoria": "Delito", "descriptor": "Violación de sellos"},
    {"categoria": "Delito", "descriptor": "Violencia de género"},
    {"categoria": "Delito", "descriptor": "Violencia intrafamiliar"},
    {"categoria": "Riesgo social", "descriptor": "Xenofobia"},
    {"categoria": "Riesgo social", "descriptor": "Zonas de prostitución"},
    {"categoria": "Riesgo social", "descriptor": "Zonas vulnerables"},
    {"categoria": "Delito", "descriptor": "Robo a transporte público con intimidación"},
    {"categoria": "Delito", "descriptor": "Robo de cable"},
    {"categoria": "Delito", "descriptor": "Explotación sexual infantil"},
    {"categoria": "Delito", "descriptor": "Explotación laboral infantil"},
    {"categoria": "Delito", "descriptor": "Tráfico ilegal de personas"},
    {"categoria": "Riesgo social", "descriptor": "Bares clandestinos"},
    {"categoria": "Delito", "descriptor": "Robo de combustible"},
    {"categoria": "Delito", "descriptor": "Femicidio"},
    {"categoria": "Delito", "descriptor": "Delitos contra la vida (homicidios, heridos)"},
    {"categoria": "Delito", "descriptor": "Venta y consumo de drogas en vía pública"},
    {"categoria": "Delito", "descriptor": "Asalto (a personas, comercio, vivienda, transporte público)"},
    {"categoria": "Delito", "descriptor": "Robo de ganado y agrícola"},
    {"categoria": "Delito", "descriptor": "Robo de equipo agrícola"},
]
# ============================================================================
# ============================== PARTE 3/10 =================================
# ====================== Utilidades base y cálculo Pareto ====================
# ============================================================================

def _map_descriptor_a_categoria() -> Dict[str, str]:
    df = pd.DataFrame(CATALOGO)
    return dict(zip(df["descriptor"], df["categoria"]))
DESC2CAT = _map_descriptor_a_categoria()


def normalizar_freq_map(freq_map: Dict[str, int]) -> Dict[str, int]:
    out = {}
    for d, v in (freq_map or {}).items():
        try:
            vv = int(pd.to_numeric(v, errors="coerce"))
            if vv > 0:
                out[d] = vv
        except Exception:
            continue
    return out


def df_desde_freq_map(freq_map: Dict[str, int]) -> pd.DataFrame:
    items = []
    for d, f in normalizar_freq_map(freq_map).items():
        items.append({
            "descriptor": d,
            "categoria": DESC2CAT.get(d, "—"),
            "frecuencia": int(f)
        })
    df = pd.DataFrame(items)
    if df.empty:
        return pd.DataFrame(columns=["descriptor", "categoria", "frecuencia"])
    return df


def combinar_maps(maps: List[Dict[str, int]]) -> Dict[str, int]:
    total = {}
    for m in maps:
        for d, f in normalizar_freq_map(m).items():
            total[d] = total.get(d, 0) + int(f)
    return total


def info_pareto(freq_map: Dict[str, int]) -> Dict[str, int]:
    d = normalizar_freq_map(freq_map)
    return {"descriptores": len(d), "total": int(sum(d.values()))}


# --- Cálculo Pareto ---
def calcular_pareto(df_in: pd.DataFrame) -> pd.DataFrame:
    df = df_in.copy()
    df["frecuencia"] = pd.to_numeric(df["frecuencia"], errors="coerce").fillna(0).astype(int)
    df = df[df["frecuencia"] > 0]
    if df.empty:
        return df.assign(porcentaje=0.0, acumulado=0, pct_acum=0.0,
                         segmento_real="20%", segmento="80%")
    df = df.sort_values("frecuencia", ascending=False)
    total = int(df["frecuencia"].sum())
    df["porcentaje"] = (df["frecuencia"] / total * 100).round(2)
    df["acumulado"]  = df["frecuencia"].cumsum()
    df["pct_acum"]   = (df["acumulado"] / total * 100).round(2)
    df["segmento_real"] = np.where(df["pct_acum"] <= 80.00, "80%", "20%")
    df["segmento"] = "80%"
    return df.reset_index(drop=True)


def _colors_for_segments(segments: List[str]) -> List[str]:
    return [VERDE if s == "80%" else AZUL for s in segments]


def _wrap_labels(labels: List[str], width: int = 22) -> List[str]:
    """
    Envuelve etiquetas largas para evitar choque visual.
    Además, si hay más de 20 etiquetas, ajusta el ancho automáticamente.
    """
    if len(labels) > 30:
        width = 15
    elif len(labels) > 20:
        width = 18
    elif len(labels) > 12:
        width = 20
    return ["\n".join(wrap(str(t), width=width)) for t in labels]
# ============================================================================
# ============================== PARTE 4/10 =================================
# ====== Gráfico Pareto (UI) + Exportación Excel con gráfico combinado ======
# ============================================================================

def _wrap_for_two_lines(labels: List[str]) -> List[str]:
    """Envuelve etiquetas en máximo 2 líneas (ajuste dinámico)."""
    if not labels:
        return labels
    max_len = max(len(str(x)) for x in labels)
    # Establecemos el ancho según longitud promedio
    if max_len > 60:
        w = 25
    elif max_len > 40:
        w = 20
    elif max_len > 25:
        w = 16
    else:
        w = 14
    wrapped = []
    for label in labels:
        parts = _wrap_labels([label], width=w)
        if len(parts) > 2:
            parts = parts[:2]
        wrapped.append("\n".join(parts))
    return wrapped

def dibujar_pareto(df_par: pd.DataFrame, titulo: str):
    if df_par.empty:
        st.info("Ingresa frecuencias (>0) para ver el gráfico.")
        return

    n_labels = len(df_par)
    x        = np.arange(n_labels)
    freqs    = df_par["frecuencia"].to_numpy()
    pct_acum = df_par["pct_acum"].to_numpy()
    colors_b = _colors_for_segments(df_par["segmento_real"].tolist())
    labels   = [str(t) for t in df_par["descriptor"].tolist()]
    labels_w = _wrap_for_two_lines(labels)

    fig_w = max(12.0, 0.60 * n_labels)
    fs    = 9 if n_labels > 28 else 10

    fig, ax1 = plt.subplots(figsize=(fig_w, 6.6))
    ax1.bar(x, freqs, color=colors_b)
    ax1.set_ylabel("Frecuencia")
    ax1.set_xticks(x)
    ax1.set_xticklabels(labels_w, rotation=90, ha="center", va="top", fontsize=fs)
    fig.subplots_adjust(bottom=0.30)

    ax1.set_title(titulo if titulo.strip() else "Diagrama de Pareto", color=TEXTO, fontsize=16)

    ax2 = ax1.twinx()
    ax2.plot(x, pct_acum, marker="o", linewidth=2, color=TEXTO)
    ax2.set_ylabel("% acumulado")
    ax2.set_ylim(0, 110)

    if (df_par["segmento_real"] == "80%").any():
        cut_idx = np.where(df_par["segmento_real"].to_numpy() == "80%")[0].max()
        ax1.axvline(cut_idx + 0.5, linestyle=":", color="k")
    ax2.axhline(80, linestyle="--", linewidth=1, color="#666666")

    fig.tight_layout()
    st.pyplot(fig)



def exportar_excel_con_grafico(df_par: pd.DataFrame, titulo: str) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        hoja = "Pareto"
        df_x = df_par.copy()
        df_x["porcentaje"] = (df_x["porcentaje"] / 100.0).round(4)
        df_x["pct_acum"]   = (df_x["pct_acum"] / 100.0).round(4)
        df_x = df_x[["categoria", "descriptor", "frecuencia",
                     "porcentaje", "pct_acum", "acumulado", "segmento"]]
        df_x.to_excel(writer, sheet_name=hoja, index=False, startrow=0, startcol=0)

        wb = writer.book
        ws = writer.sheets[hoja]
        pct_fmt = wb.add_format({"num_format": "0.00%"})
        total_fmt = wb.add_format({"bold": True})

        ws.set_column("A:A", 18)
        ws.set_column("B:B", 55)
        ws.set_column("C:C", 12)
        ws.set_column("D:D", 12, pct_fmt)
        ws.set_column("E:E", 18, pct_fmt)
        ws.set_column("F:F", 12)
        ws.set_column("G:G", 10)

        n = len(df_x)
        cats = f"=Pareto!$B$2:$B${n+1}"
        vals = f"=Pareto!$C$2:$C${n+1}"
        pcts = f"=Pareto!$E$2:$E${n+1}"
        total = int(df_par["frecuencia"].sum())

        ws.write(n + 2, 1, "TOTAL:", total_fmt)
        ws.write(n + 2, 2, total, total_fmt)

        chart = wb.add_chart({"type": "column"})
        points = [{"fill": {"color": (VERDE if s == "80%" else AZUL)}} for s in df_par["segmento_real"]]
        chart.add_series({
            "name": "Frecuencia",
            "categories": cats,
            "values": vals,
            "points": points
        })

        line = wb.add_chart({"type": "line"})
        line.add_series({
            "name": "% acumulado",
            "categories": cats,
            "values": pcts,
            "y2_axis": True,
            "marker": {"type": "circle"}
        })

        chart.combine(line)
        chart.set_y_axis({"name": "Frecuencia"})
        chart.set_y2_axis({
            "name": "Porcentaje acumulado",
            "min": 0, "max": 1.10,
            "major_unit": 0.10,
            "num_format": "0%"
        })

        title_text = titulo.strip() if titulo else ""
        chart.set_title({"name": title_text or "Diagrama de Pareto"})
        chart.set_legend({"position": "bottom"})
        chart.set_size({"width": 1180, "height": 420})
        ws.insert_chart("I2", chart)
    return output.getvalue()


# ============================================================================
# ============================== PARTE 5/10 =================================
# ======================== Conectores Google Sheets (gspread) ================
# ============================================================================

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]

def _gc():
    creds = Credentials.from_service_account_info(
        st.secrets["gcp_service_account"],
        scopes=SCOPES
    )
    return gspread.authorize(creds)


def _open_sheet():
    gc = _gc()
    return gc.open_by_url(SPREADSHEET_URL)


def _ensure_ws(sh, title: str, header: List[str]):
    try:
        ws = sh.worksheet(title)
    except gspread.exceptions.WorksheetNotFound:
        ws = sh.add_worksheet(title=title, rows=1000, cols=max(10, len(header)))
        ws.append_row(header)
        return ws

    values = ws.get_all_values()
    if not values:
        ws.append_row(header)
    else:
        first = values[0]
        if [c.strip().lower() for c in first] != [c.strip().lower() for c in header]:
            ws.clear()
            ws.append_row(header)
    return ws


def sheets_cargar_portafolio() -> Dict[str, Dict[str, int]]:
    """Lee 'paretos' (nombre, descriptor, frecuencia). Ignora vacíos y <=0."""
    try:
        sh = _open_sheet()
        ws = _ensure_ws(sh, WS_PARETOS, ["nombre", "descriptor", "frecuencia"])
        rows = ws.get_all_records()
        port: Dict[str, Dict[str, int]] = {}
        for r in rows:
            nom = str(r.get("nombre", "")).strip()
            desc = str(r.get("descriptor", "")).strip()
            freq = int(pd.to_numeric(r.get("frecuencia", 0), errors="coerce") or 0)
            if not nom or not desc or freq <= 0:
                continue
            bucket = port.setdefault(nom, {})
            bucket[desc] = bucket.get(desc, 0) + freq
        return port
    except Exception:
        return {}


def sheets_guardar_pareto(nombre: str, freq_map: Dict[str, int], sobrescribir: bool = True):
    """Guarda filas válidas. Si 'sobrescribir', elimina solo las filas del mismo nombre."""
    sh = _open_sheet()
    ws = _ensure_ws(sh, WS_PARETOS, ["nombre", "descriptor", "frecuencia"])
    if sobrescribir:
        vals = ws.get_all_values()
        header = vals[0] if vals else ["nombre", "descriptor", "frecuencia"]
        others = [r for r in vals[1:] if (len(r) > 0 and r[0].strip().lower() != nombre.strip().lower())]
        ws.clear()
        ws.update("A1", [header])
        if others:
            ws.append_rows(others, value_input_option="RAW")
    rows_new = [[nombre, d, int(f)] for d, f in normalizar_freq_map(freq_map).items()]
    if rows_new:
        ws.append_rows(rows_new, value_input_option="RAW")


def sheets_eliminar_pareto(nombre: str) -> bool:
    """
    Elimina todas las filas en Sheets cuyo campo 'nombre' coincida con el nombre indicado.
    Retorna True si se eliminaron filas.
    """
    try:
        sh = _open_sheet()
        ws = _ensure_ws(sh, WS_PARETOS, ["nombre", "descriptor", "frecuencia"])
        vals = ws.get_all_values()
        if not vals or len(vals) <= 1:
            return False
        header = vals[0]
        others = [r for r in vals[1:] if (len(r) > 0 and r[0].strip().lower() != nombre.strip().lower())]
        ws.clear()
        ws.update("A1", [header])
        if others:
            ws.append_rows(others, value_input_option="RAW")
        return True
    except Exception as e:
        st.warning(f"No se pudo eliminar '{nombre}' de Google Sheets: {e}")
        return False
# ============================================================================
# ============================== PARTE 6/10 =================================
# =================== Estado de sesión + Estilos básicos PDF =================
# ============================================================================

# ---- Estado de sesión ----
st.session_state.setdefault("freq_map", {})
st.session_state.setdefault("portafolio", {})
st.session_state.setdefault("msel", [])
st.session_state.setdefault("editor_df", pd.DataFrame(columns=["descriptor", "frecuencia"]))
st.session_state.setdefault("last_msel", [])
st.session_state.setdefault("reset_after_save", False)

# NUEVO: si cambió la URL del Sheet, vaciar portafolio para no arrastrar datos viejos
st.session_state.setdefault("sheet_url_loaded", None)
if st.session_state["sheet_url_loaded"] != SPREADSHEET_URL:
    st.session_state["portafolio"] = {}
    st.session_state["sheet_url_loaded"] = SPREADSHEET_URL

# Cargar portafolio desde Sheets solo si está vacío
if not st.session_state["portafolio"]:
    loaded = sheets_cargar_portafolio()
    if loaded:
        st.session_state["portafolio"].update(loaded)

# Reset después de guardar
if st.session_state.get("reset_after_save", False):
    st.session_state["freq_map"] = {}
    st.session_state["msel"] = []
    st.session_state["editor_df"] = pd.DataFrame(columns=["descriptor", "frecuencia"])
    st.session_state["last_msel"] = []
    st.session_state.pop("editor_freq", None)
    st.session_state["reset_after_save"] = False

# ---- Estilos PDF / páginas ----
PAGE_W, PAGE_H = A4


def _styles():
    ss = getSampleStyleSheet()
    ss.add(ParagraphStyle(
        name="CoverTitle", fontName="Helvetica-Bold",
        fontSize=30, leading=36, textColor=TEXTO, alignment=1, spaceAfter=10
    ))
    ss.add(ParagraphStyle(
        name="CoverSubtitle", parent=ss["Normal"], fontSize=12,
        leading=16, textColor=GRIS, alignment=1, spaceAfter=10
    ))
    ss.add(ParagraphStyle(
        name="CoverDate", parent=ss["Normal"], fontSize=15,
        leading=18, textColor=TEXTO, alignment=1, spaceBefore=8
    ))
    ss.add(ParagraphStyle(
        name="TitleBig", parent=ss["Title"], fontSize=24,
        leading=28, textColor=TEXTO, alignment=0, spaceAfter=10
    ))
    ss.add(ParagraphStyle(
        name="TitleBigCenter", parent=ss["Title"], fontSize=24,
        leading=28, textColor=TEXTO, alignment=1, spaceAfter=10
    ))
    ss.add(ParagraphStyle(name="H1", parent=ss["Heading1"],
                          fontSize=18, leading=22, textColor=TEXTO, spaceAfter=8))
    ss.add(ParagraphStyle(name="H1Center", parent=ss["Heading1"],
                          fontSize=18, leading=22, textColor=TEXTO, spaceAfter=8, alignment=1))
    ss.add(ParagraphStyle(name="Body", parent=ss["Normal"],
                          fontSize=11, leading=14, textColor="#111"))
    ss.add(ParagraphStyle(name="Small", parent=ss["Normal"],
                          fontSize=9.6, leading=12, textColor=GRIS))
    ss.add(ParagraphStyle(name="TableHead", parent=ss["Normal"],
                          fontSize=11, leading=13, textColor=colors.white))

    # ⚠️ No usar "Bullet" (ya existe). Creamos uno propio:
    ss.add(ParagraphStyle(
        name="BulletList", parent=ss["Body"],
        leftIndent=12, bulletIndent=0, spaceBefore=2, spaceAfter=2
    ))
    return ss


def _page_cover(canv, doc):
    canv.setFillColor(colors.HexColor(TEXTO))
    canv.rect(0, PAGE_H - 0.9 * cm, PAGE_W, 0.9 * cm, fill=1, stroke=0)


def _page_normal(_canv, _doc):
    pass


def _page_last(canv, _doc):
    canv.setFillColor(colors.HexColor(TEXTO))
    canv.rect(0, 0, PAGE_W, 0.9 * cm, fill=1, stroke=0)
# ============================================================================
# ============================== PARTE 7/10 =================================
# ============ Imágenes para PDF (Pareto/Modalidades) y textos helper =======
# ============================================================================

def _wrap_for_two_lines(labels: List[str]) -> List[str]:
    """Envuelve etiquetas en máximo 2 líneas."""
    if not labels:
        return labels
    max_len = max(len(str(x)) for x in labels)
    if max_len > 60:
        w = 25
    elif max_len > 40:
        w = 20
    elif max_len > 25:
        w = 16
    else:
        w = 14
    result = []
    for label in labels:
        parts = _wrap_labels([label], width=w)
        if len(parts) > 2:
            parts = parts[:2]
        result.append("\n".join(parts))
    return result

def _pareto_png(df_par: pd.DataFrame, titulo: str) -> bytes:
    """
    PNG del Pareto para PDF:
    - Etiquetas 90° a 2 líneas
    - bbox_inches='tight' para eliminar aire
    """
    n_labels = len(df_par)
    labels   = [str(t) for t in df_par["descriptor"].tolist()]
    labels_w = _wrap_for_two_lines(labels)

    x        = np.arange(n_labels)
    freqs    = df_par["frecuencia"].to_numpy()
    pct_acum = df_par["pct_acum"].to_numpy()
    colors_b = _colors_for_segments(df_par["segmento_real"].tolist())

    fig_w = max(12.0, 0.60 * n_labels)
    fig_h = 6.6
    fs    = 9 if n_labels > 28 else 10
    dpi   = 220

    fig, ax1 = plt.subplots(figsize=(fig_w, fig_h), dpi=dpi)
    ax1.bar(x, freqs, color=colors_b, zorder=2)
    ax1.set_ylabel("Frecuencia")
    ax1.set_xticks(x)
    ax1.set_xticklabels(labels_w, rotation=90, ha="center", va="top", fontsize=fs)
    ax1.set_title(titulo if titulo.strip() else "Diagrama de Pareto", color=TEXTO, fontsize=16)

    ax2 = ax1.twinx()
    ax2.plot(x, pct_acum, marker="o", linewidth=2, color=TEXTO, zorder=3)
    ax2.set_ylabel("% acumulado"); ax2.set_ylim(0, 110)

    if (df_par["segmento_real"] == "80%").any():
        cut_idx = np.where(df_par["segmento_real"].to_numpy() == "80%")[0].max()
        ax1.axvline(cut_idx + 0.5, linestyle=":", color="k")
    ax2.axhline(80, linestyle="--", linewidth=1, color="#666666")
    ax1.grid(True, axis="y", alpha=0.25, zorder=1)

    buf = io.BytesIO()
    fig.savefig(buf, format="PNG", dpi=dpi, bbox_inches="tight", pad_inches=0.08)
    plt.close(fig)
    return buf.getvalue()



# --- Etiquetado temático (se mantiene interno) ---
def _tema_descriptor(descriptor: str) -> str:
    d = descriptor.lower()
    if "droga" in d or "búnker" in d or "bunker" in d or "narco" in d or "venta de drogas" in d:
        return "drogas"
    if "robo" in d or "hurto" in d or "asalto" in d or "vehícul" in d or "comercio" in d:
        return "delitos contra la propiedad"
    if "violencia" in d or "lesion" in d or "homicidio" in d:
        return "violencia"
    if "infraestructura" in d or "alumbrado" in d or "lotes" in d:
        return "condiciones urbanas / entorno"
    return "seguridad y convivencia"


def _resumen_texto(df_par: pd.DataFrame) -> str:
    if df_par.empty:
        return "Sin datos disponibles."
    total = int(df_par["frecuencia"].sum())
    n = len(df_par)
    top = df_par.iloc[0]
    idx80 = int(np.where(df_par["segmento_real"].to_numpy() == "80%")[0].max() + 1) if (df_par["segmento_real"]=="80%").any() else 0
    return (f"Se registran <b>{total}</b> hechos distribuidos en <b>{n}</b> descriptores. "
            f"El descriptor de mayor incidencia es <b>{top['descriptor']}</b>, con <b>{int(top['frecuencia'])}</b> casos "
            f"({float(top['porcentaje']):.2f}%). El punto de corte del <b>80%</b> se alcanza con "
            f"<b>{idx80}</b> descriptores, útiles para la priorización operativa.")


def _texto_modalidades(descriptor: str, pares: List[Tuple[str, float]]) -> str:
    pares_filtrados = [(l, p) for l, p in pares if str(l).strip() and (p or 0) > 0]
    pares_orden = sorted(pares_filtrados, key=lambda x: x[1], reverse=True)
    if not pares_orden:
        return (f"Para <b>{descriptor}</b> no se reportaron modalidades con porcentaje. "
                "Se sugiere recolectar esta información para focalizar acciones.")
    top_txt = "; ".join([f"<b>{l}</b> ({p:.1f}%)" for l, p in pares_orden[:2]])
    return (f"En <b>{descriptor}</b> destacan: {top_txt}. "
            "Esto orienta intervenciones específicas sobre las variantes de mayor peso.")


def _modalidades_png(title: str, data_pairs: List[Tuple[str, float]], kind: str = "barh") -> bytes:
    labels = [l for l, p in data_pairs if str(l).strip()]
    vals   = [float(p or 0) for l, p in data_pairs if str(l).strip()]
    if not labels:
        labels, vals = ["Sin datos"], [100.0]

    order = np.argsort(vals)[::-1]
    labels = [labels[i] for i in order]
    vals   = [vals[i]   for i in order]
    n = len(labels)

    import matplotlib as mpl
    from matplotlib.patches import FancyBboxPatch, Circle

    cmap = mpl.cm.get_cmap("Blues")
    colors_seq = [cmap(0.35 + 0.5*(i/max(1, n-1))) for i in range(n)]

    if kind == "donut":
        fig, ax = plt.subplots(figsize=(7.8, 5.4))
        wedges, _, _ = ax.pie(
            vals, labels=None, autopct=lambda p: f"{p:.1f}%",
            startangle=90, pctdistance=0.8,
            wedgeprops=dict(width=0.4, edgecolor="white"),
            colors=colors_seq
        )
        ax.legend(wedges, [f"{l} ({v:.1f}%)" for l, v in zip(labels, vals)],
                  title="Modalidades", loc="center left",
                  bbox_to_anchor=(1.02, 0.5), fontsize=9)
        ax.set_title(title, color=TEXTO)

    elif kind == "lollipop":
        fig, ax = plt.subplots(figsize=(11.5, 5.4))
        y = np.arange(n)
        ax.hlines(y=y, xmin=0, xmax=vals, color="#94a3b8", linewidth=2)
        ax.plot(vals, y, "o", markersize=8, color=AZUL)
        ax.set_yticks(y)
        ax.set_yticklabels(_wrap_labels(labels, 35))
        ax.invert_yaxis()
        ax.set_xlabel("Porcentaje"); ax.set_xlim(0, max(100, max(vals)*1.05))
        for i, v in enumerate(vals):
            ax.text(v + 1, i, f"{v:.1f}%", va="center", fontsize=10)
        ax.set_title(title, color=TEXTO)

    elif kind == "bar":
        fig, ax = plt.subplots(figsize=(11.5, 5.4))
        x = np.arange(n)
        ax.bar(x, vals, color=colors_seq)
        ax.set_xticks(x)
        ax.set_xticklabels(_wrap_labels(labels, 20), rotation=0)
        ax.set_ylabel("Porcentaje"); ax.set_ylim(0, max(100, max(vals)*1.15))
        for i, v in enumerate(vals):
            ax.text(i, v + max(vals)*0.03, f"{v:.1f}%", ha="center", fontsize=10)
        ax.set_title(title, color=TEXTO)

    elif kind == "comp100":
        fig, ax = plt.subplots(figsize=(11.5, 3.0))
        left = 0.0
        for i, (lab, v) in enumerate(zip(labels, vals)):
            w = max(0.0, float(v))
            ax.barh(0, w, left=left, color=colors_seq[i])
            if w >= 7:
                ax.text(left + w/2, 0, f"{lab}\n{v:.1f}%", va="center", ha="center", fontsize=9, color="white")
            left += w
        ax.set_xlim(0, max(100, sum(vals)))
        ax.set_yticks([]); ax.set_xlabel("Porcentaje (composición)")
        ax.set_title(title, color=TEXTO)
        ax.grid(False)

    elif kind == "pill":
        fig_height = 0.9 + n*0.85
        fig, ax = plt.subplots(figsize=(10.8, fig_height))
        ax.set_xlim(0, 100); ax.set_ylim(0, n)
        ax.axis("off")
        track_h = 0.72
        round_r = track_h/2

        for i, (lab, v) in enumerate(zip(labels, vals)):
            y = n - 1 - i + (1 - track_h)/2
            track = FancyBboxPatch(
                (0.8, y), 98.4, track_h,
                boxstyle=f"round,pad=0,rounding_size={round_r}",
                linewidth=1, edgecolor="#9dbbd6", facecolor="#e6f0fb"
            )
            ax.add_patch(track)

            prog_w = max(0.001, min(98.4, float(v)))
            prog = FancyBboxPatch(
                (0.8, y), prog_w, track_h,
                boxstyle=f"round,pad=0,rounding_size={round_r}",
                linewidth=0, facecolor=AZUL, alpha=0.35
            )
            ax.add_patch(prog)

            ax.add_patch(Circle((0.8 + round_r*0.6, y + track_h/2), round_r*0.9, color=AZUL, alpha=0.9))
            badge_w = 12.0; badge_h = track_h*0.8
            badge_x = 5.0;  badge_y = y + (track_h - badge_h)/2
            badge = FancyBboxPatch(
                (badge_x, badge_y), badge_w, badge_h,
                boxstyle=f"round,pad=0.25,rounding_size={badge_h/2}",
                linewidth=1, edgecolor="#cfd8e3", facecolor="white"
            )
            ax.add_patch(badge)
            ax.text(badge_x + badge_w/2, y + track_h/2, f"{v:.1f}%", ha="center", va="center", fontsize=10)
            ax.text(badge_x + badge_w + 3.0, y + track_h/2, lab, va="center", ha="left",
                    fontsize=12, color="#0f172a")

        ax.set_title(title, color=TEXTO)

    else:  # 'barh'
        fig, ax = plt.subplots(figsize=(11.5, 5.4))
        y = np.arange(n)
        ax.barh(y, vals, color=colors_seq)
        ax.set_yticks(y)
        ax.set_yticklabels(_wrap_labels(labels, 35))
        ax.invert_yaxis()
        ax.set_xlabel("Porcentaje")
        ax.set_xlim(0, max(100, max(vals)*1.05))
        for i, v in enumerate(vals):
            ax.text(v + 1, i, f"{v:.1f}%", va="center", fontsize=10)
        ax.set_title(title, color=TEXTO)

    buf = io.BytesIO()
    fig.savefig(buf, format="PNG", dpi=dpi, bbox_inches="tight", pad_inches=0.08)
    plt.close(fig)
    return buf.getvalue()



# ============================================================================
# ============================== PARTE 8/10 =================================
# ========= Tabla PDF, generador de Informe PDF y UI de desgloses ===========
# ============================================================================

def _tabla_resultados_flowable(df_par: pd.DataFrame, doc_width: float) -> Table:
    """
    Cuadro simplificado: Descriptor | Frecuencia | %
    Incluye una fila final con 'Total de respuestas tratadas'.
    """
    fracs = [0.62, 0.20, 0.18]  # Descriptor, Frecuencia, %
    col_widths = [f * doc_width for f in fracs]
    stys = _styles()

    from reportlab.lib.styles import ParagraphStyle
    cell_style = ParagraphStyle(
        name="CellWrap",
        parent=stys["Normal"],
        fontSize=9.6,
        leading=12,
        textColor="#111111",
        wordWrap="CJK",
        spaceBefore=0,
        spaceAfter=0,
    )

    head = [
        Paragraph("Descriptor", stys["TableHead"]),
        Paragraph("Frecuencia", stys["TableHead"]),
        Paragraph("Porcentaje", stys["TableHead"]),
    ]
    data = [head]

    total_respuestas = int(df_par["frecuencia"].sum()) if not df_par.empty else 0
    for _, r in df_par.iterrows():
        descriptor = Paragraph(str(r["descriptor"]), cell_style)
        frecuencia = int(r["frecuencia"])
        pct = f'{float(r["porcentaje"]):.2f}%'
        data.append([descriptor, frecuencia, pct])

    total_row = [Paragraph("<b>Total de respuestas tratadas</b>", cell_style), total_respuestas, ""]
    data.append(total_row)
    total_index = len(data) - 1

    t = Table(data, colWidths=col_widths, repeatRows=1, hAlign="LEFT")
    t.setStyle(TableStyle([
        ("BACKGROUND", (0,0), (-1,0), colors.HexColor(TEXTO)),
        ("TEXTCOLOR",  (0,0), (-1,0), colors.white),
        ("FONTNAME",   (0,0), (-1,0), "Helvetica-Bold"),
        ("FONTSIZE",   (0,0), (-1,0), 11),
        ("FONTSIZE",   (0,1), (-1,-1), 9.6),
        ("ALIGN",      (0,1), (0,-1), "LEFT"),
        ("ALIGN",      (1,1), (-1,-2), "RIGHT"),
        ("LEFTPADDING",(0,0), (-1,-1), 6),
        ("RIGHTPADDING",(0,0), (-1,-1), 6),
        ("VALIGN",     (0,0), (-1,-1), "MIDDLE"),
        ("ROWBACKGROUNDS", (0,1), (-1,-2), [colors.whitesmoke, colors.Color(0.97,0.97,0.97)]),
        ("GRID", (0,0), (-1,-1), 0.25, colors.lightgrey),
        ("BACKGROUND", (0,total_index), (-1,total_index), colors.Color(0.93, 0.96, 0.99)),
        ("FONTNAME",   (0,total_index), (-1,total_index), "Helvetica-Bold"),
        ("ALIGN",      (1,total_index), (1,total_index), "RIGHT"),
        ("ALIGN",      (2,total_index), (2,total_index), "RIGHT"),
    ]))
    return t


def _altura_img_según_filas(n_filas: int) -> float:
    """(Sin uso directo ahora, mantenida por compatibilidad)."""
    if n_filas >= 28:
        return 5.8
    if n_filas >= 20:
        return 6.8
    if n_filas >= 14:
        return 7.6
    return 8.6


def generar_pdf_informe(nombre_informe: str,
                        df_par: pd.DataFrame,
                        desgloses: List[Dict]) -> bytes:
    """
    Genera el informe PDF completo: portada, introducción, gráfico, tabla,
    modalidades y conclusiones. Inserta el gráfico de Pareto a ancho completo,
    calculando la altura proporcional al PNG generado (misma apariencia que en la app).
    """
    if df_par.empty:
        st.warning("No hay datos válidos para generar el informe.")
        return b""

    buf = io.BytesIO()
    doc = BaseDocTemplate(
        buf, pagesize=A4,
        leftMargin=2*cm, rightMargin=2*cm, topMargin=2*cm, bottomMargin=2*cm
    )
    frame_std  = Frame(doc.leftMargin, doc.bottomMargin, doc.width, doc.height, id="normal")
    frame_last = Frame(doc.leftMargin, doc.bottomMargin, doc.width, doc.height, id="last")

    doc.addPageTemplates([
        PageTemplate(id="Cover",  frames=[frame_std], onPage=_page_cover),
        PageTemplate(id="Normal", frames=[frame_std], onPage=_page_normal),
        PageTemplate(id="Last",   frames=[frame_last], onPage=_page_last),
    ])
    stys = _styles()
    story: List = []

    # ---------- PORTADA ----------
    story += [NextPageTemplate("Normal")]
    story += [Spacer(1, 2.2*cm)]
    story += [Paragraph(f"Informe de Resultados Diagrama de Pareto - {nombre_informe}", stys["CoverTitle"])]
    story += [Paragraph("Estrategia Sembremos Seguridad", stys["CoverSubtitle"])]
    story += [Paragraph(datetime.now().strftime("Fecha: %d/%m/%Y"), stys["CoverDate"])]
    story += [PageBreak()]

    # ---------- INTRODUCCIÓN ----------
    story += [Paragraph("Introducción", stys["TitleBig"]), Spacer(1, 0.2*cm)]
    story += [Paragraph(
        "Este informe presenta un análisis tipo <b>Pareto (80/20)</b> sobre los descriptores seleccionados. "
        "El objetivo es identificar los elementos que concentran la mayor parte de los hechos reportados para apoyar la "
        "priorización operativa y la toma de decisiones. El documento incluye el gráfico de Pareto, un cuadro "
        "resumido con frecuencia y porcentaje por descriptor, y al final una sección de conclusiones y recomendaciones.",
        stys["Body"]
    ), Spacer(1, 0.35*cm)]

    # ---------- RESULTADOS ----------
    story += [Paragraph("Resultados generales", stys["TitleBig"]), Spacer(1, 0.2*cm)]
    story += [Paragraph(_resumen_texto(df_par), stys["Body"]), Spacer(1, 0.3*cm)]

    # --- Gráfico Pareto a ancho completo con altura proporcional ---
    from PIL import Image as PILImage
    pareto_png = _pareto_png(df_par, "Diagrama de Pareto")
    with io.BytesIO(pareto_png) as _b:
        im = PILImage.open(_b)
        w_px, h_px = im.size

    width_pts  = doc.width
    height_pts = (h_px / w_px) * width_pts

    # 1) Gráfico + descripción siempre juntos
    story.append(KeepTogether([
        RLImage(io.BytesIO(pareto_png), width=width_pts, height=height_pts),
        Spacer(1, 0.30*cm),
        Paragraph(
            "El diagrama muestra la frecuencia por descriptor (barras en verde/azul) y el <b>porcentaje acumulado</b> (línea). "
            "La línea punteada del 80% indica el <b>punto de corte</b> para priorización.",
            stys["Small"]
        ),
    ]))

    # 2) Salto de página si el gráfico es largo
    if len(df_par) >= 12:
        story.append(PageBreak())

    # 3) Tabla siempre en bloque único (no se corta)
    story.append(KeepTogether([
        Spacer(1, 0.25*cm),
        _tabla_resultados_flowable(df_par, doc.width),
    ]))


    # ---------- MODALIDADES ----------
    for sec in desgloses:
        descriptor = sec.get("descriptor", "").strip()
        rows = sec.get("rows", [])
        chart_kind = sec.get("chart", "barh")
        pares = [(r.get("Etiqueta",""), float(r.get("%", 0) or 0)) for r in rows]

        bloque = [
            Spacer(1, 0.4*cm),
            Paragraph(f"Modalidades de la problemática — {descriptor}", stys["TitleBig"]),
            Spacer(1, 0.1*cm),
            Paragraph(_texto_modalidades(descriptor, pares), stys["Small"]),
            Spacer(1, 0.2*cm),
            RLImage(io.BytesIO(_modalidades_png(descriptor or 'Modalidades', pares, kind=chart_kind)),
                    width=doc.width, height=8.5*cm),
        ]
        story.append(KeepTogether(bloque))

    # ---------- CIERRE ----------
    story += [PageBreak(), NextPageTemplate("Last")]
    story += [
        Paragraph("Conclusiones y recomendaciones", stys["TitleBigCenter"]),
        Spacer(1, 0.2*cm),
    ]

    bullets = [
        "Priorizar intervenciones sobre los descriptores que conforman el <b>80% acumulado</b>.",
        "Coordinar acciones interinstitucionales enfocadas en las <b>modalidades</b> con mayor porcentaje.",
        "Fortalecer la participación comunitaria y el control territorial en puntos críticos.",
        "Monitorear indicadores mensualmente para evaluar la efectividad de las acciones.",
    ]
    for b in bullets:
        story += [Paragraph(b, stys["BulletList"], bulletText="•")]

    story += [
        Spacer(1, 0.8*cm),
        Paragraph("Dirección de Programas Policiales Preventivos – MSP", stys["H1Center"]),
        Paragraph("Sembremos Seguridad", stys["H1Center"]),
    ]

    doc.build(story)
    return buf.getvalue()


# === UI formulario de desgloses (para editor y unificado) ===
def ui_desgloses(descriptor_list: List[str], key_prefix: str) -> List[Dict]:
    st.caption("Opcional: agrega secciones de ‘Modalidades’. Cada sección admite hasta 10 filas (Etiqueta + %).")
    max_secs = max(0, len(descriptor_list))
    default_val = 1 if max_secs > 0 else 0
    n_secs = st.number_input("Cantidad de secciones de Modalidades",
                             min_value=0, max_value=max_secs, value=default_val, step=1,
                             key=f"{key_prefix}_nsecs")
    desgloses: List[Dict] = []
    for i in range(n_secs):
        with st.expander(f"Sección Modalidades #{i+1}", expanded=(i == 0)):
            dsel = st.selectbox(f"Descriptor para la sección #{i+1}",
                                options=["(elegir)"] + descriptor_list, index=0, key=f"{key_prefix}_desc_{i}")

            chart_kind = st.selectbox(
                "Tipo de gráfico",
                options=[("Barras horizontales", "barh"),
                         ("Barras verticales", "bar"),
                         ("Lollipop (palo+punto)", "lollipop"),
                         ("Dona / Pie", "donut"),
                         ("Barra 100% (composición)", "comp100"),
                         ("Píldora (progreso redondeado)", "pill")],
                index=0, format_func=lambda x: x[0], key=f"{key_prefix}_chart_{i}"
            )[1]

            rows = [{"Etiqueta":"", "%":0.0} for _ in range(10)]
            df_rows = pd.DataFrame(rows)
            de = st.data_editor(
                df_rows, key=f"{key_prefix}_rows_{i}", use_container_width=True,
                column_config={
                    "Etiqueta": st.column_config.TextColumn("Etiqueta / Modalidad", width="large"),
                    "%": st.column_config.NumberColumn("Porcentaje", min_value=0.0, max_value=100.0, step=0.1)
                },
                num_rows="fixed"
            )
            total_pct = float(pd.to_numeric(de["%"], errors="coerce").fillna(0).sum())
            st.caption(f"Suma actual: {total_pct:.1f}% (recomendado ≈100%)")

            if dsel != "(elegir)":
                desgloses.append({"descriptor": dsel,
                                  "rows": de.to_dict(orient="records"),
                                  "chart": chart_kind})
    return desgloses

# ============================================================================
# ============================== PARTE 9/10 =================================
# ========================== Interfaz principal (UI) =========================
# ============================================================================

st.title("📊 Análisis Pareto 80/20 – Descriptores y Portafolio")

tab_editor, tab_portafolio, tab_unificado = st.tabs([
    "➕ Crear / Editar Pareto individual",
    "📁 Portafolio guardado",
    "📄 Informe PDF (unificado)"
])

# ---------------------------------------------------------------------------
# TAB 1 — Editor individual
# ---------------------------------------------------------------------------
with tab_editor:
    st.subheader("✏️ Editor de Pareto individual")

    nombre_pareto = st.text_input("Nombre del Pareto", "").strip()

    # Selector múltiple (CATÁLOGO EMBEBIDO) — ahora con key fijo
    opts = [c["descriptor"] for c in CATALOGO]
    msel = st.multiselect(
        "Selecciona los descriptores a incluir",
        options=opts,
        key="msel"  # ✅ deja que el widget maneje su estado, sin default manual
    )

    # Si hay selección, preparamos/actualizamos el DF del editor
    if msel:
        # Si la selección de descriptores cambió, reconstruimos el DF
        if msel != st.session_state.get("last_msel", []):
            data = []
            # Usamos los valores actuales de freq_map para no perder frecuencias ya digitadas
            freq_map_actual = st.session_state.get("freq_map", {})
            for d in msel:
                data.append({
                    "descriptor": d,
                    "frecuencia": freq_map_actual.get(d, 0)
                })
            st.session_state["editor_df"] = pd.DataFrame(data)
            st.session_state["last_msel"] = list(msel)

        # Editor de frecuencias (se alimenta desde session_state["editor_df"])
        df_edit = st.data_editor(
            st.session_state["editor_df"],
            num_rows="dynamic",
            use_container_width=True,
            key="editor_freq",
            column_config={
                "descriptor": st.column_config.TextColumn("Descriptor", width="large"),
                "frecuencia": st.column_config.NumberColumn("Frecuencia", min_value=0, step=1)
            }
        )

        # Actualizar el DF y el freq_map en sesión con lo que el usuario acaba de escribir
        st.session_state["editor_df"] = df_edit
        freq_map = dict(zip(df_edit["descriptor"], df_edit["frecuencia"]))
        st.session_state["freq_map"] = freq_map

        # Cálculo y vista previa del Pareto
        df_par = calcular_pareto(df_desde_freq_map(freq_map))
        st.divider()
        st.subheader("📊 Diagrama de Pareto (Vista previa)")
        dibujar_pareto(df_par, nombre_pareto)
        st.download_button(
            "📥 Exportar Excel con gráfico",
            exportar_excel_con_grafico(df_par, nombre_pareto),
            file_name=f"Pareto_{nombre_pareto or 'sin_nombre'}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        st.divider()
        desgloses = ui_desgloses(df_par["descriptor"].tolist(), key_prefix="editor")

        col1, col2 = st.columns(2)
        with col1:
            if st.button("💾 Guardar en Portafolio (y Sheets)", type="primary", use_container_width=True):
                if nombre_pareto:
                    st.session_state["portafolio"][nombre_pareto] = normalizar_freq_map(freq_map)
                    sheets_guardar_pareto(nombre_pareto, freq_map, sobrescribir=True)
                    st.success(f"Pareto '{nombre_pareto}' guardado correctamente.")
                    st.session_state["reset_after_save"] = True
                    st.rerun()  # ✅ mantiene el comportamiento
                else:
                    st.warning("Asigna un nombre al Pareto antes de guardar.")
        with col2:
            if st.button("🧾 Generar Informe PDF individual", use_container_width=True):
                if not nombre_pareto:
                    st.warning("Asigna un nombre para el informe.")
                else:
                    pdf_bytes = generar_pdf_informe(nombre_pareto, df_par, desgloses)
                    if pdf_bytes:
                        st.download_button(
                            label="📥 Descargar PDF",
                            data=pdf_bytes,
                            file_name=f"Informe_{nombre_pareto}.pdf",
                            mime="application/pdf"
                        )

    else:
        # Si no hay descriptores seleccionados, limpiamos estructuras del editor
        st.session_state["freq_map"] = {}
        st.session_state["editor_df"] = pd.DataFrame(columns=["descriptor", "frecuencia"])
        st.session_state["last_msel"] = []
        st.info("Selecciona al menos un descriptor del catálogo para comenzar.")

# ---------------------------------------------------------------------------
# TAB 2 — Portafolio de Paretos
# ---------------------------------------------------------------------------
with tab_portafolio:
    st.subheader("📁 Paretos almacenados en portafolio")

    port = st.session_state["portafolio"]
    if not port:
        st.info("No hay Paretos guardados todavía.")
    else:
        for nombre, mapa in list(port.items()):
            with st.expander(f"{nombre}", expanded=False):
                dfp = calcular_pareto(df_desde_freq_map(mapa))
                dibujar_pareto(dfp, nombre)
                st.caption(f"Total de respuestas tratadas: {int(dfp['frecuencia'].sum())}")

                # Acciones
                colA, colB, colC = st.columns([1,1,2])
                with colA:
                    st.download_button(
                        "📥 Excel con gráfico",
                        exportar_excel_con_grafico(dfp, nombre),
                        file_name=f"Pareto_{nombre}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key=f"dl_{nombre}"
                    )
                with colB:
                    if st.button(f"🗑️ Eliminar '{nombre}'", key=f"del_{nombre}"):
                        del st.session_state["portafolio"][nombre]
                        ok = sheets_eliminar_pareto(nombre)
                        if ok:
                            st.success(f"El Pareto '{nombre}' fue eliminado del sistema y de Google Sheets.")
                        else:
                            st.warning(f"El Pareto '{nombre}' se eliminó localmente, pero no pudo borrarse en Sheets.")
                        st.rerun()  # ✅ reemplazo de experimental_rerun
                with colC:
                    try:
                        pop = st.popover("📄 Informe PDF de este Pareto")
                    except Exception:
                        pop = st.expander("📄 Informe PDF de este Pareto", expanded=False)
                    with pop:
                        nombre_inf_ind = st.text_input("Nombre del informe", value=f"{nombre}", key=f"inf_nom_{nombre}")
                        desgloses_ind = ui_desgloses(dfp["descriptor"].tolist(), key_prefix=f"inf_{nombre}")
                        if st.button("Generar PDF", key=f"btn_inf_{nombre}"):
                            pdf_bytes = generar_pdf_informe(nombre_inf_ind, dfp, desgloses_ind)
                            if pdf_bytes:
                                st.download_button(
                                    "⬇️ Descargar PDF",
                                    data=pdf_bytes,
                                    file_name=f"informe_{nombre.lower().replace(' ', '_')}.pdf",
                                    mime="application/pdf",
                                    key=f"dl_inf_{nombre}",
                                )

# ---------------------------------------------------------------------------
# TAB 3 — Informe unificado
# ---------------------------------------------------------------------------
with tab_unificado:
    st.subheader("📄 Informe PDF (unificado)")

    port = st.session_state["portafolio"]
    if not port:
        st.info("Guarda al menos un Pareto para generar el informe unificado.")
    else:
        nombres = list(port.keys())
        seleccion = st.multiselect(
            "Selecciona los Paretos a incluir en el informe unificado",
            options=nombres,
            default=nombres
        )

        if seleccion:
            mapas = [port[n] for n in seleccion]
            mapa_total = combinar_maps(mapas)
            df_uni = calcular_pareto(df_desde_freq_map(mapa_total))
            st.subheader("📊 Vista previa Pareto Unificado")
            dibujar_pareto(df_uni, "Pareto Unificado")
            st.caption(f"Total de respuestas tratadas: {int(df_uni['frecuencia'].sum())}")

            desgloses_uni = ui_desgloses(df_uni["descriptor"].tolist(), key_prefix="uni")

            if st.button("📄 Generar Informe PDF (Unificado)", type="primary"):
                pdf_bytes = generar_pdf_informe("Pareto Unificado", df_uni, desgloses_uni)
                if pdf_bytes:
                    st.download_button(
                        label="📥 Descargar Informe PDF (Unificado)",
                        data=pdf_bytes,
                        file_name="Informe_Pareto_Unificado.pdf",
                        mime="application/pdf"
                    )

# ============================================================================
# ============================== PARTE 10/10 ================================
# ======================== Créditos y limpieza final ========================
# ============================================================================

st.divider()
st.markdown("""
<div style="text-align:center; font-size:14px; color:gray;">
Desarrollado para la Estrategia <b>Sembremos Seguridad</b><br>
Aplicación de análisis Pareto 80/20 con Google Sheets + ReportLab<br>
Versión 2025⚙️
</div>
""", unsafe_allow_html=True)

# Limpieza opcional de variables de sesión obsoletas
for key in ["sheet_url_loaded", "reset_after_save"]:
    if key not in st.session_state:
        st.session_state[key] = None

# Mensaje final
st.toast("✅ App lista. Puedes generar, guardar y eliminar Paretos con total integración.", icon="✅")










