# -*- coding: utf-8 -*-
"""
Moon Aerolínea · Fuel Intelligence  (v2)
Registro, revisión y auditoría de tiquetes de combustible de aviación.

Ejecutar:  streamlit run app.py
Secretos:  ANTHROPIC_API_KEY, ADMIN_PASSWORD  (.streamlit/secrets.toml o variables de entorno)
"""
import base64
import hashlib
import hmac
import io
import json
import logging
import os
import re
import sqlite3
import time
import warnings
from contextlib import contextmanager
from datetime import date, datetime, timedelta
from html import escape
from pathlib import Path
from textwrap import dedent

import anthropic
import pandas as pd
import plotly.express as px
import streamlit as st
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from PIL import Image, ImageOps
from plotly.subplots import make_subplots

st.set_page_config(
    page_title="Moon · Aerolínea",
    page_icon="🌙",
    layout="wide",
    initial_sidebar_state="expanded",
)

log = logging.getLogger("moon")
ss = st.session_state

# ============================================================
# Identidad visual
# ============================================================
NAVY, BLUE, SKY, GOLD = "#0B1D33", "#356B9A", "#7FA9CF", "#C9A96E"
INK, MUTED = "#17283D", "#5F7087"
OK, WARN, BAD = "#1F8A5B", "#C98A1B", "#B83A32"
PALETA = [BLUE, NAVY, SKY, GOLD, "#8FB3D1", "#4F6F91", "#E3D3B0", "#2E4A68"]
ESCALA = [[0, "#CFDDEB"], [1, NAVY]]
ESCALA_SUAVE = [[0, "#F1F5FA"], [1, BLUE]]

CSS = """
@import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700&family=Manrope:wght@600;700;800&display=swap');

html, body, .stApp, .stMarkdown, p, label, input, textarea, button, [data-testid="stSidebar"] {
    font-family: 'DM Sans', sans-serif;
}
.stApp { background: #F4F7FB; color: __INK__; }
#MainMenu, footer { visibility: hidden; }
header[data-testid="stHeader"] { background: transparent; }
.block-container { padding-top: 2.4rem; max-width: 1320px; }

/* Navegación lateral */
[data-testid="stSidebar"] { background: linear-gradient(180deg, __NAVY__ 0%, #12375A 100%); }
[data-testid="stSidebar"] * { color: #E6EDF5 !important; }
[data-testid="stSidebar"] hr { border-color: rgba(255,255,255,.14); }
.brand { font-family: 'Manrope', sans-serif; font-weight: 800; font-size: 20px; letter-spacing: -.4px; }
.brand-sub { color: __GOLD__ !important; font-size: 12px; margin-top: 2px; }
.side-stat { font-size: 12.5px; opacity: .85; margin: 4px 0; }

/* Títulos */
.page-title { color: __NAVY__; font-family: 'Manrope', sans-serif; font-weight: 800; font-size: clamp(27px, 3vw, 36px);
    letter-spacing: -1.2px; line-height: 1.1; margin: 0 0 14px 0; }
.section-heading { color: __NAVY__; font-family: 'Manrope', sans-serif; font-size: 17px; font-weight: 800;
    letter-spacing: -.3px; margin: 26px 0 2px 0; }
.section-sub { color: __MUTED__; font-size: 13px; margin: 0 0 12px 0; }

/* Hero con ruta de vuelo */
.hero-panel { position: relative; overflow: hidden; border-radius: 20px; padding: 28px 32px; margin: 4px 0 22px 0;
    background: linear-gradient(115deg, __NAVY__ 0%, #12375A 62%, #1A4568 100%); box-shadow: 0 10px 25px rgba(11,29,51,.14); }
.hero-title { color: #fff; font-family: 'Manrope', sans-serif; font-size: 26px; font-weight: 800; letter-spacing: -.6px; max-width: 560px; }
.hero-copy { color: #C3D0DE; font-size: 14px; margin-top: 8px; max-width: 560px; line-height: 1.6; }
.hero-route { position: absolute; right: 0; top: 0; height: 100%; width: 46%; opacity: .9; pointer-events: none; }
@media (max-width: 800px) { .hero-route { display: none; } }

/* Indicadores */
.metric-card { background: #fff; border: 1px solid #E5EAF1; border-radius: 14px; padding: 16px 18px; margin-bottom: 14px; }
.metric-card.alerta { border-left: 4px solid __WARN__; }
.metric-card.bien { border-left: 4px solid __OK__; }
.metric-label { color: __MUTED__; font-size: 13px; font-weight: 600; }
.metric-value { color: __NAVY__; font-family: 'Manrope', sans-serif; font-size: 25px; font-weight: 800; margin-top: 8px; line-height: 1.15; }
.kpi-delta { display: inline-block; margin-top: 8px; font-size: 12px; font-weight: 700; color: __BLUE__; }
.kpi-note { margin-top: 6px; font-size: 12px; color: __MUTED__; }

/* Pasos del piloto (secuencia real) */
.stepper { display: flex; gap: 8px; margin: 0 0 14px 0; flex-wrap: wrap; }
.step { padding: 7px 14px; border-radius: 999px; background: #E6ECF4; color: __MUTED__; font-size: 13px; font-weight: 700; }
.step.on { background: __NAVY__; color: #fff; }

/* Confirmación (único momento animado) */
@keyframes pop { 0% { transform: scale(.6); opacity: 0; } 70% { transform: scale(1.08); } 100% { transform: scale(1); opacity: 1; } }
.ok-card { display: flex; gap: 16px; align-items: center; background: #fff; border: 1px solid #CFE8DC; border-left: 5px solid __OK__;
    border-radius: 14px; padding: 20px 22px; margin: 8px 0 16px 0; }
.ok-icon { width: 46px; height: 46px; border-radius: 50%; background: __OK__; color: #fff; display: flex; align-items: center;
    justify-content: center; font-size: 24px; font-weight: 800; animation: pop .5s ease-out; flex: none; }
.ok-title { color: __NAVY__; font-family: 'Manrope', sans-serif; font-size: 18px; font-weight: 800; }
.ok-copy { color: __MUTED__; font-size: 14px; margin-top: 2px; }

.vacio { text-align: center; padding: 34px 20px; background: #fff; border: 1px dashed #C5D3E2; border-radius: 16px; color: __MUTED__; }
.vacio b { display: block; color: __NAVY__; font-family: 'Manrope', sans-serif; font-size: 16px; margin: 6px 0 4px 0; }
.vacio span.ico { font-size: 30px; }

/* Controles */
.stTabs [data-baseweb="tab-list"] { gap: 6px; border-bottom: 1px solid #DDE5EE; }
.stTabs [data-baseweb="tab"] { height: 46px; padding: 0 18px; border-radius: 12px 12px 0 0; font-weight: 700; color: __MUTED__; }
.stTabs [aria-selected="true"] { color: __NAVY__; background: #fff; }
.stButton > button, .stDownloadButton > button, [data-testid="stFormSubmitButton"] > button {
    border-radius: 12px; font-weight: 700; transition: transform .2s ease, box-shadow .2s ease; }
.stButton > button:hover, .stDownloadButton > button:hover, [data-testid="stFormSubmitButton"] > button:hover {
    transform: translateY(-1px); box-shadow: 0 6px 14px rgba(11,29,51,.14); }
button[kind="primary"], button[data-testid="stBaseButton-primary"], button[data-testid="stBaseButton-primaryFormSubmit"] {
    background: __NAVY__ !important; border-color: __NAVY__ !important; color: #fff !important; }
button[kind="primary"]:hover, button[data-testid="stBaseButton-primary"]:hover { background: #12375A !important; }
.stFileUploader label, .stTextInput label, .stNumberInput label, .stSelectbox label, .stDateInput label {
    color: __NAVY__ !important; font-weight: 700 !important; font-size: 14px !important; }
[data-testid="stFileUploaderDropzone"], [data-testid="stFileUploadDropzone"] {
    background: #fff !important; border: 2px dashed #AABCD0 !important; border-radius: 16px !important; }
[data-testid="stFileUploaderDropzone"]:hover { border-color: __BLUE__ !important; }
[data-testid="stFileUploaderDropzone"] *, [data-testid="stFileUploadDropzone"] * { color: __INK__ !important; }
[data-testid="stVerticalBlockBorderWrapper"] { background: #fff; border-radius: 16px; border-color: #E5EAF1 !important; }
</style>
"""
for _k, _v in {"__NAVY__": NAVY, "__BLUE__": BLUE, "__GOLD__": GOLD, "__INK__": INK, "__MUTED__": MUTED,
               "__OK__": OK, "__WARN__": WARN}.items():
    CSS = CSS.replace(_k, _v)
st.markdown("<style>" + CSS, unsafe_allow_html=True)

# ============================================================
# Configuración y utilidades generales
# ============================================================
def secret(nombre, default=None):
    """Lee st.secrets o, si no existe, una variable de entorno."""
    try:
        valor = st.secrets.get(nombre)
    except Exception:
        valor = None
    return valor or os.environ.get(nombre, default)


DATA_DIR = Path(os.environ.get("MOON_DATA_DIR", "."))
CARPETA_IMAGENES = DATA_DIR / "imagenes_tiquetes"
DB_PATH = DATA_DIR / "moon_fuel.db"
LEGACY_DIARIO = DATA_DIR / "Registro_Diario_Moon.xlsx"
LEGACY_MAESTRO = DATA_DIR / "Base_Datos_Maestra_Moon.xlsx"
MODELO = secret("ANTHROPIC_MODEL", "claude-sonnet-5-5")
PRECIO_DEFECTO = 19500.0
PRODUCTOS = ["JET A1", "AVGAS 100LL", "OTRO"]
LITROS_POR_GALON = 3.78541

# columna en base de datos -> encabezado visible / Excel
COLS = {
    "fecha_registro": "FECHA REGISTRO", "recibo": "RECIBO", "fecha_consumo": "FECHA CONSUMO",
    "matricula": "MATRICULA", "producto": "PRODUCTO", "ruta": "RUTA", "consumo_gal": "CONSUMO (GAL)",
    "valor_unit": "VALOR UNIT", "valor_total": "VALOR TOTAL", "piloto": "PILOTO",
    "imagen": "NOMENCLATURA IMAGEN", "observacion": "OBSERVACION",
}
CABECERAS = list(COLS.values())
NUM = ["CONSUMO (GAL)", "VALOR UNIT", "VALOR TOTAL"]


def html(s):
    st.markdown(dedent(s).strip(), unsafe_allow_html=True)


def cop(v):
    return f"$ {v:,.0f}"


def a_float(v, default=0.0):
    """Convierte '1.234,5', '1,234.5', 12 o None a float."""
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return default
    if isinstance(v, (int, float)):
        return float(v)
    s = re.sub(r"[^\d,.\-]", "", str(v))
    if "," in s and "." in s:
        s = s.replace(".", "").replace(",", ".") if s.rfind(",") > s.rfind(".") else s.replace(",", "")
    elif "," in s:
        s = s.replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return default


def parse_fechas(serie):
    """ISO (AAAA-MM-DD, con o sin hora) primero; dd/mm/aaaa como respaldo. Evita invertir día y mes."""
    s = serie.astype("string").str.strip()
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        iso = pd.to_datetime(s, format="ISO8601", errors="coerce")
        otras = pd.to_datetime(s.where(iso.isna()), format="mixed", dayfirst=True, errors="coerce")
    return iso.fillna(otras)


def a_fecha(v):
    if not v:
        return None
    f = parse_fechas(pd.Series([str(v)])).iloc[0]
    return None if pd.isna(f) else f.date()


def norm_matricula(m):
    m = re.sub(r"[\s_]+", "", str(m or "").upper())
    r = re.match(r"^HK-?(\d.*)$", m)
    return f"HK-{r.group(1)}" if r else m


def norm_producto(p):
    s = str(p or "").upper().replace("-", " ")
    if "JET" in s:
        return "JET A1"
    if "AVGAS" in s or "100" in s:
        return "AVGAS 100LL"
    return "OTRO"


def _nulo(v):
    if v is None or (not isinstance(v, (list, dict)) and pd.isna(v)):
        return None
    return v.item() if hasattr(v, "item") else v


# ============================================================
# Base de datos (SQLite): segura con varios pilotos a la vez
# ============================================================
@contextmanager
def conexion():
    con = sqlite3.connect(DB_PATH, timeout=30)
    try:
        yield con
        con.commit()
    finally:
        con.close()


@st.cache_resource
def init_db():
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    CARPETA_IMAGENES.mkdir(parents=True, exist_ok=True)
    with conexion() as con:
        con.execute("PRAGMA journal_mode=WAL")
        con.execute("""CREATE TABLE IF NOT EXISTS registros (
            id INTEGER PRIMARY KEY AUTOINCREMENT, estado TEXT NOT NULL DEFAULT 'PENDIENTE',
            fecha_registro TEXT, recibo TEXT, fecha_consumo TEXT, matricula TEXT, producto TEXT, ruta TEXT,
            consumo_gal REAL DEFAULT 0, valor_unit REAL DEFAULT 0, valor_total REAL DEFAULT 0,
            piloto TEXT, imagen TEXT, observacion TEXT, fecha_cierre TEXT)""")
        con.execute("CREATE TABLE IF NOT EXISTS meta (clave TEXT PRIMARY KEY, valor TEXT)")
        vacia = con.execute("SELECT COUNT(*) FROM registros").fetchone()[0] == 0
    if vacia:  # importa los Excel de la versión anterior, si existen
        for ruta, estado in ((LEGACY_MAESTRO, "ARCHIVADO"), (LEGACY_DIARIO, "PENDIENTE")):
            try:
                if ruta.exists():
                    df = pd.read_excel(ruta).reindex(columns=CABECERAS)
                    if not df.empty:
                        df = df.rename(columns={v: k for k, v in COLS.items()})
                        df["estado"] = estado
                        with conexion() as con:
                            df.to_sql("registros", con, if_exists="append", index=False)
            except Exception:
                log.exception("No se pudo importar %s", ruta)
    return True


def cargar(estado=None):
    q, p = "SELECT * FROM registros", ()
    if estado:
        q, p = q + " WHERE estado=?", (estado,)
    with conexion() as con:
        df = pd.read_sql_query(q + " ORDER BY id", con, params=p)
    df = df.rename(columns={"id": "ID", "estado": "ESTADO", **COLS})
    for c in NUM:
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0.0)
    return df[["ID", "ESTADO"] + CABECERAS + ["fecha_cierre"]]


def contar(estado):
    with conexion() as con:
        return con.execute("SELECT COUNT(*) FROM registros WHERE estado=?", (estado,)).fetchone()[0]


def get_meta(clave, default=None):
    with conexion() as con:
        r = con.execute("SELECT valor FROM meta WHERE clave=?", (clave,)).fetchone()
    return r[0] if r else default


def set_meta(clave, valor):
    with conexion() as con:
        con.execute("INSERT OR REPLACE INTO meta VALUES (?,?)", (clave, str(valor)))


def insertar_registro(d):
    cols = list(COLS)
    with conexion() as con:
        con.execute(f"INSERT INTO registros ({','.join(cols)},estado) VALUES ({','.join('?' * len(cols))},'PENDIENTE')",
                    [d.get(c) for c in cols])


def buscar_recibo(recibo, matricula):
    with conexion() as con:
        return con.execute(
            "SELECT fecha_registro FROM registros WHERE UPPER(TRIM(recibo))=? AND UPPER(TRIM(matricula))=?",
            (recibo.strip().upper(), matricula.strip().upper())).fetchone()


def _valores(fila, nuevo):
    v = {c: _nulo(fila.get(COLS[c])) for c in COLS}
    v["consumo_gal"], v["valor_unit"] = a_float(v["consumo_gal"]), a_float(v["valor_unit"])
    v["valor_total"] = round(v["consumo_gal"] * v["valor_unit"]) if v["valor_unit"] > 0 else a_float(v["valor_total"])
    if v["matricula"]:
        v["matricula"] = norm_matricula(v["matricula"])
    if nuevo:
        v["fecha_registro"] = v["fecha_registro"] or datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        v["observacion"] = v["observacion"] or "Agregado por Dirección"
    return v


def guardar_cambios(original, editado, estado):
    """Sincroniza la tabla editable con la base: borra, actualiza e inserta."""
    cols = list(COLS)
    ids_editado = {int(i) for i in editado["ID"].dropna()}
    with conexion() as con:
        for i in original["ID"]:
            if int(i) not in ids_editado:
                con.execute("DELETE FROM registros WHERE id=?", (int(i),))
        for _, fila in editado.iterrows():
            if pd.isna(fila["ID"]):
                v = _valores(fila, True)
                con.execute(f"INSERT INTO registros ({','.join(cols)},estado) VALUES ({','.join('?' * len(cols))},?)",
                            [v[c] for c in cols] + [estado])
            else:
                v = _valores(fila, False)
                con.execute(f"UPDATE registros SET {','.join(c + '=?' for c in cols)} WHERE id=?",
                            [v[c] for c in cols] + [int(fila["ID"])])


def aplicar_precios(precios, sobrescribir):
    with conexion() as con:
        for prod, p in precios.items():
            q = ("UPDATE registros SET valor_unit=?, valor_total=ROUND(consumo_gal*?,0), observacion='Validado por Dirección' "
                 "WHERE estado='PENDIENTE' AND COALESCE(NULLIF(TRIM(producto),''),'SIN PRODUCTO')=?")
            if not sobrescribir:
                q += " AND COALESCE(valor_unit,0)=0"
            con.execute(q, (p, p, prod))
    for prod, p in precios.items():
        set_meta(f"precio_{prod}", p)


def archivar_dia():
    with conexion() as con:
        return con.execute("UPDATE registros SET estado='ARCHIVADO', fecha_cierre=? WHERE estado='PENDIENTE'",
                           (datetime.now().strftime("%Y-%m-%d %H:%M:%S"),)).rowcount


def preparar(df):
    """Copia analítica: fechas reales, mes y categorías sin vacíos. No altera el original."""
    d = df.copy()
    consumo, registro = parse_fechas(d["FECHA CONSUMO"]), parse_fechas(d["FECHA REGISTRO"])
    d["REGISTRO_DT"] = registro
    d["FECHA"] = consumo.fillna(registro).fillna(pd.Timestamp.today()).dt.normalize()
    d["MES"] = d["FECHA"].dt.strftime("%Y-%m")
    for c in ["MATRICULA", "RUTA", "PRODUCTO", "PILOTO"]:
        d[c] = d[c].fillna("").astype(str).str.strip().replace("", "SIN DATO")
    return d


# ============================================================
# Imagen + lectura con IA
# ============================================================
def preparar_imagen(raw, lado_max=1800):
    """Corrige la rotación de celulares y reduce el peso (la API limita a 5 MB por imagen)."""
    img = ImageOps.exif_transpose(Image.open(io.BytesIO(raw))).convert("RGB")
    img.thumbnail((lado_max, lado_max))
    buf = io.BytesIO()
    img.save(buf, "JPEG", quality=88, optimize=True)
    return buf.getvalue()


PROMPT = """Eres el asistente de revisión de tiquetes de combustible de aviación de Moon Aerolínea (proveedor habitual: Terpel).
Analiza la imagen y responde ÚNICAMENTE con un objeto JSON, sin texto adicional ni bloques de código.

Campos a extraer:
- recibo: número del tiquete o recibo (texto)
- fecha_consumo: fecha del suministro en formato AAAA-MM-DD
- matricula: matrícula de la aeronave (por ejemplo HK-4567)
- producto: "JET A1", "AVGAS 100LL" u otro, tal como aparece
- ruta: ruta o destino SOLO si está escrito; si no aparece, ""
- consumo: volumen suministrado como número con punto decimal
- unidad: "GAL" o "LT", según aparezca en el tiquete
- piloto: nombre del piloto o firmante SOLO si aparece; si no, ""
- campos_dudosos: lista con los nombres de los campos que no estés seguro de haber leído bien

Reglas:
1. Nunca inventes datos. Si un campo no se ve o no aparece, usa "".
2. Si la imagen está borrosa, cortada, es de otro documento o NO SE LEE CLARAMENTE el volumen, responde
   {"exito": false, "motivo_error": "..."} con un mensaje breve y amable para el piloto que explique qué falló y cómo repetir la foto.
3. Si todo está bien responde:
   {"exito": true, "motivo_error": "", "recibo": "", "fecha_consumo": "", "matricula": "", "producto": "", "ruta": "",
    "consumo": 0.0, "unidad": "GAL", "piloto": "", "campos_dudosos": []}
"""


def extraer_json(texto):
    m = re.search(r"\{.*\}", texto or "", re.S)
    try:
        return json.loads(m.group(0)) if m else {}
    except json.JSONDecodeError:
        return {}


def analizar_con_ia(jpeg):
    api_key = secret("ANTHROPIC_API_KEY")
    if not api_key:
        return {"exito": False, "motivo_error": "La lectura automática no está configurada. Puedes llenar los datos a mano."}
    try:
        cliente = anthropic.Anthropic(api_key=api_key, timeout=60.0, max_retries=2)
        r = cliente.messages.create(
            model=MODELO, max_tokens=1024,
            messages=[{"role": "user", "content": [
                {"type": "image", "source": {"type": "base64", "media_type": "image/jpeg",
                                             "data": base64.b64encode(jpeg).decode()}},
                {"type": "text", "text": PROMPT}]}])
        res = extraer_json("".join(b.text for b in r.content if getattr(b, "type", "") == "text"))
        return res or {"exito": False, "motivo_error": "No pude interpretar el comprobante. Repite la foto con mejor luz."}
    except Exception:
        log.exception("Fallo al leer el tiquete con IA")
        return {"exito": False, "motivo_error": "No hubo conexión con la lectura automática. Intenta de nuevo o llena los datos a mano."}


def datos_vacios():
    return {"recibo": "", "fecha": None, "matricula": "", "producto": "JET A1", "ruta": "", "consumo": 0.0,
            "piloto": "", "dudosos": []}


def normalizar_datos(r):
    gal, dudosos = a_float(r.get("consumo")), list(r.get("campos_dudosos") or [])
    if str(r.get("unidad", "GAL")).upper().startswith("L"):
        gal /= LITROS_POR_GALON
        dudosos.append("el volumen venía en litros y se convirtió a galones")
    return {"recibo": str(r.get("recibo") or "").strip(), "fecha": a_fecha(r.get("fecha_consumo")),
            "matricula": norm_matricula(r.get("matricula")), "producto": norm_producto(r.get("producto")),
            "ruta": str(r.get("ruta") or "").strip(), "consumo": round(gal, 1),
            "piloto": str(r.get("piloto") or "").strip(), "dudosos": dudosos}


# ============================================================
# Gráficos
# ============================================================
PLOT_CFG = {"displaylogo": False, "modeBarButtonsToRemove": ["lasso2d", "select2d", "autoScale2d"]}


def estilo(fig, alto=360):
    fig.update_layout(
        height=alto, paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)", colorway=PALETA,
        font=dict(family="DM Sans, sans-serif", color=MUTED, size=12),
        title_font=dict(family="Manrope, sans-serif", size=16, color=NAVY), title_x=0,
        margin=dict(l=10, r=10, t=56, b=10), legend=dict(orientation="h", y=-0.14, x=0),
        hoverlabel=dict(bgcolor=NAVY, font_color="white", font_family="DM Sans, sans-serif"))
    fig.update_xaxes(showgrid=False, linecolor="#DDE5EE", zeroline=False)
    fig.update_yaxes(gridcolor="#E9EEF5", zeroline=False)
    return fig


def mostrar(fig, vacio_msg="Aún no hay datos para este gráfico."):
    with st.container(border=True):
        if fig is None:
            st.caption(vacio_msg)
        else:
            st.plotly_chart(fig, width="stretch", config=PLOT_CFG)


def barras_h(d, cat, val, titulo, fmt=",.1f", top=10):
    g = d.groupby(cat, as_index=False)[val].sum()
    g = g[g[val] > 0].nlargest(top, val).sort_values(val)
    if g.empty:
        return None
    fig = px.bar(g, x=val, y=cat, orientation="h", title=titulo, text=val, color=val, color_continuous_scale=ESCALA)
    fig.update_traces(texttemplate="%{text:" + fmt + "}", textposition="outside", cliponaxis=False)
    fig.update_coloraxes(showscale=False)
    fig.update_xaxes(visible=False)
    fig.update_yaxes(title=None)
    return estilo(fig, max(280, 56 + 34 * len(g)))


def dona(d, cat, val, titulo, centro):
    g = d.groupby(cat, as_index=False)[val].sum()
    g = g[g[val] > 0]
    if g.empty:
        return None
    fig = px.pie(g, names=cat, values=val, hole=0.6, title=titulo, color_discrete_sequence=PALETA)
    fig.update_traces(textinfo="percent+label", textposition="outside", marker=dict(line=dict(color="white", width=2)))
    fig.add_annotation(text=centro, showarrow=False, font=dict(family="Manrope, sans-serif", size=15, color=NAVY))
    fig = estilo(fig)
    fig.update_layout(showlegend=False)
    return fig


def tendencia(d, freq):
    regla = {"Día": "D", "Semana": "W", "Mes": "MS"}[freq]
    g = d.set_index("FECHA").resample(regla).agg({"CONSUMO (GAL)": "sum", "VALOR TOTAL": "sum"}).reset_index()
    fig = make_subplots(specs=[[{"secondary_y": True}]])
    fig.add_bar(x=g["FECHA"], y=g["CONSUMO (GAL)"], name="Galones", marker_color=BLUE, secondary_y=False)
    if g["VALOR TOTAL"].sum() > 0:
        fig.add_scatter(x=g["FECHA"], y=g["VALOR TOTAL"], name="Costo (COP)", mode="lines+markers",
                        line=dict(color=GOLD, width=3), marker=dict(size=6), secondary_y=True)
    fig.update_layout(title="Consumo y costo en el tiempo")
    fig = estilo(fig, 380)
    fig.update_yaxes(title_text="Galones", secondary_y=False)
    fig.update_yaxes(title_text="COP", showgrid=False, secondary_y=True)
    return fig


def treemap(d, val, titulo):
    g = d.groupby(["RUTA", "MATRICULA"], as_index=False)[val].sum()
    g = g[g[val] > 0]
    if g.empty:
        return None
    fig = px.treemap(g, path=["RUTA", "MATRICULA"], values=val, color=val, color_continuous_scale=ESCALA, title=titulo)
    fig.update_traces(texttemplate="%{label}<br>%{value:,.0f}", textfont=dict(color="white"))
    fig.update_coloraxes(showscale=False)
    return estilo(fig)


def mapa_calor(d, val, titulo, fmt):
    p = d.pivot_table(index="MATRICULA", columns="MES", values=val, aggfunc="sum", fill_value=0)
    if p.empty or p.to_numpy().sum() == 0:
        return None
    fig = px.imshow(p, aspect="auto", color_continuous_scale=ESCALA_SUAVE, text_auto=fmt, title=titulo)
    fig.update_coloraxes(showscale=False)
    fig.update_xaxes(title=None, side="top")
    fig.update_yaxes(title=None)
    return estilo(fig, max(280, 90 + 40 * len(p)))


def caja(d):
    if len(d) < 3:
        return None
    fig = px.box(d, x="MATRICULA", y="CONSUMO (GAL)", points="all", title="Galones por tiquete y aeronave",
                 color_discrete_sequence=[BLUE])
    fig.update_traces(marker=dict(size=5, opacity=0.6), line=dict(color=NAVY))
    fig.update_xaxes(title=None)
    return estilo(fig)


def precio_evolucion(d, freq):
    p = d[d["VALOR UNIT"] > 0]
    if p.empty:
        return None
    regla = {"Día": "D", "Semana": "W", "Mes": "MS"}[freq]
    g = p.groupby([pd.Grouper(key="FECHA", freq=regla), "PRODUCTO"]).agg(
        COSTO=("VALOR TOTAL", "sum"), GAL=("CONSUMO (GAL)", "sum")).reset_index()
    g = g[g["GAL"] > 0]
    g["PRECIO"] = g["COSTO"] / g["GAL"]
    fig = px.line(g, x="FECHA", y="PRECIO", color="PRODUCTO", markers=True, title="Precio promedio por galón",
                  color_discrete_sequence=[BLUE, GOLD, NAVY])
    fig.update_xaxes(title=None)
    fig.update_yaxes(title=None, tickprefix="$ ", tickformat=",.0f")
    return estilo(fig)


# ============================================================
# Componentes de interfaz
# ============================================================
RUTA_SVG = f"""<svg class="hero-route" viewBox="0 0 400 160" preserveAspectRatio="xMaxYMid slice" aria-hidden="true">
<path d="M10 140 C 110 20, 260 10, 392 70" fill="none" stroke="{GOLD}" stroke-width="1.6" stroke-dasharray="3 7" stroke-linecap="round"/>
<circle cx="10" cy="140" r="4" fill="none" stroke="{GOLD}" stroke-width="1.6"/>
<circle cx="392" cy="70" r="5" fill="{GOLD}"/></svg>"""


def hero(titulo, texto):
    html(f'<div class="hero-panel">{RUTA_SVG}<div class="hero-title">{titulo}</div><div class="hero-copy">{texto}</div></div>')


def kpi(label, value, delta=None, nota=None, tono=""):
    dlt = ""
    if delta is not None:
        dlt = f'<div class="kpi-delta">{"▲" if delta >= 0 else "▼"} {abs(delta):.1f}% frente al periodo anterior</div>'
    nt = f'<div class="kpi-note">{nota}</div>' if nota else ""
    html(f'<div class="metric-card {tono}"><div class="metric-label">{label}</div><div class="metric-value">{value}</div>{dlt}{nt}</div>')


def titulo(t, sub=None):
    html(f'<div class="section-heading">{t}</div>' + (f'<div class="section-sub">{sub}</div>' if sub else ""))


def vacio(icono, t, texto):
    html(f'<div class="vacio"><span class="ico">{icono}</span><b>{t}</b>{texto}</div>')


def stepper(paso):
    items = "".join(f'<div class="step {"on" if i < paso else ""}">{i + 1}. {t}</div>'
                    for i, t in enumerate(["Foto", "Revisión", "Enviado"]))
    html(f'<div class="stepper">{items}</div>')


def visor_tiquete(df, key):
    """Muestra la foto guardada junto a los datos leídos, para verificar a ojo."""
    if df.empty:
        return
    etiquetas = {int(r["ID"]): f"#{int(r['ID'])} · recibo {r['RECIBO'] or 's/n'} · {r['MATRICULA'] or 's/m'} · {r['CONSUMO (GAL)']:,.1f} gal"
                 for _, r in df.iterrows()}
    elegido = st.selectbox("Tiquete a verificar", list(etiquetas), format_func=etiquetas.get, key=key)
    fila = df[df["ID"] == elegido].iloc[0]
    c1, c2 = st.columns([1, 1], gap="large")
    ruta = CARPETA_IMAGENES / str(fila["NOMENCLATURA IMAGEN"] or "")
    with c1:
        if ruta.is_file():
            st.image(str(ruta), width="stretch")
        else:
            st.warning("Este registro no tiene imagen guardada.")
    with c2:
        for c in ["RECIBO", "FECHA CONSUMO", "MATRICULA", "PRODUCTO", "RUTA", "CONSUMO (GAL)", "PILOTO", "FECHA REGISTRO"]:
            st.markdown(f"**{c.title()}:** {fila[c]}")


# ---- Excel profesional -------------------------------------------------
def _formato_hoja(ws, df):
    for c in ws[1]:
        c.fill = PatternFill("solid", fgColor=NAVY[1:])
        c.font = Font(bold=True, color="FFFFFF")
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    for i, col in enumerate(df.columns, 1):
        letra, nombre = get_column_letter(i), str(col).upper()
        ancho = max([len(str(col))] + [len(str(v)) for v in df[col].head(200)])
        ws.column_dimensions[letra].width = min(max(ancho + 3, 12), 42)
        fmt = ("$ #,##0" if any(k in nombre for k in ("VALOR", "COSTO", "PRECIO", "INVERSION"))
               else "#,##0.0" if "GAL" in nombre else None)
        if fmt:
            for celda in ws[letra][1:]:
                celda.number_format = fmt


def excel_bytes(hojas):
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as xw:
        for nombre, df in hojas.items():
            df.to_excel(xw, sheet_name=nombre[:31], index=False)
            _formato_hoja(xw.sheets[nombre[:31]], df)
    return buf.getvalue()


def resumen(d, col):
    g = d.groupby(col).agg(TIQUETES=("ID", "count"), **{"GALONES": ("CONSUMO (GAL)", "sum"), "COSTO (COP)": ("VALOR TOTAL", "sum")}).reset_index()
    g["PRECIO PROM/GAL"] = (g["COSTO (COP)"] / g["GALONES"]).where(g["GALONES"] > 0, 0).round(0)
    return g.sort_values("GALONES", ascending=False)


# ============================================================
# Auditoría automática
# ============================================================
def hallazgos(df):
    d, filas = preparar(df), []

    def agregar(sev, tipo, sub, detalle):
        for _, r in sub.iterrows():
            filas.append({"SEVERIDAD": sev, "HALLAZGO": tipo, "ID": int(r["ID"]), "RECIBO": r["RECIBO"],
                          "MATRICULA": r["MATRICULA"], "FECHA": r["FECHA"].date(), "GALONES": r["CONSUMO (GAL)"],
                          "DETALLE": detalle(r)})

    rec = d["RECIBO"].fillna("").astype(str).str.strip()
    dup = d[(rec != "") & d.assign(_r=rec).duplicated(["_r", "MATRICULA"], keep=False)]
    agregar("Alta", "Recibo duplicado", dup, lambda r: "El mismo recibo y matrícula aparecen más de una vez")

    sin_img = d[~d["NOMENCLATURA IMAGEN"].apply(lambda n: bool(n) and (CARPETA_IMAGENES / str(n)).is_file())]
    agregar("Alta", "Sin imagen de soporte", sin_img, lambda r: "No se encontró la foto del comprobante")

    for mat, g in d.groupby("MATRICULA"):
        s = g["CONSUMO (GAL)"].std()
        if len(g) >= 5 and s and s > 0:
            z = (g["CONSUMO (GAL)"] - g["CONSUMO (GAL)"].mean()) / s
            agregar("Media", "Volumen atípico", g[z.abs() > 2.5],
                    lambda r, m=g["CONSUMO (GAL)"].mean(): f"{r['CONSUMO (GAL)']:.1f} gal frente a un promedio de {m:.1f} gal en esta aeronave")

    faltan = d[(d[["MATRICULA", "PILOTO", "RUTA"]] == "SIN DATO").any(axis=1)]
    agregar("Media", "Datos incompletos", faltan,
            lambda r: "Falta: " + ", ".join(c.lower() for c in ["MATRICULA", "PILOTO", "RUTA"] if r[c] == "SIN DATO"))

    tarde = d[(d["REGISTRO_DT"].dt.normalize() - d["FECHA"]).dt.days > 3]
    agregar("Baja", "Registro tardío", tarde, lambda r: "Se registró más de 3 días después del suministro")

    pend = d[(d["ESTADO"] == "PENDIENTE") & (d["VALOR UNIT"] <= 0)]
    agregar("Baja", "Sin precio asignado", pend, lambda r: "Pendiente de valorizar en la bandeja del día")

    orden = {"Alta": 0, "Media": 1, "Baja": 2}
    out = pd.DataFrame(filas, columns=["SEVERIDAD", "HALLAZGO", "ID", "RECIBO", "MATRICULA", "FECHA", "GALONES", "DETALLE"])
    return out.sort_values(["SEVERIDAD", "FECHA"], key=lambda s: s.map(orden) if s.name == "SEVERIDAD" else s,
                           ascending=[True, False]).reset_index(drop=True)


# ============================================================
# Vista 1 · Registro en pista
# ============================================================
def vista_pista():
    html('<h1 class="page-title">Registro en pista</h1>')
    hero("Cada abastecimiento, bajo control.",
         "Toma o sube la foto del tiquete. El sistema lee los datos y tú solo confirmas que coincidan con el papel.")
    ss.setdefault("analisis", None)

    if ss.get("ok_msg"):
        stepper(3)
        m = ss.ok_msg
        html(f'<div class="ok-card"><div class="ok-icon">✓</div><div><div class="ok-title">Tiquete enviado a Dirección</div>'
             f'<div class="ok-copy"><b>{escape(m["matricula"])}</b>: {m["galones"]:,.1f} gal de {escape(m["producto"])}, '
             f'recibo {escape(m["recibo"])}.</div></div></div>')
        if st.button("Registrar otro tiquete", type="primary"):
            ss.ok_msg, ss.analisis = None, None
            ss.upl_n = ss.get("upl_n", 0) + 1
            st.rerun()
        return

    stepper(2 if ss.analisis else 1)
    col_form, col_prev = st.columns([1.25, 0.75], gap="large")
    n = ss.get("upl_n", 0)

    with col_form:
        titulo("Nuevo registro de combustible")
        origen = st.segmented_control("Origen de la imagen", ["📁 Subir archivo", "📷 Usar cámara"],
                                      default="📁 Subir archivo", key="origen", label_visibility="collapsed")
        if origen == "📷 Usar cámara":
            archivo = st.camera_input("Fotografía del comprobante", key=f"cam_{n}")
        else:
            archivo = st.file_uploader("Comprobante de abastecimiento (Terpel)", type=["jpg", "jpeg", "png"], key=f"upl_{n}")

        with col_prev:
            if archivo is not None:
                st.image(archivo, caption="Vista previa del comprobante", width="stretch")
            else:
                vacio("🧾", "Aquí verás tu comprobante", "Verifica que se lean el número de recibo y los galones.")

        if archivo is None:
            ss.analisis = None
            return

        raw = archivo.getvalue()
        huella = hashlib.sha1(raw).hexdigest()
        an = ss.analisis
        if not an or an["huella"] != huella:
            ss.analisis = None
            b1, b2 = st.columns([1.4, 1])
            analizar = b1.button("🔍 Leer comprobante", type="primary", width="stretch")
            manual = b2.button("✍️ Llenar a mano", width="stretch")
            if analizar or manual:
                try:
                    jpeg = preparar_imagen(raw)
                except Exception:
                    st.error("No pude abrir esa imagen. Toma la foto de nuevo e inténtalo otra vez.")
                    return
                if manual:
                    ss.analisis = {"huella": huella, "jpeg": jpeg, "datos": datos_vacios()}
                    st.rerun()
                with st.spinner("Leyendo el comprobante…"):
                    res = analizar_con_ia(jpeg)
                if res.get("exito"):
                    ss.analisis = {"huella": huella, "jpeg": jpeg, "datos": normalizar_datos(res)}
                    ss.fallo = None
                    st.rerun()
                ss.fallo = {"huella": huella, "msg": res.get("motivo_error") or "La imagen no es clara. Toma la foto de nuevo."}
            if ss.get("fallo") and ss.fallo["huella"] == huella:
                st.error(f"📸 {ss.fallo['msg']}")
            return

        d = an["datos"]
        titulo("Confirma los datos", "Compáralos con el papel y corrige lo que haga falta antes de enviar.")
        if d["dudosos"]:
            st.warning("Revisa con atención: " + ", ".join(d["dudosos"]) + ".")
        with st.form("form_tiquete"):
            c1, c2 = st.columns(2)
            recibo = c1.text_input("Número de recibo", d["recibo"])
            fecha = c2.date_input("Fecha del suministro", min(d["fecha"] or date.today(), date.today()),
                                  min_value=date.today() - timedelta(days=365), max_value=date.today())
            matricula = c1.text_input("Matrícula de la aeronave", d["matricula"], placeholder="HK-1234")
            producto = c2.selectbox("Producto", PRODUCTOS, index=PRODUCTOS.index(d["producto"]))
            ruta = c1.text_input("Ruta", d["ruta"], placeholder="Ej. MDE-BOG")
            galones = c2.number_input("Galones suministrados", min_value=0.0, value=float(d["consumo"]), step=0.1, format="%.1f")
            piloto = st.text_input("Piloto", d["piloto"] or ss.get("piloto_nombre", ""))
            enviar = st.form_submit_button("✅ Confirmar y enviar a Dirección", type="primary", width="stretch")

        if enviar:
            errores = []
            if not recibo.strip():
                errores.append("Escribe el número de recibo.")
            if not matricula.strip():
                errores.append("Escribe la matrícula.")
            if galones <= 0:
                errores.append("Los galones deben ser mayores que cero.")
            if not piloto.strip():
                errores.append("Escribe tu nombre.")
            matricula = norm_matricula(matricula)
            previo = buscar_recibo(recibo, matricula) if recibo.strip() and matricula else None
            if previo:
                errores.append(f"Este recibo ya fue registrado para {matricula} ({previo[0]}). Si es otro tiquete, revisa el número.")
            if errores:
                for e in errores:
                    st.error(e)
                return
            seguro = re.sub(r"[^A-Za-z0-9_-]", "", recibo) or "SREC"
            nombre = f"Recibo_{seguro}_{datetime.now():%Y%m%d_%H%M%S}.jpg"
            (CARPETA_IMAGENES / nombre).write_bytes(an["jpeg"])
            insertar_registro({
                "fecha_registro": datetime.now().strftime("%Y-%m-%d %H:%M:%S"), "recibo": recibo.strip(),
                "fecha_consumo": fecha.isoformat(), "matricula": matricula, "producto": producto, "ruta": ruta.strip(),
                "consumo_gal": float(galones), "valor_unit": 0.0, "valor_total": 0.0, "piloto": piloto.strip(),
                "imagen": nombre, "observacion": "Enviado desde pista"})
            ss.piloto_nombre = piloto.strip()
            ss.ok_msg = {"matricula": matricula, "galones": galones, "producto": producto, "recibo": recibo.strip()}
            ss.analisis = None
            st.rerun()


# ============================================================
# Vista 2 · Panel directivo
# ============================================================
def login():
    clave_ok = secret("ADMIN_PASSWORD")
    if not clave_ok:
        st.error("🔒 Falta configurar ADMIN_PASSWORD (en .streamlit/secrets.toml o como variable de entorno).")
        return
    _, c, _ = st.columns([1, 1.2, 1])
    with c.form("login"):
        clave = st.text_input("Contraseña ejecutiva", type="password")
        ingresar = st.form_submit_button("Acceder al sistema", type="primary", width="stretch")
    if ingresar:
        espera = ss.get("bloqueo_hasta", 0) - time.time()
        if espera > 0:
            c.warning(f"Demasiados intentos. Espera {int(espera)} segundos.")
        elif hmac.compare_digest(clave.encode(), str(clave_ok).encode()):
            ss.moon_auth, ss.intentos = True, 0
            st.rerun()
        else:
            ss.intentos = ss.get("intentos", 0) + 1
            if ss.intentos >= 5:
                ss.bloqueo_hasta, ss.intentos = time.time() + 60, 0
            c.error("Contraseña incorrecta.")


@st.dialog("Cerrar y archivar el día")
def dialogo_cierre(n, gal, total, sin_precio):
    st.write(f"Se moverán **{n} tiquetes** ({gal:,.1f} gal · {cop(total)}) al histórico y la bandeja quedará vacía.")
    if sin_precio:
        st.warning(f"{sin_precio} tiquete(s) todavía no tienen precio asignado.")
    ok = st.checkbox("Revisé los tiquetes" + (" y acepto archivarlos sin precio" if sin_precio else ""))
    if st.button("📦 Archivar ahora", type="primary", disabled=not ok, width="stretch"):
        archivar_dia()
        st.balloons()
        time.sleep(2)
        ss.ver_diario = ss.get("ver_diario", 0) + 1
        ss.flash = "Día cerrado: los datos ya están en el histórico."
        st.rerun()


def hay_cambios(key):
    s = ss.get(key) or {}
    return bool(s.get("edited_rows") or s.get("added_rows") or s.get("deleted_rows"))


CFG_TABLA = {
    "FECHA REGISTRO": st.column_config.TextColumn(disabled=True),
    "CONSUMO (GAL)": st.column_config.NumberColumn(format="%.1f", min_value=0.0, step=0.1),
    "VALOR UNIT": st.column_config.NumberColumn(format="$ %.0f", min_value=0.0),
    "VALOR TOTAL": st.column_config.NumberColumn(format="$ %.0f", disabled=True, help="Se calcula: galones × valor unitario"),
    "NOMENCLATURA IMAGEN": st.column_config.TextColumn(disabled=True),
}


def tab_diario():
    df = cargar("PENDIENTE")
    if df.empty:
        vacio("🍃", "Bandeja vacía", "Cuando los pilotos envíen comprobantes aparecerán aquí para revisión.")
        return
    d = preparar(df)
    gal, total = df["CONSUMO (GAL)"].sum(), df["VALOR TOTAL"].sum()
    sin_precio = int((df["VALOR UNIT"] <= 0).sum())

    k = st.columns(5)
    with k[0]: kpi("Tiquetes por revisar", f"{len(df):,}")
    with k[1]: kpi("Volumen", f"{gal:,.1f} gal")
    with k[2]: kpi("Aeronaves", f"{d['MATRICULA'].nunique()}")
    with k[3]: kpi("Gasto estimado", cop(total))
    with k[4]: kpi("Sin precio", f"{sin_precio}", tono="alerta" if sin_precio else "bien",
                   nota="Asigna precio abajo" if sin_precio else "Todo valorizado")

    g1, g2 = st.columns(2)
    with g1: mostrar(barras_h(d, "MATRICULA", "CONSUMO (GAL)", "Galones por aeronave"))
    with g2: mostrar(barras_h(d, "RUTA", "CONSUMO (GAL)", "Galones por ruta"))
    g3, g4 = st.columns(2)
    with g3: mostrar(dona(d, "PRODUCTO", "CONSUMO (GAL)", "Mezcla de productos", f"{gal:,.0f}<br>gal"))
    with g4: mostrar(barras_h(d.assign(TIQUETES=1), "PILOTO", "TIQUETES", "Tiquetes enviados por piloto", fmt=".0f"))

    titulo("Bandeja de revisión", "Haz doble clic en una celda para corregirla. Los cambios se aplican al pulsar Guardar.")
    with st.expander("🔎 Verificar un comprobante contra la lectura automática"):
        visor_tiquete(df, "visor_diario")

    key = f"ed_diario_{ss.get('ver_diario', 0)}"
    editado = st.data_editor(df, column_order=CABECERAS, column_config=CFG_TABLA, num_rows="dynamic",
                             hide_index=True, width="stretch", key=key)
    cambios = hay_cambios(key)
    s1, s2 = st.columns([1, 3])
    if s1.button("💾 Guardar correcciones", type="primary", disabled=not cambios, width="stretch"):
        guardar_cambios(df, editado, "PENDIENTE")
        ss.ver_diario = ss.get("ver_diario", 0) + 1
        ss.flash = "Correcciones guardadas."
        st.rerun()
    if cambios:
        s2.warning("Tienes cambios sin guardar. Guárdalos antes de fijar precios o cerrar el día.")

    c_precio, c_cierre = st.columns(2, gap="large")
    with c_precio:
        titulo("Fijar precio por galón", "Cada producto se valoriza con su propio precio.")
        prod_lbl = df["PRODUCTO"].fillna("").astype(str).str.strip().replace("", "SIN PRODUCTO")
        precios = {}
        for p in sorted(prod_lbl.unique()):
            precios[p] = st.number_input(f"{p} (COP/gal)", min_value=0.0, step=100.0, key=f"precio_{p}",
                                         value=float(get_meta(f"precio_{p}", PRECIO_DEFECTO)))
        sobre = st.toggle("Reemplazar precios ya asignados", value=False)
        if st.button("Aplicar precios", type="primary", disabled=cambios, width="stretch"):
            aplicar_precios(precios, sobre)
            ss.ver_diario = ss.get("ver_diario", 0) + 1
            ss.flash = "Precios aplicados. Los indicadores se actualizaron."
            st.rerun()
    with c_cierre:
        titulo("Entregar y cerrar", "Descarga el Excel del día y luego archívalo en el histórico.")
        datos_xlsx = excel_bytes({"Registros del día": df[CABECERAS], "Por aeronave": resumen(d, "MATRICULA"),
                                  "Por producto": resumen(d, "PRODUCTO")})
        st.download_button("⬇️ Descargar Excel del día", datos_xlsx, f"Moon_Diario_{datetime.now():%Y%m%d}.xlsx",
                           "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", width="stretch")
        if st.button("📦 Archivar y cerrar el día", disabled=cambios, width="stretch"):
            dialogo_cierre(len(df), gal, total, sin_precio)


def rango_periodo(preset, fmin, fmax):
    hoy = date.today()
    return {"7 días": (hoy - timedelta(days=6), hoy), "30 días": (hoy - timedelta(days=29), hoy),
            "Este mes": (hoy.replace(day=1), hoy), "Este año": (date(hoy.year, 1, 1), hoy)}.get(preset, (fmin, fmax))


def variacion(actual, previo):
    return (actual - previo) / previo * 100 if previo else None


def tab_maestro():
    df = cargar("ARCHIVADO")
    if df.empty:
        vacio("🗂️", "El histórico está vacío", "Cierra un día en la bandeja para ver aquí los indicadores acumulados.")
        return
    d = preparar(df)
    fmin, fmax = d["FECHA"].min().date(), d["FECHA"].max().date()

    with st.container(border=True):
        a, b = st.columns([2.2, 1])
        preset = a.segmented_control("Periodo", ["7 días", "30 días", "Este mes", "Este año", "Todo", "Personalizado"],
                                     default="Todo", key="periodo") or "Todo"
        if preset == "Personalizado":
            r = b.date_input("Rango de fechas", value=(fmin, fmax), key="rango")
            ini, fin = (r[0], r[-1]) if len(r) else (fmin, fmax)
        else:
            ini, fin = rango_periodo(preset, fmin, fmax)
        with st.expander("Filtros avanzados"):
            f1, f2, f3, f4 = st.columns(4)
            sel = {c: w.multiselect(c.title(), sorted(d[c].unique()), key=f"f_{c}")
                   for c, w in zip(["MATRICULA", "RUTA", "PRODUCTO", "PILOTO"], [f1, f2, f3, f4])}

    base = d.copy()
    for c, vals in sel.items():
        if vals:
            base = base[base[c].isin(vals)]
    f = base[(base["FECHA"].dt.date >= ini) & (base["FECHA"].dt.date <= fin)]
    if f.empty:
        vacio("🔍", "Sin datos con estos filtros", "Amplía el periodo o quita algún filtro.")
        return

    previo = None
    if preset != "Todo":
        dias = (fin - ini).days + 1
        previo = base[(base["FECHA"].dt.date >= ini - timedelta(days=dias)) & (base["FECHA"].dt.date < ini)]
    gal, costo = f["CONSUMO (GAL)"].sum(), f["VALOR TOTAL"].sum()
    precio_prom = costo / f.loc[f["VALOR UNIT"] > 0, "CONSUMO (GAL)"].sum() if (f["VALOR UNIT"] > 0).any() else 0

    k = st.columns(5)
    with k[0]: kpi("Tiquetes", f"{len(f):,}", variacion(len(f), len(previo)) if previo is not None else None)
    with k[1]: kpi("Galones", f"{gal:,.1f}", variacion(gal, previo["CONSUMO (GAL)"].sum()) if previo is not None else None)
    with k[2]: kpi("Inversión", cop(costo), variacion(costo, previo["VALOR TOTAL"].sum()) if previo is not None else None)
    with k[3]: kpi("Precio promedio por galón", cop(precio_prom) if precio_prom else "Sin precios")
    with k[4]: kpi("Aeronaves activas", f"{f['MATRICULA'].nunique()}", nota=f"{gal / len(f):,.0f} gal por tiquete")

    m1, m2 = st.columns([1, 1])
    metrica = m1.segmented_control("Ver gráficos por", ["Galones", "Costo (COP)"], default="Galones", key="metrica") or "Galones"
    freq = m2.segmented_control("Agrupar tendencia por", ["Día", "Semana", "Mes"],
                                default="Mes" if (fmax - fmin).days > 120 else "Día", key="freq") or "Día"
    val = "CONSUMO (GAL)"
    if metrica == "Costo (COP)":
        if costo > 0:
            val = "VALOR TOTAL"
        else:
            st.info("Aún no hay precios asignados en este periodo; se muestran galones.")
    fmt_barra, fmt_calor = (",.1f", ".0f") if val == "CONSUMO (GAL)" else (",.0f", ".3s")

    mostrar(tendencia(f, freq))
    g1, g2 = st.columns(2)
    with g1: mostrar(barras_h(f, "MATRICULA", val, "Ranking por aeronave", fmt_barra, top=12))
    with g2: mostrar(treemap(f, val, "Ruta y aeronave"))
    g3, g4 = st.columns(2)
    with g3: mostrar(dona(f, "PRODUCTO", val, "Mezcla de productos", metrica.split(" ")[0]))
    with g4: mostrar(barras_h(f, "PILOTO", val, "Ranking por piloto", fmt_barra, top=12))
    mostrar(mapa_calor(f, val, "Aeronave por mes", fmt_calor), "Se necesita más de un mes de datos.")
    g5, g6 = st.columns(2)
    with g5: mostrar(precio_evolucion(f, freq), "Asigna precios para ver su evolución.")
    with g6: mostrar(caja(f))

    titulo("Base de datos general", "Edita cualquier error de días pasados y guarda. El filtro del periodo también aplica a la tabla.")
    crudo = df.loc[f.index].reset_index(drop=True)
    key = f"ed_maestro_{ss.get('ver_maestro', 0)}"
    editado = st.data_editor(crudo, column_order=CABECERAS, column_config=CFG_TABLA, num_rows="dynamic",
                             hide_index=True, width="stretch", key=key)
    s1, s2 = st.columns(2)
    if s1.button("💾 Guardar cambios en el histórico", type="primary", disabled=not hay_cambios(key), width="stretch"):
        guardar_cambios(crudo, editado, "ARCHIVADO")
        ss.ver_maestro = ss.get("ver_maestro", 0) + 1
        ss.flash = "Histórico actualizado."
        st.rerun()
    libro = excel_bytes({"Registros": crudo[CABECERAS], "Por aeronave": resumen(f, "MATRICULA"),
                         "Por ruta": resumen(f, "RUTA"), "Por piloto": resumen(f, "PILOTO"), "Por mes": resumen(f, "MES")})
    s2.download_button("⬇️ Descargar Excel del periodo", libro, f"Moon_Historico_{ini:%Y%m%d}_{fin:%Y%m%d}.xlsx",
                       "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", width="stretch")


def tab_auditoria():
    todo = cargar()
    if todo.empty:
        vacio("🛡️", "Nada que auditar todavía", "Los hallazgos aparecerán cuando haya tiquetes registrados.")
        return
    h = hallazgos(todo)
    k = st.columns(4)
    cuenta = h["SEVERIDAD"].value_counts()
    with k[0]: kpi("Tiquetes revisados", f"{len(todo):,}")
    with k[1]: kpi("Hallazgos graves", f"{cuenta.get('Alta', 0)}", tono="alerta" if cuenta.get("Alta", 0) else "bien")
    with k[2]: kpi("Hallazgos medios", f"{cuenta.get('Media', 0)}")
    with k[3]: kpi("Hallazgos leves", f"{cuenta.get('Baja', 0)}")

    titulo("Revisión automática", "Duplicados, soportes faltantes, volúmenes fuera de lo normal y datos incompletos.")
    if h.empty:
        st.success("Sin hallazgos: todos los tiquetes pasan las validaciones.")
        return
    nivel = st.segmented_control("Severidad", ["Todos", "Alta", "Media", "Baja"], default="Todos", key="sev") or "Todos"
    vista = h if nivel == "Todos" else h[h["SEVERIDAD"] == nivel]
    st.dataframe(vista, hide_index=True, width="stretch", column_config={
        "GALONES": st.column_config.NumberColumn(format="%.1f"), "FECHA": st.column_config.DateColumn(format="DD/MM/YYYY")})
    st.download_button("⬇️ Descargar hallazgos", excel_bytes({"Hallazgos": vista}), f"Moon_Auditoria_{datetime.now():%Y%m%d}.xlsx",
                       "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    with st.expander("🔎 Ver el comprobante de un hallazgo"):
        visor_tiquete(todo[todo["ID"].isin(vista["ID"])], "visor_audit")


def vista_directivo():
    html('<h1 class="page-title">Panel directivo</h1>')
    if not ss.get("moon_auth"):
        login()
        return
    if ss.get("flash"):
        st.toast(ss.pop("flash"), icon="✅")
    hero("Centro de mando administrativo.",
         "Revisa los tiquetes del día, fija precios, cierra la jornada y analiza el consumo de toda la flota.")
    t1, t2, t3 = st.tabs(["📥 Bandeja del día", "📈 Analítica histórica", "🛡️ Auditoría"])
    with t1: tab_diario()
    with t2: tab_maestro()
    with t3: tab_auditoria()


# ============================================================
# Navegación
# ============================================================
init_db()
with st.sidebar:
    html('<div class="brand">🌙 Moon Aerolínea</div><div class="brand-sub">Fuel Intelligence</div>')
    st.markdown("---")
    vista = st.radio("Navegación", ["✈️ Registro en pista", "📊 Panel directivo"], label_visibility="collapsed")
    st.markdown("---")
    html(f'<div class="side-stat">Pendientes de revisión: <b>{contar("PENDIENTE")}</b></div>'
         f'<div class="side-stat">{datetime.now():%d/%m/%Y}</div>')
    if vista == "📊 Panel directivo" and ss.get("moon_auth"):
        if st.button("Cerrar sesión", width="stretch"):
            ss.moon_auth = False
            st.rerun()

vista_pista() if vista == "✈️ Registro en pista" else vista_directivo()
