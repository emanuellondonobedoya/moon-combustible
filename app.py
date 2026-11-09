# -*- coding: utf-8 -*-
"""MOONFLIGHT | Fuel Control
Aplicación Streamlit para captura, revisión y análisis de comprobantes de combustible.
Persistencia: SQLite local (configurable mediante MOON_DATA_DIR).
IA opcional: Anthropic Messages API (ANTHROPIC_API_KEY).
"""

from __future__ import annotations

import io
import json
import os
import re
import sqlite3
import hmac
import hashlib
import base64
from contextlib import contextmanager
from datetime import date, datetime
from pathlib import Path

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

# ============================================================
# CONFIGURACIÓN E IDENTIDAD CORPORATIVA
# ============================================================

st.set_page_config(
    page_title="Moonflight | Fuel Control",
    page_icon="✈️",
    layout="wide",
    initial_sidebar_state="expanded",
)

NAVY = "#10243A"
NAVY_2 = "#173B5E"
BLUE = "#3977A8"
BLUE_LIGHT = "#EAF2F8"
GOLD = "#C9A66B"
INK = "#24364A"
MUTED = "#718096"
BG = "#F4F7FA"
WHITE = "#FFFFFF"
GREEN = "#21856A"
AMBER = "#B7791F"
RED = "#C44D56"

st.markdown(
    f"""
    <style>
      @import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700&family=Manrope:wght@500;600;700;800&display=swap');
      html, body, [class*="css"] {{ font-family: 'DM Sans', sans-serif; color: {INK}; }}
      .stApp {{ background: {BG}; }}
      [data-testid="stSidebar"] {{ background: #0D2034; }}
      [data-testid="stSidebar"] * {{ color: #F4F7FB; }}
      [data-testid="stSidebar"] hr {{ border-color: rgba(255,255,255,.16); }}
      .brand-kicker {{ color: {GOLD}; letter-spacing: 2.1px; font-size: 10px; font-weight: 800; text-transform: uppercase; }}
      .page-title {{ color: {NAVY}; font-family: Manrope,sans-serif; font-size: clamp(28px,3vw,38px); font-weight: 800; letter-spacing: -1.2px; margin: 0 0 3px 0; }}
      .page-subtitle {{ color: {MUTED}; font-size: 14px; margin: 0 0 18px 0; }}
      .hero {{ background: linear-gradient(115deg,{NAVY} 0%,{NAVY_2} 72%,#235477 100%); padding: 26px 30px; border-radius: 19px; margin: 10px 0 23px 0; box-shadow: 0 12px 30px rgba(16,36,58,.13); position: relative; overflow: hidden; }}
      .hero:after {{ content:'M'; position:absolute; right:30px; top:-20px; color:rgba(255,255,255,.055); font-family:Manrope,sans-serif; font-size:150px; font-weight:800; }}
      .hero-label {{ color: {GOLD}; font-size:10px; letter-spacing:2px; text-transform:uppercase; font-weight:800; }}
      .hero-title {{ color:white; font-family:Manrope,sans-serif; font-size:25px; line-height:1.2; font-weight:800; margin-top:8px; }}
      .hero-copy {{ color:#D0DCE7; font-size:13px; line-height:1.6; max-width:720px; margin-top:8px; }}
      .section-title {{ color:{NAVY}; font-family:Manrope,sans-serif; font-weight:800; font-size:17px; margin:17px 0 8px 0; }}
      .metric {{ background:{WHITE}; border:1px solid #E4EAF0; border-radius:15px; padding:17px 18px; min-height:105px; box-shadow:0 5px 18px rgba(16,36,58,.035); }}
      .metric-label {{ color:{MUTED}; font-size:10px; font-weight:800; letter-spacing:1px; text-transform:uppercase; }}
      .metric-value {{ color:{NAVY}; font-family:Manrope,sans-serif; font-size:25px; line-height:1.2; font-weight:800; margin-top:9px; }}
      .metric-note {{ color:{MUTED}; font-size:11px; margin-top:5px; }}
      .small-note {{ color:{MUTED}; font-size:11px; }}
      .stButton > button, .stFormSubmitButton > button {{ border-radius:10px; border:0; background:{NAVY}; color:white; font-weight:700; min-height:42px; }}
      .stButton > button:hover, .stFormSubmitButton > button:hover {{ background:{NAVY_2}; color:white; border:0; }}
      div[data-testid="stFileUploader"] section {{ border:1.5px dashed #9FB3C7; border-radius:13px; background:white; }}
      div[data-testid="stFileUploader"] section:hover {{ border-color:{BLUE}; }}
      div[data-testid="stTabs"] button {{ font-weight:700; }}
      div[data-testid="stDataFrame"] {{ border:1px solid #E4EAF0; border-radius:12px; overflow:hidden; }}
    </style>
    """,
    unsafe_allow_html=True,
)

# ============================================================
# BASE DE DATOS
# ============================================================

DATA_DIR = Path(os.environ.get("MOON_DATA_DIR", ".")).expanduser().resolve()
DATA_DIR.mkdir(parents=True, exist_ok=True)
DB_PATH = DATA_DIR / "moonflight_fuel.sqlite3"

LEGACY_DAILY = Path("Registro_Diario_Moon.xlsx")
LEGACY_MASTER = Path("Base_Datos_Maestra_Moon.xlsx")

FIELDS = [
    "FECHA REGISTRO", "RECIBO", "FECHA CONSUMO", "MATRICULA",
    "PRODUCTO", "RUTA", "CONSUMO (GAL)", "VALOR UNIT",
    "VALOR TOTAL", "PILOTO", "PROVEEDOR", "ESTACIÓN",
    "HORAS VUELO", "TRAMO / VUELO", "NOMENCLATURA IMAGEN",
    "OBSERVACION", "ESTADO", "VALIDADO POR",
]
NUMERIC_FIELDS = ["CONSUMO (GAL)", "VALOR UNIT", "VALOR TOTAL", "HORAS VUELO"]


def db_connect():
    conn = sqlite3.connect(DB_PATH, timeout=30, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA foreign_keys=ON")
    conn.execute("PRAGMA busy_timeout=30000")
    return conn


@contextmanager
def db_session():
    conn = db_connect()
    try:
        yield conn
    finally:
        conn.close()


def init_db():
    definitions = [
        f'"{field}" {"REAL" if field in NUMERIC_FIELDS else "TEXT"}'
        for field in FIELDS
    ]
    with db_session() as conn:
        conn.execute(
            "CREATE TABLE IF NOT EXISTS tickets ("
            "id INTEGER PRIMARY KEY AUTOINCREMENT, "
            + ", ".join(definitions)
            + ", image BLOB, image_mime TEXT, created_at TEXT NOT NULL)"
        )
        conn.execute(
            """CREATE TABLE IF NOT EXISTS app_meta (
                key TEXT PRIMARY KEY,
                value TEXT
            )"""
        )
        conn.commit()


def clean_text(value):
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except (TypeError, ValueError):
        pass
    return str(value).strip()


def to_float(value):
    if value is None:
        return None
    if isinstance(value, float) and pd.isna(value):
        return None
    if str(value).strip() == "":
        return None
    try:
        if isinstance(value, str):
            value = value.replace("$", "").replace("COP", "").replace(" ", "").replace(",", ".")
        return float(value)
    except (TypeError, ValueError):
        return None


def normalize_date(value):
    raw = clean_text(value)
    if not raw:
        return ""
    try:
        if re.fullmatch(r"\d{4}-\d{2}-\d{2}(?:[ T].*)?", raw):
            return datetime.strptime(raw[:10], "%Y-%m-%d").date().isoformat()
        return pd.to_datetime(value, dayfirst=True, errors="raise").date().isoformat()
    except Exception:
        return raw[:30]


def row_fingerprint(record):
    return (
        clean_text(record.get("RECIBO")).upper().replace(" ", ""),
        normalize_date(record.get("FECHA CONSUMO")),
        clean_text(record.get("MATRICULA")).upper().replace(" ", ""),
    )


def existing_fingerprint(fp):
    if not any(fp):
        return False
    with db_session() as conn:
        rows = conn.execute('SELECT "RECIBO", "FECHA CONSUMO", "MATRICULA" FROM tickets').fetchall()
    return any(row_fingerprint(dict(row)) == fp for row in rows)


def insert_ticket(record, image=None, image_mime=None):
    values = {}
    for field in FIELDS:
        values[field] = to_float(record.get(field)) if field in NUMERIC_FIELDS else clean_text(record.get(field))
    values["FECHA REGISTRO"] = values["FECHA REGISTRO"] or datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    values["FECHA CONSUMO"] = normalize_date(values["FECHA CONSUMO"])
    values["ESTADO"] = values["ESTADO"] or "Pendiente"
    cols = ", ".join(f'"{field}"' for field in FIELDS)
    placeholders = ", ".join(["?"] * (len(FIELDS) + 3))
    with db_session() as conn:
        cur = conn.execute(
            f"INSERT INTO tickets ({cols}, image, image_mime, created_at) VALUES ({placeholders})",
            [values[field] for field in FIELDS] + [image, image_mime, datetime.now().isoformat(timespec="seconds")],
        )
        conn.commit()
        return int(cur.lastrowid)


def load_tickets():
    with db_session() as conn:
        df = pd.read_sql_query(
            "SELECT id AS ID, " + ", ".join(f'"{field}"' for field in FIELDS) + " FROM tickets ORDER BY id DESC",
            conn,
        )
    for field in NUMERIC_FIELDS:
        df[field] = pd.to_numeric(df[field], errors="coerce")
    return df


def get_ticket_image(ticket_id):
    with db_session() as conn:
        row = conn.execute("SELECT image, image_mime FROM tickets WHERE id=?", (ticket_id,)).fetchone()
    if not row:
        return None, None
    return row["image"], row["image_mime"]


def save_editor_changes(edited):
    with db_session() as conn:
        for _, row in edited.iterrows():
            ticket_id = int(row["ID"])
            assignments = ", ".join(f'"{field}"=?' for field in FIELDS)
            vals = []
            for field in FIELDS:
                if field in NUMERIC_FIELDS:
                    vals.append(to_float(row.get(field)))
                elif field == "FECHA CONSUMO":
                    vals.append(normalize_date(row.get(field)))
                else:
                    vals.append(clean_text(row.get(field)))
            vals.append(ticket_id)
            conn.execute(f"UPDATE tickets SET {assignments} WHERE id=?", vals)
        conn.commit()


def delete_ticket(ticket_id):
    with db_session() as conn:
        conn.execute("DELETE FROM tickets WHERE id=?", (ticket_id,))
        conn.commit()


def migrate_legacy_excel_once():
    with db_session() as conn:
        done = conn.execute("SELECT value FROM app_meta WHERE key='legacy_import_done'").fetchone()
        if done:
            return 0
    imported = 0
    for path in (LEGACY_MASTER, LEGACY_DAILY):
        if not path.exists():
            continue
        try:
            old = pd.read_excel(path)
        except Exception:
            continue
        for _, row in old.iterrows():
            record = {field: row.get(field, "") for field in FIELDS}
            record["FECHA CONSUMO"] = normalize_date(record.get("FECHA CONSUMO"))
            record["ESTADO"] = record.get("ESTADO") or "Pendiente"
            if existing_fingerprint(row_fingerprint(record)):
                continue
            insert_ticket(record)
            imported += 1
    with db_session() as conn:
        conn.execute("INSERT OR REPLACE INTO app_meta(key,value) VALUES('legacy_import_done','1')")
        conn.commit()
    return imported


init_db()
legacy_imported = migrate_legacy_excel_once()

# ============================================================
# UTILIDADES DE PRESENTACIÓN Y ANÁLISIS
# ============================================================

def secret_value(key, default=""):
    try:
        return str(st.secrets.get(key, default) or default)
    except Exception:
        return os.environ.get(key, default)


def metric_card(label, value, note=""):
    st.markdown(
        f"""<div class="metric">
        <div class="metric-label">{label}</div>
        <div class="metric-value">{value}</div>
        <div class="metric-note">{note}</div></div>""",
        unsafe_allow_html=True,
    )


def hero(title, copy, label="MOONFLIGHT · FUEL CONTROL"):
    st.markdown(
        f"""<div class="hero">
        <div class="hero-label">{label}</div>
        <div class="hero-title">{title}</div>
        <div class="hero-copy">{copy}</div></div>""",
        unsafe_allow_html=True,
    )


def money(value):
    number = to_float(value) or 0
    return "$ " + f"{number:,.0f}".replace(",", ".")


def style_chart(fig, height=340):
    fig.update_layout(
        height=height,
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        font=dict(family="DM Sans", color=MUTED, size=11),
        title=dict(font=dict(family="Manrope", size=15, color=NAVY), x=0.02, xanchor="left"),
        margin=dict(l=18, r=18, t=58, b=25),
        legend=dict(orientation="h", yanchor="bottom", y=1.02, x=0, font=dict(size=10)),
        hoverlabel=dict(bgcolor=NAVY, font=dict(color="white", family="DM Sans")),
    )
    fig.update_xaxes(showgrid=False, linecolor="#E7EDF3", tickfont=dict(color=MUTED))
    fig.update_yaxes(showgrid=True, gridcolor="#E9EEF3", zeroline=False, tickfont=dict(color=MUTED))
    return fig


def df_to_excel_bytes(df, sheet_name="Moonflight"):
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.drop(columns=["ID"], errors="ignore").to_excel(writer, index=False, sheet_name=sheet_name[:31])
        ws = writer.sheets[sheet_name[:31]]
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
        for column_cells in ws.columns:
            max_len = max((len(str(cell.value)) if cell.value is not None else 0) for cell in column_cells)
            ws.column_dimensions[column_cells[0].column_letter].width = min(max(max_len + 2, 12), 30)
    return buffer.getvalue()


def prepare_numeric(df):
    data = df.copy()
    for col in NUMERIC_FIELDS:
        if col in data:
            data[col] = pd.to_numeric(data[col], errors="coerce")
    data["FECHA_DT"] = pd.to_datetime(data.get("FECHA CONSUMO"), errors="coerce")
    data["DÍA"] = data["FECHA_DT"].dt.strftime("%Y-%m-%d")
    data["MES"] = data["FECHA_DT"].dt.to_period("M").astype(str)
    data["SEMANA"] = data["FECHA_DT"].dt.to_period("W").astype(str).str[:10]
    data["COSTO CALCULADO"] = data["CONSUMO (GAL)"].fillna(0) * data["VALOR UNIT"].fillna(0)
    return data


def filter_data(df):
    if df.empty:
        return df.copy()
    data = prepare_numeric(df)
    dates = data["FECHA_DT"].dropna()
    with st.container(border=True):
        st.markdown("**Filtros del análisis**")
        f1, f2, f3 = st.columns([1.25, 1, 1])
        min_day = dates.min().date() if len(dates) else date.today()
        max_day = dates.max().date() if len(dates) else date.today()
        date_range = f1.date_input("Periodo de consumo", value=(min_day, max_day), key="filter_dates")
        aircrafts = sorted([x for x in data["MATRICULA"].dropna().astype(str).unique() if x.strip()])
        routes = sorted([x for x in data["RUTA"].dropna().astype(str).unique() if x.strip()])
        pilots = sorted([x for x in data["PILOTO"].dropna().astype(str).unique() if x.strip()])
        selected_aircraft = f2.multiselect("Aeronave", aircrafts, key="filter_aircraft")
        selected_route = f3.multiselect("Ruta", routes, key="filter_route")
        f4, f5, f6 = st.columns(3)
        selected_pilot = f4.multiselect("Piloto", pilots, key="filter_pilot")
        statuses = sorted([x for x in data["ESTADO"].dropna().astype(str).unique() if x.strip()])
        selected_status = f5.multiselect("Estado", statuses, key="filter_status")
        suppliers = sorted([x for x in data["PROVEEDOR"].dropna().astype(str).unique() if x.strip()])
        selected_supplier = f6.multiselect("Proveedor", suppliers, key="filter_supplier")
        f7, f8 = st.columns(2)
        stations = sorted([x for x in data["ESTACIÓN"].dropna().astype(str).unique() if x.strip()])
        products = sorted([x for x in data["PRODUCTO"].dropna().astype(str).unique() if x.strip()])
        selected_station = f7.multiselect("Estación / aeropuerto", stations, key="filter_station")
        selected_product = f8.multiselect("Producto", products, key="filter_product")
    if isinstance(date_range, tuple) and len(date_range) == 2:
        start, end = date_range
        data = data[(data["FECHA_DT"].dt.date >= start) & (data["FECHA_DT"].dt.date <= end)]
    elif isinstance(date_range, date):
        data = data[data["FECHA_DT"].dt.date == date_range]
    filters = [
        ("MATRICULA", selected_aircraft), ("RUTA", selected_route),
        ("PILOTO", selected_pilot), ("ESTADO", selected_status),
        ("PROVEEDOR", selected_supplier), ("ESTACIÓN", selected_station),
        ("PRODUCTO", selected_product),
    ]
    for col, selected in filters:
        if selected:
            data = data[data[col].astype(str).isin(selected)]
    return data


def build_reconciliation(df):
    cols = ["ID", "RECIBO", "MATRICULA", "FECHA CONSUMO", "ALERTA"]
    if df.empty:
        return pd.DataFrame(columns=cols)
    data = prepare_numeric(df)
    issues = []
    dupes = data.duplicated(subset=["RECIBO", "FECHA CONSUMO", "MATRICULA"], keep=False)
    for _, row in data[dupes].iterrows():
        issues.append({"ID": row["ID"], "RECIBO": row["RECIBO"], "MATRICULA": row["MATRICULA"], "FECHA CONSUMO": row["FECHA CONSUMO"], "ALERTA": "Posible recibo duplicado"})
    required = ["RECIBO", "FECHA CONSUMO", "MATRICULA", "CONSUMO (GAL)", "PILOTO"]
    for _, row in data.iterrows():
        missing = [col for col in required if pd.isna(row.get(col)) or clean_text(row.get(col)) == ""]
        if missing:
            issues.append({"ID": row["ID"], "RECIBO": row.get("RECIBO"), "MATRICULA": row.get("MATRICULA"), "FECHA CONSUMO": row.get("FECHA CONSUMO"), "ALERTA": "Campos pendientes: " + ", ".join(missing)})
        gallons, unit, total = (to_float(row.get(c)) for c in ["CONSUMO (GAL)", "VALOR UNIT", "VALOR TOTAL"])
        if gallons is not None and gallons <= 0:
            issues.append({"ID": row["ID"], "RECIBO": row.get("RECIBO"), "MATRICULA": row.get("MATRICULA"), "FECHA CONSUMO": row.get("FECHA CONSUMO"), "ALERTA": "Consumo debe ser mayor que cero"})
        if unit is not None and unit < 0:
            issues.append({"ID": row["ID"], "RECIBO": row.get("RECIBO"), "MATRICULA": row.get("MATRICULA"), "FECHA CONSUMO": row.get("FECHA CONSUMO"), "ALERTA": "Precio unitario negativo"})
        if gallons is not None and unit is not None and total is not None and total > 0:
            expected = gallons * unit
            if abs(total - expected) > max(100, expected * 0.02):
                issues.append({"ID": row["ID"], "RECIBO": row.get("RECIBO"), "MATRICULA": row.get("MATRICULA"), "FECHA CONSUMO": row.get("FECHA CONSUMO"), "ALERTA": "El total no coincide con galones × precio unitario"})
    return pd.DataFrame(issues).drop_duplicates() if issues else pd.DataFrame(columns=cols)


def extract_receipt_with_ai(image_bytes, mime_type):
    api_key = secret_value("ANTHROPIC_API_KEY")
    if not api_key:
        raise RuntimeError("No está configurado ANTHROPIC_API_KEY. Puedes continuar con el registro manual.")
    try:
        import anthropic
    except ImportError as exc:
        raise RuntimeError("Falta instalar la dependencia anthropic. Agrégala a requirements.txt.") from exc
    model = secret_value("ANTHROPIC_MODEL", "claude-haiku-5-5")
    encoded = base64.b64encode(image_bytes).decode("utf-8")
    prompt = """
Eres un asistente de lectura documental para comprobantes de abastecimiento de combustible de una aerolínea.
Lee únicamente lo que sea visible en la imagen. No inventes ni completes datos por intuición.
Devuelve exclusivamente un objeto JSON válido, sin markdown, con estas claves:
{
 "recibo":"",
 "fecha_consumo":"YYYY-MM-DD o vacío",
 "matricula":"",
 "producto":"",
 "ruta":"",
 "consumo_galones":null,
 "valor_unit_cop":null,
 "valor_total_cop":null,
 "piloto":"",
 "proveedor":"",
 "estacion":"",
 "observacion":""
}
Convierte números a formato decimal numérico, sin símbolo monetario ni separador de miles.
Si el documento muestra litros y no galones, no conviertas: deja consumo_galones null y explica la unidad en observacion.
Si un dato no aparece o no se puede leer con confianza, deja el campo vacío o null.
La fecha debe ser la fecha del abastecimiento, no la fecha actual.
"""
    client = anthropic.Anthropic(api_key=api_key)
    response = client.messages.create(
        model=model,
        max_tokens=1200,
        messages=[{"role": "user", "content": [
            {"type": "image", "source": {"type": "base64", "media_type": mime_type, "data": encoded}},
            {"type": "text", "text": prompt},
        ]}],
    )
    text_output = "\n".join(block.text for block in response.content if getattr(block, "type", None) == "text")
    text_output = re.sub(r"^```(?:json)?\s*|\s*```$", "", text_output.strip(), flags=re.IGNORECASE)
    try:
        result = json.loads(text_output)
    except json.JSONDecodeError as exc:
        raise RuntimeError("La IA no devolvió un resultado legible. Completa los datos manualmente.") from exc
    if not isinstance(result, dict):
        raise RuntimeError("La lectura no produjo un registro válido. Completa los datos manualmente.")
    return result


# ============================================================
# BARRA LATERAL Y NAVEGACIÓN
# ============================================================

with st.sidebar:
    st.markdown("<div class='brand-kicker'>REGIONAL AVIATION · COLOMBIA</div>", unsafe_allow_html=True)
    st.markdown("# MOONFLIGHT")
    st.caption("FUEL CONTROL CENTER")
    st.divider()
    view = st.radio("Módulos", ["Registro de abastecimiento", "Centro de control"], label_visibility="collapsed")
    st.divider()
    st.caption("Gestión y trazabilidad de combustible")
    st.caption("Versión 2.0 · Piloto operativo")
    if st.button("↻ Actualizar vista", use_container_width=True):
        st.rerun()


# ============================================================
# MÓDULO 1: REGISTRO DE ABASTECIMIENTO
# ============================================================

if view == "Registro de abastecimiento":
    st.markdown('<h1 class="page-title">Registro de abastecimiento</h1>', unsafe_allow_html=True)
    st.markdown('<p class="page-subtitle">Captura digital del comprobante y validación inicial de los datos del suministro.</p>', unsafe_allow_html=True)
    hero("Cada suministro, con trazabilidad.", "Carga el comprobante del proveedor, utiliza la lectura asistida por IA y confirma los datos antes de enviarlos al equipo administrativo.")
    col_upload, col_info = st.columns([1.15, 0.85], gap="large")
    with col_upload:
        st.markdown('<div class="section-title">01 · Comprobante</div>', unsafe_allow_html=True)
        uploaded = st.file_uploader("Adjuntar comprobante (JPG o PNG)", type=["jpg", "jpeg", "png"], accept_multiple_files=False, help="Carga una imagen clara, completa y legible. Tamaño máximo: 5 MB.")
        if uploaded:
            if uploaded.size > 5 * 1024 * 1024:
                st.error("El archivo supera el límite de 5 MB. Reduce su tamaño y vuelve a cargarlo.")
                uploaded = None
            else:
                current_hash = hashlib.sha256(uploaded.getvalue()).hexdigest()
                previous_hash = st.session_state.get("moon_ocr_hash")
                if previous_hash and previous_hash != current_hash:
                    for key in ("moon_ocr", "moon_ocr_name", "moon_ocr_bytes", "moon_ocr_mime"):
                        st.session_state.pop(key, None)
                st.session_state["moon_ocr_hash"] = current_hash
                st.image(uploaded, caption="Vista previa del comprobante", use_container_width=True)
        ai_col, manual_col = st.columns(2)
        with ai_col:
            run_ai = st.button("✦ Leer con inteligencia artificial", disabled=uploaded is None, use_container_width=True)
        with manual_col:
            manual = st.button("Ingresar manualmente", use_container_width=True)
    with col_info:
        st.markdown('<div class="section-title">02 · Flujo de registro</div>', unsafe_allow_html=True)
        st.markdown("""
        <div class="metric" style="min-height:0;margin-bottom:10px"><b>1. Captura</b><div class="small-note">Fotografía completa del comprobante.</div></div>
        <div class="metric" style="min-height:0;margin-bottom:10px"><b>2. Extracción</b><div class="small-note">Lectura asistida; los datos siempre deben revisarse.</div></div>
        <div class="metric" style="min-height:0;margin-bottom:10px"><b>3. Confirmación</b><div class="small-note">El usuario corrige y confirma antes de guardar.</div></div>
        <div class="metric" style="min-height:0"><b>4. Conciliación</b><div class="small-note">El registro queda pendiente de validación administrativa.</div></div>
        """, unsafe_allow_html=True)
        st.info("La lectura automática es una ayuda de captura, no una validación contable ni una autorización de pago.")
    if run_ai and uploaded is not None:
        st.session_state.pop("moon_ocr", None)
        try:
            with st.spinner("Leyendo el comprobante…"):
                st.session_state["moon_ocr"] = extract_receipt_with_ai(uploaded.getvalue(), uploaded.type or "image/jpeg")
                st.session_state["moon_ocr_name"] = uploaded.name
                st.session_state["moon_ocr_bytes"] = uploaded.getvalue()
                st.session_state["moon_ocr_mime"] = uploaded.type or "image/jpeg"
            st.success("Lectura completada. Revisa cada campo antes de guardar.")
        except Exception as exc:
            st.error(str(exc))
    if manual:
        st.session_state["moon_ocr"] = {}
        st.session_state["moon_ocr_name"] = uploaded.name if uploaded else ""
        st.session_state["moon_ocr_bytes"] = uploaded.getvalue() if uploaded else None
        st.session_state["moon_ocr_mime"] = uploaded.type if uploaded else None
    ocr = st.session_state.get("moon_ocr")
    if ocr is not None:
        st.markdown('<div class="section-title">03 · Revisión y confirmación del registro</div>', unsafe_allow_html=True)
        st.caption("Verifica la información contra el comprobante. Los campos no identificados por la IA deben completarse manualmente.")
        def ocr_text(key, fallback=""):
            return clean_text(ocr.get(key, fallback))
        raw_date = normalize_date(ocr.get("fecha_consumo"))
        try:
            default_date = pd.to_datetime(raw_date).date() if raw_date else date.today()
        except Exception:
            default_date = date.today()
        unit = to_float(ocr.get("valor_unit_cop"))
        total = to_float(ocr.get("valor_total_cop"))
        gallons = to_float(ocr.get("consumo_galones"))
        with st.form("confirm_receipt", clear_on_submit=False):
            a1, a2, a3 = st.columns(3)
            receipt_no = a1.text_input("Número de recibo / factura *", value=ocr_text("recibo"))
            consumption_date = a2.date_input("Fecha de abastecimiento *", value=default_date, format="DD/MM/YYYY")
            aircraft = a3.text_input("Matrícula de aeronave *", value=ocr_text("matricula")).upper()
            b1, b2, b3 = st.columns(3)
            product = b1.text_input("Producto / combustible", value=ocr_text("producto"))
            route = b2.text_input("Ruta / destino", value=ocr_text("ruta"))
            pilot = b3.text_input("Piloto responsable *", value=ocr_text("piloto"))
            c1, c2, c3 = st.columns(3)
            gallons_value = c1.number_input("Consumo (galones) *", min_value=0.0, value=max(gallons or 0.0, 0.0), step=1.0, format="%.2f")
            unit_value = c2.number_input("Precio unitario (COP/gal)", min_value=0.0, value=max(unit or 0.0, 0.0), step=100.0, format="%.2f")
            total_value = c3.number_input("Valor total (COP)", min_value=0.0, value=max(total or 0.0, 0.0), step=1000.0, format="%.2f")
            d1, d2, d3 = st.columns(3)
            supplier = d1.text_input("Proveedor", value=ocr_text("proveedor"))
            station = d2.text_input("Estación / aeropuerto", value=ocr_text("estacion"))
            flight_hours = d3.number_input("Horas de vuelo asociadas (opcional)", min_value=0.0, value=0.0, step=0.5, format="%.1f")
            e1, e2 = st.columns(2)
            leg = e1.text_input("Número de vuelo / tramo (opcional)")
            observation = e2.text_input("Observaciones", value=ocr_text("observacion"))
            confirm = st.form_submit_button("Guardar comprobante en bandeja administrativa", use_container_width=True)
        if confirm:
            image_bytes = st.session_state.get("moon_ocr_bytes")
            if not image_bytes and uploaded is not None:
                image_bytes = uploaded.getvalue()
            missing = []
            if not receipt_no.strip(): missing.append("número de recibo")
            if not aircraft.strip(): missing.append("matrícula")
            if not pilot.strip(): missing.append("piloto")
            if gallons_value <= 0: missing.append("consumo en galones")
            if not image_bytes: missing.append("imagen del comprobante")
            if missing:
                st.error("Completa los campos obligatorios: " + ", ".join(missing) + ".")
            else:
                record = {
                    "FECHA REGISTRO": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    "RECIBO": receipt_no,
                    "FECHA CONSUMO": consumption_date.isoformat(),
                    "MATRICULA": aircraft.strip().upper(),
                    "PRODUCTO": product,
                    "RUTA": route,
                    "CONSUMO (GAL)": gallons_value,
                    "VALOR UNIT": unit_value,
                    "VALOR TOTAL": total_value,
                    "PILOTO": pilot,
                    "PROVEEDOR": supplier,
                    "ESTACIÓN": station,
                    "HORAS VUELO": flight_hours or None,
                    "TRAMO / VUELO": leg,
                    "NOMENCLATURA IMAGEN": st.session_state.get("moon_ocr_name", uploaded.name if uploaded else "comprobante"),
                    "OBSERVACION": observation,
                    "ESTADO": "Pendiente",
                    "VALIDADO POR": "",
                }
                if existing_fingerprint(row_fingerprint(record)):
                    st.error("Ya existe un registro con el mismo número de recibo, fecha y matrícula. Revisa el histórico antes de volver a enviarlo.")
                else:
                    try:
                        ticket_id = insert_ticket(record, image_bytes, st.session_state.get("moon_ocr_mime", "image/jpeg"))
                        st.success(f"Registro #{ticket_id} guardado y enviado a revisión administrativa.")
                        for key in ("moon_ocr", "moon_ocr_name", "moon_ocr_bytes", "moon_ocr_mime", "moon_ocr_hash"):
                            st.session_state.pop(key, None)
                        st.rerun()
                    except Exception as exc:
                        st.error(f"No fue posible guardar el registro: {exc}")
    recent = load_tickets().head(5)
    if not recent.empty:
        st.markdown('<div class="section-title">Últimos registros recibidos</div>', unsafe_allow_html=True)
        st.dataframe(recent[["ID", "FECHA CONSUMO", "RECIBO", "MATRICULA", "CONSUMO (GAL)", "ESTADO"]], use_container_width=True, hide_index=True)


# ============================================================
# MÓDULO 2: CENTRO DE CONTROL DIRECTIVO
# ============================================================

else:
    st.markdown('<h1 class="page-title">Centro de control</h1>', unsafe_allow_html=True)
    st.markdown('<p class="page-subtitle">Supervisión administrativa, conciliación de comprobantes e inteligencia de consumo.</p>', unsafe_allow_html=True)
    admin_password = secret_value("ADMIN_PASSWORD")
    if not st.session_state.get("moon_admin_auth", False):
        hero("Control operativo y financiero.", "Acceso restringido para revisión de comprobantes, validación de costos y análisis histórico.", "MOONFLIGHT · MANAGEMENT ACCESS")
        if not admin_password:
            st.error("El panel directivo está bloqueado hasta configurar ADMIN_PASSWORD en los Secrets del despliegue. No existe una contraseña predeterminada.")
        else:
            with st.form("admin_login"):
                password_entry = st.text_input("Contraseña de acceso", type="password")
                login = st.form_submit_button("Ingresar al centro de control", use_container_width=True)
            if login:
                if hmac.compare_digest(password_entry, admin_password):
                    st.session_state["moon_admin_auth"] = True
                    st.rerun()
                else:
                    st.error("Credenciales incorrectas.")
    else:
        top1, top2 = st.columns([5, 1])
        with top1:
            hero("Visibilidad sobre cada galón.", "Una sola vista para revisar abastecimientos, identificar diferencias y comprender el comportamiento del consumo de combustible.", "MOONFLIGHT · FUEL INTELLIGENCE")
        with top2:
            if st.button("Cerrar sesión", use_container_width=True):
                st.session_state["moon_admin_auth"] = False
                st.rerun()
        df_all = load_tickets()
        if legacy_imported:
            st.info(f"Se importaron {legacy_imported} registros de los archivos Excel de la versión anterior.")
        if df_all.empty:
            st.info("Aún no hay comprobantes registrados. Los nuevos envíos aparecerán aquí.")
        else:
            filtered = filter_data(df_all)
            data = prepare_numeric(filtered)
            total_gallons = data["CONSUMO (GAL)"].fillna(0).sum()
            total_cost = data["VALOR TOTAL"].fillna(0).sum()
            weighted_unit = total_cost / total_gallons if total_gallons else 0
            pending = (data["ESTADO"].astype(str).str.lower() == "pendiente").sum()
            validated = (data["ESTADO"].astype(str).str.lower() == "validado").sum()
            k1, k2, k3, k4 = st.columns(4)
            with k1: metric_card("Comprobantes", f"{len(data):,}", "Registros en el periodo")
            with k2: metric_card("Combustible", f"{total_gallons:,.1f} gal", "Volumen registrado")
            with k3: metric_card("Gasto registrado", money(total_cost), "Suma de valores reportados")
            with k4: metric_card("Costo ponderado", money(weighted_unit), "COP por galón · gasto / galones")
            st.caption(f"Pendientes de revisión: {pending:,} · Comprobantes validados: {validated:,}")
            tab_ops, tab_analysis, tab_review, tab_export = st.tabs(["Resumen operativo", "Análisis de flota", "Revisión y conciliación", "Exportar información"])

            with tab_ops:
                st.markdown('<div class="section-title">Evolución del consumo y del gasto</div>', unsafe_allow_html=True)
                if not data.empty and data["FECHA_DT"].notna().any():
                    trend = data.dropna(subset=["FECHA_DT"]).groupby("DÍA", as_index=False).agg(galones=("CONSUMO (GAL)", "sum"), gasto=("VALOR TOTAL", "sum"))
                    fig = go.Figure()
                    fig.add_trace(go.Bar(x=trend["DÍA"], y=trend["galones"], name="Galones", marker_color=BLUE, opacity=.88, yaxis="y"))
                    fig.add_trace(go.Scatter(x=trend["DÍA"], y=trend["gasto"], name="Gasto COP", mode="lines+markers", line=dict(color=GOLD, width=3), marker=dict(size=7), yaxis="y2"))
                    fig.update_layout(title="Consumo diario vs. gasto registrado", yaxis=dict(title="Galones", showgrid=True, gridcolor="#E9EEF3"), yaxis2=dict(title="COP", overlaying="y", side="right", showgrid=False), barmode="group")
                    st.plotly_chart(style_chart(fig, 370), use_container_width=True)
                else:
                    st.info("No hay fechas válidas suficientes para construir la tendencia.")
                left, right = st.columns(2)
                with left:
                    by_aircraft = data.groupby("MATRICULA", as_index=False)["CONSUMO (GAL)"].sum().sort_values("CONSUMO (GAL)", ascending=True).tail(12)
                    fig_air = px.bar(by_aircraft, x="CONSUMO (GAL)", y="MATRICULA", orientation="h", text_auto=".1f", title="Volumen por aeronave", color_discrete_sequence=[NAVY_2], labels={"CONSUMO (GAL)": "Galones", "MATRICULA": "Matrícula"})
                    st.plotly_chart(style_chart(fig_air), use_container_width=True)
                with right:
                    by_route = data.groupby("RUTA", as_index=False)["VALOR TOTAL"].sum().sort_values("VALOR TOTAL", ascending=False).head(10)
                    if by_route["RUTA"].replace("", pd.NA).notna().any():
                        fig_route = px.bar(by_route, x="RUTA", y="VALOR TOTAL", text_auto=".2s", title="Gasto registrado por ruta", color_discrete_sequence=[BLUE], labels={"VALOR TOTAL": "COP", "RUTA": "Ruta"})
                        st.plotly_chart(style_chart(fig_route), use_container_width=True)
                    else:
                        st.info("Completa el campo de ruta para habilitar este análisis.")
                st.markdown('<div class="section-title">Bandeja de comprobantes</div>', unsafe_allow_html=True)
                editable_cols = ["ID"] + FIELDS
                editable = st.data_editor(
                    filtered[editable_cols], use_container_width=True, hide_index=True, num_rows="fixed", key="moon_editor_all",
                    disabled=["ID", "FECHA REGISTRO", "NOMENCLATURA IMAGEN"],
                    column_config={
                        "ID": st.column_config.NumberColumn("ID", width="small"),
                        "CONSUMO (GAL)": st.column_config.NumberColumn("Galones", format="%.2f"),
                        "VALOR UNIT": st.column_config.NumberColumn("COP/gal", format="$ %.2f"),
                        "VALOR TOTAL": st.column_config.NumberColumn("Total COP", format="$ %.0f"),
                        "HORAS VUELO": st.column_config.NumberColumn("Horas vuelo", format="%.1f"),
                        "ESTADO": st.column_config.SelectboxColumn("Estado", options=["Pendiente", "Validado", "Observado"], required=True),
                    },
                )
                if st.button("Guardar cambios de la bandeja", type="primary"):
                    try:
                        save_editor_changes(editable)
                        st.success("Cambios guardados en la base de datos.")
                        st.rerun()
                    except Exception as exc:
                        st.error(f"No fue posible guardar los cambios: {exc}")

            with tab_analysis:
                st.markdown('<div class="section-title">Composición del consumo y comportamiento de costos</div>', unsafe_allow_html=True)
                a1, a2 = st.columns(2)
                with a1:
                    product_mix = data.groupby("PRODUCTO", as_index=False)["CONSUMO (GAL)"].sum()
                    product_mix = product_mix[product_mix["PRODUCTO"].astype(str).str.strip() != ""]
                    if not product_mix.empty:
                        fig_mix = px.pie(product_mix, names="PRODUCTO", values="CONSUMO (GAL)", hole=.62, title="Consumo por tipo de producto", color_discrete_sequence=px.colors.sequential.Blues_r)
                        fig_mix.update_traces(textposition="inside", textinfo="percent+label")
                        st.plotly_chart(style_chart(fig_mix), use_container_width=True)
                    else:
                        st.info("No hay productos clasificados.")
                with a2:
                    status_mix = data.groupby("ESTADO", as_index=False).size().rename(columns={"size": "Comprobantes"})
                    fig_status = px.pie(status_mix, names="ESTADO", values="Comprobantes", hole=.62, title="Estado de revisión documental", color_discrete_sequence=[GOLD, GREEN, RED])
                    fig_status.update_traces(textposition="inside", textinfo="percent+label")
                    st.plotly_chart(style_chart(fig_status), use_container_width=True)
                b1, b2 = st.columns(2)
                with b1:
                    price_data = data[(data["CONSUMO (GAL)"] > 0) & (data["VALOR TOTAL"] > 0)].copy()
                    if not price_data.empty:
                        price_data["PRECIO REAL COP/GAL"] = price_data["VALOR TOTAL"] / price_data["CONSUMO (GAL)"]
                        price_trend = price_data.groupby("DÍA", as_index=False)["PRECIO REAL COP/GAL"].mean()
                        fig_price = px.line(price_trend, x="DÍA", y="PRECIO REAL COP/GAL", markers=True, title="Precio efectivo promedio por galón", color_discrete_sequence=[GOLD], labels={"DÍA": "Fecha", "PRECIO REAL COP/GAL": "COP/gal"})
                        st.plotly_chart(style_chart(fig_price), use_container_width=True)
                    else:
                        st.info("Registra galones y valores totales para analizar precios efectivos.")
                with b2:
                    pilot_data = data.groupby("PILOTO", as_index=False).agg(comprobantes=("ID", "count"), galones=("CONSUMO (GAL)", "sum")).sort_values("galones", ascending=False).head(10)
                    pilot_data = pilot_data[pilot_data["PILOTO"].astype(str).str.strip() != ""]
                    if not pilot_data.empty:
                        fig_pilot = px.bar(pilot_data, x="galones", y="PILOTO", orientation="h", text_auto=".1f", title="Volumen asociado a piloto", color_discrete_sequence=[BLUE], labels={"galones": "Galones", "PILOTO": "Piloto"})
                        st.plotly_chart(style_chart(fig_pilot), use_container_width=True)
                    else:
                        st.info("No hay datos de pilotos para comparar.")
                with st.expander("Indicador complementario: galones por hora de vuelo"):
                    flight = data[(data["HORAS VUELO"].fillna(0) > 0) & (data["CONSUMO (GAL)"].fillna(0) > 0)].copy()
                    if not flight.empty:
                        flight["GALONES / HORA"] = flight["CONSUMO (GAL)"] / flight["HORAS VUELO"]
                        by_flight = flight.groupby("MATRICULA", as_index=False)["GALONES / HORA"].mean().sort_values("GALONES / HORA", ascending=False)
                        fig_hour = px.bar(by_flight, x="MATRICULA", y="GALONES / HORA", title="Promedio de galones por hora registrada", color_discrete_sequence=[NAVY_2], labels={"MATRICULA": "Aeronave", "GALONES / HORA": "Gal/h"})
                        st.plotly_chart(style_chart(fig_hour), use_container_width=True)
                    else:
                        st.info("Este indicador requiere horas de vuelo asociadas. No debe interpretarse como eficiencia operacional sin conciliar horas, vuelos y carga de combustible.")

            with tab_review:
                st.markdown('<div class="section-title">Alertas de calidad y conciliación</div>', unsafe_allow_html=True)
                issues = build_reconciliation(df_all)
                pending_count = (df_all["ESTADO"].astype(str).str.lower() == "pendiente").sum()
                observed_count = (df_all["ESTADO"].astype(str).str.lower() == "observado").sum()
                r1, r2, r3 = st.columns(3)
                with r1: metric_card("Pendientes", f"{pending_count:,}", "Por validar")
                with r2: metric_card("Observados", f"{observed_count:,}", "Requieren seguimiento")
                with r3: metric_card("Alertas detectadas", f"{len(issues):,}", "Reglas documentales básicas")
                st.caption("Las alertas son controles de consistencia de datos. No reemplazan la revisión del comprobante, la validación tributaria ni la conciliación con el proveedor.")
                if issues.empty:
                    st.success("No se detectaron alertas con las reglas actuales.")
                else:
                    st.dataframe(issues, use_container_width=True, hide_index=True)
                    id_to_open = st.selectbox("Consultar comprobante por ID", options=sorted(issues["ID"].dropna().astype(int).unique().tolist(), reverse=True))
                    if id_to_open:
                        row = df_all[df_all["ID"] == id_to_open]
                        if not row.empty:
                            st.write(row[["ID", "RECIBO", "MATRICULA", "FECHA CONSUMO", "CONSUMO (GAL)", "VALOR UNIT", "VALOR TOTAL", "ESTADO"]])
                            img, mime = get_ticket_image(int(id_to_open))
                            if img:
                                st.image(img, caption=f"Comprobante asociado al registro #{id_to_open}", use_container_width=True)
                                st.download_button("Descargar imagen del comprobante", data=img, file_name=f"Moonflight_Recibo_{id_to_open}.{'png' if mime == 'image/png' else 'jpg'}", mime=mime or "image/jpeg")
                            else:
                                st.caption("Este registro proviene de una importación histórica y no tiene imagen adjunta.")
                with st.expander("Validar o cambiar estado de un comprobante"):
                    ids = df_all["ID"].astype(int).tolist()
                    chosen_id = st.selectbox("Registro", ids, format_func=lambda x: f"#{x} · " + clean_text(df_all.loc[df_all["ID"] == x, "RECIBO"].iloc[0]))
                    current_status = clean_text(df_all.loc[df_all["ID"] == chosen_id, "ESTADO"].iloc[0]) or "Pendiente"
                    status_options = ["Pendiente", "Validado", "Observado"]
                    default_idx = status_options.index(current_status) if current_status in status_options else 0
                    new_status = st.selectbox("Nuevo estado", status_options, index=default_idx)
                    validator = st.text_input("Responsable de validación")
                    review_note = st.text_input("Nota de revisión")
                    if st.button("Aplicar estado al comprobante"):
                        rec = df_all[df_all["ID"] == chosen_id].iloc[0].to_dict()
                        rec["ESTADO"] = new_status
                        rec["VALIDADO POR"] = validator
                        if review_note:
                            old_note = clean_text(rec.get("OBSERVACION"))
                            rec["OBSERVACION"] = (old_note + " | " + review_note).strip(" |")
                        save_editor_changes(pd.DataFrame([rec])[["ID"] + FIELDS])
                        st.success("Estado actualizado.")
                        st.rerun()
                with st.expander("Eliminar registro (acción administrativa)"):
                    st.warning("La eliminación es permanente en esta base de datos. Exporta un respaldo antes de continuar.")
                    delete_id = st.number_input("ID del registro a eliminar", min_value=1, step=1, value=1)
                    confirm_delete = st.checkbox("Confirmo que deseo eliminar permanentemente este registro")
                    if st.button("Eliminar registro seleccionado", disabled=not confirm_delete):
                        delete_ticket(int(delete_id))
                        st.success("Registro eliminado.")
                        st.rerun()

            with tab_export:
                st.markdown('<div class="section-title">Exportación y respaldo</div>', unsafe_allow_html=True)
                st.write("Descarga los registros filtrados o la base completa para conciliación, archivo y análisis externo.")
                exp1, exp2 = st.columns(2)
                with exp1:
                    st.download_button("Descargar selección filtrada (Excel)", data=df_to_excel_bytes(filtered, "Seleccion"), file_name=f"Moonflight_Combustible_Filtrado_{date.today().isoformat()}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
                with exp2:
                    st.download_button("Descargar histórico completo (Excel)", data=df_to_excel_bytes(df_all, "Historico"), file_name=f"Moonflight_Historico_Combustible_{date.today().isoformat()}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
                st.divider()
                st.markdown("**Resumen de la selección actual**")
                summary = pd.DataFrame([{
                    "Comprobantes": len(data),
                    "Galones": round(float(data["CONSUMO (GAL)"].fillna(0).sum()), 2),
                    "Gasto registrado COP": round(float(data["VALOR TOTAL"].fillna(0).sum()), 2),
                    "Precio ponderado COP/gal": round(float(weighted_unit), 2),
                    "Pendientes": int((data["ESTADO"].astype(str).str.lower() == "pendiente").sum()),
                    "Observados": int((data["ESTADO"].astype(str).str.lower() == "observado").sum()),
                }])
                st.dataframe(summary, use_container_width=True, hide_index=True)
        st.markdown("---")
        st.caption(f"Base de datos: {DB_PATH.name} · Registros totales: {len(df_all):,} · Última actualización al recargar la vista.")


# ============================================================
# PIE DE PÁGINA
# ============================================================

st.markdown(
    """<div style="text-align:center;color:#8A98A8;font-size:10px;padding:18px 0 8px;">
    MOONFLIGHT · FUEL CONTROL · Herramienta de apoyo a la gestión;
    no sustituye los sistemas contables, de mantenimiento ni de despacho operacional.
    </div>""",
    unsafe_allow_html=True,
)
