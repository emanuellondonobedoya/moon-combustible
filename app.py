import os
import json
import re
import hmac
import mimetypes
import shutil
from datetime import datetime

import openpyxl
import pandas as pd
import plotly.express as px
import streamlit as st
from google import genai
from google.genai import types

# ============================================================
# MOON · AEROLÍNEA | FUEL INTELLIGENCE
# ============================================================

st.set_page_config(
    page_title="Moon · Aerolínea",
    page_icon="🌙",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ----------------------------
# Identidad visual y Corrección de Títulos y Textos
# ----------------------------
NAVY = "#0B1D33"
NAVY_2 = "#142D4C"
BLUE = "#356B9A"
GOLD = "#C7A66A"
INK = "#17283D"
MUTED = "#718096"
CANVAS = "#F3F6FA"
BORDER = "#E5EAF1"
WHITE = "#FFFFFF"

st.markdown(
    f"""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700&family=Manrope:wght@400;500;600;700;800&display=swap');

    :root {{
        --moon-navy: {NAVY};
        --moon-navy-2: {NAVY_2};
        --moon-blue: {BLUE};
        --moon-gold: {GOLD};
        --moon-ink: {INK};
        --moon-muted: {MUTED};
        --moon-canvas: {CANVAS};
        --moon-border: {BORDER};
    }}

    html, body, [class*="css"] {{
        font-family: 'DM Sans', sans-serif;
        color: var(--moon-ink);
    }}
    .stApp {{ background: var(--moon-canvas); }}
    [data-testid="stHeader"] {{ background: transparent; }}
    [data-testid="stToolbar"] {{ right: 1.5rem; }}
    .block-container {{
        max-width: 1500px;
        padding-top: 2.1rem;
        padding-bottom: 3rem;
    }}
    #MainMenu, footer {{ visibility: hidden; }}

    /* Sidebar */
    [data-testid="stSidebar"] {{
        background: linear-gradient(180deg, {NAVY} 0%, #102944 100%);
        border-right: 1px solid rgba(255,255,255,.07);
    }}
    [data-testid="stSidebar"] > div:first-child {{ padding-top: 1.5rem; }}
    [data-testid="stSidebar"] [data-testid="stMarkdownContainer"] p {{ color: #B8C6D8; }}
    [data-testid="stSidebar"] [data-testid="stRadio"] label {{
        color: #DCE5EF !important;
        font-weight: 600;
        font-size: .93rem;
    }}
    [data-testid="stSidebar"] [data-testid="stRadio"] > div {{ gap: .45rem; }}
    [data-testid="stSidebar"] [data-testid="stRadio"] label > div:first-child {{
        background-color: rgba(255,255,255,.12);
        border-color: rgba(255,255,255,.30);
    }}
    [data-testid="stSidebar"] hr {{ border-color: rgba(255,255,255,.12); }}
    .brand-lockup {{
        display:flex; align-items:center; gap:12px;
        padding: 10px 8px 22px 8px;
    }}
    .brand-symbol {{
        width:42px; height:42px; border-radius:14px;
        display:flex; align-items:center; justify-content:center;
        color:#F4E7C9; font-family:'Manrope',sans-serif;
        font-weight:800; font-size:22px;
        background:linear-gradient(145deg, #1D4268, #0C2038);
        border:1px solid rgba(199,166,106,.55);
        box-shadow:0 7px 18px rgba(0,0,0,.18);
    }}
    .brand-name {{ color:#FFFFFF; font-family:'Manrope',sans-serif; font-size:17px; font-weight:800; letter-spacing:2.2px; line-height:1.1; }}
    .brand-sub {{ color:#9EB0C5; font-size:9px; letter-spacing:2.1px; margin-top:5px; }}
    .side-caption {{ color:#718AA6; font-size:10px; font-weight:700; letter-spacing:1.8px; margin: 14px 0 8px 8px; }}
    .sidebar-foot {{
        margin-top: 2.1rem; padding:14px 13px; border-radius:13px;
        border:1px solid rgba(255,255,255,.11); background:rgba(255,255,255,.045);
        color:#AFC0D3; font-size:11px; line-height:1.55;
    }}
    .sidebar-foot strong {{ color:#F4E7C9; letter-spacing:.5px; }}

    /* Encabezados y superficies (Corregido color de títulos principales) */
    .eyebrow {{
        color:{BLUE}; font-size:10px; font-weight:800;
        letter-spacing:2px; text-transform:uppercase; margin-bottom:8px;
    }}
    .page-title {{
        color: {NAVY} !important; font-family:'Manrope',sans-serif;
        font-weight:800; font-size:clamp(27px, 3vw, 38px);
        letter-spacing:-1.3px; line-height:1.12; margin:0;
    }}
    .page-subtitle {{ color:{MUTED}; font-size:14px; margin-top:9px; line-height:1.6; max-width:760px; }}
    .topline {{ display:flex; align-items:center; justify-content:space-between; gap:18px; margin-bottom:24px; }}
    .status-pill {{
        display:inline-flex; align-items:center; gap:8px;
        padding:8px 12px; border:1px solid #DCE6E2;
        border-radius:30px; background:#F8FCFA; color:#3A7461;
        font-size:11px; font-weight:700; white-space:nowrap;
    }}
    .status-dot {{ width:7px; height:7px; border-radius:50%; background:#4C9B7B; box-shadow:0 0 0 4px rgba(76,155,123,.12); }}
    .hero-panel {{
        position:relative; overflow:hidden; border-radius:22px;
        padding:27px 30px; margin: 8px 0 24px 0;
        background:linear-gradient(115deg, {NAVY} 0%, #12375A 62%, #1A4568 100%);
        box-shadow:0 14px 35px rgba(11,29,51,.11);
    }}
    .hero-panel:after {{
        content:''; position:absolute; width:320px; height:320px;
        border:1px solid rgba(226,203,160,.15); border-radius:50%;
        right:-85px; top:-190px; box-shadow:0 0 0 30px rgba(226,203,160,.035), 0 0 0 65px rgba(226,203,160,.025);
    }}
    .hero-kicker {{ color:#D9C39A; font-size:10px; font-weight:800; letter-spacing:2px; text-transform:uppercase; }}
    .hero-title {{ color:#FFFFFF; font-family:'Manrope',sans-serif; font-size:25px; font-weight:800; letter-spacing:-.6px; margin-top:8px; }}
    .hero-copy {{ color:#C3D0DE; font-size:13px; margin-top:7px; max-width:650px; line-height:1.6; }}
    .hero-mark {{ position:absolute; right:32px; top:50%; transform:translateY(-50%); color:rgba(232,214,180,.85); font-family:'Manrope',sans-serif; font-weight:800; font-size:54px; letter-spacing:-5px; opacity:.82; }}
    .section-heading {{ color:{NAVY}; font-family:'Manrope',sans-serif; font-size:17px; font-weight:800; letter-spacing:-.3px; margin: 22px 0 4px 0; }}
    .section-caption {{ color:{MUTED}; font-size:12px; margin-bottom:13px; }}
    .surface {{
        background:{WHITE}; border:1px solid {BORDER}; border-radius:17px;
        padding:19px 20px; box-shadow:0 5px 18px rgba(16,39,66,.035);
    }}
    .metric-card {{
        background:{WHITE}; border:1px solid {BORDER}; border-radius:16px;
        padding:17px 18px 16px 18px; min-height:118px;
        box-shadow:0 5px 18px rgba(16,39,66,.035);
    }}
    .metric-label {{ color:{MUTED}; font-size:11px; font-weight:700; letter-spacing:.8px; text-transform:uppercase; }}
    .metric-value {{ color:{NAVY}; font-family:'Manrope',sans-serif; font-size:25px; font-weight:800; letter-spacing:-.8px; margin-top:10px; line-height:1.15; }}
    .metric-note {{ color:#91A0B1; font-size:10px; margin-top:8px; }}
    .form-card {{ background:{WHITE}; border:1px solid {BORDER}; border-radius:18px; padding:20px 22px; }}
    .mini-label {{ color:{BLUE}; font-size:10px; font-weight:800; letter-spacing:1.5px; text-transform:uppercase; margin-bottom:5px; }}
    .footer-line {{ border-top:1px solid {BORDER}; margin-top:28px; padding-top:15px; color:#9AA7B6; font-size:10px; letter-spacing:.4px; }}

    /* Controles y visibilidad de texto en campos */
    .stTextInput label, .stNumberInput label, .stFileUploader label, .stDateInput label {{
        color:{INK} !important; font-size:12px !important; font-weight:700 !important;
    }}
    [data-testid="stFileUploader"] section {{
        border:1px dashed #AABCD0 !important; border-radius:13px !important;
        background:#F8FAFD !important; padding:20px !important;
    }}
    [data-testid="stFileUploader"] section:hover {{ border-color:{BLUE} !important; background:#F3F7FC !important; }}
    
    .stTextInput input, .stNumberInput input {{
        border-radius:10px !important; border-color:#DCE4ED !important;
        background:#FFFFFF !important; color: #0B1D33 !important; font-weight: 700 !important; min-height:43px;
    }}
    
    .stButton > button, .stDownloadButton > button {{
        border-radius:10px !important; min-height:43px;
        border:1px solid {NAVY} !important;
        background:{NAVY} !important; color:#FFFFFF !important;
        font-weight:700 !important; letter-spacing:.15px;
        box-shadow:0 5px 12px rgba(11,29,51,.12);
        transition:all .18s ease;
    }}
    .stButton > button:hover, .stDownloadButton > button:hover {{
        background:{NAVY_2} !important; border-color:{NAVY_2} !important;
        transform:translateY(-1px); box-shadow:0 8px 16px rgba(11,29,51,.16);
    }}
    [data-testid="stAlert"] {{ border-radius:12px; }}
    [data-testid="stDataFrame"] {{ border:1px solid {BORDER}; border-radius:12px; overflow:hidden; }}
    [data-testid="stExpander"] {{ border:1px solid {BORDER}; border-radius:12px; background:{WHITE}; }}
    .stPlotlyChart {{ background:{WHITE}; border:1px solid {BORDER}; border-radius:16px; padding:8px 8px 0 8px; }}

    @media (max-width: 800px) {{
        .block-container {{ padding:1.2rem 1rem 2rem 1rem; }}
        .topline {{ align-items:flex-start; flex-direction:column; }}
        .hero-panel {{ padding:22px; }}
        .hero-mark {{ display:none; }}
    }}
    </style>
    """,
    unsafe_allow_html=True,
)

# ----------------------------
# Configuración y utilidades
# ----------------------------
CARPETA_IMAGENES = "imagenes_tiquetes"
RUTA_EXCEL = "Base_Datos_Combustible_Moon.xlsx"
CABECERAS = [
    "FECHA REGISTRO", "RECIBO", "FECHA CONSUMO", "MATRICULA",
    "PRODUCTO", "RUTA", "CONSUMO (GAL)", "VALOR UNIT",
    "VALOR TOTAL", "PILOTO", "NOMENCLATURA IMAGEN", "OBSERVACION",
]
os.makedirs(CARPETA_IMAGENES, exist_ok=True)


def get_setting(key: str):
    try:
        value = st.secrets.get(key)
        if value:
            return value
    except Exception:
        pass
    return os.getenv(key)


def inicializar_excel():
    if not os.path.exists(RUTA_EXCEL):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Consumos"
        ws.append(CABECERAS)
        wb.save(RUTA_EXCEL)


def cargar_datos():
    inicializar_excel()
    try:
        df = pd.read_excel(RUTA_EXCEL)
        for col in CABECERAS:
            if col not in df.columns:
                df[col] = None
        return df[CABECERAS]
    except Exception:
        return pd.DataFrame(columns=CABECERAS)


def formato_cop(valor):
    return f"$ {valor:,.0f} COP".replace(",", ".")


def render_brand():
    st.sidebar.markdown(
        """
        <div class="brand-lockup">
            <div class="brand-symbol">M</div>
            <div>
                <div class="brand-name">MOON</div>
                <div class="brand-sub">AEROLÍNEA · OPERACIONES</div>
            </div>
        </div>
        <div class="side-caption">CENTRO DE OPERACIONES</div>
        """,
        unsafe_allow_html=True,
    )


def render_page_header(kicker, title, subtitle, show_status=True):
    status = (
        '<div class="status-pill"><span class="status-dot"></span> SISTEMA OPERATIVO</div>'
        if show_status else ""
    )
    st.markdown(
        f"""
        <div class="topline">
            <div>
                <div class="eyebrow">{kicker}</div>
                <h1 class="page-title">{title}</h1>
                <div class="page-subtitle">{subtitle}</div>
            </div>
            {status}
        </div>
        """,
        unsafe_allow_html=True,
    )


def metric_card(label, value, note):
    st.markdown(
        f"""
        <div class="metric-card">
            <div class="metric-label">{label}</div>
            <div class="metric-value">{value}</div>
            <div class="metric-note">{note}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def clean_json_response(text):
    text = (text or "").strip()
    text = re.sub(r"^```(?:json)?\s*", "", text, flags=re.IGNORECASE)
    text = re.sub(r"\s*```$", "", text)
    return json.loads(text)


inicializar_excel()

# ----------------------------
# Navegación
# ----------------------------
render_brand()
st.sidebar.markdown("<hr>", unsafe_allow_html=True)
vista = st.sidebar.radio(
    "MÓDULOS",
    ["Registro en pista", "Panel directivo"],
    label_visibility="collapsed",
)
st.sidebar.markdown(
    """
    <div class="sidebar-foot">
        <strong>MOON · AEROLÍNEA</strong><br>
        Control centralizado de consumos, soportes y costos de combustible.
    </div>
    """,
    unsafe_allow_html=True,
)
st.sidebar.markdown(
    f"<div style='color:#718AA6;font-size:10px;margin:22px 8px 0 8px;'>VERSIÓN OPERATIVA · {datetime.now().strftime('%Y')}</div>",
    unsafe_allow_html=True,
)

# ============================================================
# MÓDULO 1 | REGISTRO EN PISTA
# ============================================================
if vista == "Registro en pista":
    render_page_header(
        "GESTIÓN DE COMBUSTIBLE · OPERACIONES",
        "Registro en pista",
        "Captura y digitaliza los comprobantes de abastecimiento. La información queda centralizada para su consulta y seguimiento administrativo.",
    )

    st.markdown(
        """
        <div class="hero-panel">
            <div class="hero-kicker">MOON AEROLÍNEA · DIGITAL RECORD</div>
            <div class="hero-title">Cada abastecimiento, bajo control.</div>
            <div class="hero-copy">Registra el comprobante, extrae automáticamente los datos clave y consolida el consumo de cada aeronave en una sola fuente de información.</div>
            <div class="hero-mark">M.</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    col_form, col_context = st.columns([1.35, 0.85], gap="large")
    with col_form:
        st.markdown('<div class="section-heading">Nuevo registro de combustible</div>', unsafe_allow_html=True)
        st.markdown('<div class="section-caption">Carga el soporte del abastecimiento para iniciar el procesamiento.</div>', unsafe_allow_html=True)
        archivo_subido = st.file_uploader(
            "COMPROBANTE DE ABASTECIMIENTO",
            type=["jpg", "jpeg", "png"],
            help="Carga una fotografía clara del comprobante. Se recomienda un soporte por imagen.",
        )
        precio_unitario = st.number_input(
            "Precio estimado por galón (COP)",
            min_value=0.0,
            value=19500.0,
            step=100.0,
            format="%.0f",
        )
        procesar = st.button("Procesar y registrar comprobante", use_container_width=True, type="primary")

    with col_context:
        st.markdown('<div class="section-heading">Guía de captura</div>', unsafe_allow_html=True)
        st.markdown(
            """
            <div class="surface" style="min-height:250px;">
                <div class="mini-label">RECOMENDACIONES</div>
                <div style="color:#17283D;font-size:14px;font-weight:700;margin:10px 0 6px 0;">Calidad del soporte</div>
                <div style="color:#718096;font-size:12px;line-height:1.65;">Asegúrate de que el número del recibo, la fecha, la matrícula, el volumen y la firma sean legibles.</div>
                <div style="height:1px;background:#E5EAF1;margin:16px 0;"></div>
                <div style="color:#17283D;font-size:14px;font-weight:700;margin-bottom:6px;">Un comprobante por imagen</div>
                <div style="color:#718096;font-size:12px;line-height:1.65;">Evita fotografías con varios recibos. Una imagen individual facilita la lectura y la trazabilidad de cada registro.</div>
                <div style="height:1px;background:#E5EAF1;margin:16px 0;"></div>
                <div style="color:#17283D;font-size:14px;font-weight:700;margin-bottom:6px;">Validación</div>
                <div style="color:#718096;font-size:12px;line-height:1.65;">Verifica los datos extraídos antes de utilizar la información para reportes administrativos.</div>
            </div>
            """,
            unsafe_allow_html=True,
        )

    if archivo_subido is not None:
        st.markdown('<div class="section-heading">Vista previa del soporte</div>', unsafe_allow_html=True)
        preview_col, detail_col = st.columns([0.9, 1.1], gap="large")
        with preview_col:
            st.image(archivo_subido, use_container_width=True)
        with detail_col:
            st.markdown(
                f"""
                <div class="surface">
                    <div class="mini-label">ARCHIVO CARGADO</div>
                    <div style="color:{NAVY};font-family:Manrope,sans-serif;font-size:17px;font-weight:800;word-break:break-word;">{archivo_subido.name}</div>
                    <div style="color:{MUTED};font-size:12px;margin-top:8px;">{archivo_subido.size / 1024:,.1f} KB · {archivo_subido.type or 'Imagen'}</div>
                    <div style="height:1px;background:{BORDER};margin:17px 0;"></div>
                    <div style="color:{MUTED};font-size:12px;line-height:1.6;">Al procesar, la inteligencia artificial intentará identificar los datos del comprobante y guardar el soporte junto con el registro.</div>
                </div>
                """,
                unsafe_allow_html=True,
            )

    if procesar:
        if archivo_subido is None:
            st.warning("Primero carga una fotografía del comprobante.")
        elif precio_unitario <= 0:
            st.warning("Ingresa un precio por galón mayor que cero.")
        else:
            api_key = get_setting("GEMINI_API_KEY")
            if not api_key:
                st.error("No se encontró GEMINI_API_KEY. Configura la clave en Streamlit Secrets o en una variable de entorno para habilitar la lectura automática.")
            else:
                try:
                    client = genai.Client(api_key=api_key)
                    mime_type = archivo_subido.type or mimetypes.guess_type(archivo_subido.name)[0] or "image/jpeg"
                    image_bytes = archivo_subido.getvalue()
                    prompt = """
                    Analiza esta imagen de un tiquete de combustible de una aeronave. Extrae:
                    - "recibo": Número del tiquete.
                    - "fecha_consumo": Fecha en formato YYYY-MM-DD.
                    - "matricula": Matrícula del avión.
                    - "producto": Tipo de combustible.
                    - "ruta": Destino o ruta.
                    - "consumo": Volumen total en galones (solo número).
                    - "piloto": Nombre de quien firma, si es legible.

                    Devuelve ÚNICAMENTE un JSON válido con este formato:
                    {"exito": true, "recibo": "...", "fecha_consumo": "...", "matricula": "...", "producto": "...", "ruta": "...", "consumo": 0.0, "piloto": "..."}
                    Si la imagen no permite leer los datos principales, devuelve: {"exito": false}
                    No inventes información que no sea visible.
                    """
                    with st.spinner("Analizando el comprobante y estructurando el registro..."):
                        response = client.models.generate_content(
                            model="gemini-3.8-flash-lite",
                            contents=[types.Part.from_bytes(data=image_bytes, mime_type=mime_type), prompt],
                        )
                        datos = clean_json_response(response.text)

                    if not datos.get("exito"):
                        st.error("No fue posible leer el comprobante con suficiente claridad. Intenta cargar una fotografía más nítida.")
                    else:
                        galones = float(datos.get("consumo", 0.0) or 0.0)
                        if galones <= 0:
                            st.error("La lectura no devolvió un volumen válido. Revisa el soporte antes de registrar el consumo.")
                        else:
                            fecha_str = datetime.now().strftime("%Y%m%d_%H%M%S")
                            num_recibo = str(datos.get("recibo", "SREC")).replace("/", "-").replace(" ", "")
                            extension = os.path.splitext(archivo_subido.name)[1].lower() or ".jpg"
                            nombre_imagen = f"Recibo_{num_recibo}_{fecha_str}{extension}"
                            ruta_guardado = os.path.join(CARPETA_IMAGENES, nombre_imagen)
                            with open(ruta_guardado, "wb") as f:
                                f.write(image_bytes)

                            valor_total = galones * float(precio_unitario)
                            wb = openpyxl.load_workbook(RUTA_EXCEL)
                            hoja = wb["Consumos"] if "Consumos" in wb.sheetnames else wb.active
                            hoja.append([
                                datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                                datos.get("recibo"),
                                datos.get("fecha_consumo"),
                                datos.get("matricula"),
                                datos.get("producto"),
                                datos.get("ruta"),
                                galones,
                                float(precio_unitario),
                                valor_total,
                                datos.get("piloto"),
                                nombre_imagen,
                                "Procesado automáticamente",
                            ])
                            wb.save(RUTA_EXCEL)

                            st.success("Registro guardado correctamente.")
                            st.markdown('<div class="section-heading">Resumen del abastecimiento</div>', unsafe_allow_html=True)
                            r1, r2, r3 = st.columns(3, gap="medium")
                            with r1:
                                metric_card("Volumen registrado", f"{galones:,.1f} GAL", "Consumo identificado en el soporte")
                            with r2:
                                metric_card("Valor estimado", formato_cop(valor_total), "Volumen × precio por galón")
                            with r3:
                                metric_card("Aeronave", str(datos.get("matricula") or "Sin identificar"), "Matrícula detectada por IA")
                            with st.expander("Consultar datos extraídos"):
                                st.json(datos)
                except Exception as exc:
                    st.error(f"No fue posible completar el procesamiento: {exc}")

    st.markdown(
        '<div class="footer-line">MOON · AEROLÍNEA &nbsp; / &nbsp; REGISTRO OPERATIVO</div>',
        unsafe_allow_html=True,
    )

# ============================================================
# MÓDULO 2 | PANEL DIRECTIVO
# ============================================================
elif vista == "Panel directivo":
    render_page_header(
        "BUSINESS INTELLIGENCE · DIRECCIÓN",
        "Panel directivo",
        "Una visión consolidada del consumo, los costos y la actividad de abastecimiento de la flota.",
        show_status=False,
    )

    admin_password = get_setting("ADMIN_PASSWORD")
    if "moon_authenticated" not in st.session_state:
        st.session_state.moon_authenticated = False

    if not st.session_state.moon_authenticated:
        login_col, info_col = st.columns([1, 1], gap="large")
        with login_col:
            st.markdown('<div class="section-heading">Acceso administrativo</div>', unsafe_allow_html=True)
            st.markdown('<div class="section-caption">Ingresa tus credenciales para consultar los indicadores.</div>', unsafe_allow_html=True)
            with st.form("moon_login"):
                clave = st.text_input("CONTRASEÑA", type="password", placeholder="Ingresa tu contraseña")
                ingresar = st.form_submit_button("Ingresar al panel", use_container_width=True)
            if ingresar:
                if not admin_password:
                    st.error("El acceso aún no está configurado. Define ADMIN_PASSWORD en Streamlit Secrets o como variable de entorno.")
                elif hmac.compare_digest(str(clave), str(admin_password)):
                    st.session_state.moon_authenticated = True
                    st.rerun()
                else:
                    st.error("Credenciales incorrectas. Verifica la contraseña e inténtalo nuevamente.")
        with info_col:
            st.markdown(
                """
                <div class="hero-panel" style="margin-top:16px;min-height:205px;">
                    <div class="hero-kicker">MANAGEMENT VIEW</div>
                    <div class="hero-title">La operación, en perspectiva.</div>
                    <div class="hero-copy">Indicadores para comprender el comportamiento del consumo y fortalecer el control financiero de combustible.</div>
                    <div class="hero-mark">M.</div>
                </div>
                """,
                unsafe_allow_html=True,
            )
    else:
        head_col, logout_col = st.columns([5, 1])
        with head_col:
            st.markdown('<div class="status-pill"><span class="status-dot"></span> SESIÓN ADMINISTRATIVA ACTIVA</div>', unsafe_allow_html=True)
        with logout_col:
            if st.button("Cerrar sesión", use_container_width=True):
                st.session_state.moon_authenticated = False
                st.rerun()

        df = cargar_datos()
        if not df.empty:
            df["CONSUMO (GAL)"] = pd.to_numeric(df["CONSUMO (GAL)"], errors="coerce").fillna(0)
            df["VALOR TOTAL"] = pd.to_numeric(df["VALOR TOTAL"], errors="coerce").fillna(0)
            df["VALOR UNIT"] = pd.to_numeric(df["VALOR UNIT"], errors="coerce").fillna(0)
            df["MATRICULA"] = df["MATRICULA"].fillna("Sin identificar").astype(str)
            df["RUTA"] = df["RUTA"].fillna("Sin ruta").astype(str)

            st.markdown(
                """
                <div class="hero-panel">
                    <div class="hero-kicker">FLEET COST CONTROL · EXECUTIVE OVERVIEW</div>
                    <div class="hero-title">Visibilidad para decidir mejor.</div>
                    <div class="hero-copy">Monitorea el volumen abastecido, el gasto acumulado y la distribución de consumo entre aeronaves y rutas.</div>
                    <div class="hero-mark">M.</div>
                </div>
                """,
                unsafe_allow_html=True,
            )

            total_tiquetes = len(df)
            total_galones = df["CONSUMO (GAL)"].sum()
            total_gasto = df["VALOR TOTAL"].sum()
            promedio = total_gasto / total_tiquetes if total_tiquetes else 0

            st.markdown('<div class="section-heading">Indicadores clave</div>', unsafe_allow_html=True)
            st.markdown('<div class="section-caption">Acumulado de los registros disponibles en la base de datos.</div>', unsafe_allow_html=True)
            k1, k2, k3, k4 = st.columns(4, gap="medium")
            with k1:
                metric_card("Comprobantes", f"{total_tiquetes:,}", "Registros consolidados")
            with k2:
                metric_card("Volumen total", f"{total_galones:,.1f} GAL", "Galones registrados")
            with k3:
                metric_card("Gasto acumulado", formato_cop(total_gasto), "Valor estimado de abastecimiento")
            with k4:
                metric_card("Promedio por registro", formato_cop(promedio), "Gasto medio por comprobante")

            st.markdown('<div class="section-heading">Lectura de la operación</div>', unsafe_allow_html=True)
            st.markdown('<div class="section-caption">Distribución del consumo y del gasto según la información registrada.</div>', unsafe_allow_html=True)
            chart1, chart2 = st.columns(2, gap="large")
            with chart1:
                df_matricula = (
                    df.groupby("MATRICULA", as_index=False)["CONSUMO (GAL)"].sum()
                    .sort_values("CONSUMO (GAL)", ascending=True)
                )
                fig_mat = px.bar(
                    df_matricula,
                    x="CONSUMO (GAL)",
                    y="MATRICULA",
                    orientation="h",
                    title="Consumo por aeronave",
                    labels={"CONSUMO (GAL)": "Galones", "MATRICULA": "Aeronave"},
                    color_discrete_sequence=[BLUE],
                )
                fig_mat.update_layout(
                    template="plotly_white", height=350,
                    margin=dict(l=18, r=20, t=65, b=25),
                    title=dict(font=dict(size=15, color=NAVY, family="Manrope")),
                    font=dict(family="DM Sans", color=MUTED, size=11),
                    paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
                    xaxis=dict(showgrid=True, gridcolor="#EDF1F5", zeroline=False),
                    yaxis=dict(showgrid=False, zeroline=False),
                    showlegend=False,
                )
                fig_mat.update_traces(marker_line_width=0)
                st.plotly_chart(fig_mat, use_container_width=True, config={"displayModeBar": False})

            with chart2:
                df_ruta = df.groupby("RUTA", as_index=False)["VALOR TOTAL"].sum().sort_values("VALOR TOTAL", ascending=False)
                fig_ruta = px.pie(
                    df_ruta,
                    names="RUTA",
                    values="VALOR TOTAL",
                    title="Distribución del gasto por ruta",
                    hole=0.66,
                    color_discrete_sequence=[NAVY, BLUE, GOLD, "#8EA8C3", "#B9C9D9", "#6F879F"],
                )
                fig_ruta.update_layout(
                    template="plotly_white", height=350,
                    margin=dict(l=18, r=18, t=65, b=20),
                    title=dict(font=dict(size=15, color=NAVY, family="Manrope")),
                    font=dict(family="DM Sans", color=MUTED, size=11),
                    paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
                    legend=dict(orientation="h", y=-0.12, x=0.5, xanchor="center", font=dict(size=10)),
                )
                fig_ruta.update_traces(textposition="inside", textinfo="percent", marker=dict(line=dict(color=WHITE, width=3)))
                st.plotly_chart(fig_ruta, use_container_width=True, config={"displayModeBar": False})

            st.markdown('<div class="section-heading">Registros de abastecimiento</div>', unsafe_allow_html=True)
            st.markdown('<div class="section-caption">Base maestra de comprobantes procesados.</div>', unsafe_allow_html=True)
            st.dataframe(
                df.sort_values("FECHA REGISTRO", ascending=False, na_position="last"),
                use_container_width=True,
                hide_index=True,
            )

            download_col, spacer_col = st.columns([1.1, 3.9])
            with download_col:
                with open(RUTA_EXCEL, "rb") as f:
                    st.download_button(
                        "Descargar base de datos",
                        data=f.read(),
                        file_name="Moon_Control_Combustible.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                    )

            with st.expander("Administración y mantenimiento de datos"):
                st.warning("La eliminación es permanente. Descarga una copia de seguridad antes de continuar.")
                confirmar = st.checkbox("Confirmo que deseo eliminar todos los registros y soportes almacenados.")
                if st.button("Eliminar todos los registros", disabled=not confirmar):
                    if os.path.exists(RUTA_EXCEL):
                        os.remove(RUTA_EXCEL)
                    if os.path.exists(CARPETA_IMAGENES):
                        shutil.rmtree(CARPETA_IMAGENES)
                    inicializar_excel()
                    os.makedirs(CARPETA_IMAGENES, exist_ok=True)
                    st.success("Los registros fueron eliminados. Recarga el panel para ver la base vacía.")
                    st.rerun()
        else:
            st.markdown(
                """
                <div class="hero-panel">
                    <div class="hero-kicker">EXECUTIVE OVERVIEW</div>
                    <div class="hero-title">Tu centro de control está listo.</div>
                    <div class="hero-copy">Cuando se registren los primeros abastecimientos, aquí encontrarás los indicadores, las tendencias y la trazabilidad de los consumos de la flota.</div>
                    <div class="hero-mark">M.</div>
                </div>
                """,
                unsafe_allow_html=True,
            )
            e1, e2, e3 = st.columns(3, gap="medium")
            with e1:
                metric_card("Comprobantes", "0", "A la espera del primer registro")
            with e2:
                metric_card("Volumen total", "0 GAL", "Sin consumos registrados")
            with e3:
                metric_card("Gasto acumulado", "$ 0 COP", "Sin valores consolidados")
            st.info("Todavía no hay registros en la base de datos. Utiliza el módulo Registro en pista para comenzar.")

    st.markdown(
        '<div class="footer-line">MOON · AEROLÍNEA &nbsp; / &nbsp; EXECUTIVE DASHBOARD</div>',
        unsafe_allow_html=True,
    )
