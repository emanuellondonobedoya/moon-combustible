import os
import json
import re
import hmac
import base64
import shutil
from datetime import datetime

import openpyxl
import pandas as pd
import plotly.express as px
import streamlit as st
import anthropic

# ============================================================
# MOON · AEROLÍNEA | FUEL INTELLIGENCE (Powered by Claude)
# ============================================================

st.set_page_config(
    page_title="Moon · Aerolínea",
    page_icon="🌙",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ----------------------------
# Identidad visual y Estilos
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
    :root {{ --moon-navy: {NAVY}; --moon-navy-2: {NAVY_2}; --moon-blue: {BLUE}; --moon-ink: {INK}; --moon-muted: {MUTED}; --moon-canvas: {CANVAS}; --moon-border: {BORDER}; }}
    html, body, [class*="css"] {{ font-family: 'DM Sans', sans-serif; color: var(--moon-ink); }}
    .stApp {{ background: var(--moon-canvas); }}
    [data-testid="stHeader"] {{ background: transparent; }}
    .block-container {{ max-width: 1500px; padding-top: 2.1rem; padding-bottom: 3rem; }}
    
    /* Sidebar */
    [data-testid="stSidebar"] {{ background: linear-gradient(180deg, {NAVY} 0%, #102944 100%); border-right: 1px solid rgba(255,255,255,.07); }}
    [data-testid="stSidebar"] [data-testid="stRadio"] label {{ color: #DCE5EF !important; font-weight: 600; }}
    .brand-lockup {{ display:flex; align-items:center; gap:12px; padding: 10px 8px 22px 8px; }}
    .brand-symbol {{ width:42px; height:42px; border-radius:14px; display:flex; align-items:center; justify-content:center; color:#F4E7C9; font-family:'Manrope',sans-serif; font-weight:800; font-size:22px; background:linear-gradient(145deg, #1D4268, #0C2038); border:1px solid rgba(199,166,106,.55); }}
    .brand-name {{ color:#FFFFFF; font-family:'Manrope',sans-serif; font-size:17px; font-weight:800; letter-spacing:2.2px; line-height:1.1; }}
    .brand-sub {{ color:#9EB0C5; font-size:9px; letter-spacing:2.1px; margin-top:5px; }}
    
    /* Títulos y Tarjetas */
    .page-title {{ color: {NAVY} !important; font-family:'Manrope',sans-serif; font-weight:800; font-size:clamp(27px, 3vw, 38px); letter-spacing:-1.3px; line-height:1.12; margin:0; }}
    .hero-panel {{ position:relative; overflow:hidden; border-radius:22px; padding:27px 30px; margin: 8px 0 24px 0; background:linear-gradient(115deg, {NAVY} 0%, #12375A 62%, #1A4568 100%); }}
    .hero-kicker {{ color:#D9C39A; font-size:10px; font-weight:800; letter-spacing:2px; text-transform:uppercase; }}
    .hero-title {{ color:#FFFFFF; font-family:'Manrope',sans-serif; font-size:25px; font-weight:800; letter-spacing:-.6px; margin-top:8px; }}
    .hero-copy {{ color:#C3D0DE; font-size:13px; margin-top:7px; max-width:650px; line-height:1.6; }}
    .hero-mark {{ position:absolute; right:32px; top:50%; transform:translateY(-50%); color:rgba(232,214,180,.85); font-family:'Manrope',sans-serif; font-weight:800; font-size:54px; letter-spacing:-5px; opacity:.82; }}
    
    .section-heading {{ color:{NAVY}; font-family:'Manrope',sans-serif; font-size:17px; font-weight:800; letter-spacing:-.3px; margin: 22px 0 4px 0; }}
    .metric-card {{ background:{WHITE}; border:1px solid {BORDER}; border-radius:16px; padding:17px 18px 16px 18px; min-height:118px; }}
    .metric-label {{ color:{MUTED}; font-size:11px; font-weight:700; letter-spacing:.8px; text-transform:uppercase; }}
    .metric-value {{ color:{NAVY}; font-family:'Manrope',sans-serif; font-size:25px; font-weight:800; letter-spacing:-.8px; margin-top:10px; line-height:1.15; }}
    
    /* Controles y Visibilidad */
    .stTextInput input, .stNumberInput input {{ border-radius:10px !important; background:#FFFFFF !important; color: #0B1D33 !important; font-weight: 700 !important; }}
    .stButton > button {{ border-radius:10px !important; background:{NAVY} !important; color:#FFFFFF !important; font-weight:700 !important; }}
    </style>
    """,
    unsafe_allow_html=True,
)

# ----------------------------
# Configuración y utilidades (DOS BASES DE DATOS)
# ----------------------------
CARPETA_IMAGENES = "imagenes_tiquetes"
RUTA_DIARIO = "Registro_Diario_Moon.xlsx"
RUTA_MAESTRO = "Base_Datos_Maestra_Moon.xlsx"
CABECERAS = [
    "FECHA REGISTRO", "RECIBO", "FECHA CONSUMO", "MATRICULA",
    "PRODUCTO", "RUTA", "CONSUMO (GAL)", "VALOR UNIT",
    "VALOR TOTAL", "PILOTO", "NOMENCLATURA IMAGEN", "OBSERVACION",
]
os.makedirs(CARPETA_IMAGENES, exist_ok=True)

def inicializar_excels():
    for ruta in [RUTA_DIARIO, RUTA_MAESTRO]:
        if not os.path.exists(ruta):
            df = pd.DataFrame(columns=CABECERAS)
            df.to_excel(ruta, index=False)

def cargar_datos(ruta):
    try:
        df = pd.read_excel(ruta)
        for col in CABECERAS:
            if col not in df.columns: df[col] = None
        return df[CABECERAS]
    except Exception:
        return pd.DataFrame(columns=CABECERAS)

def formato_cop(valor):
    return f"$ {valor:,.0f} COP".replace(",", ".")

def clean_json_response(text):
    text = (text or "").strip()
    text = re.sub(r"^```(?:json)?\s*", "", text, flags=re.IGNORECASE)
    text = re.sub(r"\s*```$", "", text)
    try:
        return json.loads(text)
    except:
        return {"exito": False}

inicializar_excels()

# ----------------------------
# Barra Lateral
# ----------------------------
st.sidebar.markdown("""
    <div class="brand-lockup">
        <div class="brand-symbol">M</div>
        <div>
            <div class="brand-name">MOON</div>
            <div class="brand-sub">AEROLÍNEA · OPERACIONES</div>
        </div>
    </div>
""", unsafe_allow_html=True)
st.sidebar.markdown("<hr>", unsafe_allow_html=True)
vista = st.sidebar.radio("MÓDULOS", ["Registro en pista", "Panel directivo"], label_visibility="collapsed")

# ============================================================
# MÓDULO 1 | REGISTRO EN PISTA (PILOTOS)
# ============================================================
if vista == "Registro en pista":
    st.markdown('<h1 class="page-title">Registro en pista</h1>', unsafe_allow_html=True)
    
    st.markdown("""
        <div class="hero-panel">
            <div class="hero-kicker">MOON AEROLÍNEA · DIGITAL RECORD</div>
            <div class="hero-title">Cada abastecimiento, bajo control.</div>
            <div class="hero-copy">Sube el comprobante. La información se extraerá automáticamente para la validación administrativa al final del día.</div>
            <div class="hero-mark">M.</div>
        </div>
    """, unsafe_allow_html=True)

    col_form, col_context = st.columns([1.35, 0.85], gap="large")
    
    with col_form:
        st.markdown('<div class="section-heading">Nuevo registro de combustible</div>', unsafe_allow_html=True)
        st.info("✈️ **Pilotos:** Solo carguen la imagen del tiquete. Dirección fijará el precio al cerrar el día.")
        
        archivo_subido = st.file_uploader("COMPROBANTE DE ABASTECIMIENTO", type=["jpg", "jpeg", "png"])
        procesar = st.button("Procesar y registrar comprobante", use_container_width=True, type="primary")

    with col_context:
        st.markdown('<div class="section-heading">Soporte</div>', unsafe_allow_html=True)
        if archivo_subido is not None:
            st.image(archivo_subido, use_container_width=True, caption="Soporte cargado")

    if procesar:
        if archivo_subido is None:
            st.warning("⚠️ Primero carga una fotografía del comprobante.")
        else:
            api_key = st.secrets.get("ANTHROPIC_API_KEY")
            if not api_key:
                st.error("❌ No se encontró ANTHROPIC_API_KEY en los Secrets.")
            else:
                with st.spinner("Leyendo comprobante Terpel..."):
                    try:
                        image_bytes = archivo_subido.getvalue()
                        mime_type = archivo_subido.type or "image/jpeg"
                        base64_image = base64.b64encode(image_bytes).decode("utf-8")
                        
                        client = anthropic.Anthropic(api_key=api_key)
                        
                        prompt = """
                        Analiza esta imagen de un tiquete de combustible de una aeronave. Extrae:
                        - "recibo": Número del tiquete.
                        - "fecha_consumo": Fecha en formato YYYY-MM-DD.
                        - "matricula": Matrícula del avión.
                        - "producto": Tipo de combustible.
                        - "ruta": Destino o ruta.
                        - "consumo": Volumen total en galones (solo número).
                        - "piloto": Nombre de quien firma, si es legible.
                        Devuelve ÚNICAMENTE un objeto JSON válido.
                        {"exito": true, "recibo": "...", "fecha_consumo": "...", "matricula": "...", "producto": "...", "ruta": "...", "consumo": 0.0, "piloto": "..."}
                        """
                        
                        response = client.messages.create(
                            model="claude-3-haiku-20240307",
                            max_tokens=1024,
                            messages=[{"role": "user", "content": [{"type": "image", "source": {"type": "base64", "media_type": mime_type, "data": base64_image}}, {"type": "text", "text": prompt}]}]
                        )
                        
                        datos = clean_json_response(response.content[0].text)

                        if not datos.get("exito"):
                            st.error("❌ No fue posible leer el comprobante. Intenta con una foto más clara.")
                        else:
                            galones = float(datos.get("consumo", 0.0) or 0.0)
                            fecha_str = datetime.now().strftime("%Y%m%d_%H%M%S")
                            num_recibo = str(datos.get("recibo", "SREC")).replace("/", "-")
                            nombre_imagen = f"Recibo_{num_recibo}_{fecha_str}.jpg"
                            
                            with open(os.path.join(CARPETA_IMAGENES, nombre_imagen), "wb") as f:
                                f.write(image_bytes)

                            # Guardar en DIARIO con precios en 0 (esperando a Dirección)
                            df_diario = pd.read_excel(RUTA_DIARIO)
                            nuevo_registro = pd.DataFrame([{
                                "FECHA REGISTRO": datetime.now().strftime("%Y-%m-%d %H:%M:%S"), "RECIBO": datos.get("recibo"),
                                "FECHA CONSUMO": datos.get("fecha_consumo"), "MATRICULA": datos.get("matricula"),
                                "PRODUCTO": datos.get("producto"), "RUTA": datos.get("ruta"), "CONSUMO (GAL)": galones,
                                "VALOR UNIT": 0.0, "VALOR TOTAL": 0.0, "PILOTO": datos.get("piloto"),
                                "NOMENCLATURA IMAGEN": nombre_imagen, "OBSERVACION": "Pendiente fijar precio"
                            }])
                            df_diario = pd.concat([df_diario, nuevo_registro], ignore_index=True)
                            df_diario.to_excel(RUTA_DIARIO, index=False)

                            st.success("✅ ¡Soporte enviado a Dirección!")
                            st.info(f"Se registraron {galones:,.1f} GAL para la aeronave {datos.get('matricula')}.")

                    except Exception as exc:
                        st.error(f"❌ Error: {exc}")

# ============================================================
# MÓDULO 2 | PANEL DIRECTIVO (DIRECCIÓN)
# ============================================================
elif vista == "Panel directivo":
    st.markdown('<h1 class="page-title">Panel directivo</h1>', unsafe_allow_html=True)
    
    if "moon_authenticated" not in st.session_state:
        st.session_state.moon_authenticated = False

    if not st.session_state.moon_authenticated:
        with st.form("moon_login"):
            clave = st.text_input("CONTRASEÑA ADMINISTRATIVA", type="password")
            ingresar = st.form_submit_button("Ingresar al panel")
        if ingresar:
            if hmac.compare_digest(str(clave), str(st.secrets.get("ADMIN_PASSWORD", "adminmoon"))):
                st.session_state.moon_authenticated = True
                st.rerun()
            else:
                st.error("❌ Credenciales incorrectas.")
    else:
        if st.button("Cerrar sesión"):
            st.session_state.moon_authenticated = False
            st.rerun()

        st.markdown("""
            <div class="hero-panel">
                <div class="hero-kicker">EXECUTIVE OVERVIEW</div>
                <div class="hero-title">Gestión Financiera de Flota.</div>
                <div class="hero-copy">Asigna precios a los reportes diarios, descarga los informes y consolídalos en la base maestra mensual.</div>
            </div>
        """, unsafe_allow_html=True)
        
        tab1, tab2 = st.tabs(["📅 CIERRE DEL DÍA (Reportes Pilotos)", "📊 BASE DE DATOS MAESTRA (Histórico)"])

        # -------------------------------------------------------------------
        # PESTAÑA 1: GESTIÓN DEL DÍA
        # -------------------------------------------------------------------
        with tab1:
            df_diario = cargar_datos(RUTA_DIARIO)
            st.markdown('<div class="section-heading">Tiquetes reportados hoy en pista</div>', unsafe_allow_html=True)
            
            if df_diario.empty:
                st.info("No hay recibos pendientes de los pilotos por procesar el día de hoy.")
            else:
                st.dataframe(df_diario, use_container_width=True, hide_index=True)
                
                st.markdown('<div class="section-heading">1. Asignar precio de combustible del día</div>', unsafe_allow_html=True)
                col_p1, col_p2 = st.columns([1, 2])
                with col_p1:
                    precio_dia = st.number_input("Precio por Galón (COP):", value=19500.0, step=100.0)
                    if st.button("Aplicar Precio al Diario", type="primary"):
                        df_diario["VALOR UNIT"] = precio_dia
                        df_diario["VALOR TOTAL"] = pd.to_numeric(df_diario["CONSUMO (GAL)"], errors='coerce').fillna(0) * precio_dia
                        df_diario["OBSERVACION"] = "Cerrado por Dirección"
                        df_diario.to_excel(RUTA_DIARIO, index=False)
                        st.success("✅ Precio aplicado a todos los registros de hoy.")
                        st.rerun()
                with col_p2:
                    st.write("Al aplicar el precio, se actualizará el 'VALOR TOTAL' de todos los recibos de la tabla superior.")

                st.markdown("---")
                st.markdown('<div class="section-heading">2. Descargar y Consolidar</div>', unsafe_allow_html=True)
                col_d1, col_d2 = st.columns(2)
                
                with col_d1:
                    with open(RUTA_DIARIO, "rb") as f:
                        st.download_button("⬇️ 1. Descargar Excel del Día", data=f.read(), file_name=f"Moon_Combustible_Diario_{datetime.now().strftime('%Y%m%d')}.xlsx", use_container_width=True)
                
                with col_d2:
                    if st.button("📦 2. Archivar al Histórico Mensual y Limpiar Día", use_container_width=True):
                        # Pasar a la maestra
                        df_maestro = cargar_datos(RUTA_MAESTRO)
                        df_maestro = pd.concat([df_maestro, df_diario], ignore_index=True)
                        df_maestro.to_excel(RUTA_MAESTRO, index=False)
                        # Limpiar el diario
                        pd.DataFrame(columns=CABECERAS).to_excel(RUTA_DIARIO, index=False)
                        st.success("✅ Día consolidado con éxito. El panel diario está limpio para mañana.")
                        st.rerun()

        # -------------------------------------------------------------------
        # PESTAÑA 2: BASE DE DATOS HISTÓRICA (MAESTRA)
        # -------------------------------------------------------------------
        with tab2:
            df_maestro = cargar_datos(RUTA_MAESTRO)
            if df_maestro.empty:
                st.info("La base de datos general está vacía. Consolida un cierre de día para ver los datos históricos.")
            else:
                df_maestro["CONSUMO (GAL)"] = pd.to_numeric(df_maestro["CONSUMO (GAL)"], errors="coerce").fillna(0)
                df_maestro["VALOR TOTAL"] = pd.to_numeric(df_maestro["VALOR TOTAL"], errors="coerce").fillna(0)
                
                st.markdown('<div class="section-heading">Indicadores Acumulados (Global)</div>', unsafe_allow_html=True)
                k1, k2, k3 = st.columns(3)
                with k1: metric_card("Tiquetes Totales", f"{len(df_maestro):,}", "")
                with k2: metric_card("Galones Históricos", f"{df_maestro['CONSUMO (GAL)'].sum():,.1f} GAL", "")
                with k3: metric_card("Gasto Histórico", formato_cop(df_maestro["VALOR TOTAL"].sum()), "")

                st.markdown('<div class="section-heading">Base de datos consolidada (Todos los días)</div>', unsafe_allow_html=True)
                st.dataframe(df_maestro, use_container_width=True, hide_index=True)
                
                with open(RUTA_MAESTRO, "rb") as f:
                    st.download_button("⬇️ Descargar Base de Datos Consolidada (Mensual)", data=f.read(), file_name="Moon_Base_Maestra_General.xlsx", type="primary")

                with st.expander("⚠️ Mantenimiento Avanzado"):
                    st.warning("Formateo de la Base Maestra")
                    if st.button("Eliminar Todo el Historial y Base de Datos"):
                        if os.path.exists(RUTA_MAESTRO): os.remove(RUTA_MAESTRO)
                        if os.path.exists(RUTA_DIARIO): os.remove(RUTA_DIARIO)
                        if os.path.exists(CARPETA_IMAGENES): shutil.rmtree(CARPETA_IMAGENES)
                        inicializar_excels()
                        st.success("Bases de datos reiniciadas de fábrica.")
                        st.rerun()
