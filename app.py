import os
import json
import re
import hmac
import base64
import time
from datetime import datetime

import openpyxl
import pandas as pd
import plotly.express as px
import streamlit as st
import anthropic

# ============================================================
# MOON · AEROLÍNEA | FUEL INTELLIGENCE (Interactive & BI Dashboards)
# ============================================================

st.set_page_config(
    page_title="Moon · Aerolínea",
    page_icon="🌙",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ----------------------------
# Identidad visual, CSS y Animaciones
# ----------------------------
NAVY = "#0B1D33"
BLUE = "#356B9A"
INK = "#17283D"
WHITE = "#FFFFFF"

st.markdown(
    f"""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@400;500;600;700&family=Manrope:wght@400;500;600;700;800&display=swap');
    
    html, body, [class*="css"] {{ font-family: 'DM Sans', sans-serif; color: {INK}; }}
    .stApp {{ background-color: #F4F7FB; }}
    
    /* CORRECCIÓN: Textos blancos invisibles y Caja de subida */
    .stFileUploader label, .stTextInput label, .stNumberInput label {{
        color: {NAVY} !important; font-weight: 800 !important; font-size: 14px !important;
    }}
    [data-testid="stFileUploadDropzone"] {{
        background-color: #FFFFFF !important; border: 2px dashed #AABCD0 !important;
        border-radius: 16px !important; transition: all 0.3s ease;
    }}
    [data-testid="stFileUploadDropzone"]:hover {{
        border-color: {BLUE} !important; background-color: #F0F4F8 !important;
    }}
    [data-testid="stFileUploadDropzone"] * {{ color: {INK} !important; }}
    
    /* ANIMACIÓN: Marca Hero flotante */
    @keyframes float {{
        0% {{ transform: translateY(-50%); }}
        50% {{ transform: translateY(-56%); }}
        100% {{ transform: translateY(-50%); }}
    }}
    .hero-mark {{
        position:absolute; right:32px; top:50%;
        color:rgba(232,214,180,.85); font-family:'Manrope',sans-serif; 
        font-weight:800; font-size:54px; opacity:.82; animation: float 4s ease-in-out infinite;
    }}

    /* DISEÑO DE PANELES Y TARJETAS BI */
    .hero-panel {{ position:relative; overflow:hidden; border-radius:22px; padding:27px 30px; margin: 8px 0 24px 0; background:linear-gradient(115deg, {NAVY} 0%, #12375A 62%, #1A4568 100%); box-shadow: 0 10px 25px rgba(11,29,51,0.15); }}
    .hero-kicker {{ color:#D9C39A; font-size:10px; font-weight:800; letter-spacing:2px; text-transform:uppercase; }}
    .hero-title {{ color:#FFFFFF; font-family:'Manrope',sans-serif; font-size:25px; font-weight:800; letter-spacing:-.6px; margin-top:8px; }}
    .hero-copy {{ color:#C3D0DE; font-size:13px; margin-top:7px; max-width:650px; line-height:1.6; }}
    
    .page-title {{ color: {NAVY} !important; font-family:'Manrope',sans-serif; font-weight:800; font-size:clamp(27px, 3vw, 38px); letter-spacing:-1.3px; line-height:1.12; margin:0; }}
    .section-heading {{ color:{NAVY}; font-family:'Manrope',sans-serif; font-size:17px; font-weight:800; letter-spacing:-.3px; margin: 22px 0 4px 0; }}
    
    .metric-card {{ background:{WHITE}; border:1px solid #E5EAF1; border-radius:16px; padding:17px 18px 16px 18px; box-shadow:0 5px 18px rgba(16,39,66,.035); margin-bottom: 15px; }}
    .metric-label {{ color:#718096; font-size:11px; font-weight:700; letter-spacing:.8px; text-transform:uppercase; }}
    .metric-value {{ color:{NAVY}; font-family:'Manrope',sans-serif; font-size:25px; font-weight:800; margin-top:10px; line-height:1.15; }}
    
    /* BOTONES INTERACTIVOS */
    .stButton > button {{ border-radius:12px !important; background:{NAVY} !important; color:{WHITE} !important; font-weight:700 !important; border: none !important; transition: all 0.3s ease !important; }}
    .stButton > button:hover {{ background:#12375A !important; transform: translateY(-2px); box-shadow: 0 8px 15px rgba(11,29,51,0.2); }}
    </style>
    """,
    unsafe_allow_html=True,
)

# ----------------------------
# Configuración (Bases de Datos y Utilidades)
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
            pd.DataFrame(columns=CABECERAS).to_excel(ruta, index=False)

def cargar_datos(ruta):
    try:
        df = pd.read_excel(ruta)
        for col in CABECERAS:
            if col not in df.columns: df[col] = None
        return df[CABECERAS]
    except:
        return pd.DataFrame(columns=CABECERAS)

def clean_json_response(text):
    text = (text or "").strip()
    text = re.sub(r"^```(?:json)?\s*", "", text, flags=re.IGNORECASE)
    text = re.sub(r"\s*```$", "", text)
    try:
        return json.loads(text)
    except:
        return {"exito": False, "motivo_error": "Error interno al leer el comprobante."}

def metric_card(label, value):
    st.markdown(f'<div class="metric-card"><div class="metric-label">{label}</div><div class="metric-value">{value}</div></div>', unsafe_allow_html=True)

def style_plotly(fig):
    fig.update_layout(
        paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
        font=dict(family="DM Sans", color="#718096", size=11),
        title_font=dict(family="Manrope", size=16, color=NAVY),
        margin=dict(l=20, r=20, t=50, b=20)
    )
    return fig

inicializar_excels()

# ----------------------------
# Barra Lateral
# ----------------------------
with st.sidebar:
    st.markdown("### 🌙 MOON · AEROLÍNEA")
    st.markdown("**Centro de Operaciones**")
    st.markdown("---")
    vista = st.radio("Navegación:", ["✈️ Registro en pista", "📊 Panel directivo"], label_visibility="collapsed")

# ============================================================
# MÓDULO 1 | REGISTRO EN PISTA (PILOTOS)
# ============================================================
if vista == "✈️ Registro en pista":
    st.markdown('<h1 class="page-title">Registro en pista</h1>', unsafe_allow_html=True)
    
    st.markdown("""
        <div class="hero-panel">
            <div class="hero-kicker">MOON AEROLÍNEA · DIGITAL RECORD</div>
            <div class="hero-title">Cada abastecimiento, bajo control.</div>
            <div class="hero-copy">Sube el comprobante. Nuestro sistema inteligente analizará la imagen al instante.</div>
            <div class="hero-mark">M.</div>
        </div>
    """, unsafe_allow_html=True)

    col_form, col_context = st.columns([1.35, 0.85], gap="large")
    
    with col_form:
        st.markdown('<div class="section-heading">Nuevo registro de combustible</div>', unsafe_allow_html=True)
        archivo_subido = st.file_uploader("COMPROBANTE DE ABASTECIMIENTO (Terpel)", type=["jpg", "jpeg", "png"])
        procesar = st.button("Procesar y enviar a Dirección", use_container_width=True)

    with col_context:
        if archivo_subido is not None:
            st.image(archivo_subido, use_container_width=True, caption="Vista previa del recibo")

    if procesar:
        if archivo_subido is None:
            st.warning("⚠️ Debes cargar una fotografía del comprobante primero.")
        else:
            api_key = st.secrets.get("ANTHROPIC_API_KEY")
            if not api_key:
                st.error("❌ Falla técnica: No se encontró la llave de la IA.")
            else:
                st.toast("Iniciando escaneo de imagen...", icon="🔍")
                with st.spinner("La Inteligencia Artificial está validando el tiquete..."):
                    try:
                        image_bytes = archivo_subido.getvalue()
                        mime_type = archivo_subido.type or "image/jpeg"
                        base64_image = base64.b64encode(image_bytes).decode("utf-8")
                        
                        client = anthropic.Anthropic(api_key=api_key)
                        
                        prompt = """
                        Eres el asistente de revisión de tiquetes de combustible de Moon Aerolínea.
                        Analiza la imagen. Debes extraer: recibo, fecha_consumo, matricula, producto, ruta, consumo, piloto.
                        
                        REGLA ESTRICTA:
                        Si la imagen es muy borrosa, está cortada, o NO SE LEE CLARAMENTE el volumen de galones ("consumo"), 
                        debes devolver "exito": false y redactar un "motivo_error" amigable explicando al piloto por qué debe tomar de nuevo la foto.
                        
                        Ejemplo de error: {"exito": false, "motivo_error": "¡Hola piloto! La parte inferior del recibo está cortada y no logro ver los galones. ¿Podrías tomar la foto un poco más lejos?"}
                        
                        Si todo está perfecto, devuelve:
                        {"exito": true, "motivo_error": "", "recibo": "...", "fecha_consumo": "...", "matricula": "...", "producto": "...", "ruta": "...", "consumo": 0.0, "piloto": "..."}
                        """
                        
                        response = client.messages.create(
                            model="claude-3-haiku-20240307",
                            max_tokens=1024,
                            messages=[{"role": "user", "content": [{"type": "image", "source": {"type": "base64", "media_type": mime_type, "data": base64_image}}, {"type": "text", "text": prompt}]}]
                        )
                        
                        datos = clean_json_response(response.content[0].text)

                        if not datos.get("exito"):
                            st.error(f"📸 **Problema con la foto:** {datos.get('motivo_error', 'La imagen no es clara. Por favor, tómala de nuevo.')}")
                        else:
                            galones = float(datos.get("consumo", 0.0) or 0.0)
                            fecha_str = datetime.now().strftime("%Y%m%d_%H%M%S")
                            num_recibo = str(datos.get("recibo", "SREC")).replace("/", "-")
                            nombre_imagen = f"Recibo_{num_recibo}_{fecha_str}.jpg"
                            
                            with open(os.path.join(CARPETA_IMAGENES, nombre_imagen), "wb") as f:
                                f.write(image_bytes)

                            df_diario = pd.read_excel(RUTA_DIARIO)
                            nuevo_registro = pd.DataFrame([{
                                "FECHA REGISTRO": datetime.now().strftime("%Y-%m-%d %H:%M:%S"), "RECIBO": datos.get("recibo"),
                                "FECHA CONSUMO": datos.get("fecha_consumo"), "MATRICULA": datos.get("matricula"),
                                "PRODUCTO": datos.get("producto"), "RUTA": datos.get("ruta"), "CONSUMO (GAL)": galones,
                                "VALOR UNIT": 0.0, "VALOR TOTAL": 0.0, "PILOTO": datos.get("piloto"),
                                "NOMENCLATURA IMAGEN": nombre_imagen, "OBSERVACION": "Enviado desde Pista"
                            }])
                            df_diario = pd.concat([df_diario, nuevo_registro], ignore_index=True)
                            df_diario.to_excel(RUTA_DIARIO, index=False)

                            st.balloons()
                            st.success(f"✅ ¡Excelente! Tiquete procesado correctamente ({galones:,.1f} GAL para {datos.get('matricula')}).")
                            st.info("La información ya está en la bandeja de Dirección para asignación de precios.")

                    except Exception as exc:
                        st.error("❌ Ocurrió un error de conexión con la IA. Intenta de nuevo.")

# ============================================================
# MÓDULO 2 | PANEL DIRECTIVO (DASHBOARDS & EDICIÓN)
# ============================================================
elif vista == "📊 Panel directivo":
    st.markdown('<h1 class="page-title">Panel directivo</h1>', unsafe_allow_html=True)
    
    if "moon_auth" not in st.session_state:
        st.session_state.moon_auth = False

    if not st.session_state.moon_auth:
        with st.form("login"):
            clave = st.text_input("🔐 CONTRASEÑA EJECUTIVA", type="password")
            ingresar = st.form_submit_button("Acceder al Sistema")
        if ingresar:
            if hmac.compare_digest(str(clave), str(st.secrets.get("ADMIN_PASSWORD", "adminmoon"))):
                st.session_state.moon_auth = True
                st.rerun()
            else:
                st.error("❌ Acceso denegado.")
    else:
        col_header, col_logout = st.columns([5,1])
        with col_logout:
            if st.button("Cerrar Sesión"):
                st.session_state.moon_auth = False
                st.rerun()

        st.markdown("""
            <div class="hero-panel">
                <div class="hero-kicker">FINANCIAL & OPERATIONS HUB</div>
                <div class="hero-title">Centro de Mando Administrativo.</div>
                <div class="hero-copy">Monitorea los indicadores de Power BI en tiempo real, <strong>edita los datos haciendo doble clic en las tablas</strong>, asigna precios y consolida reportes.</div>
            </div>
        """, unsafe_allow_html=True)
        
        tab1, tab2 = st.tabs(["📅 DASHBOARD DIARIO (Cierre)", "📊 DASHBOARD MAESTRO (Mensual)"])

        # -------------------------------------
        # PESTAÑA 1: DASHBOARD DIARIO
        # -------------------------------------
        with tab1:
            df_diario = cargar_datos(RUTA_DIARIO)
            
            if df_diario.empty:
                st.info("🍃 No hay tiquetes pendientes de procesar hoy.")
            else:
                # Asegurar formatos numéricos para los gráficos
                df_diario["CONSUMO (GAL)"] = pd.to_numeric(df_diario["CONSUMO (GAL)"], errors='coerce').fillna(0)
                df_diario["VALOR TOTAL"] = pd.to_numeric(df_diario["VALOR TOTAL"], errors='coerce').fillna(0)
                
                # --- KPI DIARIOS ---
                c1, c2, c3 = st.columns(3)
                with c1: metric_card("Tiquetes Hoy", f"{len(df_diario):,}")
                with c2: metric_card("Volumen Hoy", f"{df_diario['CONSUMO (GAL)'].sum():,.1f} GAL")
                with c3: metric_card("Gasto Diario Estimado", f"$ {df_diario['VALOR TOTAL'].sum():,.0f} COP")
                
                # --- GRÁFICOS DIARIOS (Power BI style) ---
                g1, g2 = st.columns(2)
                with g1:
                    df_ruta = df_diario.groupby("RUTA", as_index=False)["CONSUMO (GAL)"].sum().sort_values("CONSUMO (GAL)", ascending=False)
                    fig_ruta = px.bar(df_ruta, x="RUTA", y="CONSUMO (GAL)", title="Galones por Ruta (Hoy)", color_discrete_sequence=[BLUE], text_auto='.1f')
                    st.plotly_chart(style_plotly(fig_ruta), use_container_width=True)
                with g2:
                    # Mostrar qué pilotos enviaron facturas hoy y cuántos tiquetes
                    df_pilotos = df_diario["PILOTO"].value_counts().reset_index()
                    df_pilotos.columns = ["Piloto", "Facturas"]
                    fig_pilotos = px.pie(df_pilotos, names="Piloto", values="Facturas", title="Facturas enviadas por Piloto (Hoy)", hole=0.5, color_discrete_sequence=px.colors.sequential.Blues_r)
                    st.plotly_chart(style_plotly(fig_pilotos), use_container_width=True)

                st.markdown('<div class="section-heading">Bandeja de Entrada: Tabla Editable</div>', unsafe_allow_html=True)
                st.write("✏️ *Puedes hacer doble clic en cualquier celda para corregir datos antes de guardar.*")
                
                # TABLA INTERACTIVA Y EDITABLE
                df_diario_editado = st.data_editor(df_diario, num_rows="dynamic", use_container_width=True, key="editor_diario")
                
                if st.button("💾 Guardar correcciones manuales"):
                    df_diario_editado.to_excel(RUTA_DIARIO, index=False)
                    st.success("¡Base diaria actualizada!")
                
                st.markdown("---")
                st.markdown('<div class="section-heading">Fijar Precio y Consolidar</div>', unsafe_allow_html=True)
                c_p1, c_p2 = st.columns(2)
                
                with c_p1:
                    precio_dia = st.number_input("💵 Precio del combustible hoy (COP):", value=19500.0, step=100.0)
                    if st.button("Aplicar Precio al Excel y Gráficos"):
                        df_diario_editado["VALOR UNIT"] = precio_dia
                        df_diario_editado["VALOR TOTAL"] = pd.to_numeric(df_diario_editado["CONSUMO (GAL)"], errors='coerce').fillna(0) * precio_dia
                        df_diario_editado["OBSERVACION"] = "Validado por Dirección"
                        df_diario_editado.to_excel(RUTA_DIARIO, index=False)
                        st.success("✅ Cálculos generados. Los gráficos de Gasto Diario se han actualizado.")
                        st.rerun()

                with c_p2:
                    with open(RUTA_DIARIO, "rb") as f:
                        st.download_button("⬇️ 1. Descargar Excel del Día", data=f.read(), file_name=f"Moon_Diario_{datetime.now().strftime('%Y%m%d')}.xlsx", use_container_width=True)
                    
                    if st.button("📦 2. Archivar en Histórico y Limpiar Día", use_container_width=True):
                        df_maestro = cargar_datos(RUTA_MAESTRO)
                        df_maestro = pd.concat([df_maestro, df_diario_editado], ignore_index=True)
                        df_maestro.to_excel(RUTA_MAESTRO, index=False)
                        pd.DataFrame(columns=CABECERAS).to_excel(RUTA_DIARIO, index=False)
                        st.balloons()
                        st.success("Día cerrado exitosamente. Datos movidos al Maestro.")
                        time.sleep(2)
                        st.rerun()

        # -------------------------------------
        # PESTAÑA 2: DASHBOARD MAESTRO (Histórico)
        # -------------------------------------
        with tab2:
            df_maestro = cargar_datos(RUTA_MAESTRO)
            if df_maestro.empty:
                st.info("El historial está vacío por ahora. Cierra un día para ver los gráficos históricos.")
            else:
                df_maestro["CONSUMO (GAL)"] = pd.to_numeric(df_maestro["CONSUMO (GAL)"], errors='coerce').fillna(0)
                df_maestro["VALOR TOTAL"] = pd.to_numeric(df_maestro["VALOR TOTAL"], errors='coerce').fillna(0)
                
                # --- KPI MAESTROS ---
                st.markdown('<div class="section-heading">Indicadores Globales (Consolidado)</div>', unsafe_allow_html=True)
                k1, k2, k3 = st.columns(3)
                with k1: metric_card("Tiquetes Históricos", f"{len(df_maestro):,}")
                with k2: metric_card("Galones Históricos", f"{df_maestro['CONSUMO (GAL)'].sum():,.1f} GAL")
                with k3: metric_card("Inversión Histórica", f"$ {df_maestro['VALOR TOTAL'].sum():,.0f} COP")

                # --- GRÁFICOS MAESTROS (Power BI style) ---
                mg1, mg2 = st.columns(2)
                with mg1:
                    df_mat = df_maestro.groupby("MATRICULA", as_index=False)["CONSUMO (GAL)"].sum().sort_values("CONSUMO (GAL)")
                    fig_mat = px.bar(df_mat, x="CONSUMO (GAL)", y="MATRICULA", orientation='h', title="Consumo Histórico por Aeronave", color_discrete_sequence=[NAVY], text_auto='.1f')
                    st.plotly_chart(style_plotly(fig_mat), use_container_width=True)
                with mg2:
                    df_gasto_ruta = df_maestro.groupby("RUTA", as_index=False)["VALOR TOTAL"].sum()
                    fig_gasto = px.pie(df_gasto_ruta, names="RUTA", values="VALOR TOTAL", title="Distribución de Gasto por Destino", hole=0.4, color_discrete_sequence=px.colors.sequential.Blues_r)
                    fig_gasto.update_traces(textposition='inside', textinfo='percent+label')
                    st.plotly_chart(style_plotly(fig_gasto), use_container_width=True)

                st.markdown('<div class="section-heading">Base de Datos General (Tabla Editable)</div>', unsafe_allow_html=True)
                st.write("✏️ *Si notas un error de días pasados, puedes editarlo directamente en esta tabla y guardar.*")
                
                # TABLA HISTÓRICA EDITABLE
                df_maestro_editado = st.data_editor(df_maestro, num_rows="dynamic", use_container_width=True, key="editor_maestro")
                
                col_save, col_down = st.columns(2)
                with col_save:
                    if st.button("💾 Guardar cambios en Histórico"):
                        df_maestro_editado.to_excel(RUTA_MAESTRO, index=False)
                        st.success("¡Historial actualizado! Los gráficos generales se han recalculado.")
                with col_down:
                    with open(RUTA_MAESTRO, "rb") as f:
                        st.download_button("⬇️ Descargar Excel Histórico Mensual", data=f.read(), file_name="Moon_Base_Maestra_General.xlsx", use_container_width=True)
