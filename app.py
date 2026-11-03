import os
import json
import openpyxl
import pandas as pd
import streamlit as st
import plotly.express as px
import shutil
from datetime import datetime
from google import genai
from google.genai import types

# ==========================================
# 1. CONFIGURACIÓN INICIAL Y ESTÉTICA (MOON)
# ==========================================
st.set_page_config(page_title="Ops Combustible - Moon", page_icon="🌙", layout="wide")

# CSS Personalizado: Colores Azul/Blanco y Animación de Avión Flotando
st.markdown("""
<style>
    /* Animación del avión flotando */
    @keyframes flotar {
        0% { transform: translateY(0px) translateX(0px); }
        50% { transform: translateY(-10px) translateX(0px); }
        100% { transform: translateY(0px) translateX(0px); }
    }
    .avion-animado {
        font-size: 3rem;
        display: inline-block;
        animation: flotar 3s ease-in-out infinite;
        margin-right: 15px;
    }
    
    /* Tema de colores Azul y Blanco (Marca Moon) */
    h1, h2, h3, h4 {
        color: #003366 !important; 
    }
    
    /* Estilo de los botones principales */
    .stButton>button {
        background-color: #004080 !important;
        color: white !important;
        border-radius: 8px;
        border: 2px solid #0059b3;
        transition: all 0.3s ease;
    }
    .stButton>button:hover {
        background-color: #0059b3 !important;
        box-shadow: 0px 4px 10px rgba(0, 51, 102, 0.3);
        transform: scale(1.02);
    }
</style>
""", unsafe_allow_html=True)

# Conexión Segura con Google AI (Gemini)
try:
    client = genai.Client(api_key=st.secrets["GEMINI_API_KEY"])
except Exception:
    client = genai.Client()

CARPETA_IMAGENES = "imagenes_tiquetes"
os.makedirs(CARPETA_IMAGENES, exist_ok=True)
RUTA_EXCEL = "Base_Datos_Combustible_Moon.xlsx"

def inicializar_excel():
    if not os.path.exists(RUTA_EXCEL):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Consumos"
        cabeceras = [
            "FECHA REGISTRO", "RECIBO", "FECHA CONSUMO", "MATRICULA", 
            "PRODUCTO", "RUTA", "CONSUMO (GAL)", "VALOR UNIT", 
            "VALOR TOTAL", "PILOTO", "NOMENCLATURA IMAGEN", "OBSERVACION"
        ]
        ws.append(cabeceras)
        wb.save(RUTA_EXCEL)

inicializar_excel()

# ==========================================
# 2. BARRA DE NAVEGACIÓN
# ==========================================
st.sidebar.markdown("<h2 style='text-align: center; color: #003366;'>🌙 Moon Ops</h2>", unsafe_allow_html=True)
st.sidebar.markdown("---")
vista = st.sidebar.radio("Navegación:", ["📸 Registro en Pista", "📊 Panel Directivo"])

# ==========================================
# VISTA 1: PILOTOS (Registro en Pista)
# ==========================================
if vista == "📸 Registro en Pista":
    # Encabezado interactivo con avión animado
    st.markdown("""
        <div style="display: flex; align-items: center; justify-content: center; background-color: #f0f8ff; padding: 20px; border-radius: 10px; margin-bottom: 20px; border-left: 5px solid #003366;">
            <div class="avion-animado">✈️</div>
            <h1 style="margin: 0; color: #003366;">Registro de Combustible</h1>
        </div>
    """, unsafe_allow_html=True)
    
    st.info("⚠️ **IMPORTANTE:** Sube **UN SOLO** recibo por fotografía para evitar confusiones en el sistema.")

    archivo_subido = st.file_uploader("Capturar foto del recibo Terpel", type=["jpg", "jpeg", "png"], label_visibility="collapsed")

    if archivo_subido is not None:
        st.image(archivo_subido, caption="Foto lista para enviar", use_container_width=True)
        
        precio_unitario = st.number_input("Precio por galón (COP) hoy:", min_value=0.0, value=19500.0, step=100.0)

        if st.button("🚀 Enviar y Procesar Tiquete", use_container_width=True, type="primary"):
            with st.spinner("La IA de Moon está analizando el tiquete..."):
                try:
                    image_bytes = archivo_subido.getvalue()

                    prompt = """
                    Analiza esta imagen de un tiquete de combustible Terpel firmado a mano. Extrae:
                    - "recibo": Número del Tiquete (ej. 02-1243808).
                    - "fecha_consumo": Fecha en formato YYYY-MM-DD.
                    - "matricula": Matrícula del avión.
                    - "producto": Tipo de combustible (ej. JETA1).
                    - "ruta": Destino (ej. EOH, PSO).
                    - "consumo": Volumen total en galones (SOLO el número, ej: 320.5).
                    - "piloto": Nombre manuscrito de quien firma abajo.

                    Devuelve UNICAMENTE un JSON válido con este formato exacto:
                    {"exito": true, "recibo": "...", "fecha_consumo": "...", "matricula": "...", "producto": "...", "ruta": "...", "consumo": 0.0, "piloto": "..."}
                    Si la foto es ilegible, devuelve: {"exito": false}
                    """

                    response = client.models.generate_content(
                        model='gemini-3.8-flash',
                        contents=[types.Part.from_bytes(data=image_bytes, mime_type='image/jpeg'), prompt]
                    )
                    
                    texto_limpio = response.text.replace('```json', '').replace('```', '').strip()
                    datos = json.loads(texto_limpio)

                    if not datos.get("exito"):
                        st.error("⚠️ La foto está muy borrosa. Por favor, tómala de nuevo.")
                    else:
                        fecha_str = datetime.now().strftime("%Y%m%d_%H%M%S")
                        num_recibo = str(datos.get("recibo", "SREC")).replace("/", "-").replace(" ", "")
                        nombre_imagen = f"Recibo_{num_recibo}_{fecha_str}.jpg"
                        ruta_guardado = os.path.join(CARPETA_IMAGENES, nombre_imagen)
                        
                        with open(ruta_guardado, "wb") as f:
                            f.write(image_bytes)

                        galones = float(datos.get("consumo", 0.0))
                        valor_total = galones * float(precio_unitario)

                        wb = openpyxl.load_workbook(RUTA_EXCEL)
                        hoja = wb.active
                        
                        nueva_fila = [
                            datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                            datos.get("recibo"),
                            datos.get("fecha_consumo"),
                            datos.get("matricula"),
                            datos.get("producto"),
                            datos.get("ruta"),
                            galones,
                            precio_unitario,
                            valor_total,
                            datos.get("piloto"),
                            nombre_imagen,
                            "Procesado automáticamente"
                        ]
                        
                        hoja.append(nueva_fila)
                        wb.save(RUTA_EXCEL)

                        st.success("✅ ¡Tiquete registrado y guardado con éxito!")
                        
                        col_r1, col_r2 = st.columns(2)
                        col_r1.metric("Galones Registrados", f"{galones} GAL")
                        col_r2.metric("Valor Estimado", f"$ {valor_total:,.0f} COP")

                except Exception as e:
                    st.error(f"❌ Ocurrió un error inesperado al leer: {e}")

# ==========================================
# VISTA 2: PANEL DIRECTIVO (Protegido con Clave)
# ==========================================
elif vista == "📊 Panel Directivo":
    st.markdown("<h1 style='color: #003366;'>🔒 Acceso Administrativo - Moon</h1>", unsafe_allow_html=True)
    
    # Sistema de Seguridad (Clave)
    clave = st.text_input("Ingrese la contraseña para acceder al panel de control:", type="password")
    
    if clave != "adminmoon" and clave != "":
        st.error("❌ Contraseña incorrecta. Acceso denegado.")
    
    elif clave == "adminmoon":
        st.success("✅ Acceso concedido.")
        st.markdown("---")
        
        try:
            df = pd.read_excel(RUTA_EXCEL)
        except Exception:
            df = pd.DataFrame()

        if len(df) > 0:
            df["CONSUMO (GAL)"] = pd.to_numeric(df["CONSUMO (GAL)"], errors='coerce').fillna(0)
            df["VALOR TOTAL"] = pd.to_numeric(df["VALOR TOTAL"], errors='coerce').fillna(0)

            st.subheader("📊 Resumen Financiero y Operativo")
            col1, col2, col3 = st.columns(3)
            
            col1.metric("Tiquetes Totales", len(df))
            col2.metric("Total Galones", f"{df['CONSUMO (GAL)'].sum():,.1f} GAL")
            col3.metric("Gasto Estimado", f"$ {df['VALOR TOTAL'].sum():,.0f} COP")
            
            st.markdown("---")

            col_g1, col_g2 = st.columns(2)
            with col_g1:
                df_matricula = df.groupby("MATRICULA")["CONSUMO (GAL)"].sum().reset_index()
                # Gráfico azul corporativo
                fig_mat = px.bar(df_matricula, x="MATRICULA", y="CONSUMO (GAL)", 
                                 title="Galones por Aeronave",
                                 color_discrete_sequence=['#004080'])
                st.plotly_chart(fig_mat, use_container_width=True)

            with col_g2:
                df_ruta = df.groupby("RUTA")["VALOR TOTAL"].sum().reset_index()
                # Gráfico de pastel en tonos azules
                fig_ruta = px.pie(df_ruta, names="RUTA", values="VALOR TOTAL", 
                                  title="Dinero Invertido por Ruta",
                                  color_discrete_sequence=px.colors.sequential.Blues_r)
                st.plotly_chart(fig_ruta, use_container_width=True)

            st.subheader("📑 Base de Datos Maestra (Excel)")
            st.dataframe(df, use_container_width=True)

            with open(RUTA_EXCEL, "rb") as f:
                st.download_button(
                    label="📥 Descargar Base de Datos Completa",
                    data=f,
                    file_name="Moon_Control_Combustible.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            
            st.markdown("---")
            
            # ZONA DE PELIGRO: Borrar Datos Específicos o Totales
            with st.expander("⚙️ Zona de Mantenimiento (Borrar Datos)"):
                
                # OPCIÓN 1: Borrar un registro específico
                st.markdown("#### 🗑️ Borrar un tiquete específico (Por error o prueba)")
                opciones_borrar = df["NOMENCLATURA IMAGEN"].tolist()
                
                if opciones_borrar:
                    registro_a_borrar = st.selectbox(
                        "Selecciona el archivo del tiquete que deseas eliminar:", 
                        opciones_borrar,
                        help="Busca el nombre del archivo en la tabla de arriba en la columna 'NOMENCLATURA IMAGEN'"
                    )
                    
                    if st.button("Borrar Tiquete Seleccionado"):
                        # Borrar la imagen física
                        ruta_img_borrar = os.path.join(CARPETA_IMAGENES, registro_a_borrar)
                        if os.path.exists(ruta_img_borrar):
                            os.remove(ruta_img_borrar)
                        
                        # Borrar la fila del Excel
                        wb = openpyxl.load_workbook(RUTA_EXCEL)
                        hoja = wb.active
                        # Recorremos el excel para buscar esa nomenclatura exacta (Columna K es la 11)
                        for row in range(2, hoja.max_row + 1):
                            if str(hoja.cell(row=row, column=11).value) == str(registro_a_borrar):
                                hoja.delete_rows(row)
                                break # Rompe el ciclo cuando lo encuentra y lo borra
                        wb.save(RUTA_EXCEL)
                        
                        st.success(f"✅ Se eliminó correctamente el recibo '{registro_a_borrar}'. Por favor recarga la página para actualizar las gráficas y la tabla.")
                else:
                    st.info("No hay recibos individuales para borrar.")

                st.markdown("---")
                
                # OPCIÓN 2: Borrar todo (Formateo Total)
                st.markdown("#### 🚨 Formateo Total del Sistema")
                st.warning("⚠️ Si haces clic aquí, se borrarán TODOS los registros y la base de datos quedará en cero.")
                if st.button("Borrar Todos los Registros del Sistema"):
                    if os.path.exists(RUTA_EXCEL):
                        os.remove(RUTA_EXCEL)
                    if os.path.exists(CARPETA_IMAGENES):
                        shutil.rmtree(CARPETA_IMAGENES)
                    
                    inicializar_excel()
                    os.makedirs(CARPETA_IMAGENES, exist_ok=True)
                    st.success("✅ Base de datos borrada exitosamente desde cero. Por favor, recarga la página web.")
        else:
            st.info("🟢 El panel está listo. No hay registros en la base de datos aún.")
