import os
import json
import openpyxl
import pandas as pd
import streamlit as st
import plotly.express as px
from datetime import datetime
from google import genai
from google.genai import types

# ==========================================
# 1. CONFIGURACIÓN INICIAL DEL SISTEMA
# ==========================================
st.set_page_config(page_title="Ops Combustible - Moon Flight", page_icon="✈️", layout="wide")

# Conexión Segura con Google AI (Gemini)
try:
    cliente_genai = genai.Client(api_key=st.secrets["GEMINI_API_KEY"])
except Exception:
    # Si corre localmente y no hay secrets, intenta leer las variables de entorno de la PC
    cliente_genai = genai.Client()

CARPETA_IMAGENES = "imagenes_tiquetes"
os.makedirs(CARPETA_IMAGENES, exist_ok=True)
RUTA_EXCEL = "Base_Datos_Combustible_Moon.xlsx"

# Función robusta para crear el Excel desde cero si no existe
def inicializar_excel():
    if not os.path.exists(RUTA_EXCEL):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Consumos"
        # Exactamente 12 columnas ordenadas como solicitaste
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
st.sidebar.title("✈️ Moon Flight Ops")
st.sidebar.markdown("---")
vista = st.sidebar.radio("Navegación:", ["📸 Registro en Pista", "📊 Panel Directivo"])

# ==========================================
# VISTA 1: PILOTOS (Registro en Pista)
# ==========================================
if vista == "📸 Registro en Pista":
    st.title("Registro Rápido de Tiquetes")
    st.markdown("Sube la foto del tiquete Terpel. La IA completará el reporte en Excel automáticamente.")

    archivo_subido = st.file_uploader("Capturar foto del recibo", type=["jpg", "jpeg", "png"], label_visibility="collapsed")

    if archivo_subido is not None:
        st.image(archivo_subido, caption="Foto lista para enviar", use_container_width=True)
        
        # Precio configurable por el operario antes de enviar
        precio_unitario = st.number_input("Precio por galón (COP) hoy:", min_value=0.0, value=19500.0, step=100.0)

        if st.button("🚀 Enviar y Procesar Tiquete", use_container_width=True, type="primary"):
            with st.spinner("La Inteligencia Artificial está leyendo los datos..."):
                try:
                    # Leer bytes de la imagen
                    image_bytes = archivo_subido.getvalue()

                    # Prompt blindado
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

                    respuesta = cliente_genai.models.generate_content(
                        model='gemini-2.5-flash',
                        contents=[types.Part.from_bytes(data=image_bytes, mime_type='image/jpeg'), prompt]
                    )
                    
                    # Limpieza extrema del texto para evitar errores de formato
                    texto_limpio = respuesta.text.replace('```json', '').replace('```', '').strip()
                    datos = json.loads(texto_limpio)

                    if not datos.get("exito"):
                        st.error("⚠️ La foto está muy borrosa. Por favor, tómala de nuevo asegurándote de enfocar bien el papel.")
                    else:
                        # 1. Crear Nomenclatura Única
                        fecha_str = datetime.now().strftime("%Y%m%d_%H%M%S")
                        num_recibo = str(datos.get("recibo", "SREC")).replace("/", "-").replace(" ", "")
                        nombre_imagen = f"Recibo_{num_recibo}_{fecha_str}.jpg"
                        ruta_guardado = os.path.join(CARPETA_IMAGENES, nombre_imagen)
                        
                        # 2. Guardar Imagen Físicamente
                        with open(ruta_guardado, "wb") as f:
                            f.write(image_bytes)

                        # 3. Cálculos Financieros Asegurados (forzando formato float numérico)
                        galones = float(datos.get("consumo", 0.0))
                        valor_total = galones * float(precio_unitario)

                        # 4. Guardar en Excel
                        wb = openpyxl.load_workbook(RUTA_EXCEL)
                        hoja = wb.active
                        
                        # Exactamente 12 datos para 12 columnas
                        nueva_fila = [
                            datetime.now().strftime("%Y-%m-%d %H:%M:%S"), # 1. FECHA REGISTRO
                            datos.get("recibo"),                          # 2. RECIBO
                            datos.get("fecha_consumo"),                   # 3. FECHA CONSUMO
                            datos.get("matricula"),                       # 4. MATRICULA
                            datos.get("producto"),                        # 5. PRODUCTO
                            datos.get("ruta"),                            # 6. RUTA
                            galones,                                      # 7. CONSUMO (GAL) -> Numérico
                            precio_unitario,                              # 8. VALOR UNIT -> Numérico
                            valor_total,                                  # 9. VALOR TOTAL -> Numérico
                            datos.get("piloto"),                          # 10. PILOTO
                            nombre_imagen,                                # 11. NOMENCLATURA IMAGEN
                            "Procesado exitosamente por IA"               # 12. OBSERVACION
                        ]
                        
                        hoja.append(nueva_fila)
                        wb.save(RUTA_EXCEL)

                        # 5. Confirmación Visual al Piloto
                        st.success("✅ ¡Tiquete registrado y guardado en la base de datos maestra!")
                        st.info(f"💾 Código de Archivo: **{nombre_imagen}**")
                        
                        col_r1, col_r2 = st.columns(2)
                        col_r1.metric("Galones Registrados", f"{galones} GAL")
                        col_r2.metric("Valor Estimado", f"$ {valor_total:,.0f} COP")

                except json.JSONDecodeError:
                    st.error("⚠️ La IA no pudo estructurar la información. Intenta con una foto más recta.")
                except Exception as e:
                    st.error(f"❌ Ocurrió un error inesperado: {e}")

# ==========================================
# VISTA 2: PANEL DIRECTIVO Y DASHBOARDS
# ==========================================
elif vista == "📊 Panel Directivo":
    st.title("📊 Panel de Inteligencia de Negocios (BI) - Moon Flight")
    
    # Intentar leer la base de datos
    try:
        df = pd.read_excel(RUTA_EXCEL)
    except Exception:
        st.error("Error crítico leyendo el archivo de base de datos.")
        df = pd.DataFrame()

    if len(df) > 0:
        # Forzar formatos numéricos para los gráficos (Auditoría de seguridad)
        df["CONSUMO (GAL)"] = pd.to_numeric(df["CONSUMO (GAL)"], errors='coerce').fillna(0)
        df["VALOR TOTAL"] = pd.to_numeric(df["VALOR TOTAL"], errors='coerce').fillna(0)

        # --- 1. KPI GLOBALES ---
        st.subheader("Resumen Financiero y Operativo")
        col1, col2, col3, col4 = st.columns(4)
        
        tiquetes_hoy = len(df[df["FECHA REGISTRO"].astype(str).str.contains(datetime.now().strftime("%Y-%m-%d"))])
        
        col1.metric("Tiquetes Totales", len(df))
        col2.metric("Operaciones de Hoy", tiquetes_hoy)
        col3.metric("Total Galones", f"{df['CONSUMO (GAL)'].sum():,.1f} GAL")
        col4.metric("Costo Total Invertido", f"$ {df['VALOR TOTAL'].sum():,.0f} COP")
        
        st.markdown("---")

        # --- 2. GRÁFICOS INTERACTIVOS ---
        st.subheader("Análisis de Consumo")
        col_g1, col_g2 = st.columns(2)
        
        with col_g1:
            # Gráfico de barras: Consumo por avión
            df_matricula = df.groupby("MATRICULA")["CONSUMO (GAL)"].sum().reset_index()
            fig_mat = px.bar(df_matricula, x="MATRICULA", y="CONSUMO (GAL)", 
                             title="Galones consumidos por Aeronave",
                             labels={"CONSUMO (GAL)": "Galones", "MATRICULA": "Avión"},
                             color="MATRICULA")
            st.plotly_chart(fig_mat, use_container_width=True)

        with col_g2:
            # Gráfico circular: Gasto por destino
            df_ruta = df.groupby("RUTA")["VALOR TOTAL"].sum().reset_index()
            fig_ruta = px.pie(df_ruta, names="RUTA", values="VALOR TOTAL", 
                              title="Distribución de Dinero Invertido por Ruta",
                              hole=0.4)
            st.plotly_chart(fig_ruta, use_container_width=True)

        st.markdown("---")

        # --- 3. TABLA Y DESCARGAS ---
        st.subheader("Base de Datos Maestra (Excel)")
        st.dataframe(df, use_container_width=True)

        with open(RUTA_EXCEL, "rb") as f:
            st.download_button(
                label="📥 Descargar Archivo Excel Actualizado",
                data=f,
                file_name="MoonFlight_Control_Combustible.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary"
            )
        
        st.markdown("---")
        
        # --- 4. AUDITORÍA VISUAL ---
        st.subheader("🔍 Auditoría Visual de Recibos y Firmas")
        imagenes = os.listdir(CARPETA_IMAGENES)
        
        if imagenes:
            img_sel = st.selectbox("Buscar recibo físico guardado (Nomenclatura):", reversed(imagenes))
            ruta_img = os.path.join(CARPETA_IMAGENES, img_sel)
            st.image(ruta_img, caption=f"Evidencia: {img_sel}", width=500)
            
            # Botón de descarga individual
            with open(ruta_img, "rb") as f:
                st.download_button(f"📥 Descargar {img_sel}", data=f, file_name=img_sel)
        else:
            st.info("La base de datos tiene registros, pero las imágenes fuente fueron limpiadas o movidas.")
            
    else:
        # Estado 0: Cuando el Excel apenas se creó y no hay ningún tiquete procesado
        st.info("🟢 El sistema está operativo y el archivo Excel maestro está inicializado, pero aún no se ha subido el primer tiquete de combustible.")
        st.markdown("¡Ve a la vista **Registro en Pista** para subir la primera foto de prueba!")
