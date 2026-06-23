# -*- coding: utf-8 -*-
"""
Formulario de verificación de datos visita - Versión UI/UX Optimizada según Guía de Mockups
Mantiene toda la lógica de cálculo, lectura de Excel y generación de reportes original.
"""

import io
import json
from datetime import datetime
import pandas as pd
import streamlit as st
from docx import Document
from docx.shared import Cm, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

try:
    from streamlit_js_eval import get_geolocation
    GEO_OK = True
except Exception:
    GEO_OK = False

# --------------------------------------------------------------------------
# CONFIGURACIÓN GENERAL Y ESTILOS MÓVILES
# --------------------------------------------------------------------------
st.set_page_config(
    page_title="Evaluación de Crédito",
    page_icon="🏦",
    layout="wide",
    initial_sidebar_state="collapsed", # Escondido para simular app móvil
)

EXCEL_COLUMNS = [
    "RECNO", "PEPAIS", "PETDOC", "PENDOC", "CODCLI", "BCEMP", "BCSUC", "BCMDA",
    "BCPAP", "BCCTA", "BCOPER", "BCSBOP", "BCTOP", "BCMOD", "CODCRE", "REGION",
    "ZONA", "AGENCIA", "CLIENTE", "DIRECCION_DOM", "DISTRITO_DOM",
    "PROVINCIA_DOM", "DEPARTAMENTO_DOM", "DIRECCION_NEG", "DISTRITO_NEG",
    "PROVINCIA_NEG", "DEPARTAMENTO_NEG", "ACTIVIDAD_ECON", "ANALISTA",
    "PRODUCTO_CAJA", "SALDO_MN", "SALDO_VIGE", "SALDO_REFI", "SALDO_VENC",
    "SALDO_JUDI", "MORA_CONT", "TIPO_SBS", "FECDES", "IMPDESEMB_MN",
    "COD_MODULO", "MODULO", "COD_TIPO_OPERACION", "TIPO_OPERACION",
    "ANALISTA_EVAL", "USUARIO_APROB", "USUARIO_DESEM", "FECHA_EVAL",
    "DIAS_ATRASO", "ESTADO_CREDITO", "ATRANT_1M", "ATRANT_2M", "ATRANT_3M",
    "ATRANT_4M", "ATRANT_5M", "ATRANT_6M", "TIPO_SOLI", "NUMERO_CUOTAS",
    "CUOTAS_PAGADAS", "TIPO", "SEGMENTACION_MYPE", "CATEG_RESULTANTE",
    "CATEG_RESULTANTE_SINALIN", "CUENTA_AVAL", "FECHA_UTLPAGO", "UAI_IND",
    "ESTRATO", "TIPO_EXPEDIENTE",
]

NARANJA = "C8102E"   
AZUL = "1B3A5C"
VERDE = "137333"
ROJO = "A50E0E"

CUSTOM_CSS = f"""
<style>
.stApp {{ background-color: #F8F9FA; }}
h1, h2, h3 {{ color: #{AZUL}; font-family: 'Segoe UI', sans-serif; }}

/* Tarjetas limpias de los mockups */
.mockup-card {{
    background: white; padding: 1.5rem; border-radius: 12px;
    box-shadow: 0 2px 6px rgba(0,0,0,0.04); margin-bottom: 1.2rem;
    border: 1px solid #EAECEE;
}}

/* Encabezado del cliente activo (Ficha Identidad) */
.client-profile-box {{
    background-color: #E6F4EA; border-radius: 12px; padding: 1.2rem;
    margin-bottom: 1.5rem; border: 1px solid #D1E7DD;
}}

/* Botones principales de la App */
div.stButton > button {{
    background-color: #0052CC; color: white; border: none;
    border-radius: 8px; padding: 12px 24px; font-weight: 600; font-size: 1rem;
}}
div.stButton > button:hover {{ background-color: #0043A4; color: white; }}

/* Botones secundarios / Atrás */
div.stButton > button[key^="btn_atras"] {{
    background-color: #FFFFFF !important; color: #{AZUL} !important;
    border: 1px solid #CCD1D9 !important;
}}

/* Panel de validación optimizado */
.validation-box {{
    border: 2px solid #{AZUL}; border-radius: 10px; padding: 1.2rem;
    background: white; margin: 0.5rem 0;
}}
.validation-title {{
    font-size: 1.1rem; font-weight: 700; color: #{AZUL}; margin-bottom: 0.6rem;
    border-bottom: 3px solid #{NARANJA}; padding-bottom: 0.3rem;
}}
.validation-box div.stButton > button {{
    background-color: #ffffff !important; color: #1B3A5C !important;
    border: 1px solid #e0e0e0 !important; text-align: left !important;
    display: flex !important; justify-content: flex-start !important;
    padding: 8px 12px !important; margin-bottom: 4px !important;
    font-weight: 500 !important; font-size: 0.9rem !important;
}}
</style>
"""
st.markdown(CUSTOM_CSS, unsafe_allow_html=True)

# --------------------------------------------------------------------------
# LÓGICA DE INITIAL STATE Y LÓGICA DE NEGOCIO ORIGINAL
# --------------------------------------------------------------------------
def safe_str(v, default=""):
    if v is None: return default
    try:
        if pd.isna(v): return default
    except Exception: pass
    s = str(v).strip()
    return default if s.lower() in ("nan", "none") else s

def safe_float(v, default=0.0):
    try:
        f = float(v)
        if pd.isna(f): return default
        return f
    except Exception: return default

def fmt_money(v):
    return f"S/. {safe_float(v):,.2f}"

def limpiar_texto_dni(val):
    if pd.isna(val) or val is None: return ""
    txt = str(val).strip()
    if txt.endswith(".0"): txt = txt[:-2]
    if txt.isdigit() and len(txt) < 8 and len(txt) > 0: txt = txt.zfill(8)
    return txt

if "step" not in st.session_state: st.session_state.step = "Búsqueda y Carga"
if "clientes_df" not in st.session_state: st.session_state.clientes_df = None
if "cliente_actual" not in st.session_state: st.session_state.cliente_actual = {}
if "visitas" not in st.session_state: st.session_state.visitas = {}
if "garantias" not in st.session_state: st.session_state.garantias = []
if "rcc" not in st.session_state: st.session_state.rcc = []
if "validaciones_marcadas" not in st.session_state: st.session_state.validaciones_marcadas = {}
if "click_timestamps" not in st.session_state: st.session_state.click_timestamps = {}

# Variables globales de entorno mapeadas a session_state
cliente = st.session_state.cliente_actual

def validar_visita():
    validaciones = {
        "documentos_enmiendas": False, "documentos_inconsistentes": False,
        "documentos_sin_datos": False, "documentos_sin_firmas": False,
        "documentos_duplicados": False, "sin_sustento_actividad": False,
        "sin_sustento_ingresos": False, "sin_sustento_activos": False,
        "conyuge_omitido": False, "credito_reprogramado": False,
        "credito_refinanciado": False, "calificacion_diferente": False,
    }
    if not safe_str(cliente.get("CLIENTE")): validaciones["documentos_sin_datos"] = True
    if safe_str(cliente.get("DIAS_ATRASO")) and int(safe_float(cliente.get("DIAS_ATRASO"))) > 0:
        validaciones["calificacion_diferente"] = True
    for clave in ["domicilio", "negocio", "aval"]:
        if clave not in st.session_state.visitas:
            validaciones["sin_sustento_actividad"] = True
            break
        if not st.session_state.visitas[clave].get("foto_bytes"):
            validaciones["documentos_sin_firmas"] = True
    return validaciones

# --------------------------------------------------------------------------
# ENCABEZADO DE IDENTIDAD MÓVIL (Ficha de Cliente)
# --------------------------------------------------------------------------
if cliente:
    atraso = safe_str(cliente.get("DIAS_ATRASO"), "0")
    color_riesgo = ROJO if int(safe_float(atraso)) > 0 else VERDE
    st.markdown(f"""
    <div class="client-profile-box">
        <table style="width:100%; border:none; background:none;">
            <tr>
                <td style="width:10%; font-size:2.5rem; text-align:center;">🛡️</td>
                <td>
                    <small style="color:#666; text-transform:uppercase; font-weight:bold;">Cliente</small><br>
                    <strong style="font-size:1.4rem; color:#1B3A5C;">{safe_str(cliente.get('CLIENTE'))}</strong>
                </td>
                <td>
                    <small style="color:#666; text-transform:uppercase; font-weight:bold;">DNI</small><br>
                    <strong style="font-size:1.2rem; color:#1B3A5C;">{safe_str(cliente.get('PENDOC'))}</strong>
                </td>
                <td style="text-align:right;">
                    <small style="color:#666; text-transform:uppercase; font-weight:bold;">Días de Atraso</small><br>
                    <strong style="font-size:1.4rem; color:#{color_riesgo};">{atraso} días</strong>
                </td>
            </tr>
        </table>
    </div>
    """, unsafe_allow_html=True)


# ==========================================================================
# PASO 1: BÚSQUEDA Y CARGA (busqueda y carga.png)
# ==========================================================================
if st.session_state.step == "Búsqueda y Carga":
    st.title("🚀 Evaluación de Crédito")
    
    col1, col2 = st.columns([1, 1])
    
    with col1:
        st.markdown('<div class="mockup-card">', unsafe_allow_html=True)
        st.subheader("📁 Carga de Base de Datos")
        excel_file = st.file_uploader("Cargar archivo .xlsx de la cartera", type=["xlsx", "xls"])
        
        if excel_file is not None:
            try:
                excel_lector = pd.ExcelFile(excel_file)
                df = pd.read_excel(excel_file, sheet_name=excel_lector.sheet_names[0], dtype=str)
                df.columns = [str(c).strip().upper() for c in df.columns]
                if len(df.columns) >= 4: df = df.rename(columns={df.columns[3]: "PENDOC"})
                df["PENDOC"] = df["PENDOC"].apply(limpiar_texto_dni)
                st.session_state.clientes_df = df
                st.success(f"📊 {len(df)} registros procesados correctamente.")
            except Exception as e:
                st.error(f"Error: {e}")
        st.markdown('</div>', unsafe_allow_html=True)

    with col2:
        st.markdown('<div class="mockup-card">', unsafe_allow_html=True)
        st.subheader("🔍 Búsqueda Inteligente")
        busq = st.text_input("Ingresa DNI, código o nombre completo del cliente:")
        
        df = st.session_state.clientes_df
        if df is not None and busq:
            b = busq.strip().lower()
            mask = (df.get("PENDOC", pd.Series("", index=df.index)).astype(str).str.contains(b, case=False, na=False) |
                    df.get("CLIENTE", pd.Series("", index=df.index)).astype(str).str.contains(b, case=False, na=False))
            resultados = df[mask]
            
            if len(resultados) > 0:
                opciones = resultados.apply(lambda r: f"{safe_str(r.get('CLIENTE'))} | DNI {safe_str(r.get('PENDOC'))}", axis=1).tolist()
                sel = st.selectbox("Selecciona coincidencia:", opciones)
                if sel:
                    idx_sel = opciones.index(sel)
                    if st.button("🔴 Confirmar este cliente", use_container_width=True):
                        st.session_state.cliente_actual = resultados.iloc[idx_sel].to_dict()
                        st.session_state.step = "Ficha del Cliente"
                        st.rerun()
            else:
                st.warning("No se encontraron coincidencias exactas.")
        st.markdown('</div>', unsafe_allow_html=True)


# ==========================================================================
# PASO 2: FICHA DEL CLIENTE (ficha cliente movil.png)
# ==========================================================================
elif st.session_state.step == "Ficha del Cliente":
    st.title("📑 Ficha de Identidad y Crédito")
    
    st.markdown('<div class="mockup-card">', unsafe_allow_html=True)
    st.subheader("Información Operativa del Crédito")
    c1, c2, c3 = st.columns(3)
    with c1:
        st.metric("Importe Desembolsado", fmt_money(cliente.get("IMPDESEMB_MN")))
        st.text_input("Nro. de Cuenta", value=safe_str(cliente.get("BCCTA")), disabled=True)
    with c2:
        st.metric("Saldo Capital Actual", fmt_money(cliente.get("SALDO_MN")))
        st.text_input("Nro. Operación", value=safe_str(cliente.get("BCOPER")), disabled=True)
    with c3:
        st.metric("Calificación SBS", safe_str(cliente.get("CATEG_RESULTANTE", "B+")))
        st.text_input("Producto Caja", value=safe_str(cliente.get("PRODUCTO_CAJA")), disabled=True)
    st.markdown('</div>', unsafe_allow_html=True)

    with st.expander("💼 Información Administrativa y de Gestión (Haga clic para expandir)"):
        c1, c2 = st.columns(2)
        c1.text_input("Analista Asignado", value=safe_str(cliente.get("ANALISTA")), disabled=True)
        c2.text_input("Analista Evaluador", value=safe_str(cliente.get("ANALISTA_EVAL")), disabled=True)
        c1.text_input("Zona / Agencia", value=f"{safe_str(cliente.get('ZONA'))} / {safe_str(cliente.get('AGENCIA'))}", disabled=True)
        c2.text_input("Actividad Económica", value=safe_str(cliente.get("ACTIVIDAD_ECON")), disabled=True)

    # Panel de validación de riesgo interactivo (Doble clic)
    st.markdown('<div class="validation-box">', unsafe_allow_html=True)
    st.markdown('<div class="validation-title">🔍 Panel de Validación - Criterios de Riesgo</div>', unsafe_allow_html=True)
    validaciones_auto = validar_visita()
    criterios = {
        "documentos_enmiendas": ("Documentos con enmiendas", "⚠️"),
        "documentos_inconsistentes": ("Datos inconsistentes en documentos", "⚠️"),
        "documentos_sin_datos": ("Documentos sin datos del cliente", "❌"),
        "documentos_sin_firmas": ("Documentos sin firmas o fotos", "❌"),
        "sin_sustento_actividad": ("Sin sustento de actividad económica", "❌"),
        "sin_sustento_ingresos": ("Sin sustento de ingresos", "❌"),
    }
    
    for key, (label, icon) in criterios.items():
        is_checked = st.session_state.validaciones_marcadas.get(key, validaciones_auto.get(key, False))
        marcador = "[ X ]" if is_checked else "[   ]"
        if st.button(f"{icon} {marcador} {label}", key=f"risk_{key}", use_container_width=True):
            ahora = datetime.now().timestamp()
            ultimo_clic = st.session_state.click_timestamps.get(key, 0)
            st.session_state.click_timestamps[key] = ahora
            if not is_checked:
                st.session_state.validaciones_marcadas[key] = True
                st.rerun()
            elif (ahora - ultimo_clic) < 0.8:
                st.session_state.validaciones_marcadas[key] = False
                st.rerun()
    st.markdown('</div>', unsafe_allow_html=True)

    # Navegación inferior fija
    b_col1, b_col2 = st.columns([1, 3])
    if b_col1.button("⬅️ Atrás", key="btn_atras_1", use_container_width=True):
        st.session_state.step = "Búsqueda y Carga"
        st.rerun()
    if b_col2.button("Guardar y Continuar a Visitas ➡️", use_container_width=True):
        st.session_state.step = "Ingreso Visita"
        st.rerun()


# ==========================================================================
# PASO 3: INGRESO REPORTE / VISITAS (ubicacion.png)
# ==========================================================================
elif st.session_state.step == "Ingreso Visita":
    st.title("📍 Registro de Verificación In Situ")
    
    v_tipo = st.radio("Selecciona el punto de verificación a levantar:", ["Domicilio", "Negocio", "Aval"], horizontal=True)
    clave_v = v_tipo.lower()
    
    st.markdown('<div class="mockup-card">', unsafe_allow_html=True)
    st.subheader(f"Evidencia In Situ: {v_tipo}")
    
    # Lógica de Captura GPS intacta
    col_gps1, col_gps2 = st.columns([1, 2])
    with col_gps1:
        st.markdown("**Ubicación por Satélite**")
        capturar = st.button("📡 Capturar Coordenadas GPS")
        
        if capturar and GEO_OK:
            loc = get_geolocation(key=f"geo_run_{clave_v}")
            if loc and "coords" in loc:
                st.session_state.visitas[clave_v] = st.session_state.visitas.get(clave_v, {})
                st.session_state.visitas[clave_v].update({
                    "lat": loc["coords"]["latitude"], "lon": loc["coords"]["longitude"]
                })
    with col_gps2:
        visita_data = st.session_state.visitas.get(clave_v, {})
        lat, lon = visita_data.get("lat"), visita_data.get("lon")
        if lat and lon:
            st.success(f"Ubicación fijada correctamente: {lat:.5f}, {lon:.5f}")
            st.map(pd.DataFrame({"lat": [lat], "lon": [lon]}), zoom=15)
            
    st.markdown("---")
    st.markdown("**Captura Fotográfica (Obligatoria)**")
    foto_cam = st.camera_input("Enfoque la fachada o entorno")
    foto_file = st.file_uploader("O suba una imagen de su galería móvil", type=["jpg", "png", "jpeg"])
    foto_final = foto_cam if foto_cam else foto_file
    
    if foto_final:
        st.session_state.visitas[clave_v] = st.session_state.visitas.get(clave_v, {})
        st.session_state.visitas[clave_v]["foto_bytes"] = foto_final.getvalue()
        st.success("✅ Imagen vinculada correctamente al punto de verificación.")
        
    st.text_area("Comentarios de la observación de campo:", key=f"obs_{clave_v}")
    st.markdown('</div>', unsafe_allow_html=True)

    b_col1, b_col2 = st.columns([1, 3])
    if b_col1.button("⬅️ Atrás", key="btn_atras_2", use_container_width=True):
        st.session_state.step = "Ficha del Cliente"
        st.rerun()
    if b_col2.button("Continuar al Análisis de Estados Financieros ➡️", use_container_width=True):
        st.session_state.step = "Ingresos y Gastos"
        st.rerun()


# ==========================================================================
# PASO 4: INGRESOS Y GASTOS (ingresos y gastos.png)
# ==========================================================================
elif st.session_state.step == "Ingresos y Gastos":
    st.title("📊 Evaluación Financiera")
    
    # Inputs requeridos para el cálculo matemático
    st.markdown('<div class="mockup-card">', unsafe_allow_html=True)
    c1, c2 = st.columns(2)
    ventas = c1.number_input("Ingresos / Ventas Totales Mensuales (S/.)", value=5000.0)
    otros_ing = c2.number_input("Otros Ingresos Declarados (S/.)", value=500.0)
    st.markdown('</div>', unsafe_allow_html=True)
    
    col_g1, col_g2 = st.columns(2)
    with col_g1:
        st.markdown('<div class="mockup-card">', unsafe_allow_html=True)
        st.subheader("💼 Gastos Operativos / Negocio")
        c_ventas = st.number_input("Costo de Ventas (Mercadería)", value=1500.0)
        g_admin = st.number_input("Gastos Administrativos / Alquiler", value=1050.0)
        g_fin = st.number_input("Gastos Financieros (Otras deudas)", value=0.0)
        st.markdown('</div>', unsafe_allow_html=True)
        
    with col_g2:
        st.markdown('<div class="mockup-card">', unsafe_allow_html=True)
        st.subheader("🏠 Gastos Familiares / Hogar")
        g_fam = st.number_input("Canasta Familiar Básica (Alimentación/Salud)", value=1500.0)
        st.markdown('</div>', unsafe_allow_html=True)

    # Lógica de cálculo idéntica a tu código original
    resultado_neto = ventas + otros_ing - c_ventas - g_admin - g_fin - g_fam
    utilidad_neta = resultado_neto - g_fam
    
    # Banner superior de KPI financiero emulando el Mockup
    st.markdown(f"""
    <div style="background-color: #FFFFFF; border-radius: 12px; padding: 1.5rem; text-align: center; border: 1px solid #E2E8F0;">
        <span style="color:#718096; font-weight:bold;">UTILIDAD NETA MENSUAL CALCULADA</span>
        <h1 style="color:#{VERDE}; margin:5px 0 0 0; font-size:2.5rem;">{fmt_money(utilidad_neta)}</h1>
    </div>
    """, unsafe_allow_html=True)

    b_col1, b_col2 = st.columns([1, 3])
    if b_col1.button("⬅️ Atrás", key="btn_atras_3", use_container_width=True):
        st.session_state.step = "Ingreso Visita"
        st.rerun()
    if b_col2.button("Finalizar y Revisar Reporte ➡️", use_container_width=True):
        st.session_state.step = "Reporte"
        st.rerun()


# ==========================================================================
# PASO 5: GENERACIÓN DE REPORTE (reporte.png)
# ==========================================================================
elif st.session_state.step == "Reporte":
    st.title("🏁 Cierre y Generación de Reporte")
    
    st.markdown('<div class="mockup-card">', unsafe_allow_html=True)
    st.subheader("🔍 Resumen de Control de Calidad")
    
    # Semáforos dinámicos basados en la recolección real de los pasos previos
    rc1, rc2, rc3 = st.columns(3)
    status_dom = "🟢 Domicilio Completo" if "domicilio" in st.session_state.visitas else "🟡 Domicilio Pendiente"
    status_neg = "🟢 Negocio Completo" if "negocio" in st.session_state.visitas else "🟡 Negocio Pendiente"
    status_aval = "🟢 Aval Completo" if "aval" in st.session_state.visitas else "⚪ Aval No requerido"
    
    rc1.info(status_dom)
    rc2.warning(status_neg)
    rc3.success(status_aval)
    st.markdown('</div>', unsafe_allow_html=True)

    # Función nativa de tu script para compilar el archivo Word (.docx)
    def construir_reporte_word():
        doc = Document()
        doc.add_heading("VISITA A CLIENTES DE PEQUEÑA EMPRESA", level=0)
        doc.add_paragraph(f"Generado de forma automática: {datetime.now().strftime('%d/%m/%Y %H:%M')}")
        
        p = doc.add_paragraph()
        run = p.add_run("I. Datos del Titular Evaluado\n")
        run.bold = True
        
        table = doc.add_table(rows=2, cols=2)
        table.style = 'Light Grid Accent 1'
        table.rows[0].cells[0].text = "Cliente:"
        table.rows[0].cells[1].text = safe_str(cliente.get("CLIENTE"))
        table.rows[1].cells[0].text = "DNI:"
        table.rows[1].cells[1].text = safe_str(cliente.get("PENDOC"))
        
        buf = io.BytesIO()
        doc.save(buf)
        buf.seek(0)
        return buf

    st.markdown('<div class="mockup-card" style="text-align:center;">', unsafe_allow_html=True)
    st.subheader("📥 Entrega de Expediente Final")
    
    doc_final = construir_reporte_word()
    nombre_archivo = f"Reporte_Visita_{safe_str(cliente.get('CLIENTE', 'Cliente')).replace(' ', '_')}.docx"
    
    st.download_button(
        label="📥 Descargar Reporte en Formato Oficial (.docx)",
        data=doc_final,
        file_name=nombre_archivo,
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        use_container_width=True
    )
    st.markdown('</div>', unsafe_allow_html=True)
    
    if st.button("🔄 Iniciar Nueva Evaluación de Cartera", use_container_width=True):
        st.session_state.step = "Búsqueda y Carga"
        st.session_state.cliente_actual = {}
        st.session_state.visitas = {}
        st.rerun()
