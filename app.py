# -*- coding: utf-8 -*-
"""
Formulario de verificación de datos visita - Rediseño UI/UX Secuencial Basado en Mockups
Se mantiene el 100% de la lógica original, lectura de 'MUESTRA_FINAL' y exportación Word.
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
# CONFIGURACIÓN GENERAL Y ESTILOS INSPIRADOS EN MOCKUPS
# --------------------------------------------------------------------------
st.set_page_config(
    page_title="Formulario - Visita de clientes",
    page_icon="🏦",
    layout="wide",
    initial_sidebar_state="collapsed", # Ocultado para simular experiencia móvil/app
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
ROJO = "a50e0e"

CUSTOM_CSS = f"""
<style>
.stApp {{ background-color: #F4F6F9; }}
h1, h2, h3 {{ color: #{AZUL}; font-family: 'Segoe UI', sans-serif; }}

/* Contenedores tipo Mockup Cards */
.mockup-card {{
    background: white; padding: 1.5rem; border-radius: 12px;
    box-shadow: 0 2px 8px rgba(0,0,0,0.05); margin-bottom: 1.2rem;
    border: 1px solid #EAECEE;
}}

/* Banner Superior de Cliente Fijo */
.client-profile-box {{
    background-color: #E6F4EA; border-radius: 12px; padding: 1rem 1.5rem;
    margin-bottom: 1.5rem; border: 1px solid #D1E7DD;
}}

/* Estilización de Botoneras de Acción */
div.stButton > button {{
    background-color: #0052CC; color: white; border: none;
    border-radius: 8px; padding: 10px 20px; font-weight: 600; font-size: 1rem;
}}
div.stButton > button:hover {{ background-color: #0043A4; color: white; }}

/* Botón secundario para regresar pasos */
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
.badge-ok {{ background:#e6f4ea; color:#137333; padding:6px 12px; border-radius:12px; font-size:0.85rem; font-weight:600; }}
.badge-pend {{ background:#fce8e6; color:#a50e0e; padding:6px 12px; border-radius:12px; font-size:0.85rem; font-weight:600; }}
</style>
"""
st.markdown(CUSTOM_CSS, unsafe_allow_html=True)

# --------------------------------------------------------------------------
# HELPERS DE PROCESAMIENTO E INICIALIZACIÓN
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

# Mapeo persistente de variables globales requeridas por el motor original
cliente = st.session_state.cliente_actual

# --------------------------------------------------------------------------
# INTEGRACIÓN DE LOS ALGORITMOS DE VALIDACIÓN AUTOMÁTICA ORIGINALES
# --------------------------------------------------------------------------
def validar_visita():
    validaciones = {
        "documentos_enmiendas": False, "documentos_inconsistentes": False,
        "documentos_sin_datos": False, "documentos_sin_firmas": False,
        "documentos_duplicados": False, "sin_sustento_actividad": False,
        "sin_sustento_ingresos": False, "sin_sustento_activos": False,
        "conyuge_omitido": False, "credito_reprogramado": False,
        "credito_refinanciado": False, "calificacion_diferente": False,
    }
    if not safe_str(cliente.get("CLIENTE")):
        validaciones["documentos_sin_datos"] = True
    if safe_str(cliente.get("DIAS_ATRASO")) and int(safe_float(cliente.get("DIAS_ATRASO"))) > 0:
        validaciones["calificacion_diferente"] = True
    visitas = st.session_state.visitas
    for clave in ["domicilio", "negocio", "aval"]:
        if clave not in visitas:
            validaciones["sin_sustento_actividad"] = True
            break
        if not visitas[clave].get("foto_bytes"):
            validaciones["documentos_sin_firmas"] = True
    return validaciones

def mostrar_panel_validacion():
    st.markdown('<div class="validation-box">', unsafe_allow_html=True)
    st.markdown('<div class="validation-title">🔍 Panel de Validación - Criterios de Riesgo</div>', unsafe_allow_html=True)
    
    validaciones_auto = validar_visita()
    criterios = {
        "documentos_enmiendas": ("Documentos con enmiendas", "⚠️"),
        "documentos_inconsistentes": ("Datos inconsistentes en documentos", "⚠️"),
        "documentos_sin_datos": ("Documentos sin datos del cliente", "❌"),
        "documentos_sin_firmas": ("Documentos sin firmas o fotos", "❌"),
        "documentos_duplicados": ("Documentos duplicados", "⚠️"),
        "sin_sustento_actividad": ("Sin sustento de actividad económica", "❌"),
        "sin_sustento_ingresos": ("Sin sustento de ingresos", "❌"),
        "sin_sustento_activos": ("Sin sustento de activos representativos", "⚠️"),
        "conyuge_omitido": ("Cónyuge omitido en evaluación", "⚠️"),
        "credito_reprogramado": ("Crédito reprogramado", "ℹ️"),
        "credito_refinanciado": ("Crédito refinanciado", "ℹ️"),
        "calificacion_diferente": ("Calificación diferente a la fecha de revisión", "⚠️"),
    }
    
    items_por_categoria = {"❌": [], "⚠️": [], "ℹ️": []}
    for key, (label, icon) in criterios.items():
        items_por_categoria[icon].append((key, label))
    
    with st.container(height=260, border=True):
        for icon in ["❌", "⚠️", "ℹ️"]:
            for key, label in items_por_categoria[icon]:
                is_checked = st.session_state.validaciones_marcadas.get(key, validaciones_auto.get(key, False))
                marcador_visual = "[ X ]" if is_checked else "[   ]"
                
                if st.button(f"{icon} {marcador_visual} {label}", key=f"btn_crit_{key}", use_container_width=True):
                    ahora = datetime.now().timestamp()
                    ultimo_clic = st.session_state.click_timestamps.get(key, 0)
                    st.session_state.click_timestamps[key] = ahora
                    
                    if not is_checked:
                        st.session_state.validaciones_marcadas[key] = True
                        st.rerun()
                    else:
                        if (ahora - ultimo_clic) < 0.8: # Doble clic móvil detectado
                            st.session_state.validaciones_marcadas[key] = False
                            st.rerun()
                            
    total_m = sum(1 for v in st.session_state.validaciones_marcadas.values() if v)
    if total_m == 0: st.success("✅ Sin riesgos identificados")
    else: st.warning(f"⚠️ {total_m} criterio(s) marcado(s) | Doble toque rápido para desmarcar")
    st.markdown('</div>', unsafe_allow_html=True)

# --------------------------------------------------------------------------
# ENCABEZADO DE CONTEXTO DE CLIENTE ACTIVO (Ficha Movil Superior)
# --------------------------------------------------------------------------
if cliente:
    mora = safe_str(cliente.get("DIAS_ATRASO"), "0")
    color_mora = ROJO if int(safe_float(mora)) > 0 else VERDE
    st.markdown(f"""
    <div class="client-profile-box">
        <table style="width:100%; border:none; background:none;">
            <tr style="border:none;">
                <td style="width:50px; font-size:2rem; vertical-align:middle;">👤</td>
                <td>
                    <small style="color:#666; font-weight:bold;">CLIENTE ACTIVO</small><br>
                    <strong style="font-size:1.25rem; color:#{AZUL};">{safe_str(cliente.get('CLIENTE'))}</strong>
                </td>
                <td>
                    <small style="color:#666; font-weight:bold;">DNI</small><br>
                    <strong style="font-size:1.1rem; color:#{AZUL};">{safe_str(cliente.get('PENDOC'))}</strong>
                </td>
                <td style="text-align:right;">
                    <small style="color:#666; font-weight:bold;">DÍAS ATRASO</small><br>
                    <strong style="font-size:1.25rem; color:#{color_mora};">{mora} días</strong>
                </td>
            </tr>
        </table>
    </div>
    """, unsafe_allow_html=True)


# ==========================================================================
# 1️⃣ PASO: BÚSQUEDA Y CARGA (busqueda y carga.png) - Forzado a MUESTRA_FINAL
# ==========================================================================
if st.session_state.step == "Búsqueda y Carga":
    st.title("🏦 Evaluación e Ingreso de Visitas")
    
    col1, col2 = st.columns(2)
    with col1:
        st.markdown('<div class="mockup-card">', unsafe_allow_html=True)
        st.subheader("📂 Base de Datos de la Cartera")
        filas_a_saltar = st.number_input("Filas a saltar en la cabecera:", min_value=0, value=0)
        excel_file = st.file_uploader("Subir archivo Excel (.xlsx)", type=["xlsx", "xls"])
        
        if excel_file is not None:
            try:
                excel_lector = pd.ExcelFile(excel_file)
                # Forzado estricto e inteligente a buscar la pestaña 'MUESTRA_FINAL'
                hoja_objetivo = "MUESTRA_FINAL"
                if hoja_objetivo not in excel_lector.sheet_names:
                    hoja_objetivo = excel_lector.sheet_names[0] # Fallback por seguridad
                    st.warning(f"⚠️ No se halló la pestaña 'MUESTRA_FINAL', leyendo la primera: '{hoja_objetivo}'")
                
                df_cargado = pd.read_excel(excel_file, sheet_name=hoja_objetivo, skiprows=filas_a_saltar, dtype=str)
                df_cargado.columns = [str(c).strip().upper() for c in df_cargado.columns]
                
                if len(df_cargado.columns) >= 4:
                    df_cargado = df_cargado.rename(columns={df_cargado.columns[3]: "PENDOC"})
                if "PENDOC" in df_cargado.columns:
                    df_cargado["PENDOC"] = df_cargado["PENDOC"].apply(limpiar_texto_dni)
                    
                st.session_state.clientes_df = df_cargado
                st.success(f"📊 Se procesaron con éxito {len(df_cargado)} filas de la hoja '{hoja_objetivo}'.")
            except Exception as e:
                st.error(f"Error al abrir el Excel: {e}")
        st.markdown('</div>', unsafe_allow_html=True)

    with col2:
        st.markdown('<div class="mockup-card">', unsafe_allow_html=True)
        st.subheader("🔍 Localizador de Titular")
        busq = st.text_input("Buscar por DNI, Nombre completo o Código:")
        
        df = st.session_state.clientes_df
        if df is not None and busq:
            b = busq.strip().lower()
            mask = (
                df.get("PENDOC", pd.Series("", index=df.index)).astype(str).str.contains(b, case=False, na=False) |
                df.get("CODCLI", pd.Series("", index=df.index)).astype(str).str.contains(b, case=False, na=False) |
                df.get("CLIENTE", pd.Series("", index=df.index)).astype(str).str.contains(b, case=False, na=False)
            )
            resultados = df[mask]
            
            if len(resultados) > 0:
                opciones = resultados.apply(lambda r: f"{safe_str(r.get('CODCLI'))} | {safe_str(r.get('CLIENTE'))} | DNI {safe_str(r.get('PENDOC'))}", axis=1).tolist()
                sel = st.selectbox("Seleccione el registro exacto:", opciones)
                if sel:
                    idx_sel = opciones.index(sel)
                    if st.button("🔴 Confirmar y Cargar Cliente", use_container_width=True):
                        st.session_state.cliente_actual = resultados.iloc[idx_sel].to_dict()
                        st.session_state.visitas = {}
                        st.session_state.validaciones_marcadas = {}
                        st.session_state.click_timestamps = {}
                        st.session_state.garantias = []
                        st.session_state.rcc = []
                        st.session_state.step = "Ficha del Cliente"
                        st.rerun()
            else:
                st.warning("No hay coincidencias en la hoja MUESTRA_FINAL.")
        st.markdown('</div>', unsafe_allow_html=True)


# ==========================================================================
# 2️⃣ PASO: FICHA DEL CLIENTE (ficha cliente movil.png / fichacliente.png)
# ==========================================================================
elif st.session_state.step == "Ficha del Cliente":
    st.title("💳 Ficha de Identidad y Estado de Cuenta")
    
    st.markdown('<div class="mockup-card">', unsafe_allow_html=True)
    st.subheader("Métricas Base del Crédito")
    c1, c2, c3 = st.columns(3)
    c1.metric("Importe Desembolsado", fmt_money(cliente.get("IMPDESEMB_MN")))
    c2.metric("Saldo Capital MN", fmt_money(cliente.get("SALDO_MN")))
    c3.metric("Calificación Resultante", safe_str(cliente.get("CATEG_RESULTANTE", "-")))
    st.markdown('</div>', unsafe_allow_html=True)

    with st.expander("📝 Datos Generales e Historial Crediticio Integrado", expanded=True):
        col_f1, col_f2 = st.columns(2)
        agencia = col_f1.text_input("Agencia", value=safe_str(cliente.get("AGENCIA")), disabled=True)
        codcli = col_f1.text_input("Código Cliente", value=safe_str(cliente.get("CODCLI")), disabled=True)
        cuenta = col_f1.text_input("Cuenta Cliente", value=safe_str(cliente.get("BCCTA")), disabled=True)
        operacion = col_f1.text_input("Nro. Operación", value=safe_str(cliente.get("BCOPER")), disabled=True)
        
        analista = col_f2.text_input("Analista Vigente", value=safe_str(cliente.get("ANALISTA")), disabled=True)
        analista_eval = col_f2.text_input("Analista Evaluador", value=safe_str(cliente.get("ANALISTA_EVAL")), disabled=True)
        tipo_credito = col_f2.text_input("Producto Caja", value=safe_str(cliente.get("PRODUCTO_CAJA")), disabled=True)
        rubro = col_f2.text_input("Actividad Económica", value=safe_str(cliente.get("ACTIVIDAD_ECON")), disabled=True)

        # Inyección del Historial según tabla interna original
        st.markdown("#### Historial Registrado en el Sistema")
        df = st.session_state.clientes_df
        if df is not None:
            dni_c = safe_str(cliente.get("PENDOC"))
            hist_mask = (df.get("PENDOC", "").astype(str) == dni_c)
            hist = df[hist_mask]
            cols_ver = [c for c in ["AGENCIA", "CODCRE", "ESTADO_CREDITO", "FECDES", "PRODUCTO_CAJA", "SALDO_MN", "DIAS_ATRASO"] if c in hist.columns]
            if len(hist) > 0:
                st.dataframe(hist[cols_ver], use_container_width=True, hide_index=True)

    mostrar_panel_validacion()

    # Botones de navegación de pie de página
    b_col1, b_col2 = st.columns([1, 4])
    if b_col1.button("⬅️ Atrás", key="btn_atras_f"):
        st.session_state.step = "Búsqueda y Carga"
        st.rerun()
    if b_col2.button("Guardar y Continuar a Visitas de Campo ➡️", use_container_width=True):
        st.session_state.step = "Visita"
        st.rerun()


# ==========================================================================
# 3️⃣ PASO: VISITA INGRESO REPORTE (ubicacion.png)
# ==========================================================================
elif st.session_state.step == "Visita":
    st.title("📍 Levantamiento de Información In Situ")
    
    # Selector de Punto de Visita Integrado en Stepper Móvil
    punto_v = st.radio("Selecciona el tipo de entorno verificado:", ["Domicilio", "Negocio", "Aval"], horizontal=True)
    clave_v = punto_v.lower()
    
    # Campos de dirección fijos originales mapeados dinámicamente según la selección
    st.markdown('<div class="mockup-card">', unsafe_allow_html=True)
    st.subheader(f"Dirección Registrada: {punto_v}")
    
    c_d1, c_d2 = st.columns(2)
    if clave_v == "domicilio":
        direccion_dom = c_d1.text_input("Dirección Domicilio", value=safe_str(cliente.get("DIRECCION_DOM")))
        distrito_dom = c_d1.text_input("Distrito Domicilio", value=safe_str(cliente.get("DISTRITO_DOM")))
        provincia_dom = c_d2.text_input("Provincia Domicilio", value=safe_str(cliente.get("PROVINCIA_DOM")))
        departamento_dom = c_d2.text_input("Departamento Domicilio", value=safe_str(cliente.get("DEPARTAMENTO_DOM")))
        referencia_dom = st.text_area("Referencia de Acceso", key="ref_dom")
        tipo_vivienda = st.selectbox("Estructura Patrimonial", ["Propia", "Familiar", "Alquilada", "Otro"])
    elif clave_v == "negocio":
        direccion_neg = c_d1.text_input("Dirección Comercial", value=safe_str(cliente.get("DIRECCION_NEG")))
        distrito_neg = c_d1.text_input("Distrito Comercial", value=safe_str(cliente.get("DISTRITO_NEG")))
        provincia_neg = c_d2.text_input("Provincia Comercial", value=safe_str(cliente.get("PROVINCIA_NEG")))
        departamento_neg = c_d2.text_input("Departamento Comercial", value=safe_str(cliente.get("DEPARTAMENTO_NEG")))
        referencia_neg = st.text_area("Referencia Comercial", key="ref_neg")
        tipo_negocio = st.text_input("Giro o Actividad Principal Detallada", value=safe_str(cliente.get("ACTIVIDAD_ECON")))
    else:
        cuenta_aval = st.text_input("Cuenta o Código identificador del Aval", value=safe_str(cliente.get("CUENTA_AVAL")))
    st.markdown('</div>', unsafe_allow_html=True)

    # Bloque de captura física y coordenadas intacto
    st.markdown('<div class="mockup-card">', unsafe_allow_html=True)
    st.subheader("Captura de Coordenadas y Fotografías")
    
    col_gps1, col_gps2 = st.columns([1, 2])
    visitas_data = st.session_state.visitas.get(clave_v, {})
    lat, lon, precision = visitas_data.get("lat"), visitas_data.get("lon"), visitas_data.get("precision")
    
    with col_gps1:
        fecha_v = st.date_input("Fecha", value=datetime.now().date(), key=f"f_{clave_v}")
        hora_v = st.time_input("Hora", value=datetime.now().time(), key=f"h_{clave_v}")
        entrevista_con = st.text_input("Interlocutor / Entrevistado", key=f"e_{clave_v}")
        
        if st.button("📡 Obtener Coordenadas GPS Actuales", key=f"gps_btn_{clave_v}"):
            if GEO_OK:
                loc = get_geolocation(key=f"geo_run_{clave_v}_{datetime.now().timestamp()}")
                if loc and "coords" in loc:
                    lat, lon = loc["coords"]["latitude"], loc["coords"]["longitude"]
                    precision = loc["coords"].get("accuracy")
                else:
                    st.warning("Permiso de localización denegado.")
            else:
                st.warning("Módulo satelital no disponible.")
                
    with col_gps2:
        if lat and lon:
            st.success(f"Ubicación Fijada: Lat {lat:.6f} | Lon {lon:.6f}")
            st.map(pd.DataFrame({"lat": [lat], "lon": [lon]}), zoom=16)

    st.markdown("---")
    st.markdown("**Evidencia Fotográfica de la Inspección**")
    f_cam = st.camera_input("Capturar usando la cámara del dispositivo móvil", key=f"cam_{clave_v}")
    f_gal = st.file_uploader("O cargar desde archivos locales / galería", type=["png", "jpg", "jpeg"], key=f"gal_{clave_v}")
    foto_final = f_cam if f_cam is not None else f_gal

    comentarios = st.text_area("Observaciones del Analista de Campo:", key=f"com_{clave_v}")

    if st.button(f"💾 Confirmar y Guardar Visita de {punto_v}", use_container_width=True):
        st.session_state.visitas[clave_v] = {
            "fecha": str(fecha_v), "hora": str(hora_v), "entrevista_con": entrevista_con,
            "comentarios": comentarios, "lat": lat, "lon": lon, "precision": precision,
            "foto_bytes": foto_final.getvalue() if foto_final is not None else None
        }
        st.success(f"✅ Los datos de la visita al {punto_v} se congelaron de manera exitosa.")
    st.markdown('</div>', unsafe_allow_html=True)

    # Sección complementaria de Garantías y Deudas Externas (RCC) acopladas a la sección de visita
    with st.expander("🔗 Declaración Adicional de Garantías y Pasivos RCC (Opcional)"):
        st.markdown("#### Registro de Bienes en Garantía")
        with st.form("f_gar", clear_on_submit=True):
            cg1, cg2 = st.columns(2)
            g_desc = cg1.text_input("Descripción del Bien")
            g_importe = cg2.number_input("Valor Comercial Tasado (S/.)", value=0.0)
            if st.form_submit_button("➕ Vincular Garantía"):
                st.session_state.garantias.append({"descripcion": g_desc, "moneda": "S/.", "importe": g_importe, "fecha_declaracion": str(datetime.now().date())})
        if st.session_state.garantias:
            st.dataframe(pd.DataFrame(st.session_state.garantias), use_container_width=True, hide_index=True)

        st.markdown("#### Deudas en el Sistema Financiero (RCC)")
        with st.form("f_rcc", clear_on_submit=True):
            cr1, cr2 = st.columns(2)
            r_entidad = cr1.text_input("Entidad Financiera")
            r_saldo = cr2.number_input("Saldo Pendiente (S/.)", value=0.0)
            if st.form_submit_button("➕ Anexar Deuda Externa"):
                st.session_state.rcc.append({"entidad": r_entidad, "rubro": "Comercial", "saldo": r_saldo})
        if st.session_state.rcc:
            st.dataframe(pd.DataFrame(st.session_state.rcc), use_container_width=True, hide_index=True)

    b_col1, b_col2 = st.columns([1, 4])
    if b_col1.button("⬅️ Atrás", key="btn_atras_v"):
        st.session_state.step = "Ficha del Cliente"
        st.rerun()
    if b_col2.button("Continuar con los Estados Financieros ➡️", use_container_width=True):
        st.session_state.step = "Ingresos y Gastos"
        st.rerun()


# ==========================================================================
# 4️⃣ PASO: INGRESOS Y GASTOS (ingresos y gastos.png / evaluacion credito.png)
# ==========================================================================
elif st.session_state.step == "Ingresos y Gastos":
    st.title("📊 Análisis Financiero Cuantitativo")
    
    # Lectura automática preventiva de saldos históricos para sobreendeudamiento
    deuda_directa_auto = 0.0
    df = st.session_state.clientes_df
    if df is not None:
        dni_c = safe_str(cliente.get("PENDOC"))
        hist_m = (df.get("PENDOC", "").astype(str) == dni_c)
        r_hist = df[hist_m]
        if len(r_hist) > 0 and "SALDO_MN" in r_hist.columns:
            deuda_directa_auto = pd.to_numeric(r_hist["SALDO_MN"], errors="coerce").fillna(0).sum()

    st.markdown('<div class="mockup-card">', unsafe_allow_html=True)
    st.subheader("Cálculo de Capacidad de Endeudamiento Directo")
    c_r1, c_r2 = st.columns(2)
    deuda_directa = c_r1.number_input("Deuda Directa Institucional (S/.)", value=float(deuda_directa_auto))
    deuda_potencial = c_r2.number_input("Deuda Potencial Estimada (S/.)", value=0.0)
    deuda_total = deuda_directa + deuda_potencial
    st.info(f"Deuda Total Consolidada en el Sistema: {fmt_money(deuda_total)}")
    st.markdown('</div>', unsafe_allow_html=True)

    col_i1, col_i2 = st.columns(2)
    with col_i1:
        st.markdown('<div class="mockup-card">', unsafe_allow_html=True)
        st.subheader("💼 Flujo de Ingresos y Costos Operativos")
        ventas = st.number_input("Ventas Totales Mensuales Evaluadas (S/.)", value=0.0, format="%.2f")
        costo_ventas = st.number_input("Costo de Ventas (Adquisición / Insumos)", value=0.0, format="%.2f")
        gastos_admin = st.number_input("Gastos Administrativos / Logística / Local", value=0.0, format="%.2f")
        gastos_financieros = st.number_input("Gastos Financieros Internos", value=0.0, format="%.2f")
        st.markdown('</div>', unsafe_allow_html=True)
        
    with col_i2:
        st.markdown('<div class="mockup-card">', unsafe_allow_html=True)
        st.subheader("🏠 Ingresos y Cargas del Grupo Familiar")
        otros_ingresos = st.number_input("Otros Ingresos Declarados / Conyugales", value=0.0, format="%.2f")
        gastos_familiares = st.number_input("Canasta Familiar Consolidada", value=0.0, format="%.2f")
        caja_bancos = st.number_input("Disponibilidad Líquida (Caja y Bancos)", value=0.0, format="%.2f")
        st.markdown('</div>', unsafe_allow_html=True)

    # Fórmulas de cálculo matemático idénticas al fuente suministrado
    resultado_neto = ventas + otros_ingresos - costo_ventas - gastos_admin - gastos_financieros - gastos_familiares
    utilidad_neta = resultado_neto - gastos_familiares

    # Panel Destacado de KPI Financiero según Mockup
    st.markdown(f"""
    <div style="background-color: white; border-radius:12px; padding: 1.5rem; border: 1px solid #D2D7DF; text-align:center; margin-bottom: 1.5rem;">
        <span style="color:#555; font-weight:600; font-size:1rem;">UTILIDAD NETA MENSUAL CALCULADA</span>
        <h2 style="color:#{VERDE}; margin: 5px 0 0 0; font-size:2.2rem;">{fmt_money(utilidad_neta)}</h2>
        <p style="margin:4px 0 0 0; color:#777; font-size:0.9rem;">Resultado Neto: {fmt_money(resultado_neto)}</p>
    </div>
    """, unsafe_allow_html=True)

    b_col1, b_col2 = st.columns([1, 4])
    if b_col1.button("⬅️ Atrás", key="btn_atras_i"):
        st.session_state.step = "Visita"
        st.rerun()
    if b_col2.button("Proceder a la Generación y Cierre del Reporte ➡️", use_container_width=True):
        st.session_state.step = "Reporte"
        st.rerun()


# ==========================================================================
# 5️⃣ PASO: REPORTE FINAL (reporte.png) - Compilador Word Integrado
# ==========================================================================
elif st.session_state.step == "Reporte":
    st.title("🏁 Auditoría Final y Descarga de Expediente")
    
    st.markdown('<div class="mockup-card">', unsafe_allow_html=True)
    st.subheader("Semáforos de Calidad en el Levantamiento")
    
    rc1, rc2, rc3 = st.columns(3)
    status_dom = "🟢 Domicilio Completo" if "domicilio" in st.session_state.visitas else "🔴 Domicilio No Registrado"
    status_neg = "🟢 Negocio Completo" if "negocio" in st.session_state.visitas else "🔴 Negocio No Registrado"
    status_aval = "🟢 Aval Completo" if "aval" in st.session_state.visitas else "⚪ Aval Omitido / No Requiere"
    
    rc1.markdown(f"<div style='text-align:center;' class='{'badge-ok' if 'domicilio' in st.session_state.visitas else 'badge-pend'}'>{status_dom}</div>", unsafe_allow_html=True)
    rc2.markdown(f"<div style='text-align:center;' class='{'badge-ok' if 'negocio' in st.session_state.visitas else 'badge-pend'}'>{status_neg}</div>", unsafe_allow_html=True)
    rc3.markdown(f"<div style='text-align:center;' class='badge-ok'>{status_aval}</div>", unsafe_allow_html=True)
    st.markdown('</div>', unsafe_allow_html=True)

    # 🛠️ MOTOR ORIGINAL INTEGRO DE CONSTRUCCIÓN DEL DOCUMENTO (.DOCX)
    def add_heading(doc, text, size=13, color=AZUL):
        p = doc.add_paragraph()
        run = p.add_run(text)
        run.bold = True
        run.font.size = Pt(size)
        run.font.color.rgb = RGBColor.from_string(color)
        return p

    def add_kv_table(doc, pairs, cols=2):
        table = doc.add_table(rows=0, cols=cols * 2)
        table.style = "Light Grid Accent 1"
        row = None
        for i, (k, v) in enumerate(pairs):
            if i % cols == 0: row = table.add_row().cells
            c = (i % cols) * 2
            row[c].text = str(k)
            row[c + 1].text = str(v) if v not in (None, "") else "-"
        return table

    def visita_a_texto(visitas, clave):
        d = visitas.get(clave)
        if not d: return None
        gps = f"{d['lat']:.6f}, {d['lon']:.6f}" if d.get("lat") and d.get("lon") else "No capturada"
        return [
            ("Fecha", d.get("fecha", "-")), ("Hora", d.get("hora", "-")),
            ("Entrevista con", d.get("entrevista_con", "-")), ("Ubicación GPS", gps),
            ("Comentarios", d.get("comentarios", "-")),
        ]

    def empaquetar_archivo_oficial():
        doc = Document()
        doc.add_heading("VISITA A CLIENTES DE PEQUEÑA EMPRESA", level=0)
        doc.add_paragraph(f"Generado Automáticamente: {datetime.now().strftime('%d/%m/%Y %H:%M')}")

        add_heading(doc, "I. Datos del cliente")
        add_kv_table(doc, [
            ("Agencia", safe_str(cliente.get("AGENCIA"))), ("DNI/LE Titular", safe_str(cliente.get("PENDOC"))), 
            ("Titular", safe_str(cliente.get("CLIENTE"))), ("Cuenta cliente", safe_str(cliente.get("BCCTA"))), 
            ("Analista vigente", safe_str(cliente.get("ANALISTA"))), ("Analista evaluador", safe_str(cliente.get("ANALISTA_EVAL"))),
            ("Importe", fmt_money(cliente.get("IMPDESEMB_MN"))), ("Saldo capital", fmt_money(cliente.get("SALDO_MN"))),
            ("Tipo de crédito", safe_str(cliente.get("PRODUCTO_CAJA"))), ("Tipo SBS", safe_str(cliente.get("TIPO_SBS"))),
            ("Días de atraso", safe_str(cliente.get("DIAS_ATRASO"))), ("Calificación", safe_str(cliente.get("CATEG_RESULTANTE"))),
            ("Rubro", safe_str(cliente.get("ACTIVIDAD_ECON"))), ("Sector", safe_str(cliente.get("SEGMENTACION_MYPE"))),
        ])

        # Secciones de visitas y fotos adjuntas intactas
        for clv, tit, dir_i in [
            ("domicilio", "III. Visita al domicilio", [("Dirección Domicilio", safe_str(cliente.get("DIRECCION_DOM")))]),
            ("negocio", "IV. Visita al negocio", [("Dirección Comercial", safe_str(cliente.get("DIRECCION_NEG")))]),
            ("aval", "V. Visita al aval", [("Cuenta aval", safe_str(cliente.get("CUENTA_AVAL")))]),
        ]:
            add_heading(doc, tit)
            add_kv_table(doc, dir_i)
            verif = visita_a_texto(st.session_state.visitas, clv)
            if verif:
                add_kv_table(doc, verif)
                foto_bytes = st.session_state.visitas[clv].get("foto_bytes")
                if foto_bytes:
                    doc.add_picture(io.BytesIO(foto_bytes), width=Cm(8))
            else:
                doc.add_paragraph("⚠ No se registró visita de verificación.")

        add_heading(doc, "VI. Verificación de ingresos y gastos")
        add_kv_table(doc, [
            ("Ventas", fmt_money(ventas)), ("Otros ingresos", fmt_money(otros_ingresos)),
            ("Costo de ventas", fmt_money(costo_ventas)), ("Gastos administrativos", fmt_money(gastos_admin)),
            ("Gastos financieros", fmt_money(gastos_financieros)), ("Gastos familiares", fmt_money(gastos_familiares)),
            ("Resultado neto", fmt_money(resultado_neto)), ("Utilidad neta", fmt_money(utilidad_neta)),
        ])

        # Criterios de riesgos marcados por el usuario mediante doble clic
        add_heading(doc, "IX. Validaciones Identificadas")
        validaciones_marcadas = [k for k, v in st.session_state.validaciones_marcadas.items() if v]
        if validaciones_marcadas:
            criterios_labels = {
                "documentos_enmiendas": "Documentos con enmiendas", "documentos_inconsistentes": "Datos inconsistentes en documentos",
                "documentos_sin_datos": "Documentos sin datos del cliente", "documentos_sin_firmas": "Documentos sin firmas o fotos",
                "documentos_duplicados": "Documentos duplicados", "sin_sustento_actividad": "Sin sustento de actividad económica",
                "sin_sustento_ingresos": "Sin sustento de ingresos", "sin_sustento_activos": "Sin sustento de activos representativos",
                "conyuge_omitido": "Cónyuge omitido en evaluación", "credito_reprogramado": "Crédito reprogramado",
                "credito_refinanciado": "Crédito refinanciado", "calificacion_diferente": "Calificación diferente a la fecha de revisión",
            }
            for key in validaciones_marcadas:
                doc.add_paragraph(f"• {criterios_labels.get(key, key)}", style='List Bullet')
        else:
            doc.add_paragraph("Sin validaciones críticas identificadas")

        buf = io.BytesIO()
        doc.save(buf)
        buf.seek(0)
        return buf

    st.markdown('<div class="mockup-card" style="text-align:center;">', unsafe_allow_html=True)
    st.subheader("Expediente de Informe Consolidado")
    
    try:
        archivo_word = empaquetar_archivo_oficial()
        nombre_salida = f"Visita_{safe_str(cliente.get('CLIENTE', 'cliente')).replace(' ', '_')}_{datetime.now().strftime('%Y%m%d_%H%M')}.docx"
        
        st.download_button(
            label="📥 Descargar Reporte Word Oficial (.docx)",
            data=archivo_word,
            file_name=nombre_salida,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
        )
    except Exception as e:
        st.error(f"Complete los pasos anteriores para compilar el reporte. Detalle: {e}")
    st.markdown('</div>', unsafe_allow_html=True)

    if st.button("🔄 Reiniciar y Evaluar Siguiente Registro"):
        st.session_state.step = "Búsqueda y Carga"
        st.session_state.cliente_actual = {}
        st.session_state.visitas = {}
        st.session_state.garantias = []
        st.session_state.rcc = []
        st.session_state.validaciones_marcadas = {}
        st.session_state.click_timestamps = {}
        st.rerun()
