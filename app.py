# -*- coding: utf-8 -*-
"""
Formulario de Verificación de Datos Visita - Optimizado y Corregido de Errores
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
# CONFIGURACIÓN GENERAL
# --------------------------------------------------------------------------
st.set_page_config(
    page_title="Formulario - Visita de clientes",
    page_icon="🏦",
    layout="wide",
    initial_sidebar_state="expanded",
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

# CSS PERSONALIZADO CORREGIDO: Soluciona las letras invisibles/blancas en las cajas de entrada
CUSTOM_CSS = f"""
<style>
.stApp {{ background-color: #f8fafc; }}
section[data-testid="stSidebar"] {{ background-color: #{AZUL}; }}

/* Forzar que títulos y etiquetas del Sidebar sean blancos */
section[data-testid="stSidebar"] h1, 
section[data-testid="stSidebar"] h2, 
section[data-testid="stSidebar"] h3, 
section[data-testid="stSidebar"] label, 
section[data-testid="stSidebar"] p, 
section[data-testid="stSidebar"] span {{ 
    color: #ffffff !important; 
}}

/* CORRECCIÓN CRÍTICA: Mantener texto oscuro y legible dentro de inputs y selectores del sidebar */
section[data-testid="stSidebar"] input, 
section[data-testid="stSidebar"] select, 
section[data-testid="stSidebar"] div[data-baseweb="select"] * {{ 
    color: #111827 !important; 
}}

h1, h2, h3 {{ color: #{AZUL}; font-weight: 700; }}

div.stButton > button {{
    background-color: #{NARANJA}; color: white; border: none;
    border-radius: 6px; font-weight: 600; padding: 0.4rem 1rem;
}}
div.stButton > button:hover {{ background-color: #a30d24; color: white; }}

.card {{
    background: white; padding: 1.2rem 1.4rem; border-radius: 10px;
    box-shadow: 0 1px 4px rgba(0,0,0,0.06); margin-bottom: 1rem;
    border: 1px solid #e2e8f0;
}}
.badge-ok {{ background:#e6f4ea; color:#137333; padding:5px 10px; border-radius:12px; font-size:0.85rem; font-weight:600; }}
.badge-pend {{ background:#fce8e6; color:#a50e0e; padding:5px 10px; border-radius:12px; font-size:0.85rem; font-weight:600; }}
</style>
"""
st.markdown(CUSTOM_CSS, unsafe_allow_html=True)


# --------------------------------------------------------------------------
# HELPERS Y NORMALIZACIÓN DE DATOS
# --------------------------------------------------------------------------
def safe_str(v, default=""):
    if v is None:
        return default
    try:
        if pd.isna(v):
            return default
    except Exception:
        pass
    s = str(v).strip()
    return default if s.lower() in ("nan", "none") else s


def safe_float(v, default=0.0):
    try:
        f = float(v)
        if pd.isna(f):
            return default
        return f
    except Exception:
        return default


def fmt_money(v):
    return f"S/. {safe_float(v):,.2f}"


def limpiar_formato_dni(val):
    """Elimina espacios y remueve el '.0' residual generado por Excel"""
    s = safe_str(val).strip()
    if s.endswith(".0"):
        s = s[:-2]
    return s


def init_state():
    defaults = {
        "clientes_df": None,
        "cliente_actual": {},
        "visitas": {},   
        "garantias": [],
        "rcc": [],
        "validaciones_marcadas": {},  
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v


init_state()


# --------------------------------------------------------------------------
# LÓGICA DE BÚSQUEDA ROBUSTA DE DNI
# --------------------------------------------------------------------------
def buscar_cliente_por_dni(dni_input, df):
    if not dni_input or df is None or len(df) == 0:
        return None
    
    target = limpiar_formato_dni(dni_input)
    if not target:
        return None

    # Normalizar temporalmente la columna para la comparación
    dni_series = df.get("PENDOC", pd.Series("", index=df.index)).fillna("").astype(str).str.strip()
    dni_series_limpia = dni_series.apply(lambda x: x[:-2] if x.endswith(".0") else x)
    
    mask = (dni_series_limpia == target)
    resultados = df[mask]
    
    if len(resultados) > 0:
        return resultados.iloc[0].to_dict()
    return None


# --------------------------------------------------------------------------
# VALIDACIÓN AUTOMÁTICA
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
    cliente_data = st.session_state.cliente_actual
    if not cliente_data:
        return validaciones

    if not safe_str(cliente_data.get("CLIENTE")):
        validaciones["documentos_sin_datos"] = True
    if safe_str(cliente_data.get("DIAS_ATRASO")) and int(safe_float(cliente_data.get("DIAS_ATRASO"))) > 0:
        validaciones["calificacion_diferente"] = True
    
    visitas = st.session_state.visitas
    for clave in ["domicilio", "negocio", "aval"]:
        if clave not in visitas:
            validaciones["sin_sustento_actividad"] = True
            break
        if not visitas[clave].get("foto_bytes"):
            validaciones["documentos_sin_firmas"] = True
            
    return validaciones


# --------------------------------------------------------------------------
# CUADRO DE ALERTAS OPTIMIZADO (COMPACTO Y MULTI-OPCIÓN)
# --------------------------------------------------------------------------
def mostrar_panel_validacion():
    st.subheader("📋 Criterios de Riesgo / Alertas de la Visita")
    
    opciones_riesgo = {
        "Documentos con enmiendas o tachaduras": "documentos_enmiendas",
        "Datos inconsistentes detectados en documentos": "documentos_inconsistentes",
        "Expediente digital sin datos completos del cliente": "documentos_sin_datos",
        "Falta de firmas mandatorias o registros fotográficos": "documentos_sin_firmas",
        "Documentos duplicados en más de un expediente": "documentos_duplicados",
        "No se evidenció sustento real de la actividad económica": "sin_sustento_actividad",
        "Flujo e ingresos declarados sin sustento verídico": "sin_sustento_ingresos",
        "Activos comerciales declarados sin sustento físico": "sin_sustento_activos",
        "Se omitió registrar la firma/presencia del cónyuge": "conyuge_omitido",
        "El crédito vigente se encuentra Reprogramado": "credito_reprogramado",
        "El crédito vigente se encuentra Refinanciado": "credito_refinanciado",
        "Calificación SBS desmejorada frente a la fecha de revisión": "calificacion_diferente",
    }
    
    validaciones_auto = validar_visita()
    
    # Pre-cargar los elementos que el sistema detecte o el usuario ya haya marcado
    seleccionados_defecto = []
    for label, key in opciones_riesgo.items():
        if st.session_state.validaciones_marcadas.get(key, validaciones_auto.get(key, False)):
            seleccionados_defecto.append(label)
            
    # COMPONENTE COMPACTO RESONSIVO: Reemplaza las 12 filas gigantes por una sola caja dinámica
    seleccion_usuario = st.multiselect(
        "Marque o añada las desviaciones encontradas en la inspección de campo:",
        options=list(opciones_riesgo.keys()),
        default=seleccionados_defecto,
        help="Haga clic para expandir o remover alertas rápidamente."
    )
    
    # Sincronizar de vuelta al estado de la aplicación
    for label, key in opciones_riesgo.items():
        st.session_state.validaciones_marcadas[key] = label in seleccion_usuario

    if len(seleccion_usuario) == 0:
        st.success("✅ Operación Conforme: No se registraron hallazgos críticos de riesgo.")
    else:
        st.error(f"⚠️ Alerta: Se han anexado {len(seleccion_usuario)} criterio(s) de control al informe final.")


# --------------------------------------------------------------------------
# SIDEBAR REPOSITORIO DE DATOS
# --------------------------------------------------------------------------
with st.sidebar:
    st.markdown("### 📂 Repositorio de Datos")
    excel_file = st.file_uploader(
        "Cargar matriz Excel (.xlsx)",
        type=["xlsx", "xls"],
        help="Sube la base de datos para habilitar búsquedas y cruces automáticos.",
    )
    if excel_file is not None:
        try:
            df_upload = pd.read_excel(excel_file, dtype=str)
            df_upload.columns = [c.strip().upper() for c in df_upload.columns]
            
            # Limpieza preventiva de IDs al cargar el dataframe
            for c in ["PENDOC", "CODCLI", "BCCTA", "BCOPER"]:
                if c in df_upload.columns:
                    df_upload[c] = df_upload[c].fillna("").astype(str).str.strip().apply(
                        lambda x: x[:-2] if x.endswith(".0") else x
                    )
            st.session_state.clientes_df = df_upload
            st.success(f"📊 Base indexada ({len(df_upload)} registros)")
        except Exception as e:
            st.error(f"Error de lectura: {e}")

    st.divider()
    df = st.session_state.clientes_df
    if df is not None:
        st.markdown("### 🔍 Buscador Secundario")
        busq = st.text_input("Buscar por Nombre o Código", key="busqueda_sidebar")
        if busq:
            b = busq.strip().lower()
            mask = (
                df.get("CODCLI", pd.Series("", index=df.index)).astype(str).str.contains(b, case=False, na=False) |
                df.get("CLIENTE", pd.Series("", index=df.index)).astype(str).str.contains(b, case=False, na=False)
            )
            resultados = df[mask]
        else:
            resultados = df

        if len(resultados) > 0:
            opciones = resultados.apply(
                lambda r: f"{safe_str(r.get('CODCLI'))} | {safe_str(r.get('CLIENTE'))}", axis=1
            ).tolist()
            sel = st.selectbox("Seleccione cliente alternativo:", opciones)
            if sel and st.button("📥 Forzar Importación", use_container_width=True):
                idx_sel = opciones.index(sel)
                st.session_state.cliente_actual = resultados.iloc[idx_sel].to_dict()
                st.session_state.visitas = {}
                st.session_state.validaciones_marcadas = {}
                st.rerun()

    st.divider()
    if st.session_state.cliente_actual:
        st.warning(f"Titular Seleccionado:\n**{safe_str(st.session_state.cliente_actual.get('CLIENTE'))}**")
        if st.button("🗑️ Reiniciar Formulario Vacío", use_container_width=True):
            st.session_state.cliente_actual = {}
            st.session_state.visitas = {}
            st.session_state.garantias = []
            st.session_state.rcc = []
            st.session_state.validaciones_marcadas = {}
            st.rerun()


# --------------------------------------------------------------------------
# CONTENEDOR PRINCIPAL
# --------------------------------------------------------------------------
cliente = st.session_state.cliente_actual

st.title("Gestión Integrada de Visitas")
st.caption("Módulo de mitigación y control de riesgos de crédito en campo")

tabs = st.tabs([
    "1️⃣ Cliente y Crédito", "2️⃣ Historial y Riesgo", "3️⃣ Visita Domicilio",
    "4️⃣ Visita Negocio", "5️⃣ Ingresos y Gastos", "6️⃣ Garantías y Aval", "7️⃣ Generar Reporte"
])

# TAB 1: CLIENTE Y CRÉDITO (BÚSQUEDA INTERACTIVA DIRECTA REPARADA)
with tabs[0]:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.subheader("Búsqueda Directa del Titular")
    
    # Estructura limpia de entrada de DNI con botón de acción dedicado
    col_dni1, col_dni2 = st.columns([0.7, 0.3])
    with col_dni1:
        dni_digitado = st.text_input(
            "💳 Ingrese Número de DNI / LE del Cliente",
            value=limpiar_formato_dni(cliente.get("PENDOC")),
            key="dni_search_input_field",
            help="Escriba el DNI y presione el botón de la derecha para extraer los datos."
        )
    with col_dni2:
        st.markdown("<div style='padding-top: 1.7rem;'></div>", unsafe_allow_html=True)
        if st.button("🔍 Extraer Datos por DNI", use_container_width=True):
            if st.session_state.clientes_df is not None:
                res = buscar_cliente_por_dni(dni_digitado, st.session_state.clientes_df)
                if res:
                    st.session_state.cliente_actual = res
                    st.session_state.visitas = {}
                    st.session_state.validaciones_marcadas = {}
                    st.success(f"✅ Encontrado: {res.get('CLIENTE')}")
                    st.rerun()
                else:
                    st.error("❌ El número de DNI ingresado no figura en la base de datos.")
            else:
                st.info("💡 Primero cargue el archivo Excel en el repositorio (Panel Izquierdo).")

    # Campos abiertos (Sin disabled=True para permitir llenado manual si el cliente lo prefiere)
    c1, c2, c3 = st.columns(3)
    with c1:
        agencia = st.text_input("Agencia de Registro", value=safe_str(cliente.get("AGENCIA")))
        dni_display = st.text_input("Documento Identidad Activo", value=limpiar_formato_dni(cliente.get("PENDOC")))
        codcli = st.text_input("Código único de Cliente (CODCLI)", value=safe_str(cliente.get("CODCLI")))
    with c2:
        titular = st.text_input("Apellidos y Nombres Completos", value=safe_str(cliente.get("CLIENTE")))
        cuenta = st.text_input("Código de Cuenta Relacionada", value=safe_str(cliente.get("BCCTA")))
        operacion = st.text_input("Nro. de Operación SBS", value=safe_str(cliente.get("BCOPER")))
    with c3:
        analista = st.text_input("Analista Administrador", value=safe_str(cliente.get("ANALISTA")))
        analista_eval = st.text_input("Analista Evaluador Originador", value=safe_str(cliente.get("ANALISTA_EVAL")))
        aprobado_por = st.text_input("Usuario Otorgante / Aprobador", value=safe_str(cliente.get("USUARIO_APROB")))
    st.markdown("</div>", unsafe_allow_html=True)

    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.subheader("Estructura de la Operación de Crédito")
    c1, c2, c3 = st.columns(3)
    with c1:
        importe = st.number_input("Monto Desembolsado (S/.)", value=safe_float(cliente.get("IMPDESEMB_MN")), format="%.2f")
        saldo_capital = st.number_input("Saldo de Capital Pendiente (S/.)", value=safe_float(cliente.get("SALDO_MN")), format="%.2f")
        tipo_credito = st.text_input("Sub-producto Caja", value=safe_str(cliente.get("PRODUCTO_CAJA")))
    with c2:
        tipo_sbs = st.text_input("Tipo de Crédito según SBS", value=safe_str(cliente.get("TIPO_SBS")))
        fecha_desembolso = st.text_input("Fecha de Activación / Desembolso", value=safe_str(cliente.get("FECDES")))
        cuotas_pagadas = st.text_input("Cuotas Amortizadas / Canceladas", value=safe_str(cliente.get("CUOTAS_PAGADAS")))
    with c3:
        dias_atraso = st.text_input("Días de Atraso a la Fecha", value=safe_str(cliente.get("DIAS_ATRASO")))
        prom_mora = st.text_input("Días de Mora Histórica Promedio", value=safe_str(cliente.get("MORA_CONT")))
        calificacion = st.text_input("Clasificación de Riesgo Interna", value=safe_str(cliente.get("CATEG_RESULTANTE")))
        
    c1, c2, c3 = st.columns(3)
    with c1:
        rubro = st.text_input("CIIU / Actividad Desarrollada", value=safe_str(cliente.get("ACTIVIDAD_ECON")))
    with c2:
        sector = st.text_input("Segmento MYPE Asignado", value=safe_str(cliente.get("SEGMENTACION_MYPE")))
    with c3:
        modulo = st.text_input("Módulo de Negocio Corporativo", value=safe_str(cliente.get("MODULO")))
    st.markdown("</div>", unsafe_allow_html=True)
    
    # Cuadro de alertas unificado e integrado en la parte inferior
    st.markdown('<div class="card">', unsafe_allow_html=True)
    mostrar_panel_validacion()
    st.markdown("</div>", unsafe_allow_html=True)


# TAB 2: HISTORIAL CREDITICIO
with tabs[1]:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.subheader("3. Historial de Créditos Internos Detectados")
    df = st.session_state.clientes_df
    current_dni = dni_display if 'dni_display' in locals() else safe_str(cliente.get("PENDOC"))
    
    if df is not None and (current_dni or codcli):
        mask = pd.Series(False, index=df.index)
        if current_dni:
            mask = mask | (df.get("PENDOC", "").astype(str) == current_dni)
        if codcli:
            mask = mask | (df.get("CODCLI", "").astype(str) == codcli)
        hist = df[mask]
        cols_show = [c for c in [
            "AGENCIA", "CODCRE", "ESTADO_CREDITO", "FECDES", "FECHA_UTLPAGO",
            "PRODUCTO_CAJA", "SALDO_MN", "DIAS_ATRASO", "ANALISTA"
        ] if c in hist.columns]
        if len(hist) > 0:
            st.dataframe(hist[cols_show], use_container_width=True, hide_index=True)
        else:
            st.info("No se registran deudas paralelas o históricas para este rut en el Excel.")
    else:
        st.info("Seleccione o busque un cliente válido para procesar el historial financiero.")
    st.markdown("</div>", unsafe_allow_html=True)

    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.subheader("4. Nivel de Sobreendeudamiento")
    deuda_directa_auto = 0.0
    if df is not None and 'hist' in locals() and len(hist) > 0:
        for col in ["SALDO_VIGE", "SALDO_REFI"]:
            if col in hist.columns:
                deuda_directa_auto += pd.to_numeric(hist[col], errors="coerce").fillna(0).sum()

    c1, c2 = st.columns(2)
    with c1:
        deuda_directa = st.number_input("a) Consolidado de Deuda Directa (S/.)", value=float(deuda_directa_auto), format="%.2f")
        deuda_potencial = st.number_input("b) Líneas de Crédito / Deuda Potencial (S/.)", value=0.0, format="%.2f")
        deuda_total = st.number_input("c) Exposición de Deuda Total (S/.)", value=float(deuda_directa) + float(deuda_potencial), format="%.2f")
    with c2:
        resultado_neto_rs = st.number_input("e) Ratio Cuota / Resultado Neto Mensual (%)", value=0.0, format="%.2f")
        pasivo_patrimonio = st.number_input("f) Apalancamiento Pasivo / Patrimonio (%)", value=0.0, format="%.2f")
        tipo_deuda = st.selectbox("d) Denominación de la Divisa", ["Sin identificación", "MN", "ME"])
    st.markdown("</div>", unsafe_allow_html=True)


# BLOQUE DE VERIFICACIÓN TECNOLÓGICA REUTILIZABLE
def bloque_verificacion(clave, etiqueta):
    st.markdown("##### 📍 Evidencias Tecnológicas de Campo")
    visitas = st.session_state.visitas
    data = visitas.get(clave, {})

    colf1, colf2 = st.columns(2)
    with colf1:
        fecha_v = st.date_input("Fecha Ajustada de Visita", value=datetime.now().date(), key=f"fecha_{clave}")
    with colf2:
        hora_v = st.time_input("Hora de Ingreso de Visita", value=datetime.now().time(), key=f"hora_{clave}")

    entrevista_con = st.text_input("Persona Contactada / Entrevistada:", key=f"entrevista_{clave}")
    comentarios = st.text_area("Conclusiones de Verificación In Situ:", key=f"comentarios_{clave}")

    st.markdown("**Localización por Satélite (GPS)**")
    cgps1, cgps2 = st.columns([1, 2])
    with cgps1:
        capturar = st.button("📡 Capturar Ubicación GPS en Vivo", key=f"btn_gps_{clave}", use_container_width=True)
    lat, lon, precision = data.get("lat"), data.get("lon"), data.get("precision")
    if capturar:
        if GEO_OK:
            loc = get_geolocation(key=f"geo_{clave}_{datetime.now().timestamp()}")
            if loc and "coords" in loc:
                lat = loc["coords"]["latitude"]
                lon = loc["coords"]["longitude"]
                precision = loc["coords"].get("accuracy")
            else:
                st.warning("Habilite los permisos de posicionamiento en su celular/PC para continuar.")
        else:
            st.warning("Geolocalización deshabilitada en este entorno web.")
            
    with cgps2:
        if lat and lon:
            st.success(f"Fijado correctamente (Precisión: ±{precision:.1f} metros)")
            st.map(pd.DataFrame({"lat": [lat], "lon": [lon]}), zoom=15, height=170)

    st.markdown("**Registro Fotográfico Obligatorio**")
    cfoto1, cfoto2 = st.columns(2)
    with cfoto1:
        foto_camara = st.camera_input("Capturar con Cámara del Celular / Laptop", key=f"camara_{clave}")
    with cfoto2:
        foto_archivo = st.file_uploader("Subir Fotografía desde Galería", type=["jpg", "jpeg", "png"], key=f"upload_{clave}")
    foto_final = foto_camara if foto_camara is not None else foto_archivo

    if st.button(f"💾 Archivar Inspección de {etiqueta}", key=f"guardar_{clave}", use_container_width=True):
        st.session_state.visitas[clave] = {
            "fecha": str(fecha_v), "hora": str(hora_v), "entrevista_con": entrevista_con,
            "comentarios": comentarios, "lat": lat, "lon": lon, "precision": precision,
            "foto_bytes": foto_final.getvalue() if foto_final is not None else None,
        }
        st.success(f"✅ Los datos de {etiqueta} se archivaron temporalmente.")


# TAB 3: DOMICILIO
with tabs[2]:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.subheader("5. Verificación Domiciliaria")
    c1, c2 = st.columns(2)
    with c1:
        direccion_dom = st.text_input("Dirección Domicilio Declarada", value=safe_str(cliente.get("DIRECCION_DOM")), key="dir_dom")
        distrito_dom = st.text_input("Distrito", value=safe_str(cliente.get("DISTRITO_DOM")), key="dist_dom")
    with c2:
        provincia_dom = st.text_input("Provincia", value=safe_str(cliente.get("PROVINCIA_DOM")), key="prov_dom")
        departamento_dom = st.text_input("Departamento", value=safe_str(cliente.get("DEPARTAMENTO_DOM")), key="depto_dom")
    referencia_dom = st.text_area("Puntos de referencia urbanos:", key="ref_dom")
    tipo_vivienda = st.selectbox("Condición de Ocupación de la Vivienda", ["Propia", "Familiar", "Alquilada", "Otro"], key="tipo_viv")
    st.divider()
    bloque_verificacion("domicilio", "Domicilio")
    st.markdown("</div>", unsafe_allow_html=True)

# TAB 4: NEGOCIO
with tabs[3]:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.subheader("6. Verificación de Unidad del Negocio / Local Comercial")
    c1, c2 = st.columns(2)
    with c1:
        direccion_neg = st.text_input("Dirección Comercial Declarada", value=safe_str(cliente.get("DIRECCION_NEG")), key="dir_neg")
        distrito_neg = st.text_input("Distrito Comercial", value=safe_str(cliente.get("DISTRITO_NEG")), key="dist_neg")
    with c2:
        provincia_neg = st.text_input("Provincia Comercial", value=safe_str(cliente.get("PROVINCIA_NEG")), key="prov_neg")
        departamento_neg = st.text_input("Departamento Comercial", value=safe_str(cliente.get("DEPARTAMENTO_NEG")), key="depto_neg")
    referencia_neg = st.text_area("Rótulo, fachada o accesos de referencia:", key="ref_neg")
    tipo_negocio = st.text_input("Giro o Actividad Económica Constatada", value=safe_str(cliente.get("ACTIVIDAD_ECON")), key="tipo_neg")
    st.divider()
    bloque_verificacion("negocio", "Negocio")
    st.markdown("</div>", unsafe_allow_html=True)

# TAB 5: EVALUACIÓN ECONÓMICA
with tabs[4]:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.subheader("Evaluación Comercial de Ingresos")
    c1, c2 = st.columns(2)
    with c1:
        actividad_principal = st.text_input("Giro Comercial Evaluado Principal", value=safe_str(cliente.get("ACTIVIDAD_ECON")), key="act_princ")
    with c2:
        otras_actividades = st.text_input("Líneas de Negocio Secundarias", key="otras_act")
    ventas = st.number_input("Ventas Promedio Mensuales Consolidadas (S/.)", value=0.0, format="%.2f", key="ventas")

    st.subheader("Detalle de Egresos y Márgenes de Utilidad")
    c1, c2, c3 = st.columns(3)
    with c1:
        costo_ventas = st.number_input("Costo Directo de Mercadería (Ventas)", value=0.0, format="%.2f", key="costo_ventas")
        gastos_admin = st.number_input("Gastos de Operación / Personal", value=0.0, format="%.2f", key="gastos_admin")
    with c2:
        gastos_financieros = st.number_input("Cuotas Créditos Sistema Financiero", value=0.0, format="%.2f", key="gastos_fin")
        gastos_familiares = st.number_input("Egresos de Canasta Familiar / Carga", value=0.0, format="%.2f", key="gastos_fam")
    with c3:
        otros_ingresos = st.number_input("Otros Ingresos Comprobables", value=0.0, format="%.2f", key="otros_ing")
        caja_bancos = st.number_input("Saldo disponible Líquido (Caja/Bancos)", value=0.0, format="%.2f", key="caja_bancos")

    resultado_neto = ventas + otros_ingresos - costo_ventas - gastos_admin - gastos_financieros - gastos_familiares
    utilidad_neta = resultado_neto - gastos_familiares
    c1, c2 = st.columns(2)
    c1.metric("Resultado Neto de Operación", fmt_money(resultado_neto))
    c2.metric("Utilidad Excedente Disponible Líquida", fmt_money(utilidad_neta))
    st.markdown("</div>", unsafe_allow_html=True)

# TAB 6: GARANTÍAS Y AVALES
with tabs[5]:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.subheader("7. Registro e Indexación de Garantías Propuestas")
    with st.form("form_garantia", clear_on_submit=True):
        c1, c2, c3 = st.columns(3)
        with c1:
            g_desc = st.text_input("Detalle o Descripción del Activo en Garantía")
            g_moneda = st.selectbox("Moneda de Tasación", ["S/.", "US$"])
        with c2:
            g_importe = st.number_input("Valor Comercial de Tasación", value=0.0, format="%.2f")
            g_perito = st.text_input("Código / Nombre del Perito Tasador")
        with c3:
            g_fecha_tasacion = st.date_input("Fecha del Informe de Tasación REPEV", value=None)
            g_fecha_decl = st.date_input("Fecha de Firma de DJ", value=datetime.now().date())
        if st.form_submit_button("➕ Vincular Garantía al Expediente"):
            st.session_state.garantias.append({
                "descripcion": g_desc, "moneda": g_moneda, "importe": g_importe,
                "perito": g_perito, "fecha_tasacion": str(g_fecha_tasacion),
                "fecha_declaracion": str(g_fecha_decl),
            })
    if st.session_state.garantias:
        st.dataframe(pd.DataFrame(st.session_state.garantias), use_container_width=True, hide_index=True)
    st.markdown("</div>", unsafe_allow_html=True)

    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.subheader("VI. Inspección de Obligados Solidarios / Avales")
    cuenta_aval = st.text_input("Código de Aval Relacionado", value=safe_str(cliente.get("CUENTA_AVAL")), key="cuenta_aval")
    bloque_verificacion("aval", "Aval")
    st.markdown("</div>", unsafe_allow_html=True)


# --------------------------------------------------------------------------
# EXPORTACIÓN ESTRUCTURADA A WORD (.DOCX)
# --------------------------------------------------------------------------
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
        if i % cols == 0:
            row = table.add_row().cells
        c = (i % cols) * 2
        row[c].text = str(k)
        row[c + 1].text = str(v) if v not in (None, "") else "-"
    return table


def visita_a_texto(visitas, clave):
    d = visitas.get(clave)
    if not d:
        return None
    gps = f"{d['lat']:.6f}, {d['lon']:.6f}" if d.get("lat") and d.get("lon") else "No capturada"
    return [
        ("Fecha", d.get("fecha", "-")), ("Hora", d.get("hora", "-")),
        ("Entrevista con", d.get("entrevista_con", "-")), ("Ubicación GPS", gps),
        ("Comentarios", d.get("comentarios", "-")),
    ]


def generar_reporte():
    doc = Document()
    doc.add_heading("VISITA A CLIENTES DE PEQUEÑA EMPRESA", level=0)
    p = doc.add_paragraph("")
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_paragraph(f"Emitido el: {datetime.now().strftime('%d/%m/%Y %H:%M')}")

    add_heading(doc, "I. Datos de Identificación del Cliente")
    add_kv_table(doc, [
        ("Agencia", agencia), ("DNI/LE Titular", dni_display),
        ("Titular", titular), ("Cuenta cliente", cuenta),
        ("Analista vigente", analista), ("Analista evaluador", analista_eval),
        ("Importe", fmt_money(importe)), ("Saldo capital", fmt_money(saldo_capital)),
        ("Tipo de crédito", tipo_credito), ("Tipo SBS", tipo_sbs),
        ("Días de atraso", dias_atraso), ("Calificación", calificacion),
        ("Rubro", rubro), ("Sector", sector),
    ])

    add_heading(doc, "II. Mitigación del Riesgo de Sobreendeudamiento")
    add_kv_table(doc, [
        ("Deuda directa", fmt_money(deuda_directa)), ("Deuda potencial", fmt_money(deuda_potencial)),
        ("Deuda total", fmt_money(deuda_total)), ("Cuota/Resultado neto", f"{resultado_neto_rs:.2f}%"),
        ("Pasivo/Patrimonio", f"{pasivo_patrimonio:.2f}%"),
    ])

    for clave, titulo, direccion_info in [
        ("domicilio", "III. Informe de Visita al Domicilio", [
            ("Dirección", direccion_dom), ("Distrito", distrito_dom),
            ("Provincia", provincia_dom), ("Departamento", departamento_dom), ("Tipo de vivienda", tipo_vivienda),
        ]),
        ("negocio", "IV. Informe de Visita a la Unidad de Negocio", [
            ("Dirección", direccion_neg), ("Distrito", distrito_neg),
            ("Provincia", provincia_neg), ("Departamento", departamento_neg), ("Actividad", tipo_negocio),
        ]),
        ("aval", "V. Informe de Visita al Aval / Coobligado", [("Cuenta aval", cuenta_aval)]),
    ]:
        add_heading(doc, titulo)
        add_kv_table(doc, direccion_info)
        verif = visita_a_texto(st.session_state.visitas, clave)
        if verif:
            doc.add_paragraph("Declaración de Verificación de Campo:").bold = True
            add_kv_table(doc, verif)
            foto_bytes = st.session_state.visitas[clave].get("foto_bytes")
            if foto_bytes:
                doc.add_picture(io.BytesIO(foto_bytes), width=Cm(8))
        else:
            doc.add_paragraph("⚠ No se cargaron cierres de inspección para esta sección.")

    add_heading(doc, "VI. Consolidado Financiero de Ingresos y Egresos")
    add_kv_table(doc, [
        ("Ventas", fmt_money(ventas)), ("Otros ingresos", fmt_money(otros_ingresos)),
        ("Costo de ventas", fmt_money(costo_ventas)), ("Gastos administrativos", fmt_money(gastos_admin)),
        ("Gastos financieros", fmt_money(gastos_financieros)), ("Gastos familiares", fmt_money(gastos_familiares)),
        ("Resultado neto", fmt_money(resultado_neto)), ("Utilidad neta", fmt_money(utilidad_neta)),
    ])

    add_heading(doc, "VII. Matriz de Hallazgos y Riesgos Seleccionados")
    validaciones_marcadas = [k for k, v in st.session_state.validaciones_marcadas.items() if v]
    if validaciones_marcadas:
        criterios_labels = {
            "documentos_enmiendas": "Documentos con enmiendas o tachaduras",
            "documentos_inconsistentes": "Datos inconsistentes detectados en documentos",
            "documentos_sin_datos": "Expediente digital sin datos completos del cliente",
            "documentos_sin_firmas": "Falta de firmas mandatorias o registros fotográficos",
            "documentos_duplicados": "Documentos duplicados en más de un expediente",
            "sin_sustento_actividad": "No se evidenció sustento real de la actividad económica",
            "sin_sustento_ingresos": "Flujo e ingresos declarados sin sustento verídico",
            "sin_sustento_activos": "Activos comerciales declarados sin sustento físico",
            "conyuge_omitido": "Se omitió registrar la firma/presencia del cónyuge",
            "credito_reprogramado": "El crédito vigente se encuentra Reprogramado",
            "credito_refinanciado": "El crédito vigente se encuentra Refinanciado",
            "calificacion_diferente": "Calificación SBS desmejorada frente a la fecha de revisión",
        }
        for key in validaciones_marcadas:
            doc.add_paragraph(f"⚠ {criterios_labels.get(key, key)}", style='List Bullet')
    else:
        doc.add_paragraph("Operación Conforme: Sin observaciones críticas ni hallazgos de riesgo marcados.")

    add_heading(doc, "Firmas y Cierre de Acta")
    add_kv_table(doc, [
        ("Emitido por (Firma)", ""), ("Fecha de Validación", datetime.now().strftime("%d/%m/%Y")),
        ("Visto Bueno (Jefatura)", ""), ("Fecha de Revisión", ""),
    ])

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf


# TAB 7: PANEL DE DESCARGAS
with tabs[6]:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.subheader("Semáforo de Controles Obligatorios")
    for clave, etiqueta in [("domicilio", "Domicilio"), ("negocio", "Negocio"), ("aval", "Aval")]:
        if clave in st.session_state.visitas:
            d = st.session_state.visitas[clave]
            st.markdown(f"**{etiqueta}** — <span class='badge-ok'>Visita Guardada</span>", unsafe_allow_html=True)
        else:
            st.markdown(f"**{etiqueta}** — <span class='badge-pend'>Pendiente de Inspección</span>", unsafe_allow_html=True)
    st.markdown("</div>", unsafe_allow_html=True)

    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.subheader("Generación de Documentación y Actas")
    if st.button("Compilar y Descargar Informe Ejecutivo (.docx)", type="primary"):
        buf = generar_reporte()
        nombre = f"Informe_Visita_{safe_str(titular, 'cliente').replace(' ', '_')}_{datetime.now().strftime('%Y%m%d_%H%M')}.docx"
        st.download_button(
            "⬇️ Hacer clic aquí para Descargar Reporte Word",
            data=buf,
            file_name=nombre,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
        )
        st.success("Acta compilada con éxito.")
    st.markdown("</div>", unsafe_allow_html=True)
