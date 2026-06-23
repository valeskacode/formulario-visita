# -*- coding: utf-8 -*-
"""
formulario de verificacion de datos visita 
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

NARANJA = "C8102E"   # color institucional aproximado
AZUL = "1B3A5C"
VERDE = "137333"
ROJO = "a50e0e"

CUSTOM_CSS = f"""
<style>
.stApp {{ background-color: #f7f7f9; }}
section[data-testid="stSidebar"] {{ background-color: #1B3A5C; }}
section[data-testid="stSidebar"] * {{ color: #ffffff !important; }}
h1, h2, h3 {{ color: #{AZUL}; }}
div.stButton > button {{
    background-color: #{NARANJA}; color: white; border: none;
    border-radius: 6px; font-weight: 600;
}}
div.stButton > button:hover {{ background-color: #a30d24; color: white; }}
.card {{
    background: white; padding: 1.2rem 1.4rem; border-radius: 10px;
    box-shadow: 0 1px 4px rgba(0,0,0,0.08); margin-bottom: 1rem;
}}
.badge-ok {{ background:#e6f4ea; color:#137333; padding:6px 12px; border-radius:12px; font-size:0.85rem; font-weight:600; }}
.badge-pend {{ background:#fce8e6; color:#a50e0e; padding:6px 12px; border-radius:12px; font-size:0.85rem; font-weight:600; }}
.validation-box {{
    border: 2px solid #{AZUL}; border-radius: 10px; padding: 1.5rem;
    background: linear-gradient(135deg, rgba(27,58,92,0.05) 0%, rgba(200,16,46,0.05) 100%);
    margin: 1rem 0;
}}
.validation-title {{
    font-size: 1.1rem; font-weight: 700; color: #{AZUL}; margin-bottom: 1rem;
    border-bottom: 3px solid #{NARANJA}; padding-bottom: 0.5rem;
}}
.validation-item {{
    padding: 0.8rem; margin-bottom: 0.6rem; border-radius: 6px;
    display: flex; align-items: center; gap: 0.8rem;
    border-left: 4px solid #ddd;
}}
.validation-ok {{
    background: #e6f4ea; border-left-color: #{VERDE};
}}
.validation-warning {{
    background: #fff3cd; border-left-color: #ff9800;
}}
.validation-error {{
    background: #fce8e6; border-left-color: #{ROJO};
}}
.validation-icon {{ font-size: 1.3rem; min-width: 30px; text-align: center; }}
.validation-text {{ flex: 1; }}
.responsive-grid {{
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(280px, 1fr));
    gap: 1rem;
}}
@media (max-width: 768px) {{
    .card {{ padding: 1rem 1.2rem; }}
    .responsive-grid {{ grid-template-columns: 1fr; }}
    h1 {{ font-size: 1.5rem; }}
    h2 {{ font-size: 1.2rem; }}
}}
@media (max-width: 480px) {{
    .card {{ padding: 0.8rem 1rem; }}
    section[data-testid="stSidebar"] {{ width: 100% !important; }}
}}
</style>
"""
st.markdown(CUSTOM_CSS, unsafe_allow_html=True)


# --------------------------------------------------------------------------
# HELPERS
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


def init_state():
    defaults = {
        "clientes_df": None,
        "cliente_actual": {},
        "visitas": {},   # domicilio / negocio / aval -> dict con foto, gps, hora, etc
        "garantias": [],
        "rcc": [],
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v


init_state()


# --------------------------------------------------------------------------
# VALIDACIÓN DE CRITERIOS
# --------------------------------------------------------------------------
def validar_visita():
    """Retorna diccionario con validaciones según criterios de la tabla."""
    validaciones = {
        "documentos_enmiendas": False,
        "documentos_inconsistentes": False,
        "documentos_sin_datos": False,
        "documentos_sin_firmas": False,
        "documentos_duplicados": False,
        "sin_sustento_actividad": False,
        "sin_sustento_ingresos": False,
        "sin_sustento_activos": False,
        "conyuge_omitido": False,
        "credito_reprogramado": False,
        "credito_refinanciado": False,
        "calificacion_diferente": False,
    }
    
    # Validar campos críticos
    if not safe_str(cliente.get("CLIENTE")):
        validaciones["documentos_sin_datos"] = True
    
    if safe_str(cliente.get("DIAS_ATRASO")) and int(safe_float(cliente.get("DIAS_ATRASO"))) > 0:
        validaciones["calificacion_diferente"] = True
    
    # Validar visitas
    visitas = st.session_state.visitas
    for clave in ["domicilio", "negocio", "aval"]:
        if clave not in visitas:
            validaciones["sin_sustento_actividad"] = True
            break
        visita = visitas[clave]
        if not visita.get("foto_bytes"):
            validaciones["documentos_sin_firmas"] = True
    
    return validaciones


def mostrar_panel_validacion():
    """Muestra el panel de validación con criterios de riesgo."""
    st.markdown('<div class="validation-box">', unsafe_allow_html=True)
    st.markdown('<div class="validation-title">🔍 Panel de Validación</div>', unsafe_allow_html=True)
    
    validaciones = validar_visita()
    
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
    
    items_criticos = []
    items_advertencia = []
    items_info = []
    
    for key, (label, icon) in criterios.items():
        if validaciones[key]:
            if icon == "❌":
                items_criticos.append((label, icon, "validation-error"))
            elif icon == "⚠️":
                items_advertencia.append((label, icon, "validation-warning"))
            else:
                items_info.append((label, icon, "validation-error"))
    
    # Mostrar críticos primero
    for label, icon, clase in items_criticos + items_advertencia + items_info:
        st.markdown(
            f'<div class="validation-item {clase}">'
            f'<div class="validation-icon">{icon}</div>'
            f'<div class="validation-text">{label}</div>'
            f'</div>',
            unsafe_allow_html=True
        )
    
    if not any(validaciones.values()):
        st.markdown(
            '<div class="validation-item validation-ok">'
            '<div class="validation-icon">✅</div>'
            '<div class="validation-text">Todas las validaciones OK</div>'
            '</div>',
            unsafe_allow_html=True
        )
    
    st.markdown('</div>', unsafe_allow_html=True)


# --------------------------------------------------------------------------
# SIDEBAR: carga de Excel + búsqueda de cliente
# --------------------------------------------------------------------------
with st.sidebar:
    st.markdown("")
    st.caption("")
    st.divider()
    st.markdown("### 📂 Base de clientes (Excel)")
    excel_file = st.file_uploader(
        "Cargar archivo .xlsx exportado del sistema",
        type=["xlsx", "xls"],
        help="Debe tener las columnas: RECNO, CODCLI, CLIENTE, PENDOC, etc.",
    )
    if excel_file is not None:
        try:
            df = pd.read_excel(excel_file, dtype=str)
            df.columns = [c.strip().upper() for c in df.columns]
            faltantes = [c for c in EXCEL_COLUMNS if c not in df.columns]
            st.session_state.clientes_df = df
            st.success(f"✅ {len(df)} registros cargados")
            if faltantes:
                st.warning(
                    "Columnas no encontradas (se usarán vacías): "
                    + ", ".join(faltantes[:6])
                    + ("..." if len(faltantes) > 6 else "")
                )
        except Exception as e:
            st.error(f"No se pudo leer el archivo: {e}")

    st.divider()
    df = st.session_state.clientes_df
    if df is not None:
        st.markdown("### Buscar cliente")
        busq = st.text_input("DNI, código de cliente o nombre", key="busqueda_cliente")
        if busq:
            b = busq.strip().lower()
            mask = (
                df.get("PENDOC", pd.Series("", index=df.index)).astype(str).str.contains(b, case=False, na=False)
                | df.get("CODCLI", pd.Series("", index=df.index)).astype(str).str.contains(b, case=False, na=False)
                | df.get("CLIENTE", pd.Series("", index=df.index)).astype(str).str.contains(b, case=False, na=False)
            )
            resultados = df[mask]
        else:
            resultados = df

        if len(resultados) == 0:
            st.info("Sin resultados.")
        else:
            opciones = resultados.apply(
                lambda r: f"{safe_str(r.get('CODCLI'))} | {safe_str(r.get('CLIENTE'))} | DNI {safe_str(r.get('PENDOC'))}",
                axis=1,
            ).tolist()
            sel = st.selectbox("Selecciona el cliente", opciones, key="select_cliente")
            if sel:
                idx_sel = opciones.index(sel)
                fila = resultados.iloc[idx_sel].to_dict()
                if st.button("➡️ Usar este cliente", use_container_width=True):
                    st.session_state.cliente_actual = fila
                    st.session_state.visitas = {}
                    st.rerun()
    else:
        st.info("Sube el Excel para buscar clientes, o llena los datos manualmente en la pestaña 1.")

    st.divider()
    if st.session_state.cliente_actual:
        st.success(f"Cliente activo:\n**{safe_str(st.session_state.cliente_actual.get('CLIENTE'))}**")
        if st.button("🗑️ Limpiar visita actual", use_container_width=True):
            st.session_state.cliente_actual = {}
            st.session_state.visitas = {}
            st.session_state.garantias = []
            st.session_state.rcc = []
            st.rerun()


cliente = st.session_state.cliente_actual

st.title("Visita a Clientes")
st.caption("Formulario digital de verificación")

tabs = st.tabs([
    "1️⃣ Cliente y Crédito",
    "2️⃣ Historial y Riesgo",
    "3️⃣ Visita Domicilio",
    "4️⃣ Visita Negocio",
    "5️⃣ Ingresos y Gastos",
    "6️⃣ Garantías y Aval",
    "7️⃣ Generar Reporte",
])

# --------------------------------------------------------------------------
# TAB 1 — Datos del cliente y crédito vigente
# --------------------------------------------------------------------------
with tabs[0]:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.subheader("Titular")
    c1, c2, c3 = st.columns([1, 1, 1])
    with c1:
        agencia = st.text_input("Agencia", value=safe_str(cliente.get("AGENCIA")))
        dni = st.text_input("DNI / LE Titular", value=safe_str(cliente.get("PENDOC")))
        codcli = st.text_input("Código de cliente", value=safe_str(cliente.get("CODCLI")))
    with c2:
        titular = st.text_input("Nombre del titular", value=safe_str(cliente.get("CLIENTE")))
        cuenta = st.text_input("Cuenta cliente", value=safe_str(cliente.get("BCCTA")))
        operacion = st.text_input("Nro. de operación", value=safe_str(cliente.get("BCOPER")))
    with c3:
        analista = st.text_input("Analista vigente", value=safe_str(cliente.get("ANALISTA")))
        analista_eval = st.text_input("Analista evaluador", value=safe_str(cliente.get("ANALISTA_EVAL")))
        aprobado_por = st.text_input("Aprobado por", value=safe_str(cliente.get("USUARIO_APROB")))
    st.markdown("</div>", unsafe_allow_html=True)

    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.subheader("Datos del crédito")
    c1, c2, c3 = st.columns([1, 1, 1])
    with c1:
        importe = st.number_input("Importe desembolsado (S/.)", value=safe_float(cliente.get("IMPDESEMB_MN")), format="%.2f")
        saldo_capital = st.number_input("Saldo capital (S/.)", value=safe_float(cliente.get("SALDO_MN")), format="%.2f")
        tipo_credito = st.text_input("Tipo de crédito", value=safe_str(cliente.get("PRODUCTO_CAJA")))
    with c2:
        tipo_sbs = st.text_input("Tipo según SBS", value=safe_str(cliente.get("TIPO_SBS")))
        fecha_desembolso = st.text_input("Fecha de desembolso", value=safe_str(cliente.get("FECDES")))
        cuotas_pagadas = st.text_input("Nro. cuotas pagadas", value=safe_str(cliente.get("CUOTAS_PAGADAS")))
    with c3:
        dias_atraso = st.text_input("Días de atraso", value=safe_str(cliente.get("DIAS_ATRASO")))
        prom_mora = st.text_input("Promedio de mora", value=safe_str(cliente.get("MORA_CONT")))
        calificacion = st.text_input("Calificación", value=safe_str(cliente.get("CATEG_RESULTANTE")))
    c1, c2, c3 = st.columns([1, 1, 1])
    with c1:
        rubro = st.text_input("Rubro / Actividad económica", value=safe_str(cliente.get("ACTIVIDAD_ECON")))
    with c2:
        sector = st.text_input("Segmentación MYPE", value=safe_str(cliente.get("SEGMENTACION_MYPE")))
    with c3:
        modulo = st.text_input("Módulo", value=safe_str(cliente.get("MODULO")))
    st.markdown("</div>", unsafe_allow_html=True)
    
    # Panel de validación
    mostrar_panel_validacion()

# --------------------------------------------------------------------------
# TAB 2 — Historial crediticio y riesgo de sobreendeudamiento
# --------------------------------------------------------------------------
with tabs[1]:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.subheader("3. Historial crediticio")
    df = st.session_state.clientes_df
    if df is not None and (dni or codcli):
        mask = pd.Series(False, index=df.index)
        if dni:
            mask = mask | (df.get("PENDOC", "").astype(str) == dni)
        if codcli:
            mask = mask | (df.get("CODCLI", "").astype(str) == codcli)
        hist = df[mask]
        cols_show = [c for c in [
            "AGENCIA", "CODCRE", "ESTADO_CREDITO", "FECDES", "FECHA_UTLPAGO",
            "PRODUCTO_CAJA", "SALDO_MN", "DIAS_ATRASO", "ANALISTA",
        ] if c in hist.columns]
        if len(hist):
            st.dataframe(hist[cols_show], use_container_width=True, hide_index=True)
        else:
            st.info("No se encontraron otros créditos para este cliente en el Excel cargado.")
    else:
        st.info("Carga el Excel y selecciona un cliente para ver su historial automáticamente.")
    st.markdown("</div>", unsafe_allow_html=True)

    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.subheader("4. Riesgo de sobreendeudamiento")
    deuda_directa_auto = 0.0
    if df is not None and (dni or codcli):
        for col in ["SALDO_VIGE", "SALDO_REFI"]:
            if col in hist.columns:
                deuda_directa_auto += pd.to_numeric(hist[col], errors="coerce").fillna(0).sum()

    c1, c2 = st.columns(2)
    with c1:
        deuda_directa = st.number_input("a) Deuda directa (S/.)", value=float(deuda_directa_auto), format="%.2f")
        deuda_potencial = st.number_input("b) Deuda potencial (S/.)", value=0.0, format="%.2f")
        deuda_total = st.number_input("c) Deuda total (S/.)", value=float(deuda_directa) + float(deuda_potencial), format="%.2f")
    with c2:
        resultado_neto_rs = st.number_input("e) Cuota / Resultado neto (%)", value=0.0, format="%.2f")
        pasivo_patrimonio = st.number_input("f) Pasivo / Patrimonio (%)", value=0.0, format="%.2f")
        tipo_deuda = st.selectbox("d) Tipo", ["Sin identificación", "MN", "ME"])
    st.markdown("</div>", unsafe_allow_html=True)


# --------------------------------------------------------------------------
# Bloque reutilizable: registro de visita con foto + hora + GPS
# --------------------------------------------------------------------------
def bloque_verificacion(clave, etiqueta):
    """Dibuja el bloque de 'Fecha/Hora/Lugar/Foto' para domicilio, negocio o aval."""
    st.markdown("##### 📍 Registro de verificación in situ")
    visitas = st.session_state.visitas
    data = visitas.get(clave, {})

    colf1, colf2 = st.columns(2)
    with colf1:
        fecha_v = st.date_input("Fecha de visita", value=datetime.now().date(), key=f"fecha_{clave}")
    with colf2:
        hora_v = st.time_input("Hora de visita", value=datetime.now().time(), key=f"hora_{clave}")

    entrevista_con = st.text_input("Entrevista con", key=f"entrevista_{clave}")
    comentarios = st.text_area("Comentarios", key=f"comentarios_{clave}")

    # Ubicación GPS
    st.markdown("**Ubicación GPS**")
    cgps1, cgps2 = st.columns([1, 2])
    with cgps1:
        capturar = st.button("📡 Capturar ubicación actual", key=f"btn_gps_{clave}")
    lat, lon, precision = data.get("lat"), data.get("lon"), data.get("precision")
    if capturar:
        if GEO_OK:
            loc = get_geolocation(key=f"geo_{clave}_{datetime.now().timestamp()}")
            if loc and "coords" in loc:
                lat = loc["coords"]["latitude"]
                lon = loc["coords"]["longitude"]
                precision = loc["coords"].get("accuracy")
            else:
                st.warning("No se pudo obtener la ubicación. Acepta el permiso de ubicación en el navegador e inténtalo de nuevo.")
        else:
            st.warning("El módulo de geolocalización no está instalado en este entorno. Ingresa la dirección manualmente abajo.")
    with cgps2:
        if lat and lon:
            st.success(f"Lat: {lat:.6f}  |  Lon: {lon:.6f}" + (f"  (±{precision:.0f} m)" if precision else ""))
            st.map(pd.DataFrame({"lat": [lat], "lon": [lon]}), zoom=15, height=180)
        else:
            st.caption("Sin ubicación capturada todavía. También puedes dejar solo la dirección escrita.")

    # Foto de verificación
    st.markdown("**Foto de verificación**")
    cfoto1, cfoto2 = st.columns(2)
    with cfoto1:
        foto_camara = st.camera_input("Tomar foto ahora (recomendado)", key=f"camara_{clave}")
    with cfoto2:
        foto_archivo = st.file_uploader("...o subir desde galería", type=["jpg", "jpeg", "png"], key=f"upload_{clave}")
    foto_final = foto_camara if foto_camara is not None else foto_archivo

    if st.button(f"💾 Registrar visita de {etiqueta}", key=f"guardar_{clave}", use_container_width=True):
        st.session_state.visitas[clave] = {
            "fecha": str(fecha_v),
            "hora": str(hora_v),
            "entrevista_con": entrevista_con,
            "comentarios": comentarios,
            "lat": lat,
            "lon": lon,
            "precision": precision,
            "foto_bytes": foto_final.getvalue() if foto_final is not None else None,
        }
        st.success(f"✅ Visita de {etiqueta} registrada a las {hora_v} del {fecha_v}.")


# --------------------------------------------------------------------------
# TAB 3 — Visita al domicilio
# --------------------------------------------------------------------------
with tabs[2]:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.subheader("5. Dirección del domicilio")
    c1, c2 = st.columns(2)
    with c1:
        direccion_dom = st.text_input("Dirección", value=safe_str(cliente.get("DIRECCION_DOM")), key="dir_dom")
        distrito_dom = st.text_input("Distrito", value=safe_str(cliente.get("DISTRITO_DOM")), key="dist_dom")
    with c2:
        provincia_dom = st.text_input("Provincia", value=safe_str(cliente.get("PROVINCIA_DOM")), key="prov_dom")
        departamento_dom = st.text_input("Departamento", value=safe_str(cliente.get("DEPARTAMENTO_DOM")), key="depto_dom")
    referencia_dom = st.text_area("Referencia", key="ref_dom")
    tipo_vivienda = st.selectbox("Tipo de vivienda", ["Propia", "Familiar", "Alquilada", "Otro"], key="tipo_viv")
    st.divider()
    bloque_verificacion("domicilio", "Domicilio")
    st.markdown("</div>", unsafe_allow_html=True)

# --------------------------------------------------------------------------
# TAB 4 — Visita al negocio
# --------------------------------------------------------------------------
with tabs[3]:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.subheader("6. Dirección del negocio")
    c1, c2 = st.columns(2)
    with c1:
        direccion_neg = st.text_input("Dirección", value=safe_str(cliente.get("DIRECCION_NEG")), key="dir_neg")
        distrito_neg = st.text_input("Distrito", value=safe_str(cliente.get("DISTRITO_NEG")), key="dist_neg")
    with c2:
        provincia_neg = st.text_input("Provincia", value=safe_str(cliente.get("PROVINCIA_NEG")), key="prov_neg")
        departamento_neg = st.text_input("Departamento", value=safe_str(cliente.get("DEPARTAMENTO_NEG")), key="depto_neg")
    referencia_neg = st.text_area("Referencia", key="ref_neg")
    tipo_negocio = st.text_input("Tipo de negocio / Actividad principal", value=safe_str(cliente.get("ACTIVIDAD_ECON")), key="tipo_neg")
    st.divider()
    bloque_verificacion("negocio", "Negocio")
    st.markdown("</div>", unsafe_allow_html=True)

# --------------------------------------------------------------------------
# TAB 5 — Verificación de ingresos y gastos
# --------------------------------------------------------------------------
with tabs[4]:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.subheader("8-9. Actividad y ventas")
    c1, c2 = st.columns(2)
    with c1:
        actividad_principal = st.text_input("Actividad principal", value=safe_str(cliente.get("ACTIVIDAD_ECON")), key="act_princ")
    with c2:
        otras_actividades = st.text_input("Otras actividades", key="otras_act")
    ventas = st.number_input("Ventas mensuales (S/.)", value=0.0, format="%.2f", key="ventas")

    st.subheader("10. Costos y gastos")
    c1, c2, c3 = st.columns(3)
    with c1:
        costo_ventas = st.number_input("Costo de ventas", value=0.0, format="%.2f", key="costo_ventas")
        gastos_admin = st.number_input("Gastos administrativos", value=0.0, format="%.2f", key="gastos_admin")
    with c2:
        gastos_financieros = st.number_input("Gastos financieros", value=0.0, format="%.2f", key="gastos_fin")
        gastos_familiares = st.number_input("Gastos familiares", value=0.0, format="%.2f", key="gastos_fam")
    with c3:
        otros_ingresos = st.number_input("Otros ingresos", value=0.0, format="%.2f", key="otros_ing")
        caja_bancos = st.number_input("Caja y bancos", value=0.0, format="%.2f", key="caja_bancos")

    resultado_neto = ventas + otros_ingresos - costo_ventas - gastos_admin - gastos_financieros - gastos_familiares
    utilidad_neta = resultado_neto - gastos_familiares
    c1, c2 = st.columns(2)
    c1.metric("Resultado neto calculado", fmt_money(resultado_neto))
    c2.metric("Utilidad neta calculada", fmt_money(utilidad_neta))
    st.markdown("</div>", unsafe_allow_html=True)

# --------------------------------------------------------------------------
# TAB 6 — Garantías y visita al aval
# --------------------------------------------------------------------------
with tabs[5]:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.subheader("7. Garantías")
    with st.form("form_garantia", clear_on_submit=True):
        c1, c2, c3 = st.columns(3)
        with c1:
            g_desc = st.text_input("Descripción de la garantía")
            g_moneda = st.selectbox("Moneda", ["S/.", "US$"])
        with c2:
            g_importe = st.number_input("Importe", value=0.0, format="%.2f")
            g_perito = st.text_input("Perito")
        with c3:
            g_fecha_tasacion = st.date_input("Fecha de tasación", value=None)
            g_fecha_decl = st.date_input("Fecha de declaración", value=datetime.now().date())
        if st.form_submit_button("➕ Agregar garantía"):
            st.session_state.garantias.append({
                "descripcion": g_desc, "moneda": g_moneda, "importe": g_importe,
                "perito": g_perito, "fecha_tasacion": str(g_fecha_tasacion),
                "fecha_declaracion": str(g_fecha_decl),
            })
    if st.session_state.garantias:
        st.dataframe(pd.DataFrame(st.session_state.garantias), use_container_width=True, hide_index=True)
    st.markdown("</div>", unsafe_allow_html=True)

    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.subheader("VI. Visita al aval / cónyuge")
    cuenta_aval = st.text_input("Cuenta aval", value=safe_str(cliente.get("CUENTA_AVAL")), key="cuenta_aval")
    bloque_verificacion("aval", "Aval")
    st.markdown("</div>", unsafe_allow_html=True)

    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.subheader("VII. Otros datos — Deuda RCC")
    with st.form("form_rcc", clear_on_submit=True):
        c1, c2, c3 = st.columns(3)
        with c1:
            r_entidad = st.text_input("Entidad")
        with c2:
            r_rubro = st.text_input("Rubro")
        with c3:
            r_saldo = st.number_input("Saldo (S/.)", value=0.0, format="%.2f")
        if st.form_submit_button("➕ Agregar entidad RCC"):
            st.session_state.rcc.append({"entidad": r_entidad, "rubro": r_rubro, "saldo": r_saldo})
    if st.session_state.rcc:
        st.dataframe(pd.DataFrame(st.session_state.rcc), use_container_width=True, hide_index=True)
    st.markdown("</div>", unsafe_allow_html=True)


# --------------------------------------------------------------------------
# Generación del reporte Word
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
        ("Fecha", d.get("fecha", "-")),
        ("Hora", d.get("hora", "-")),
        ("Entrevista con", d.get("entrevista_con", "-")),
        ("Ubicación GPS", gps),
        ("Comentarios", d.get("comentarios", "-")),
    ]


def generar_reporte():
    doc = Document()
    doc.add_heading("VISITA A CLIENTES DE PEQUEÑA EMPRESA", level=0)
    p = doc.add_paragraph("")
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_paragraph(f"Generado: {datetime.now().strftime('%d/%m/%Y %H:%M')}")

    add_heading(doc, "I. Datos del cliente")
    add_kv_table(doc, [
        ("Agencia", agencia), ("DNI/LE Titular", dni),
        ("Titular", titular), ("Cuenta cliente", cuenta),
        ("Analista vigente", analista), ("Analista evaluador", analista_eval),
        ("Importe", fmt_money(importe)), ("Saldo capital", fmt_money(saldo_capital)),
        ("Tipo de crédito", tipo_credito), ("Tipo SBS", tipo_sbs),
        ("Días de atraso", dias_atraso), ("Calificación", calificacion),
        ("Rubro", rubro), ("Sector", sector),
    ])

    add_heading(doc, "II. Riesgo de sobreendeudamiento")
    add_kv_table(doc, [
        ("Deuda directa", fmt_money(deuda_directa)),
        ("Deuda potencial", fmt_money(deuda_potencial)),
        ("Deuda total", fmt_money(deuda_total)),
        ("Cuota/Resultado neto", f"{resultado_neto_rs:.2f}%"),
        ("Pasivo/Patrimonio", f"{pasivo_patrimonio:.2f}%"),
    ])

    for clave, titulo, direccion_info in [
        ("domicilio", "III. Visita al domicilio", [
            ("Dirección", direccion_dom), ("Distrito", distrito_dom),
            ("Provincia", provincia_dom), ("Departamento", departamento_dom),
            ("Tipo de vivienda", tipo_vivienda),
        ]),
        ("negocio", "IV. Visita al negocio", [
            ("Dirección", direccion_neg), ("Distrito", distrito_neg),
            ("Provincia", provincia_neg), ("Departamento", departamento_neg),
            ("Actividad", tipo_negocio),
        ]),
        ("aval", "V. Visita al aval", [("Cuenta aval", cuenta_aval)]),
    ]:
        add_heading(doc, titulo)
        add_kv_table(doc, direccion_info)
        verif = visita_a_texto(st.session_state.visitas, clave)
        if verif:
            doc.add_paragraph("Registro de verificación:").bold = True
            add_kv_table(doc, verif)
            foto_bytes = st.session_state.visitas[clave].get("foto_bytes")
            if foto_bytes:
                doc.add_picture(io.BytesIO(foto_bytes), width=Cm(8))
        else:
            doc.add_paragraph("⚠ No se registró visita de verificación para esta sección.")

    add_heading(doc, "VI. Verificación de ingresos y gastos")
    add_kv_table(doc, [
        ("Ventas", fmt_money(ventas)), ("Otros ingresos", fmt_money(otros_ingresos)),
        ("Costo de ventas", fmt_money(costo_ventas)), ("Gastos administrativos", fmt_money(gastos_admin)),
        ("Gastos financieros", fmt_money(gastos_financieros)), ("Gastos familiares", fmt_money(gastos_familiares)),
        ("Resultado neto", fmt_money(resultado_neto)), ("Utilidad neta", fmt_money(utilidad_neta)),
    ])

    if st.session_state.garantias:
        add_heading(doc, "VII. Garantías")
        for g in st.session_state.garantias:
            add_kv_table(doc, list(g.items()))

    if st.session_state.rcc:
        add_heading(doc, "VIII. Deuda RCC")
        for r in st.session_state.rcc:
            add_kv_table(doc, list(r.items()))

    add_heading(doc, "Conformidad")
    add_kv_table(doc, [
        ("Hecho por", ""), ("Fecha", datetime.now().strftime("%d/%m/%Y")),
        ("Revisado por", ""), ("Fecha", ""),
    ])

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf


with tabs[6]:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.subheader("Resumen de visitas registradas")
    for clave, etiqueta in [("domicilio", "Domicilio"), ("negocio", "Negocio"), ("aval", "Aval")]:
        if clave in st.session_state.visitas:
            d = st.session_state.visitas[clave]
            st.markdown(f"**{etiqueta}** — <span class='badge-ok'>Registrada</span>", unsafe_allow_html=True)
            st.caption(f"{d['fecha']} {d['hora']} · {d.get('entrevista_con') or 'sin entrevistado'} · "
                       f"{'con foto' if d.get('foto_bytes') else 'sin foto'} · "
                       f"{'con GPS' if d.get('lat') else 'sin GPS'}")
        else:
            st.markdown(f"**{etiqueta}** — <span class='badge-pend'>Pendiente</span>", unsafe_allow_html=True)
    st.markdown("</div>", unsafe_allow_html=True)

    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.subheader(" Generar reporte Word de la visita")
    st.caption("Incluye todos los datos llenados y las fotos/GPS de las verificaciones registradas.")
    if st.button("Generar documento", type="primary"):
        buf = generar_reporte()
        nombre = f"Visita_{safe_str(titular, 'cliente').replace(' ', '_')}_{datetime.now().strftime('%Y%m%d_%H%M')}.docx"
        st.download_button(
            "⬇️ Descargar reporte (.docx)",
            data=buf,
            file_name=nombre,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
        )
        st.success("Reporte generado. Descárgalo antes de cerrar la app: en el plan gratuito de hosting los archivos no quedan guardados permanentemente en el servidor.")
    st.markdown("</div>", unsafe_allow_html=True)
