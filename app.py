# -*- coding: utf-8 -*-
"""
Formulario de verificación de datos visita - Versión Muestra Final Automática
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

# Colores institucionales
NARANJA = "C8102E"   
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
</style>
"""
st.markdown(CUSTOM_CSS, unsafe_allow_html=True)


# --------------------------------------------------------------------------
# HELPERS DE LIMPIEZA
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


def limpiar_texto_dni(val):
    """Limpia cadenas de DNI quitando decimales flotantes y rellenando ceros."""
    if pd.isna(val) or val is None:
        return ""
    txt = str(val).strip()
    if txt.endswith(".0"):
        txt = txt[:-2]
    # Si son números y le faltan ceros a la izquierda para completar 8 dígitos
    if txt.isdigit() and len(txt) < 8 and len(txt) > 0:
        txt = txt.zfill(8)
    return txt


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
# BÚSQUEDA DE CLIENTE POR DNI
# --------------------------------------------------------------------------
def buscar_cliente_por_dni(dni_input, df):
    if not dni_input or df is None or len(df) == 0:
        return None
    
    txt_buscar = limpiar_texto_dni(dni_input)
    mask = (df.get("PENDOC", pd.Series("", index=df.index)).astype(str) == txt_buscar)
    resultados = df[mask]
    
    if len(resultados) > 0:
        return resultados.iloc[0].to_dict()
    return None


# --------------------------------------------------------------------------
# VALIDACIÓN DE CRITERIOS
# --------------------------------------------------------------------------
def validar_visita():
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
    if not safe_str(cliente.get("CLIENTE")):
        validaciones["documentos_sin_datos"] = True
    if safe_str(cliente.get("DIAS_ATRASO")) and int(safe_float(cliente.get("DIAS_ATRASO"))) > 0:
        validaciones["calificacion_diferente"] = True
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
    st.markdown('<div class="validation-box">', unsafe_allow_html=True)
    st.markdown('<div class="validation-title">🔍 Panel de Validación - Criterios de Riesgo</div>', unsafe_allow_html=True)
    
    validaciones_auto = validar_visita()
    criterios = {
        "documentos_enmiendas": ("Documentos con enmiendas", "⚠️"),
        "documentos_inconsistentes": ("Datos inconsistentes en documentos", "⚠️"),
        "documentos_sin_datos": ("Documentos sin datos del cliente", "❌"),
        "documentos_sin_firmas": ("Documentos sin firmas o fotos", "❌"),
        "documentos_duplicados": ("Documentos duplicados", "⚠️"),
        "sin_sustento_actividad": ("No se evidenció sustento de actividad económica", "❌"),
        "sin_sustento_ingresos": ("No se evidenció sustento de ingresos", "❌"),
        "sin_sustento_activos": ("No se evidenció sustento de activos representativos", "⚠️"),
        "conyuge_omitido": ("Se omitió al cónyuge", "⚠️"),
        "credito_reprogramado": ("Crédito reprogramado", "ℹ️"),
        "credito_refinanciado": ("Crédito refinanciado", "ℹ️"),
        "calificacion_diferente": ("Calificación diferente a la fecha de revisión", "⚠️"),
    }
    
    items_por_categoria = {"❌": [], "⚠️": [], "ℹ️": []}
    for key, (label, icon) in criterios.items():
        items_por_categoria[icon].append((key, label))
    
    for icon in ["❌", "⚠️", "ℹ️"]:
        for key, label in items_por_categoria[icon]:
            is_checked = st.session_state.validaciones_marcadas.get(key, validaciones_auto.get(key, False))
            col1, col2, col3 = st.columns([0.08, 0.08, 0.84])
            with col1:
                st.markdown(f"<div>{icon}</div>", unsafe_allow_html=True)
            with col2:
                st.session_state.validaciones_marcadas[key] = st.checkbox(
                    label="", value=is_checked, key=f"check_{key}", label_visibility="collapsed"
                )
            with col3:
                st.markdown(label)
    st.markdown('</div>', unsafe_allow_html=True)


# --------------------------------------------------------------------------
# SIDEBAR: INTERFAZ DE CARGA OPTIMIZADA (AUTO-MUESTRA FINAL)
# --------------------------------------------------------------------------
with st.sidebar:
    st.markdown("### 📂 Base de clientes (Excel)")
    
    filas_a_saltar = st.number_input(
        "Filas a saltar (Modifica si lee pocos registros):", 
        min_value=0, max_value=20, value=0, step=1,
        help="Si tu Excel exportado tiene títulos arriba, pon 1, 2 o 3 para saltarlos."
    )
    
    excel_file = st.file_uploader(
        "Cargar archivo .xlsx", type=["xlsx", "xls"]
    )
    
    if excel_file is not None:
        try:
            excel_lector = pd.ExcelFile(excel_file)
            lista_hojas = excel_lector.sheet_names
            
            # --- SELECCIÓN INTELIGENTE DE MUESTRA_FINAL ---
            indice_defecto = 0
            for idx, nombre_hoja in enumerate(lista_hojas):
                if str(nombre_hoja).strip().upper() == "MUESTRA_FINAL":
                    indice_defecto = idx
                    break
            
            hoja_seleccionada = st.selectbox(
                "Selecciona la pestaña:", 
                lista_hojas, 
                index=indice_defecto,
                help="Se ha pre-seleccionado automáticamente 'MUESTRA_FINAL' si fue detectada."
            )

            # Carga limpia forzando formato texto
            df = pd.read_excel(excel_file, sheet_name=hoja_seleccionada, skiprows=filas_a_saltar, dtype=str)
            df.columns = [str(c).strip().upper() for c in df.columns]
            
            # --- ASIGNACIÓN CRÍTICA POR POSICIÓN (COLUMNA D) ---
            if len(df.columns) >= 4:
                nombre_original_col_d = df.columns[3] # Índice 3 = 4ta columna (D)
                df = df.rename(columns={nombre_original_col_d: "PENDOC"})
            
            # Normalizar y limpiar todos los DNIs leídos de la columna D
            if "PENDOC" in df.columns:
                df["PENDOC"] = df["PENDOC"].apply(limpiar_texto_dni)
                
            st.session_state.clientes_df = df
            st.success(f"✅ ¡{len(df)} registros cargados de la pestaña '{hoja_seleccionada}'!")
            
            with st.expander("👀 Ver datos leídos (Primeras 5 filas)"):
                st.dataframe(df.head(5))
                
        except Exception as e:
            st.error(f"Error al procesar el archivo: {e}")

    st.divider()
    df = st.session_state.clientes_df
    if df is not None:
        st.markdown("### Búsqueda Manual")
        busq = st.text_input("Ingresa DNI, Código o Nombre:", key="busqueda_cliente")
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

        if len(resultados) > 0:
            opciones = resultados.apply(
                lambda r: f"{safe_str(r.get('CODCLI'))} | {safe_str(r.get('CLIENTE'))} | DNI: {safe_str(r.get('PENDOC'))}",
                axis=1,
            ).tolist()
            sel = st.selectbox("Selecciona el cliente exacto:", opciones, key="select_cliente")
            if sel:
                idx_sel = opciones.index(sel)
                fila = resultados.iloc[idx_sel].to_dict()
                if st.button("➡️ Cargar este cliente", use_container_width=True):
                    st.session_state.cliente_actual = fila
                    st.session_state.visitas = {}
                    st.rerun()
        else:
            st.info("No se encontraron coincidencias en la búsqueda.")

    if st.session_state.cliente_actual:
        st.divider()
        if st.button("🗑️ Limpiar y reiniciar formulario", use_container_width=True):
            st.session_state.cliente_actual = {}
            st.session_state.visitas = {}
            st.session_state.garantias = []
            st.session_state.rcc = []
            st.session_state.validaciones_marcadas = {}
            st.rerun()


cliente = st.session_state.cliente_actual
dni = safe_str(cliente.get("PENDOC"))
codcli = safe_str(cliente.get("CODCLI"))
titular = safe_str(cliente.get("CLIENTE"))
agencia = safe_str(cliente.get("AGENCIA"))
cuenta = safe_str(cliente.get("BCCTA"))
operacion = safe_str(cliente.get("BCOPER"))
analista = safe_str(cliente.get("ANALISTA"))
analista_eval = safe_str(cliente.get("ANALISTA_EVAL"))
aprobado_por = safe_str(cliente.get("USUARIO_APROB"))

# Variables numéricas del crédito
importe = safe_float(cliente.get("IMPDESEMB_MN"))
saldo_capital = safe_float(cliente.get("SALDO_MN"))
tipo_credito = safe_str(cliente.get("PRODUCTO_CAJA"))
tipo_sbs = safe_str(cliente.get("TIPO_SBS"))
fecha_desembolso = safe_str(cliente.get("FECDES"))
cuotas_pagadas = safe_str(cliente.get("CUOTAS_PAGADAS"))
dias_atraso = safe_str(cliente.get("DIAS_ATRASO"))
prom_mora = safe_str(cliente.get("MORA_CONT"))
calificacion = safe_str(cliente.get("CATEG_RESULTANTE"))
rubro = safe_str(cliente.get("ACTIVIDAD_ECON"))
sector = safe_str(cliente.get("SEGMENTACION_MYPE"))
modulo = safe_str(cliente.get("MODULO"))

tabs = st.tabs([
    "1️⃣ Cliente y Crédito", "2️⃣ Historial y Riesgo", "3️⃣ Visita Domicilio",
    "4️⃣ Visita Negocio", "5️⃣ Ingresos y Gastos", "6️⃣ Garantías y Aval", "7️⃣ Generar Reporte"
])

# --------------------------------------------------------------------------
# TAB 1 — DATOS DEL CLIENTE
# --------------------------------------------------------------------------
with tabs[0]:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.subheader("Titular")
    
    dni_input = st.text_input(
        "🔍 Buscar automáticamente por DNI / LE del Titular:", 
        value=dni,
        key="dni_search"
    )
    
    if dni_input and not cliente:
        cliente_encontrado = buscar_cliente_por_dni(dni_input, st.session_state.clientes_df)
        if cliente_encontrado:
            st.session_state.cliente_actual = cliente_encontrado
            st.success(f"✅ ¡Cliente encontrado e importado con éxito!")
            st.rerun()
        elif st.session_state.clientes_df is not None:
            st.error("❌ El número de DNI ingresado no figura en la pestaña del Excel cargada.")
    
    c1, c2, c3 = st.columns(3)
    with c1:
        st.text_input("Agencia", value=agencia, disabled=True)
        st.text_input("DNI / LE Titular", value=dni, disabled=True)
        st.text_input("Código de cliente", value=codcli, disabled=True)
    with c2:
        st.text_input("Nombre del titular", value=titular, disabled=True)
        st.text_input("Cuenta cliente", value=cuenta, disabled=True)
        st.text_input("Nro. de operación", value=operacion, disabled=True)
    with c3:
        st.text_input("Analista vigente", value=analista, disabled=True)
        st.text_input("Analista evaluador", value=analista_eval, disabled=True)
        st.text_input("Aprobado por", value=aprobado_por, disabled=True)
    st.markdown("</div>", unsafe_allow_html=True)

    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.subheader("Datos del crédito")
    c1, c2, c3 = st.columns(3)
    with c1:
        st.number_input("Importe desembolsado (S/.)", value=importe, format="%.2f", disabled=True)
        st.number_input("Saldo capital (S/.)", value=saldo_capital, format="%.2f", disabled=True)
        st.text_input("Tipo de crédito", value=tipo_credito, disabled=True)
    with c2:
        st.text_input("Tipo según SBS", value=tipo_sbs, disabled=True)
        st.text_input("Fecha de desembolso", value=fecha_desembolso, disabled=True)
        st.text_input("Nro. cuotas pagadas", value=cuotas_pagadas, disabled=True)
    with c3:
        st.text_input("Días de atraso", value=dias_atraso, disabled=True)
        st.text_input("Promedio de mora", value=prom_mora, disabled=True)
        st.text_input("Calificación", value=calificacion, disabled=True)
    st.markdown("</div>", unsafe_allow_html=True)
    
    st.markdown('<div class="card">', unsafe_allow_html=True)
    mostrar_panel_validacion()
    st.markdown("</div>", unsafe_allow_html=True)


# --------------------------------------------------------------------------
# TAB 2 — HISTORIAL Y RIESGO
# --------------------------------------------------------------------------
with tabs[1]:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.subheader("3. Historial crediticio")
    df = st.session_state.clientes_df
    if df is not None and (dni or codcli):
        mask = pd.Series(False, index=df.index)
        if dni:
            mask = mask | (df["PENDOC"].astype(str) == dni)
        if codcli:
            mask = mask | (df["CODCLI"].astype(str) == codcli)
        hist = df[mask]
        cols_show = [c for c in ["AGENCIA", "CODCRE", "ESTADO_CREDITO", "FECDES", "PRODUCTO_CAJA", "SALDO_MN", "DIAS_ATRASO"] if c in hist.columns]
        if len(hist) > 0:
            st.dataframe(hist[cols_show], use_container_width=True, hide_index=True)
        else:
            st.info("No se registraron otros créditos históricos.")
    else:
        st.info("Carga un cliente desde la barra lateral para calcular el historial.")
    st.markdown("</div>", unsafe_allow_html=True)

    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.subheader("4. Riesgo de sobreendeudamiento")
    deuda_directa_auto = 0.0
    if df is not None and (dni or codcli) and 'hist' in locals() and len(hist) > 0:
        for col in ["SALDO_VIGE", "SALDO_REFI", "SALDO_MN"]:
            if col in hist.columns:
                deuda_directa_auto = pd.to_numeric(hist[col], errors="coerce").fillna(0).sum()
                break

    c1, c2 = st.columns(2)
    with c1:
        deuda_directa = st.number_input("a) Deuda directa (S/.)", value=float(deuda_directa_auto), format="%.2f")
        deuda_potencial = st.number_input("b) Deuda potencial (S/.)", value=0.0, format="%.2f")
        deuda_total = st.number_input("c) Deuda total (S/.)", value=float(deuda_directa) + float(deuda_potencial), format="%.2f")
    with c2:
        resultado_neto_rs = st.number_input("e) Cuota / Resultado neto (%)", value=0.0, format="%.2f")
        pasivo_patrimonio = st.number_input("f) Pasivo / Patrimonio (%)", value=0.0, format="%.2f")
        tipo_deuda = st.selectbox("d) Tipo Moneda", ["Sin identificación", "MN", "ME"])
    st.markdown("</div>", unsafe_allow_html=True)


# --------------------------------------------------------------------------
# BLOQUE DINÁMICO DE VISITA IN SITU
# --------------------------------------------------------------------------
def bloque_verificacion(clave, etiqueta):
    st.markdown("##### 📍 Registro de verificación in situ")
    visitas = st.session_state.visitas
    data = visitas.get(clave, {})

    colf1, colf2 = st.columns(2)
    with colf1:
        fecha_v = st.date_input("Fecha de visita", value=datetime.now().date(), key=f"fecha_{clave}")
    with colf2:
        hora_v = st.time_input("Hora de visita", value=datetime.now().time(), key=f"hora_{clave}")

    entrevista_con = st.text_input("Entrevista con", key=f"entrevista_{clave}")
    comentarios = st.text_area("Comentarios / Observaciones", key=f"comentarios_{clave}")

    st.markdown("**Ubicación GPS**")
    capturar = st.button("📡 Capturar coordenadas", key=f"btn_gps_{clave}")
    lat, lon, precision = data.get("lat"), data.get("lon"), data.get("precision")
    
    if capturar:
        if GEO_OK:
            loc = get_geolocation(key=f"geo_{clave}_{datetime.now().timestamp()}")
            if loc and "coords" in loc:
                lat = loc["coords"]["latitude"]
                lon = loc["coords"]["longitude"]
                precision = loc["coords"].get("accuracy")
            else:
                st.warning("Permiso de ubicación denegado por el navegador.")
        else:
            st.warning("Geolocalización no disponible.")
            
    if lat and lon:
        st.success(f"Lat: {lat:.6f} | Lon: {lon:.6f} (±{precision:.0f}m)")
        st.map(pd.DataFrame({"lat": [lat], "lon": [lon]}), zoom=15, height=150)

    st.markdown("**Fotografía de Respaldo**")
    cfoto1, cfoto2 = st.columns(2)
    with cfoto1:
        foto_camara = st.camera_input("Capturar con cámara", key=f"camara_{clave}")
    with cfoto2:
        foto_archivo = st.file_uploader("Subir archivo de imagen", type=["jpg", "png"], key=f"upload_{clave}")
    foto_final = foto_camara if foto_camara is not None else foto_archivo

    if st.button(f"💾 Guardar Visita de {etiqueta}", key=f"guardar_{clave}", use_container_width=True):
        st.session_state.visitas[clave] = {
            "fecha": str(fecha_v), "hora": str(hora_v), "entrevista_con": entrevista_con,
            "comentarios": comentarios, "lat": lat, "lon": lon, "precision": precision,
            "foto_bytes": foto_final.getvalue() if foto_final is not None else None,
        }
        st.success(f"✅ Cambios guardados para {etiqueta}.")


# --------------------------------------------------------------------------
# TABS FORMULARIOS COMPLETOS
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
    referencia_dom = st.text_area("Referencia de acceso", key="ref_dom")
    tipo_vivienda = st.selectbox("Condición de vivienda", ["Propia", "Familiar", "Alquilada", "Otro"], key="tipo_viv")
    st.divider()
    bloque_verificacion("domicilio", "Domicilio")
    st.markdown("</div>", unsafe_allow_html=True)

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
    referencia_neg = st.text_area("Referencia del local comercial", key="ref_neg")
    tipo_negocio = st.text_input("Giro / Actividad principal", value=rubro, key="tipo_neg")
    st.divider()
    bloque_verificacion("negocio", "Negocio")
    st.markdown("</div>", unsafe_allow_html=True)

with tabs[4]:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.subheader("8-9. Flujo Financiero Mensual")
    c1, c2 = st.columns(2)
    with c1:
        actividad_principal = st.text_input("Actividad Declarada", value=rubro, key="act_princ")
        ventas = st.number_input("Ventas totales del mes (S/.)", value=0.0, format="%.2f", key="ventas")
        costo_ventas = st.number_input("Costo directo de ventas", value=0.0, format="%.2f", key="costo_ventas")
    with c2:
        otras_actividades = st.text_input("Otras fuentes de ingresos", key="otras_act")
        gastos_admin = st.number_input("Gastos operativos y administrativos", value=0.0, format="%.2f", key="gastos_admin")
        gastos_financieros = st.number_input("Obligaciones financieras externas", value=0.0, format="%.2f", key="gastos_fin")
    
    st.divider()
    c3, c4 = st.columns(2)
    with c3:
        gastos_familiares = st.number_input("Gastos familiares / Canasta básica", value=0.0, format="%.2f", key="gastos_fam")
    with c4:
        otros_ingresos = st.number_input("Otros ingresos demostrables", value=0.0, format="%.2f", key="otros_ing")

    resultado_neto = ventas + otros_ingresos - costo_ventas - gastos_admin - gastos_financieros - gastos_familiares
    utilidad_neta = resultado_neto - gastos_familiares
    
    cc1, cc2 = st.columns(2)
    cc1.metric("Resultado Neto de Operación", fmt_money(resultado_neto))
    cc2.metric("Excedente / Utilidad Disponible", fmt_money(utilidad_neta))
    st.markdown("</div>", unsafe_allow_html=True)

with tabs[5]:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.subheader("7. Garantías Reales / Mobiliarias")
    with st.form("form_garantia", clear_on_submit=True):
        c1, c2, c3 = st.columns(3)
        with c1:
            g_desc = st.text_input("Descripción de la garantía")
            g_moneda = st.selectbox("Moneda", ["S/.", "US$"])
        with c2:
            g_importe = st.number_input("Valor Comercial", value=0.0, format="%.2f")
            g_perito = st.text_input("Perito Valuador")
        with c3:
            g_fecha_tasacion = st.date_input("Fecha Tasación", value=None)
            g_fecha_decl = st.date_input("Fecha Declarativa", value=datetime.now().date())
        if st.form_submit_button("➕ Registrar Garantía"):
            st.session_state.garantias.append({
                "descripcion": g_desc, "moneda": g_moneda, "importe": g_importe,
                "perito": g_perito, "fecha_tasacion": str(g_fecha_tasacion), "fecha_declaracion": str(g_fecha_decl),
            })
    if st.session_state.garantias:
        st.dataframe(pd.DataFrame(st.session_state.garantias), use_container_width=True, hide_index=True)
    st.markdown("</div>", unsafe_allow_html=True)

    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.subheader("VI. Aval / Codeudor Solidario")
    cuenta_aval = st.text_input("Código o Cuenta Aval:", value=safe_str(cliente.get("CUENTA_AVAL")), key="cuenta_aval")
    bloque_verificacion("aval", "Aval")
    st.markdown("</div>", unsafe_allow_html=True)


# --------------------------------------------------------------------------
# TAB 7 — DESCARGA DE REPORTE WORD
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

def generar_reporte():
    doc = Document()
    doc.add_heading("INFORME DIGITAL - REVISIÓN DE CAMPO PEQUEÑA EMPRESA", level=0)
    doc.add_paragraph(f"Fecha Impresión: {datetime.now().strftime('%d/%m/%Y %H:%M')}")

    add_heading(doc, "I. Datos de Identificación")
    add_kv_table(doc, [
        ("Agencia Asignada", agencia), ("DNI/LE Titular", dni),
        ("Razón Social / Nombre", titular), ("Código Único", codcli),
        ("Analista Evaluador", analista_eval), ("Monto Otorgado", fmt_money(importe)),
        ("Saldo Insoluto", fmt_money(saldo_capital)), ("Calificación Interna", calificacion),
    ])

    for clave, titulo, dir_info in [
        ("domicilio", "II. Verificación Domiciliaria", [("Dirección", direccion_dom), ("Distrito", distrito_dom), ("Vivienda", tipo_vivienda)]),
        ("negocio", "III. Verificación de Local de Negocio", [("Dirección", direccion_neg), ("Giro Comercial", tipo_negocio)]),
    ]:
        add_heading(doc, titulo)
        add_kv_table(doc, dir_info)
        vis = st.session_state.visitas.get(clave)
        if vis:
            add_kv_table(doc, [
                ("Entrevistado", vis.get("entrevista_con")),
                ("Coordenadas GPS", f"{vis.get('lat')}, {vis.get('lon')}"),
                ("Observaciones", vis.get("comentarios")),
            ])
            if vis.get("foto_bytes"):
                doc.add_picture(io.BytesIO(vis["foto_bytes"]), width=Cm(7))
        else:
            doc.add_paragraph("⚠️ Sección sin validación física en campo.")

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf

with tabs[6]:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.subheader("Estado de Carga de Formularios")
    for k, v in [("domicilio", "Domicilio"), ("negocio", "Negocio"), ("aval", "Aval")]:
        if k in st.session_state.visitas:
            st.markdown(f"✅ **{v}** — <span class='badge-ok'>Completado</span>", unsafe_allow_html=True)
        else:
            st.markdown(f"❌ **{v}** — <span class='badge-pend'>Falta Registro</span>", unsafe_allow_html=True)
    
    st.divider()
    if st.button("💾 Compilar Informe Final (.docx)", type="primary", use_container_width=True):
        doc_buf = generar_reporte()
        st.download_button(
            "⬇️ Hacer click aquí para descargar Reporte",
            data=doc_buf,
            file_name=f"Informe_Visita_{dni}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True
        )
    st.markdown("</div>", unsafe_allow_html=True)
