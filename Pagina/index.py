import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import io
import re
import os
from datetime import datetime, date
import warnings
warnings.filterwarnings("ignore")

# ─────────────────────────────────────────────
# RUTA FIJA: Dataset de empresas reguladas
# ─────────────────────────────────────────────
# El archivo se carga automáticamente al iniciar la aplicación.
# Cambia esta ruta si mueves el archivo a otra carpeta.
_EMPRESAS_REG_FILE = "Empresas_reguladas_verificación_cumplimiento.xlsm"
_LOGO_FILE = "logoSena.png"

def obtener_ruta_logo():
    """Obtiene la ruta correcta del logo."""
    candidatos = [
        os.path.join(os.path.dirname(os.path.abspath(__file__)), _LOGO_FILE),
        os.path.join(os.getcwd(), _LOGO_FILE),
        _LOGO_FILE,
    ]
    for ruta in candidatos:
        if os.path.exists(ruta):
            return ruta
    return None

@st.cache_data(show_spinner=False)
def _cargar_empresas_reguladas():
    """Carga el dataset de empresas reguladas desde la ruta fija."""
    # Buscar en el mismo directorio que index.py o en el directorio actual
    candidatos = [
        os.path.join(os.path.dirname(os.path.abspath(__file__)), _EMPRESAS_REG_FILE),
        os.path.join(os.getcwd(), _EMPRESAS_REG_FILE),
        _EMPRESAS_REG_FILE,
    ]
    for ruta in candidatos:
        if os.path.exists(ruta):
            try:
                df_reg = pd.read_excel(ruta, engine="openpyxl")
                # Normalizar nombres de columnas clave
                df_reg.columns = [str(c).strip() for c in df_reg.columns]
                # Limpiar espacios en columnas de actividad y razón social
                for col in ["Actividad económica", "Descripción actividad económica", "Razon social"]:
                    if col in df_reg.columns:
                        df_reg[col] = df_reg[col].astype(str).str.strip()
                return df_reg, ruta
            except Exception as e:
                return None, str(e)
    return None, f"Archivo no encontrado: {_EMPRESAS_REG_FILE}"

# ─────────────────────────────────────────────
# CONFIGURACIÓN GLOBAL
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="SISCON Gestión Contratos de Aprendizaje",
    page_icon="logoSena.png",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─────────────────────────────────────────────
# ESTILOS CSS SENA (blanco y verde)
# ─────────────────────────────────────────────
st.markdown("""
<style>
/* Fuente principal */
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');

html, body, [class*="css"] { font-family: 'Inter', sans-serif; }

/* Sidebar */
section[data-testid="stSidebar"] {
    background: linear-gradient(180deg, #27AE60 0%, #1E8449 60%, #196F3D 100%);
    color: white;
    overflow: hidden !important;
}
section[data-testid="stSidebar"] * { color: white !important; }
section[data-testid="stSidebar"] .stRadio label { color: white !important; }

/* Encabezado lateral */
.sena-logo {
    text-align: center;
    padding: 18px 10px 10px 10px;
    border-bottom: 1px solid rgba(255,255,255,0.25);
    margin-bottom: 10px;
}
.sena-logo h2 { color: white !important; font-size: 1.3rem; margin: 0; }
.sena-logo p  { color: rgba(255,255,255,0.75) !important; font-size: 0.78rem; margin: 0; }

/* Tarjetas métricas */
.metric-card {
    background: white;
    border-left: 5px solid #27AE60;
    border-radius: 8px;
    padding: 16px 20px;
    box-shadow: 0 2px 8px rgba(0,0,0,0.08);
    margin-bottom: 12px;
}
.metric-card .metric-value { font-size: 2rem; font-weight: 700; color: #27AE60; }
.metric-card .metric-label { font-size: 0.82rem; color: #555; font-weight: 500; }
.metric-card .metric-delta { font-size: 0.78rem; color: #888; margin-top: 2px; }

/* Alertas */
.alert-warning {
    background: #fff8e1; border-left: 4px solid #F39C12;
    border-radius: 6px; padding: 12px 16px; margin: 8px 0;
    font-size: 0.88rem; color: #555;
}
.alert-info {
    background: #EAFAF1; border-left: 4px solid #27AE60;
    border-radius: 6px; padding: 12px 16px; margin: 8px 0;
    font-size: 0.88rem; color: #333;
}
.alert-error {
    background: #ffebee; border-left: 4px solid #c62828;
    border-radius: 6px; padding: 12px 16px; margin: 8px 0;
    font-size: 0.88rem; color: #555;
}

/* Sección titular */
.section-header {
    background: linear-gradient(90deg, #27AE60, #2ECC71);
    color: white;
    padding: 14px 22px;
    border-radius: 8px;
    margin-bottom: 20px;
    display: flex; align-items: center; gap: 10px;
}
.section-header h2 { margin: 0; font-size: 1.2rem; color: white; }
.section-header p  { margin: 0; font-size: 0.8rem; color: rgba(255,255,255,0.85); }

/* Tablas de datos */
.dataframe thead tr th {
    background-color: #27AE60 !important;
    color: white !important;
    font-weight: 600 !important;
}
.dataframe tbody tr:nth-child(even) { background-color: #f0faf4 !important; }

/* Botones */
.stButton > button {
    background: #27AE60; color: white; border: none;
    border-radius: 6px; font-weight: 600;
    transition: background 0.2s;
}
.stButton > button:hover { background: #1E8449; }

/* Selectbox y multiselect */
.stSelectbox label, .stMultiSelect label { font-weight: 600; color: #333; }

/* Tabs */
button[data-baseweb="tab"] { font-weight: 600; }
button[data-baseweb="tab"][aria-selected="true"] {
    color: #27AE60 !important;
    border-bottom-color: #27AE60 !important;
}

/* Chips de estado */
.badge-green  { background:#d4edda; color:#155724; border-radius:12px; padding:2px 10px; font-size:0.78rem; font-weight:600; }
.badge-yellow { background:#fff3cd; color:#856404; border-radius:12px; padding:2px 10px; font-size:0.78rem; font-weight:600; }
.badge-red    { background:#f8d7da; color:#721c24; border-radius:12px; padding:2px 10px; font-size:0.78rem; font-weight:600; }
.badge-blue   { background:#cce5ff; color:#004085; border-radius:12px; padding:2px 10px; font-size:0.78rem; font-weight:600; }

/* Scrollbar */
::-webkit-scrollbar { width: 6px; height: 6px; }
::-webkit-scrollbar-track { background: #f1f1f1; }
::-webkit-scrollbar-thumb { background: #27AE60; border-radius: 3px; }

/* Bloquear scroll del panel lateral */
[data-testid="stAppViewContainer"] {
    overflow-y: hidden;
}
</style>
""", unsafe_allow_html=True)

# ─────────────────────────────────────────────
# CONSTANTES DE NEGOCIO
# ─────────────────────────────────────────────
NIVEL_CONFIG = {
    "TÉCNICO":    {"meses": 15, "trimestres": 5, "lectiva": [1,2,3], "productiva": [4,5]},
    "TECNÓLOGO":  {"meses": 24, "trimestres": 8, "lectiva": [1,2,3,4,5,6], "productiva": [7,8]},
    "OPERARIO":   {"meses": 6,  "trimestres": 2, "lectiva": [1], "productiva": [2]},
    "AUXILIAR":   {"meses": 12, "trimestres": 4, "lectiva": [1,2,3], "productiva": [4]},
}

COLORES_SENA = {
    "primario":   "#27AE60",   # verde claro SENA
    "secundario": "#2980B9",   # azul
    "acento":     "#F39C12",   # naranja
    "rojo":       "#E74C3C",   # rojo
    "azul":       "#2980B9",   # azul
    "morado":     "#8E44AD",   # morado
    "cyan":       "#16A085",   # teal
    "gris":       "#7F8C8D",   # gris
    "bg":         "#F4F6F8",
}

# Paleta variada: verde, azul, naranja, morado, teal, rojo, amarillo, cyan, rosa, gris
PALETA = [
    "#27AE60",  # verde
    "#2980B9",  # azul
    "#F39C12",  # naranja
    "#8E44AD",  # morado
    "#16A085",  # teal
    "#E74C3C",  # rojo
    "#F1C40F",  # amarillo
    "#1ABC9C",  # cyan
    "#E91E8C",  # rosa
    "#7F8C8D",  # gris
]

# Mapa fijo por nivel y etapa (garantiza colores consistentes en todas las páginas)
COLOR_NIVEL = {
    "TÉCNICO":    "#27AE60",
    "TECNÓLOGO":  "#2980B9",
    "OPERARIO":   "#F39C12",
    "AUXILIAR":   "#8E44AD",
    "OTRO":       "#7F8C8D",
}
COLOR_ETAPA = {
    "LECTIVA":    "#27AE60",
    "PRODUCTIVA": "#F39C12",
    "COMPLETADO": "#2980B9",
    "DESCONOCIDA":"#7F8C8D",
}

# ─────────────────────────────────────────────
# FUNCIONES DE LIMPIEZA Y ENRIQUECIMIENTO
# ─────────────────────────────────────────────
def limpiar_texto(s):
    """
    Normaliza una cadena de texto: elimina espacios múltiples y aplica strip.
    Retorna cadena vacía si el valor es NaN.
    Útil para limpiar columnas de texto antes de comparaciones o agrupaciones.

    Parámetros
    ----------
    s : cualquier valor (str, float, NaN)

    Retorna
    -------
    str : texto normalizado o "" si era NaN
    """
    if pd.isna(s):
        return ""  # ← Si llega NaN (celda vacía en Excel), retorna vacío
    return re.sub(r'\s+', ' ', str(s)).strip()

def detectar_nivel(especialidad):
    """
    Infiere el nivel de formación SENA a partir del nombre de la especialidad.
    Orden de evaluación: TECNÓLOGO > TÉCNICO > OPERARIO > AUXILIAR > OTRO.
    Si ninguna palabra clave coincide, retorna "OTRO".

    DIAGNÓSTICO: Si muchos registros quedan como "OTRO", revisar que la
    columna "Especialidad" exista y no esté vacía o con nombres inesperados.

    Parámetros
    ----------
    especialidad : str — nombre del programa de formación

    Retorna
    -------
    str : uno de "TÉCNICO", "TECNÓLOGO", "OPERARIO", "AUXILIAR", "OTRO"
    """
    esp = limpiar_texto(especialidad).upper()
    if "TECNÓLOGO" in esp or "TECNOLOGO" in esp:
        return "TECNÓLOGO"
    if "TÉCNICO" in esp or "TECNICO" in esp:
        return "TÉCNICO"
    if "OPERARIO" in esp or "OPERARIA" in esp:
        return "OPERARIO"
    if "AUXILIAR" in esp:
        return "AUXILIAR"
    return "OTRO"

def calcular_trimestre(fecha_inicio_lectiva, fecha_ref):
    """
    Calcula en qué trimestre académico se encuentra un aprendiz.
    Divide los días transcurridos desde el inicio lectivo en bloques de ~91 días.

    DIAGNÓSTICO: Si retorna None masivamente, verificar que las columnas
    "Fecha lectiva" y "Fecha inicio contrato" existen y tienen fechas válidas.
    Un resultado None implica que el aprendiz no aparecerá en análisis
    por trimestre.

    Parámetros
    ----------
    fecha_inicio_lectiva : datetime | NaT
    fecha_ref            : datetime | NaT — fecha de referencia (contrato o hoy)

    Retorna
    -------
    int : número de trimestre (≥1) o None si no se puede calcular
    """
    if pd.isna(fecha_inicio_lectiva) or pd.isna(fecha_ref):
        return None
    try:
        delta = (pd.to_datetime(fecha_ref) - pd.to_datetime(fecha_inicio_lectiva)).days
        if delta < 0: return None
        trimestre = int(delta / 91) + 1   # ~91 días por trimestre
        return max(1, trimestre)
    except:
        return None

def detectar_etapa(nivel, trimestre):
    if nivel not in NIVEL_CONFIG or trimestre is None:
        return "DESCONOCIDA"
    cfg = NIVEL_CONFIG[nivel]
    if trimestre > cfg["trimestres"]:
        return "COMPLETADO"
    if trimestre in cfg["lectiva"]:
        return "LECTIVA"
    return "PRODUCTIVA"

def inferir_duracion_programa(nivel, f_lectiva, f_productiva):
    if pd.notna(f_lectiva) and pd.notna(f_productiva):
        meses = round((pd.to_datetime(f_productiva) - pd.to_datetime(f_lectiva)).days / 30.44)
        return int(meses)
    if nivel in NIVEL_CONFIG:
        return NIVEL_CONFIG[nivel]["meses"]
    return None

def enriquecer_df(df):
    """Aplica toda la lógica de negocio SENA al dataframe."""
    df = df.copy()

    # Normalizar columnas de texto
    text_cols = ["Estado aprendiz","Estado Contrato","Especialidad","Regional","Centro","Ciudad"]
    for c in text_cols:
        if c in df.columns:
            df[c] = df[c].apply(limpiar_texto)

    # Nivel de formación
    if "Especialidad" in df.columns:
        df["Nivel formación"] = df["Especialidad"].apply(detectar_nivel)

    # Parsear fechas
    date_cols = ["Fecha lectiva","Fecha productiva","Fecha inicio contrato",
                 "Fecha fin contrato","Fecha creación","Fecha"]
    for c in date_cols:
        if c in df.columns:
            df[c] = pd.to_datetime(df[c], errors="coerce")

    # Trimestre actual (tomando la fecha de inicio de contrato o fecha productiva)
    hoy = pd.Timestamp.today()
    ref = df.get("Fecha inicio contrato", pd.Series([None]*len(df)))
    f_lect = df.get("Fecha lectiva", pd.Series([None]*len(df)))
    df["Trimestre actual"] = [
        calcular_trimestre(fl, fc if pd.notna(fc) else hoy)
        for fl, fc in zip(f_lect, ref)
    ]

    # Etapa
    df["Etapa formación"] = [
        detectar_etapa(niv, tri)
        for niv, tri in zip(df.get("Nivel formación", [""]*len(df)),
                            df.get("Trimestre actual", [None]*len(df)))
    ]

    # Duración programa (meses)
    df["Duración programa (meses)"] = [
        inferir_duracion_programa(niv, fl, fp)
        for niv, fl, fp in zip(
            df.get("Nivel formación", [""]*len(df)),
            df.get("Fecha lectiva", [None]*len(df)),
            df.get("Fecha productiva", [None]*len(df))
        )
    ]

    # Duración contrato (días)
    if "Fecha inicio contrato" in df.columns and "Fecha fin contrato" in df.columns:
        df["Duración contrato (días)"] = (
            df["Fecha fin contrato"] - df["Fecha inicio contrato"]
        ).dt.days

    # Año del contrato
    if "Fecha inicio contrato" in df.columns:
        df["Año contrato"] = df["Fecha inicio contrato"].dt.year
        df["Mes contrato"]  = df["Fecha inicio contrato"].dt.month
        df["Mes-Año contrato"] = df["Fecha inicio contrato"].dt.to_period("M").astype(str)

    # ── Columna derivada: Tiene contrato ────────────────────────────────────
    # Un aprendiz "tiene contrato" si la columna "Estado Contrato" existe,
    # no es NaN y no es cadena vacía. Si esta columna falta en el Excel,
    # todos los aprendices quedarán sin contrato → revisar nombre de columna.
    has_contract_col = "Estado Contrato"
    if has_contract_col in df.columns:
        df["Tiene contrato"] = df[has_contract_col].notna() & (df[has_contract_col] != "")
    else:
        # DIAGNÓSTICO: La columna esperada no existe. Columnas disponibles:
        import logging
        logging.warning(
            f"[enriquecer_df] Columna '{has_contract_col}' no encontrada. "
            f"Columnas disponibles: {list(df.columns)}"
        )
        df["Tiene contrato"] = False  # Fallback seguro: nadie tiene contrato

    return df

def detectar_columnas(df):
    """Devuelve un dict con los nombres detectados de columnas clave."""
    cols = {c.lower().strip(): c for c in df.columns}
    mapping = {}
    keywords = {
        "especialidad":      ["especialidad","programa","carrera","nombre programa"],
        "estado_aprendiz":   ["estado aprendiz","estado del aprendiz","estado"],
        "fecha_lectiva":     ["fecha lectiva","inicio lectiva","inicio formacion"],
        "fecha_productiva":  ["fecha productiva","inicio productiva","inicio practica"],
        "fecha_ini_contrato":["fecha inicio contrato","inicio contrato","fecha contrato"],
        "fecha_fin_contrato":["fecha fin contrato","fin contrato"],
        "estado_contrato":   ["estado contrato","estado del contrato"],
        "regional":          ["regional"],
        "centro":            ["centro","centro de formacion"],
        "nombre":            ["nombres","nombre"],
        "apellido":          ["apellidos","apellido"],
        "documento":         ["numero documento","documento","cedula","nit"],
        "empresa":           ["razon social","empresa","razon_social"],
        "nit_empresa":       ["nit"],
        "ficha":             ["ficha","codigo ficha"],
        "ciudad":            ["ciudad","municipio"],
    }
    for key, kws in keywords.items():
        for kw in kws:
            if kw in cols:
                mapping[key] = cols[kw]
                break
    return mapping

def reporte_calidad(df):
    """
    Genera un diccionario con métricas básicas de calidad del dataset.
    Permite identificar columnas con alta tasa de nulos o valores inesperados
    antes de ejecutar el análisis principal.

    DIAGNÓSTICO: Si "pct_nulos" de una columna clave supera el 30%,
    el análisis relacionado puede ser poco representativo.
    Genera un resumen de calidad del dataset."""
    total = len(df)
    reporte = {
        "total_filas": total,
        "columnas": len(df.columns),
        "nulos_por_columna": df.isnull().sum().to_dict(),
        "duplicados": df.duplicated().sum(),
        "pct_nulos_total": round(df.isnull().sum().sum() / (total * len(df.columns)) * 100, 1),
    }
    return reporte

def obtener_hojas(archivo_bytes, nombre):
    """
    Retorna la lista de hojas de un Excel sin hacer ningún widget.
    Si es CSV o tiene una sola hoja, retorna lista vacía.
    """
    nombre = nombre.lower()
    if not nombre.endswith((".xlsx", ".xls", ".xlsm")):
        return []
    try:
        buf = io.BytesIO(archivo_bytes)
        xl = pd.ExcelFile(buf)
        return xl.sheet_names if len(xl.sheet_names) > 1 else []
    except:
        return []

def cargar_dataset(archivo_bytes, nombre, sheet_name=None):
    """
    Carga dataset desde xlsx, xls, csv con detección de encoding y limpieza.
    Recibe bytes directamente (no un objeto archivo con .name) para ser
    compatible con funciones cacheadas que no aceptan widgets internos.

    Parámetros
    ----------
    archivo_bytes : bytes  — contenido del archivo
    nombre        : str    — nombre del archivo (para detectar extensión)
    sheet_name    : str|None — hoja a leer (ya elegida fuera del caché)
    """
    nombre_lower = nombre.lower()
    try:
        if nombre_lower.endswith(".csv"):
            encodings = ["utf-8","latin-1","iso-8859-1","cp1252"]
            separators = [",",";","\t","|"]
            df = None
            for enc in encodings:
                for sep in separators:
                    try:
                        buf = io.BytesIO(archivo_bytes)
                        df = pd.read_csv(buf, encoding=enc, sep=sep, on_bad_lines="skip")
                        if len(df.columns) > 1:
                            break
                    except:
                        continue
                if df is not None and len(df.columns) > 1:
                    break
        elif nombre_lower.endswith((".xlsx",".xls",".xlsm")):
            buf = io.BytesIO(archivo_bytes)
            # sheet_name ya viene decidido desde fuera (sin widget aquí)
            df = pd.read_excel(buf, sheet_name=sheet_name or 0)
        else:
            return None, "Formato no soportado. Use .xlsx, .xls o .csv"

        # Limpiar columnas sin nombre
        df.columns = [
            str(c).strip() if not str(c).startswith("Unnamed") else f"Col_{i}"
            for i, c in enumerate(df.columns)
        ]
        # Eliminar filas completamente vacías
        df.dropna(how="all", inplace=True)
        # Detectar si la primera fila es en realidad el encabezado
        if len(df) > 0 and df.iloc[0].astype(str).str.contains("fecha|nombre|estado|codigo", case=False, na=False).any():
            df.columns = df.iloc[0].astype(str).str.strip()
            df = df.iloc[1:].reset_index(drop=True)

        return df, None
    except Exception as e:
        return None, str(e)

# ─────────────────────────────────────────────
# SIDEBAR – NAVEGACIÓN
# ─────────────────────────────────────────────
with st.sidebar:
    st.markdown("""
<style>
/* Oculta el área grande */
[data-testid="stFileUploader"] {
    border: none;
    padding: 0;
    background: transparent;
}

/* Estilo del botón */
[data-testid="stFileUploader"] button {
    background-color: #28a745 !important;
    color: white !important;
    border: none !important;
    border-radius: 8px !important;
    padding: 10px 20px !important;
    font-weight: bold !important;
}

/* Oculta textos extra */
[data-testid="stFileUploaderDropzone"] {
    padding: 0 !important;
    border: none !important;
    background: transparent !important;
}
</style>
""", unsafe_allow_html=True)

    # Mostrar logo del SENA
    logo_ruta = obtener_ruta_logo()
    if logo_ruta:
        col1, col2, col3 = st.columns([1, 1, 1])
        with col2:
            st.image(logo_ruta, width=72)
    
    st.markdown("""
    <div class="sena-logo">
        <h2>SISCON</h2>
        <p>Gestión Contratos de Aprendizaje</p>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("### 📂 Cargar Dataset")
    archivo_upload = st.file_uploader(
        "Sube tu archivo de datos",
        type=["xlsx","xls","csv","xlsm"],
        help="Formatos: Excel (.xlsx, .xls, .xlsm) o CSV"
    )

    st.markdown("---")
    st.markdown("### 📌 Menú Principal")

    paginas = {
        "🏠 Inicio / Resumen":           "inicio",
        "📊 Dashboard General":          "dashboard",
        "🎓 Análisis por Nivel":         "niveles",
        "📅 Análisis por Trimestre":     "trimestres",
        "🏢 Empresas y Contratos":       "empresas",
        "🏭 Actividad Económica":        "actividad_economica",
        "🗺️ Mapa de Colombia":          "mapa",
        "⚠️ Deserción y Riesgo":         "desercion",
        "🔍 Explorar Datos":             "explorar",
        "📁 Carga de Datasets":          "carga",
    }

    seleccion = st.radio(
        "Navegar a:",
        list(paginas.keys()),
        label_visibility="collapsed"
    )
    pagina = paginas[seleccion]

    st.markdown("---")
    st.markdown("""
    <div style="padding:10px; font-size:0.75rem; color:rgba(255,255,255,0.65); text-align:center;">
        Sistema Administrativo SENA Contratos de Aprendizaje v2.0
        <span style="color:rgba(255,255,255,0.45);"></span>
    </div>
    """, unsafe_allow_html=True)

# ─────────────────────────────────────────────
# CARGA Y CACHÉ DEL DATASET
# ─────────────────────────────────────────────
@st.cache_data(show_spinner=False)
def procesar_archivo(bytes_data, nombre, sheet_name=None):
    df_raw, err = cargar_dataset(bytes_data, nombre, sheet_name=sheet_name)
    if err:
        return None, None, err
    df_enr = enriquecer_df(df_raw)
    calidad = reporte_calidad(df_raw)
    return df_enr, calidad, None

df_global = None
calidad_global = None
col_map = {}

if archivo_upload is not None:
    bytes_data = archivo_upload.read()

    # ── Selección de hoja FUERA del caché ──────────────────────────────────
    # obtener_hojas() no usa widgets; el st.selectbox vive aquí, en el
    # flujo principal de Streamlit, donde los widgets son válidos.
    hojas = obtener_hojas(bytes_data, archivo_upload.name)
    sheet_elegida = None
    if hojas:
        sheet_elegida = st.selectbox(
            "📋 El archivo tiene varias hojas. Selecciona la hoja a analizar:",
            hojas,
            key="sheet_selector"
        )

    with st.spinner("⏳ Procesando dataset..."):
        df_global, calidad_global, error_carga = procesar_archivo(
            bytes_data, archivo_upload.name, sheet_name=sheet_elegida
        )
    if error_carga:
        st.error(f"❌ Error al cargar: {error_carga}")
    elif df_global is not None:
        col_map = detectar_columnas(df_global)
        st.session_state["df"] = df_global
        st.session_state["calidad"] = calidad_global
        st.session_state["col_map"] = col_map
        st.session_state["archivo_nombre"] = archivo_upload.name

# Recuperar de session_state si ya fue cargado
if "df" in st.session_state and df_global is None:
    df_global = st.session_state["df"]
    calidad_global = st.session_state.get("calidad", {})
    col_map = st.session_state.get("col_map", {})

# ─────────────────────────────────────────────
# HELPERS DE GRÁFICAS
# ─────────────────────────────────────────────
def fig_bar(df, x, y, color=None, title="", color_discrete=None, **kwargs):
    cd = color_discrete or {**COLOR_NIVEL, **COLOR_ETAPA}
    fig = px.bar(df, x=x, y=y, color=color, title=title,
                 color_discrete_map=cd,
                 color_discrete_sequence=PALETA, **kwargs)
    fig.update_layout(
        plot_bgcolor="white", paper_bgcolor="white",
        font_family="Inter", title_font_size=14,
        legend_title_font_size=11,
        margin=dict(t=40,b=20,l=10,r=10),
    )
    fig.update_traces(marker_line_width=0)
    return fig

def fig_pie(labels, values, title="", colors=None):
    colors = colors or PALETA
    fig = go.Figure(go.Pie(
        labels=labels, values=values,
        marker_colors=colors,
        hole=0.4,
        textinfo="percent+label",
        textfont_size=12,
    ))
    fig.update_layout(
        title=title, font_family="Inter",
        plot_bgcolor="white", paper_bgcolor="white",
        margin=dict(t=40,b=10,l=10,r=10),
        showlegend=True,
    )
    return fig

def tarjeta_metrica(col, valor, etiqueta, delta=None, color="#27AE60"):
    delta_html = f'<div class="metric-delta">{delta}</div>' if delta else ""
    col.markdown(f"""
    <div class="metric-card" style="border-left-color:{color}">
        <div class="metric-value" style="color:{color}">{valor}</div>
        <div class="metric-label">{etiqueta}</div>
        {delta_html}
    </div>
    """, unsafe_allow_html=True)

# ─────────────────────────────────────────────
# ═══════════ PÁGINAS ═══════════
# ─────────────────────────────────────────────

# ── PÁGINA: CARGA DE DATASETS ──────────────────────────────────────────────
if pagina == "carga":
    st.markdown("""
    
    <div class="section-header">
        <div>
            <h2>📁 Carga de Datasets</h2>
            <p>Importa archivos Excel o CSV para análisis</p>
        </div>
    </div>
    """, unsafe_allow_html=True)

    col1, col2 = st.columns([1,1])
    with col1:
        st.markdown("""
        <div class="alert-info">
        <strong>📌 Formatos soportados</strong><br>
        • Excel: <code>.xlsx</code>, <code>.xls</code>, <code>.xlsm</code><br>
        • CSV: <code>.csv</code> (separadores: coma, punto y coma, tabulación)<br>
        • Múltiples hojas: selección interactiva<br>
        • Encodings: UTF-8, Latin-1, ISO-8859-1, CP1252
        </div>
        """, unsafe_allow_html=True)

        st.markdown("""
        <div class="alert-warning">
        <strong>⚠️ Limpieza automática incluida</strong><br>
        • Detección de encabezados desplazados<br>
        • Eliminación de filas completamente vacías<br>
        • Normalización de espacios en blanco<br>
        • Conversión automática de tipos de fecha<br>
        • Mapeo inteligente de columnas SENA
        </div>
        """, unsafe_allow_html=True)

    with col2:
        st.markdown("""
        <div class="alert-info">
        <strong>🎓 Columnas que el sistema detecta automáticamente</strong><br>
        • Especialidad / Programa<br>
        • Fechas lectiva y productiva<br>
        • Fechas de inicio/fin de contrato<br>
        • Estado del aprendiz y del contrato<br>
        • Regional, Centro, Ciudad<br>
        • Razón social de empresa / NIT<br>
        • Ficha, Número documento
        </div>
        """, unsafe_allow_html=True)

    st.markdown("---")
    st.markdown("### 📤 Cargar Archivo")
    st.info("👈 Usa el panel lateral para subir tu archivo de datos.")

    if df_global is not None:
        st.success(f"✅ Dataset activo: **{st.session_state.get('archivo_nombre','archivo')}** — {len(df_global):,} registros · {len(df_global.columns)} columnas")

        with st.expander("🔍 Vista previa del dataset cargado", expanded=True):
            st.dataframe(df_global.head(50), use_container_width=True)

        with st.expander("📋 Reporte de Calidad del Dato"):
            if calidad_global:
                c1, c2, c3, c4 = st.columns(4)
                tarjeta_metrica(c1, f"{calidad_global['total_filas']:,}", "Filas totales")
                tarjeta_metrica(c2, calidad_global['columnas'], "Columnas")
                tarjeta_metrica(c3, calidad_global['duplicados'], "Filas duplicadas", color="#E74C3C")
                tarjeta_metrica(c4, f"{calidad_global['pct_nulos_total']}%", "% nulos global", color="#F39C12")

                st.markdown("#### 🧮 Valores nulos por columna")
                nulos = pd.DataFrame.from_dict(calidad_global["nulos_por_columna"], orient="index", columns=["Nulos"])
                nulos["% Nulo"] = (nulos["Nulos"] / calidad_global["total_filas"] * 100).round(1)
                nulos = nulos[nulos["Nulos"] > 0].sort_values("Nulos", ascending=False)
                if not nulos.empty:
                    fig_nulos = px.bar(nulos.reset_index(), x="index", y="% Nulo",
                                       title="% de valores nulos por columna",
                                       color="% Nulo",
                                       color_continuous_scale=["#7F8C8D","#F39C12","#E74C3C"])
                    fig_nulos.update_layout(xaxis_title="Columna", yaxis_title="% Nulo",
                                            plot_bgcolor="white", paper_bgcolor="white",
                                            font_family="Inter")
                    st.plotly_chart(fig_nulos, use_container_width=True)

        with st.expander("🗺️ Columnas detectadas automáticamente"):
            if col_map:
                col_df = pd.DataFrame(list(col_map.items()), columns=["Campo SENA", "Columna en Dataset"])
                st.table(col_df)
            else:
                st.warning("No se detectaron columnas estándar SENA.")

        # Descarga del dataset enriquecido
        st.markdown("#### 💾 Descargar Dataset Enriquecido")
        buffer = io.BytesIO()
        df_global.to_excel(buffer, index=False, engine="openpyxl")
        st.download_button(
            label="⬇️ Descargar Excel con columnas SENA",
            data=buffer.getvalue(),
            file_name="dataset_sena_enriquecido.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.markdown("""
        <div class="alert-warning">
        <strong>⚠️ No hay dataset cargado.</strong> Sube un archivo usando el panel lateral izquierdo.
        </div>
        """, unsafe_allow_html=True)
    st.stop()

# ── PAGINA: MODELO DE DATOS ─────────────────────────────────────────────────
    st.stop()

# ── REQUIERE DATASET PARA EL RESTO DE PÁGINAS ──────────────────────────────
if df_global is None:
    st.markdown("""
    <div class="section-header">
        <div>
            <h2>🏠 Bienvenido al Sistema SENA</h2>
            <p>Gestión Administrativa – Contratos de Aprendizaje</p>
        </div>
    </div>
    """, unsafe_allow_html=True)

    col1, col2, col3 = st.columns(3)
    with col1:
        st.markdown("""
        <div class="metric-card">
            <div class="metric-value">📊</div>
            <div class="metric-label">Dashboard interactivo por nivel, trimestre y etapa</div>
        </div>
        """, unsafe_allow_html=True)
    with col2:
        st.markdown("""
        <div class="metric-card">
            <div class="metric-value">🎓</div>
            <div class="metric-label">Análisis Técnico vs Tecnólogo, deserción y riesgo</div>
        </div>
        """, unsafe_allow_html=True)
    with col3:
        st.markdown("""
        <div class="metric-card">
            <div class="metric-value">🏢</div>
            <div class="metric-label">Top empresas contratantes y análisis de contratos</div>
        </div>
        """, unsafe_allow_html=True)

    st.markdown("""
    <div class="alert-info">
    <strong>👈 Para comenzar:</strong> Sube tu archivo de datos usando el panel lateral izquierdo.<br>
    Formatos aceptados: <code>.xlsx</code>, <code>.xls</code>, <code>.xlsm</code>, <code>.csv</code>
    </div>
    """, unsafe_allow_html=True)
    st.stop()

# Alias conveniente
df = df_global

# ─────────────────────────────────────────────
# ── FILTRO GLOBAL (afecta a todas las páginas) ─
# ─────────────────────────────────────────────
with st.sidebar:
    st.markdown("---")
    with st.expander("🔎 Filtro global", expanded=False):
        st.caption("Aplica a todas las páginas del sistema.")

        filtros_globales = {}

        if "Nivel formación" in df.columns:
            opts_niv = sorted(df["Nivel formación"].dropna().unique().tolist())
            sel_niv = st.multiselect("Nivel de formación", opts_niv, default=[], key="gf_nivel")
            if sel_niv:
                filtros_globales["Nivel formación"] = sel_niv

        if "Regional" in df.columns:
            opts_reg = sorted(df["Regional"].dropna().unique().tolist())
            sel_reg = st.multiselect("Regional", opts_reg, default=[], key="gf_regional")
            if sel_reg:
                filtros_globales["Regional"] = sel_reg

        if "Centro" in df.columns:
            opts_centro = sorted(df["Centro"].dropna().unique().tolist())
            sel_centro = st.multiselect("Centro", opts_centro, default=[], key="gf_centro")
            if sel_centro:
                filtros_globales["Centro"] = sel_centro

        if "Ciudad" in df.columns:
            opts_ciudad = sorted(df["Ciudad"].dropna().unique().tolist())
            sel_ciudad = st.multiselect("Ciudad", opts_ciudad, default=[], key="gf_ciudad")
            if sel_ciudad:
                filtros_globales["Ciudad"] = sel_ciudad

        if "Etapa formación" in df.columns:
            opts_etapa = [e for e in ["LECTIVA", "PRODUCTIVA"] if e in df["Etapa formación"].unique()]
            sel_etapa = st.multiselect("Etapa de formación", opts_etapa, default=[], key="gf_etapa")
            if sel_etapa:
                filtros_globales["Etapa formación"] = sel_etapa

        if "Tiene contrato" in df.columns:
            sel_contrato = st.radio(
                "Estado de contrato",
                ["Todos", "Con contrato", "Sin contrato"],
                index=0, key="gf_contrato",
            )
        else:
            sel_contrato = "Todos"

        if st.button("🔄 Limpiar filtros", use_container_width=True, key="gf_clear"):
            for k in ["gf_nivel", "gf_regional", "gf_centro", "gf_ciudad", "gf_etapa", "gf_contrato"]:
                if k in st.session_state:
                    del st.session_state[k]
            st.rerun()

# Aplicar filtros globales sobre df (afecta a TODAS las páginas)
for col_f, valores_f in filtros_globales.items():
    df = df[df[col_f].isin(valores_f)]

if sel_contrato == "Con contrato" and "Tiene contrato" in df.columns:
    df = df[df["Tiene contrato"] == True]
elif sel_contrato == "Sin contrato" and "Tiene contrato" in df.columns:
    df = df[df["Tiene contrato"] == False]

# Aviso visible si hay filtros activos y el resultado quedó vacío o reducido
_filtros_activos = bool(filtros_globales) or (sel_contrato != "Todos")
if _filtros_activos:
    st.sidebar.info(f"🔎 Filtro activo: **{len(df):,}** de **{len(df_global):,}** registros")
    if df.empty:
        st.warning("⚠️ El filtro global no arrojó resultados. Ajusta los criterios en el panel lateral.")
        st.stop()

# ─────────────────────────────────────────────
# ── PÁGINA: INICIO / RESUMEN ─────────────────
# ─────────────────────────────────────────────
if pagina == "inicio":
    st.markdown("""
    <div class="section-header">
        <div>
            <h2>🏠 Resumen Ejecutivo</h2>
            <p>Indicadores clave del sistema de contratos de aprendizaje</p>
        </div>
    </div>
    """, unsafe_allow_html=True)

    hoy = pd.Timestamp.today().normalize()
    total = len(df)
    contratados = int(df.get("Tiene contrato", pd.Series([False]*total)).sum())
    sin_contrato = total - contratados
    tasa_contratacion = round(contratados / total * 100, 1) if total > 0 else 0

    niveles = df["Nivel formación"].value_counts() if "Nivel formación" in df.columns else pd.Series()
    tecnicos = int(niveles.get("TÉCNICO", 0))
    tecnologos = int(niveles.get("TECNÓLOGO", 0))

    # ── FILA 1: KPIs principales ────────────────────────────────────────
    c1, c2, c3, c4, c5 = st.columns(5)
    tarjeta_metrica(c1, f"{total:,}", "Total Aprendices", f"Dataset: {st.session_state.get('archivo_nombre','')}")
    tarjeta_metrica(c2, f"{contratados:,}", "Con Contrato", f"Tasa: {tasa_contratacion}%")
    tarjeta_metrica(c3, f"{sin_contrato:,}", "Sin Contrato", f"{round(sin_contrato/total*100,1) if total else 0}% del total", color="#E74C3C")
    tarjeta_metrica(c4, f"{tecnicos:,}", "Técnicos", "15 meses · 5 trimestres")
    tarjeta_metrica(c5, f"{tecnologos:,}", "Tecnólogos", "24 meses · 8 trimestres", color="#2980B9")

    st.markdown("---")

    # ── FILA 2: Gauge de tasa + Panel de alertas ────────────────────────
    col_gauge, col_alertas = st.columns([1, 2])

    df_fc = pd.DataFrame()
    v30 = v60 = disp = 0
    if "Fecha fin contrato" in df.columns:
        df_fc = df[df["Fecha fin contrato"].notna()].copy()
        df_fc["Días restantes"] = (df_fc["Fecha fin contrato"] - hoy).dt.days
        v30  = int(df_fc["Días restantes"].between(0, 30).sum())
        v60  = int(df_fc["Días restantes"].between(31, 60).sum())
        disp = int((df_fc["Días restantes"] < 0).sum())

    with col_gauge:
        color_aguja = "#27AE60" if tasa_contratacion >= 75 else "#F39C12" if tasa_contratacion >= 50 else "#E74C3C"
        fig_gauge = go.Figure(go.Indicator(
            mode="gauge+number+delta",
            value=tasa_contratacion,
            delta={"reference": 75, "suffix": "%",
                   "increasing": {"color": "#27AE60"}, "decreasing": {"color": "#E74C3C"}},
            number={"suffix": "%", "font": {"size": 34, "color": color_aguja}},
            title={"text": "Tasa de Contratación", "font": {"size": 14}},
            gauge={
                "axis": {"range": [0, 100], "tickwidth": 1},
                "bar":  {"color": color_aguja, "thickness": 0.3},
                "steps": [
                    {"range": [0,  50], "color": "#FADBD8"},
                    {"range": [50, 75], "color": "#FEF9E7"},
                    {"range": [75, 100], "color": "#EAFAF1"},
                ],
                "threshold": {"line": {"color": "#2C3E50", "width": 3}, "value": 75},
            },
        ))
        fig_gauge.update_layout(height=260, margin=dict(t=40, b=10, l=20, r=20),
                                font_family="Inter", paper_bgcolor="white")
        st.plotly_chart(fig_gauge, use_container_width=True)
        st.caption("🎯 Meta referencial: 75% — línea negra en el medidor")

    with col_alertas:
        st.markdown("#### 🚨 Panel de Alertas")
        alerta_alguna = False

        if v30 > 0:
            alerta_alguna = True
            st.markdown(f"""<div class="alert-error">🔴 <strong>{v30} contratos</strong> vencen en los próximos 30 días — atención urgente</div>""", unsafe_allow_html=True)
        if v60 > 0:
            alerta_alguna = True
            st.markdown(f"""<div class="alert-warning">🟡 <strong>{v60} contratos</strong> vencen entre 31 y 60 días</div>""", unsafe_allow_html=True)
        if disp > 0:
            alerta_alguna = True
            st.markdown(f"""<div class="alert-info">🔵 <strong>{disp} aprendices</strong> disponibles (contrato ya finalizado)</div>""", unsafe_allow_html=True)
        if sin_contrato > 0:
            alerta_alguna = True
            st.markdown(f"""<div class="alert-warning">⚠️ <strong>{sin_contrato:,} aprendices</strong> aún sin contrato ({round(sin_contrato/total*100,1) if total else 0}% del total)</div>""", unsafe_allow_html=True)
        if calidad_global and calidad_global.get("duplicados", 0) > 0:
            alerta_alguna = True
            st.markdown(f"""<div class="alert-warning">⚠️ <strong>{calidad_global['duplicados']} filas duplicadas</strong> detectadas en el dataset</div>""", unsafe_allow_html=True)
        if calidad_global and calidad_global.get("pct_nulos_total", 0) > 20:
            alerta_alguna = True
            st.markdown(f"""<div class="alert-error">🔴 El dataset tiene un <strong>{calidad_global['pct_nulos_total']}%</strong> de valores nulos — algunos análisis pueden no ser representativos</div>""", unsafe_allow_html=True)

        if not alerta_alguna:
            st.markdown("""<div class="alert-info">✅ No hay alertas activas en este momento.</div>""", unsafe_allow_html=True)

    st.markdown("---")

    # ── FILA 3: Tendencia mensual + Tasa por nivel ──────────────────────
    col_tend, col_niv = st.columns([3, 2])

    with col_tend:
        if "Mes-Año contrato" in df.columns and "Tiene contrato" in df.columns:
            tend = (df[df["Tiene contrato"]]
                    .groupby("Mes-Año contrato").size()
                    .reset_index(name="Contratos")
                    .sort_values("Mes-Año contrato")
                    .tail(18))
            if not tend.empty:
                fig_tend = px.area(
                    tend, x="Mes-Año contrato", y="Contratos",
                    title="📈 Contratos firmados — últimos 18 meses",
                    color_discrete_sequence=["#27AE60"],
                )
                fig_tend.update_traces(line_color="#27AE60", fillcolor="rgba(39,174,96,0.15)",
                                       mode="lines+markers", line_width=2)
                fig_tend.update_layout(
                    plot_bgcolor="white", paper_bgcolor="white", font_family="Inter",
                    xaxis_tickangle=-40, xaxis_title="", yaxis_title="Contratos",
                    margin=dict(t=50, b=40, l=10, r=10), height=300,
                )
                st.plotly_chart(fig_tend, use_container_width=True)
            else:
                st.info("No hay suficientes datos de fecha para mostrar la tendencia mensual.")
        else:
            st.info("No se encontró la columna 'Mes-Año contrato' para mostrar la tendencia.")

    with col_niv:
        if "Nivel formación" in df.columns and "Tiene contrato" in df.columns:
            tasa_niv = (df.groupby("Nivel formación")
                        .agg(total=("Nivel formación", "count"),
                             contratados=("Tiene contrato", "sum"))
                        .assign(tasa=lambda x: (x["contratados"]/x["total"]*100).round(1))
                        .reset_index()
                        .sort_values("tasa"))

            fig_niv = go.Figure()
            for _, row in tasa_niv.iterrows():
                color = "#27AE60" if row["tasa"] >= 75 else "#F39C12" if row["tasa"] >= 50 else "#E74C3C"
                fig_niv.add_trace(go.Bar(
                    x=[row["tasa"]], y=[row["Nivel formación"]],
                    orientation="h", marker_color=color,
                    text=f"{row['tasa']}%", textposition="outside",
                    showlegend=False,
                    hovertemplate=f"<b>{row['Nivel formación']}</b><br>Tasa: {row['tasa']}%<br>Total: {int(row['total']):,}<extra></extra>",
                ))
            fig_niv.add_vline(x=75, line_dash="dash", line_color="#2C3E50",
                              annotation_text="Meta 75%", annotation_position="top right")
            fig_niv.update_layout(
                title="🎯 Tasa de contratación por nivel",
                plot_bgcolor="white", paper_bgcolor="white", font_family="Inter",
                xaxis=dict(range=[0, 110], title="Tasa (%)"), yaxis_title="",
                height=300, margin=dict(t=50, b=10, l=10, r=60),
            )
            st.plotly_chart(fig_niv, use_container_width=True)

    # ── FILA 4: Tablas expandibles de seguimiento ───────────────────────
    if not df_fc.empty:
        cols_urg = [c for c in ["Nombres", "Apellidos", "Especialidad", "Razón social",
                                "Empresa", "Fecha fin contrato", "Días restantes"] if c in df_fc.columns]

        with st.expander("📋 Ver contratos por vencer en los próximos 60 días", expanded=False):
            df_urgentes = df_fc[df_fc["Días restantes"].between(0, 60)].sort_values("Días restantes")
            if not df_urgentes.empty:
                st.dataframe(df_urgentes[cols_urg], use_container_width=True, hide_index=True)
            else:
                st.success("✅ No hay contratos por vencer en los próximos 60 días.")

        with st.expander("🔵 Ver aprendices disponibles (contrato finalizado)", expanded=False):
            df_disp = df_fc[df_fc["Días restantes"] < 0].sort_values("Días restantes", ascending=False)
            if not df_disp.empty:
                st.dataframe(df_disp[cols_urg], use_container_width=True, hide_index=True)
            else:
                st.info("No hay aprendices con contrato finalizado en el dataset.")
        # ─────────────────────────────────────────────


# ─────────────────────────────────────────────
# ── PÁGINA: DASHBOARD GENERAL ────────────────
# ─────────────────────────────────────────────
elif pagina == "dashboard":
    st.markdown("""
    <div class="section-header">
        <div>
            <h2>📊 Dashboard General</h2>
            <p>Visión completa del estado de contratos</p>
        </div>
    </div>
    """, unsafe_allow_html=True)

    # Filtros globales
    with st.expander("🔧 Filtros", expanded=True):
        fc1, fc2 = st.columns(2)
        niveles_disp = ["Todos"] + sorted(df["Nivel formación"].dropna().unique().tolist()) if "Nivel formación" in df.columns else ["Todos"]
        filtro_nivel = fc1.selectbox("Nivel de formación", niveles_disp)

        años_disp = ["Todos"] + sorted([str(int(x)) for x in df["Año contrato"].dropna().unique()]) if "Año contrato" in df.columns else ["Todos"]
        filtro_año = fc2.selectbox("Año contrato", años_disp)

    dff = df.copy()
    if filtro_nivel != "Todos" and "Nivel formación" in dff.columns:
        dff = dff[dff["Nivel formación"] == filtro_nivel]
    if filtro_año != "Todos" and "Año contrato" in df.columns:
        mascara_año = dff["Año contrato"].apply(lambda x: str(int(x)) if pd.notna(x) else "") == filtro_año
        mascara_sin_contrato = dff["Año contrato"].isna()
        dff = dff[mascara_año | mascara_sin_contrato]

    total_f = len(dff)
    contratados_f = int(dff.get("Tiene contrato", pd.Series([False]*total_f)).sum())
    st.info(f"Mostrando **{total_f:,}** registros tras aplicar filtros.")

    c1, c2, c3, c4 = st.columns(4)
    tarjeta_metrica(c1, f"{total_f:,}", "Aprendices filtrados")
    tarjeta_metrica(c2, f"{contratados_f:,}", "Con contrato")
    tarjeta_metrica(c3, f"{round(contratados_f/total_f*100,1) if total_f>0 else 0}%", "Tasa de contratación")
    if "Duración contrato (días)" in dff.columns:
        tarjeta_metrica(c4, f"{int(dff['Duración contrato (días)'].median()):,} días", "Mediana duración contrato")

    st.markdown("---")
    col1, col2 = st.columns(2)

    # Estado del contrato — gráfico horizontal interactivo
    with col1:
        if "Estado Contrato" in dff.columns:
            vc_ec = dff["Estado Contrato"].value_counts().reset_index()
            vc_ec.columns = ["Estado Contrato", "count"]
            vc_ec = vc_ec.head(10)

            # Paleta de colores por estado
            color_estado = {
                "TERMINADO": "#27AE60",
                "VIGENTE": "#2980B9",
                "CADENA DE FORMACION": "#F39C12",
                "TERMINADO POR JUSTA CAUSA (FALTA DEL APRENDIZ)": "#E74C3C",
                "TERMINADO POR PRACTICA O ACUERDO": "#8E44AD",
                "COMPENSADO TERMINADO": "#16A085",
                "COMPENSADO TERMINADO POR JUSTA CAUSA (FALTA DEL APRENDIZ)": "#c0392b",
            }
            vc_ec["Color"] = vc_ec["Estado Contrato"].map(
                lambda x: color_estado.get(x, "#7F8C8D")
            )

            fig_ec = go.Figure(go.Bar(
                x=vc_ec["count"],
                y=vc_ec["Estado Contrato"],
                orientation="h",
                marker_color=vc_ec["Color"],
                text=vc_ec["count"],
                textposition="outside",
                hovertemplate="<b>%{y}</b><br>Contratos: %{x}<extra></extra>",
            ))
            fig_ec.update_layout(
                title="Contratos por Estado",
                plot_bgcolor="white",
                paper_bgcolor="white",
                font_family="Inter",
                xaxis_title="Cantidad de Contratos",
                yaxis_title="",
                yaxis=dict(autorange="reversed", tickfont=dict(size=11)),
                height=max(350, len(vc_ec) * 42),
                margin=dict(l=10, r=80, t=50, b=10),
            )
            st.plotly_chart(fig_ec, use_container_width=True)

    # Estado del aprendiz
      # Estado del aprendiz
    with col2:
        if "Estado aprendiz" in dff.columns:
            vc_ea = dff["Estado aprendiz"].value_counts().head(20).reset_index()
            vc_ea.columns = ["Estado aprendiz", "count"]
            fig_ea = px.bar(
                vc_ea,
                x="count",
                y="Estado aprendiz",
                orientation="h",
                title="Aprendices por Estado",
                labels={"count": "Cantidad de Aprendices", "Estado aprendiz": "Estado aprendiz"},
                color_discrete_sequence=["#1F77B4"],
            )
            fig_ea.update_layout(
                yaxis=dict(autorange="reversed"),
                xaxis_title="Cantidad de Aprendices",
                yaxis_title="Estado aprendiz",
            )
            st.plotly_chart(fig_ea, use_container_width=True)

    # Evolución temporal
    if "Mes-Año contrato" in dff.columns:
         vc_tiempo = (dff[dff["Tiene contrato"]]
                     .groupby("Mes-Año contrato").size()
                     .reset_index(name="Contratos")
                     .sort_values("Mes-Año contrato"))
         vc_tiempo = vc_tiempo[vc_tiempo["Mes-Año contrato"] >= "2022-01"]
         if not vc_tiempo.empty:
             fig_line = px.line(vc_tiempo, x="Mes-Año contrato", y="Contratos",
                               title="📈 Evolución mensual de contratos",
                               color_discrete_sequence=["#0CE814"])
             fig_line.update_traces(line_width=2.5, mode="lines+markers")
             fig_line.update_layout(plot_bgcolor="white", paper_bgcolor="white",
                                   font_family="Inter", xaxis_tickangle=-45)
             st.plotly_chart(fig_line, use_container_width=True)

    # Contratos por regional


    # ── GRÁFICO: Programas con más y menos contratados ──────────────────────
    # Se agrupa por Especialidad contando solo aprendices con contrato.
    # Se toman los 10 programas TOP y los 10 con menos contratos para
    # mostrarlos en dos barras horizontales con colores diferenciados.
    if "Especialidad" in dff.columns and "Tiene contrato" in dff.columns:
        df_con = dff[dff["Tiene contrato"]]  # Solo filas con contrato activo

        # Conteo de contratos por programa
        prog_cnt = (
            df_con.groupby("Especialidad")
            .size()
            .reset_index(name="Contratos")
            .sort_values("Contratos", ascending=False)
        )

        if len(prog_cnt) >= 2:
            top_n    = min(10, len(prog_cnt))
            top_prog = prog_cnt.head(top_n).copy()
            bot_prog = prog_cnt.tail(top_n).copy()

            top_prog["Categoría"] = "🔝 Mayor contratación"
            bot_prog["Categoría"] = "🔻 Menor contratación"

            # Truncar nombres largos para que el gráfico sea legible
            def truncar(nombre, n=45):
                return nombre if len(nombre) <= n else nombre[:n] + "…"

            top_prog["Programa"] = top_prog["Especialidad"].apply(truncar)
            bot_prog["Programa"] = bot_prog["Especialidad"].apply(truncar)

            col_prog1, col_prog2 = st.columns(2)

            with col_prog1:
                fig_top = px.bar(
                    top_prog.sort_values("Contratos"),
                    x="Contratos", y="Programa",
                    orientation="h",
                    title=f"🔝 Top {top_n} programas — Mayor contratación",
                    color_discrete_sequence=["#7F8C8D"],
                    text="Contratos",
                )
                fig_top.update_traces(textposition="outside")
                fig_top.update_layout(
                    plot_bgcolor="white", paper_bgcolor="white",
                    font_family="Inter", yaxis_title="", xaxis_title="Contratos",
                    margin=dict(l=10, r=30, t=50, b=10),
                )
                st.plotly_chart(fig_top, use_container_width=True)

            with col_prog2:
                fig_bot = px.bar(
                    bot_prog.sort_values("Contratos", ascending=False),
                    x="Contratos", y="Programa",
                    orientation="h",
                    title=f"🔻 Top {top_n} programas — Menor contratación",
                    color_discrete_sequence=["#F39C12"],
                    text="Contratos",
                )
                fig_bot.update_traces(textposition="outside")
                fig_bot.update_layout(
                    plot_bgcolor="white", paper_bgcolor="white",
                    font_family="Inter", yaxis_title="", xaxis_title="Contratos",
                    margin=dict(l=10, r=30, t=50, b=10),
                )
                st.plotly_chart(fig_bot, use_container_width=True)


# ─────────────────────────────────────────────
# ── PÁGINA: ANÁLISIS POR NIVEL ───────────────
# ─────────────────────────────────────────────
elif pagina == "niveles":
    st.markdown("""
    <div class="section-header">
        <div>
            <h2>🎓 Análisis por Nivel de Formación</h2>
            <p>Técnico · Tecnólogo · Operario · Auxiliar</p>
        </div>
    </div>
    """, unsafe_allow_html=True)

    if "Nivel formación" not in df.columns:
        st.error("No se detectó columna de Especialidad en el dataset.")
        st.stop()

    tabs = st.tabs(["📊 Comparativo","🔵 Técnico","🟢 Tecnólogo","🟡 Operario / Auxiliar","📋 Tabla Detalle"])

    # ─ Tab 0: Comparativo ─
    with tabs[0]:
        col1, col2 = st.columns(2)

        with col1:
            # Distribución por nivel
            vc_niv = df["Nivel formación"].value_counts().reset_index()
            vc_niv.columns = ["Nivel","Cantidad"]
            fig_niv = px.bar(vc_niv, x="Nivel", y="Cantidad",
                             color="Nivel",
                             color_discrete_sequence=PALETA,
                             title="Aprendices por Nivel",
                             text="Cantidad")
            fig_niv.update_layout(plot_bgcolor="white", paper_bgcolor="white",
                                  font_family="Inter", showlegend=False)
            st.plotly_chart(fig_niv, use_container_width=True)

        with col2:
            # Tasa de contratación por nivel
            if "Tiene contrato" in df.columns:
                tasa_niv = (df.groupby("Nivel formación")
                              .agg(total=("Nivel formación","count"),
                                   contratados=("Tiene contrato","sum"))
                              .assign(tasa=lambda x: (x["contratados"]/x["total"]*100).round(1))
                              .reset_index())
                fig_tasa = px.bar(tasa_niv, x="Nivel formación", y="tasa",
                                  color="Nivel formación",
                                  color_discrete_sequence=PALETA,
                                  title="Tasa de Contratación por Nivel (%)",
                                  text="tasa")
                fig_tasa.update_traces(texttemplate="%{text}%")
                fig_tasa.update_layout(plot_bgcolor="white", paper_bgcolor="white",
                                       font_family="Inter", showlegend=False)
                st.plotly_chart(fig_tasa, use_container_width=True)

        # Distribución por etapa
        if "Etapa formación" in df.columns:
            pivot_etapa = (df[df["Nivel formación"].isin(["TÉCNICO","TECNÓLOGO","OPERARIO","AUXILIAR"])]
                           .groupby(["Nivel formación","Etapa formación"])
                           .size()
                           .reset_index(name="Cantidad"))
            fig_etapa = px.bar(pivot_etapa, x="Nivel formación", y="Cantidad",
                               color="Etapa formación",
                               barmode="group",
                               title="Distribución Lectiva vs Productiva por Nivel",
                               color_discrete_map={
                                   "LECTIVA": "#7F8C8D",
                                   "PRODUCTIVA": "#F39C12",
                                   "COMPLETADO": "#2980B9",
                               })
            fig_etapa.update_layout(plot_bgcolor="white", paper_bgcolor="white",
                                    font_family="Inter")
            st.plotly_chart(fig_etapa, use_container_width=True)

        # Top especialidades
        st.markdown("#### 🏆 Top 15 Especialidades más frecuentes")
        top_esp = df["Especialidad"].value_counts().head(15).reset_index()
        top_esp.columns = ["Especialidad","Aprendices"]
        top_esp["Especialidad"] = top_esp["Especialidad"].str[:60]
        fig_esp = px.bar(top_esp, y="Especialidad", x="Aprendices", orientation="h",
                         color="Aprendices", color_continuous_scale=["#ECF0F1","#7F8C8D"],
                         title="")
        fig_esp.update_layout(plot_bgcolor="white", paper_bgcolor="white",
                              font_family="Inter", height=450,
                              yaxis={"categoryorder":"total ascending"})
        st.plotly_chart(fig_esp, use_container_width=True)

    # ─ Tab 1: Técnico ─
    with tabs[1]:
        st.markdown("### 📐 Estructura Técnico")
        st.markdown("""
        | | |
        |---|---|
        | **Duración total** | 15 meses |
        | **Trimestres** | 5 |
        | **Etapa Lectiva** | Trimestres 1, 2 y 3 |
        | **Etapa Productiva** | Trimestres 4 y 5 |
        """)
        df_tec = df[df["Nivel formación"] == "TÉCNICO"]
        if df_tec.empty:
            st.warning("No hay datos de aprendices Técnicos en el dataset.")
        else:
            c1, c2, c3 = st.columns(3)
            tarjeta_metrica(c1, f"{len(df_tec):,}", "Aprendices Técnicos")
            tasa_t = round(int(df_tec.get("Tiene contrato", pd.Series([False]*len(df_tec))).sum()) / len(df_tec) * 100, 1)
            tarjeta_metrica(c2, f"{tasa_t}%", "Tasa de contratación")
            if "Duración contrato (días)" in df_tec.columns:
                tarjeta_metrica(c3, f"{int(df_tec['Duración contrato (días)'].mean()):,} días", "Duración media contrato")

            if "Trimestre actual" in df_tec.columns:
                vc_trim = df_tec["Trimestre actual"].value_counts().sort_index()
                colores_trim = []
                for t in vc_trim.index:
                    if t in [1,2,3]: colores_trim.append("#E74C3C")
                    elif t in [4,5]: colores_trim.append("#F39C12")
                    else: colores_trim.append("#E74C3C")

                fig_trim_tec = go.Figure(go.Bar(
                    x=[f"T{t}" for t in vc_trim.index],
                    y=vc_trim.values,
                    marker_color=colores_trim,
                    text=vc_trim.values,
                    textposition="outside"
                ))
                fig_trim_tec.update_layout(
                    title="Aprendices Técnicos por Trimestre (🔴=Lectiva · 🟡=Productiva)",
                    plot_bgcolor="white", paper_bgcolor="white",
                    font_family="Inter", xaxis_title="Trimestre", yaxis_title="Cantidad"
                )
                st.plotly_chart(fig_trim_tec, use_container_width=True)

    # ─ Tab 2: Tecnólogo ─
    with tabs[2]:
        st.markdown("### 📐 Estructura Tecnólogo")
        st.markdown("""
        | | |
        |---|---|
        | **Duración total** | 24 meses |
        | **Trimestres** | 8 |
        | **Etapa Lectiva** | Trimestres 1 al 6 |
        | **Etapa Productiva** | Trimestres 7 y 8 |
        """)
        df_tec2 = df[df["Nivel formación"] == "TECNÓLOGO"]
        if df_tec2.empty:
            st.warning("No hay datos de aprendices Tecnólogos en el dataset.")
        else:
            c1, c2, c3 = st.columns(3)
            tarjeta_metrica(c1, f"{len(df_tec2):,}", "Aprendices Tecnólogos", color="#2980B9")
            tasa_t2 = round(int(df_tec2.get("Tiene contrato", pd.Series([False]*len(df_tec2))).sum()) / len(df_tec2) * 100, 1)
            tarjeta_metrica(c2, f"{tasa_t2}%", "Tasa de contratación", color="#2980B9")
            if "Duración contrato (días)" in df_tec2.columns:
                tarjeta_metrica(c3, f"{int(df_tec2['Duración contrato (días)'].mean()):,} días", "Duración media contrato", color="#2980B9")

            if "Trimestre actual" in df_tec2.columns:
                vc_trim2 = df_tec2["Trimestre actual"].value_counts().sort_index()
                colores_trim2 = ["#2980B9" if t<=6 else "#F39C12" for t in vc_trim2.index]
                fig_trim_tec2 = go.Figure(go.Bar(
                    x=[f"T{t}" for t in vc_trim2.index],
                    y=vc_trim2.values,
                    marker_color=colores_trim2,
                    text=vc_trim2.values,
                    textposition="outside"
                ))
                fig_trim_tec2.update_layout(
                    title="Aprendices Tecnólogos por Trimestre (🟢=Lectiva T1-T6 · 🟡=Productiva T7-T8)",
                    plot_bgcolor="white", paper_bgcolor="white",
                    font_family="Inter", xaxis_title="Trimestre", yaxis_title="Cantidad"
                )
                st.plotly_chart(fig_trim_tec2, use_container_width=True)

    # ─ Tab 3: Operario / Auxiliar ─
    with tabs[3]:
        col1, col2 = st.columns(2)
        with col1:
            st.markdown("### 🔧 Operario")
            st.markdown("""
            | | |
            |---|---|
            | **Duración aprox** | 6 meses |
            | **Trimestres** | 2 |
            | **Lectiva** | Trimestre 1 |
            | **Productiva** | Trimestre 2 |
            """)
            df_op = df[df["Nivel formación"] == "OPERARIO"]
            tarjeta_metrica(st, f"{len(df_op):,}", "Aprendices Operarios", color="#F39C12")

        with col2:
            st.markdown("### 🛠️ Auxiliar")
            st.markdown("""
            | | |
            |---|---|
            | **Duración aprox** | 12 meses |
            | **Trimestres** | 4 |
            | **Lectiva** | Trimestres 1, 2, 3 |
            | **Productiva** | Trimestre 4 |
            """)
            df_aux = df[df["Nivel formación"] == "AUXILIAR"]
            tarjeta_metrica(st, f"{len(df_aux):,}", "Aprendices Auxiliares", color="#2980B9")

        if len(df_op) > 0 or len(df_aux) > 0:
            combined = pd.concat([
                df_op.assign(Nivel="OPERARIO"),
                df_aux.assign(Nivel="AUXILIAR")
            ])
            if "Estado aprendiz" in combined.columns and not combined.empty:
                vc_comb = combined.groupby(["Nivel","Estado aprendiz"]).size().reset_index(name="Cantidad")
                fig_comb = px.bar(vc_comb, x="Estado aprendiz", y="Cantidad", color="Nivel",
                                  barmode="group", title="Estado Aprendiz: Operario vs Auxiliar",
                                  color_discrete_sequence=["#F39C12", "#2980B9"])
                fig_comb.update_layout(plot_bgcolor="white", paper_bgcolor="white",
                                       font_family="Inter", xaxis_tickangle=-30)
                st.plotly_chart(fig_comb, use_container_width=True)

    # ─ Tab 4: Tabla ─
    with tabs[3 if len(tabs)==4 else 4]:
        st.markdown("### 📋 Resumen estadístico por nivel")
        if "Nivel formación" in df.columns and "Tiene contrato" in df.columns:
            resumen = (df.groupby("Nivel formación")
                         .agg(
                             Total=("Nivel formación","count"),
                             Contratados=("Tiene contrato","sum"),
                         )
                         .assign(
                             Sin_contrato=lambda x: x["Total"] - x["Contratados"],
                             Tasa_contratacion=lambda x: (x["Contratados"]/x["Total"]*100).round(1)
                         )
                         .reset_index())
            resumen.columns = ["Nivel","Total","Contratados","Sin Contrato","Tasa (%)"]
            st.dataframe(resumen, use_container_width=True, hide_index=True)

elif pagina == "trimestres":
    st.markdown("""
    <div class="section-header">
        <div>
            <h2>📅 Análisis por Trimestre y Etapa de Formación</h2>
            <p>¿En qué trimestre y etapa se generan más contratos?</p>
        </div>
    </div>
    """, unsafe_allow_html=True)

    if "Trimestre actual" not in df.columns:
        st.error("No se pudo calcular el trimestre. Verifica que el dataset tenga 'Fecha lectiva'.")
        st.stop()

    # ── FILTROS INTERACTIVOS ──────────────────────────────────────────────
    with st.expander("🔧 Filtros", expanded=True):
        fc1, fc2, fc3 = st.columns(3)

        niveles_disp = ["Todos"] + sorted(df["Nivel formación"].dropna().unique().tolist()) \
            if "Nivel formación" in df.columns else ["Todos"]
        filtro_niv_tr = fc1.selectbox("Nivel de formación", niveles_disp, key="tr_nivel")

        solo_contratados = fc2.checkbox("Solo aprendices CON contrato", value=True, key="tr_contrato")

        max_trim = st.session_state.get("tr_max", 10)
        filtro_max_trim = fc3.slider("Máx. trimestres a mostrar", 4, 15, 10, key="tr_max")

    # Aplicar filtros
    dft = df.copy()
    if filtro_niv_tr != "Todos" and "Nivel formación" in dft.columns:
        dft = dft[dft["Nivel formación"] == filtro_niv_tr]
    if solo_contratados and "Tiene contrato" in dft.columns:
        dft = dft[dft["Tiene contrato"] == True]

    total_fil = len(dft)
    st.info(f"Mostrando **{total_fil:,}** registros con los filtros actuales.")

    # ── MÉTRICAS RÁPIDAS ─────────────────────────────────────────────────
    if "Trimestre actual" in dft.columns and not dft.empty:
        trim_counts = dft["Trimestre actual"].dropna()
        trim_max_val = int(trim_counts.mode()[0]) if not trim_counts.empty else 0
        lectiva_cnt  = int((dft.get("Etapa formación", pd.Series()) == "LECTIVA").sum())
        prod_cnt     = int((dft.get("Etapa formación", pd.Series()) == "PRODUCTIVA").sum())

        cm1, cm2, cm3, cm4 = st.columns(4)
        tarjeta_metrica(cm1, f"T{trim_max_val}", "Trimestre más frecuente")
        tarjeta_metrica(cm2, f"{lectiva_cnt:,}", "En Etapa Lectiva", color="#E74C3C")
        tarjeta_metrica(cm3, f"{prod_cnt:,}", "En Etapa Productiva", color="#F39C12")
        tasa_prod = round(prod_cnt / total_fil * 100, 1) if total_fil > 0 else 0
        tarjeta_metrica(cm4, f"{tasa_prod}%", "% en Etapa Productiva", color="#F39C12")

    st.markdown("---")

    # ── GRÁFICO 1: Barras por trimestre (coloreadas por etapa) ───────────
    col1, col2 = st.columns([3, 2])

    with col1:
        vc_trim = (dft["Trimestre actual"]
                   .dropna()
                   .astype(int)
                   .value_counts()
                   .reset_index())
        vc_trim.columns = ["Trimestre", "Cantidad"]
        vc_trim = vc_trim[vc_trim["Trimestre"] <= filtro_max_trim].sort_values("Trimestre")

        if not vc_trim.empty:
            # Determinar etapa por nivel seleccionado para colorear barras
            def etapa_de_trimestre(t, nivel):
                if nivel in NIVEL_CONFIG:
                    cfg = NIVEL_CONFIG[nivel]
                    if t in cfg["lectiva"]:    return "LECTIVA"
                    if t in cfg["productiva"]: return "PRODUCTIVA"
                return "LECTIVA" if t <= 6 else "PRODUCTIVA"

            if filtro_niv_tr != "Todos":
                vc_trim["Etapa"] = vc_trim["Trimestre"].apply(
                    lambda t: etapa_de_trimestre(t, filtro_niv_tr)
                )
            else:
                # Sin filtro: colorear T1-T6 como lectiva, T7+ productiva
                vc_trim["Etapa"] = vc_trim["Trimestre"].apply(
                    lambda t: "LECTIVA" if t <= 6 else "PRODUCTIVA"
                )

            trim_max_bar = vc_trim.loc[vc_trim["Cantidad"].idxmax(), "Trimestre"]

            fig_tr = px.bar(
                vc_trim,
                x=vc_trim["Trimestre"].apply(lambda t: f"T{t}"),
                y="Cantidad",
                color="Etapa",
                color_discrete_map={"LECTIVA": "#E74C3C", "PRODUCTIVA": "#F39C12"},
                text="Cantidad",
                title="Contratos generados por Trimestre",
                labels={"x": "Trimestre", "Cantidad": "Contratos"},
                hover_data={"Etapa": True},
            )
            fig_tr.update_traces(textposition="outside")
            fig_tr.update_layout(
                plot_bgcolor="white", paper_bgcolor="white", font_family="Inter",
                xaxis_title="Trimestre", yaxis_title="Contratos",
                legend_title="Etapa",
                bargap=0.2,
                annotations=[dict(
                    x=f"T{trim_max_bar}", y=vc_trim[vc_trim["Trimestre"]==trim_max_bar]["Cantidad"].values[0],
                    text="⭐ Pico", showarrow=True, arrowhead=2,
                    font=dict(color="#E74C3C", size=12), ax=0, ay=-35
                )],
            )
            st.plotly_chart(fig_tr, use_container_width=True)

    # ── GRÁFICO 2: Donut de etapas ────────────────────────────────────────
    with col2:
        if "Etapa formación" in dft.columns:
            vc_etapa = (dft[dft["Etapa formación"].isin(["LECTIVA", "PRODUCTIVA"])]
                        ["Etapa formación"].value_counts().reset_index())
            vc_etapa.columns = ["Etapa", "Cantidad"]

            if not vc_etapa.empty:
                etapa_max = vc_etapa.loc[vc_etapa["Cantidad"].idxmax(), "Etapa"]
                st.markdown(f"""
                <div class="alert-info">
                📌 <strong>Etapa con más registros: {etapa_max}</strong>
                </div>
                """, unsafe_allow_html=True)

                color_etapa_map = {"LECTIVA": "#E74C3C", "PRODUCTIVA": "#F39C12"}
                fig_donut = go.Figure(go.Pie(
                    labels=vc_etapa["Etapa"],
                    values=vc_etapa["Cantidad"],
                    hole=0.45,
                    marker_colors=[color_etapa_map.get(e, "#E74C3C") for e in vc_etapa["Etapa"]],
                    textinfo="percent+label",
                    textfont_size=13,
                    hovertemplate="<b>%{label}</b><br>Aprendices: %{value:,}<br>%{percent}<extra></extra>",
                ))
                fig_donut.update_layout(
                    title="Distribución por Etapa",
                    font_family="Inter",
                    plot_bgcolor="white", paper_bgcolor="white",
                    legend=dict(orientation="v", x=1.05),
                    margin=dict(t=50, b=10, l=10, r=10),
                )
                st.plotly_chart(fig_donut, use_container_width=True)

    # ── GRÁFICO 3: Heatmap Nivel × Trimestre ─────────────────────────────
    st.markdown("### 🔥 Concentración: Nivel × Trimestre")
    if "Nivel formación" in dft.columns:
        pivot_heat = (
            dft[dft["Trimestre actual"].notna() & (dft["Trimestre actual"] <= filtro_max_trim)]
            .groupby(["Nivel formación", "Trimestre actual"])
            .size()
            .reset_index(name="Cantidad")
        )
        pivot_heat["Trimestre actual"] = pivot_heat["Trimestre actual"].astype(int)
        pivot_heat["T"] = pivot_heat["Trimestre actual"].apply(lambda t: f"T{t}")

        if not pivot_heat.empty:
            # Construir matriz para heatmap
            pivot_matrix = pivot_heat.pivot(index="Nivel formación", columns="T", values="Cantidad").fillna(0)
            # Ordenar columnas T1, T2 ... T10
            col_order = sorted(pivot_matrix.columns, key=lambda x: int(x[1:]))
            pivot_matrix = pivot_matrix[col_order]

            fig_heat = go.Figure(go.Heatmap(
                z=pivot_matrix.values,
                x=pivot_matrix.columns.tolist(),
                y=pivot_matrix.index.tolist(),
                colorscale=[[0, "#EAF4FB"], [0.5, "#2980B9"], [1, "#1A5276"]],
                text=pivot_matrix.values.astype(int),
                texttemplate="%{text}",
                textfont_size=12,
                hovertemplate="<b>%{y}</b> — %{x}<br>Aprendices: %{z}<extra></extra>",
            ))
            fig_heat.update_layout(
                plot_bgcolor="white", paper_bgcolor="white", font_family="Inter",
                xaxis_title="Trimestre", yaxis_title="",
                height=max(200, len(pivot_matrix) * 70),
                margin=dict(l=10, r=10, t=30, b=10),
            )
            st.plotly_chart(fig_heat, use_container_width=True)

    # ── GRÁFICO 4: Barras apiladas Trimestre × Nivel ─────────────────────
    st.markdown("### 📊 Contratos por Trimestre y Nivel de Formación")
    if "Nivel formación" in dft.columns:
        df_apil = dft[dft["Trimestre actual"].notna() & (dft["Trimestre actual"] <= filtro_max_trim)].copy()
        df_apil["Trimestre actual"] = df_apil["Trimestre actual"].astype(int)
        df_apil["T"] = df_apil["Trimestre actual"].apply(lambda t: f"T{t}")

        pivot2 = (df_apil.groupby(["T", "Nivel formación"])
                  .size().reset_index(name="Contratos"))
        # Orden correcto T1…T10
        orden_t = sorted(pivot2["T"].unique(), key=lambda x: int(x[1:]))

        if not pivot2.empty:
            fig_tn = px.bar(
                pivot2,
                x="T", y="Contratos",
                color="Nivel formación",
                barmode="stack",
                category_orders={"T": orden_t},
                color_discrete_map={
                    "TÉCNICO":   "#F39C12",
                    "TECNÓLOGO": "#2980B9",
                    "OPERARIO":  "#8E44AD",
                    "AUXILIAR":  "#27AE60",
                    "OTRO":      "#E74C3C",
                },
                title="",
                labels={"T": "Trimestre", "Contratos": "Aprendices"},
                hover_data={"Nivel formación": True},
                text="Contratos",
            )
            fig_tn.update_traces(textposition="inside", textfont_size=10)
            fig_tn.update_layout(
                plot_bgcolor="white", paper_bgcolor="white",
                font_family="Inter", xaxis_title="Trimestre",
                legend_title="Nivel",
                bargap=0.15,
            )
            st.plotly_chart(fig_tn, use_container_width=True)

    # ── GRÁFICO 5: Línea de evolución por trimestre ───────────────────────
    if "Nivel formación" in dft.columns and not dft.empty:
        st.markdown("### 📈 Evolución acumulada por trimestre y nivel")
        df_lin = dft[dft["Trimestre actual"].notna() & (dft["Trimestre actual"] <= filtro_max_trim)].copy()
        df_lin["Trimestre actual"] = df_lin["Trimestre actual"].astype(int)

        pivot_lin = (df_lin.groupby(["Trimestre actual", "Nivel formación"])
                     .size().reset_index(name="Cantidad"))

        fig_lin = px.line(
            pivot_lin, x="Trimestre actual", y="Cantidad",
            color="Nivel formación",
            markers=True,
            color_discrete_map={
                "TÉCNICO":   "#F39C12",
                "TECNÓLOGO": "#2980B9",
                "OPERARIO":  "#8E44AD",
                "AUXILIAR":  "#27AE60",
            },
            labels={"Trimestre actual": "Trimestre", "Cantidad": "Aprendices"},
            hover_data={"Nivel formación": True},
        )
        fig_lin.update_traces(line_width=2.5)
        fig_lin.update_layout(
            plot_bgcolor="white", paper_bgcolor="white",
            font_family="Inter", xaxis_title="Trimestre",
            legend_title="Nivel",
            xaxis=dict(tickmode="linear", dtick=1,
                       ticktext=[f"T{i}" for i in range(1, filtro_max_trim+1)],
                       tickvals=list(range(1, filtro_max_trim+1))),
        )
        st.plotly_chart(fig_lin, use_container_width=True)
# ─────────────────────────────────────────────
# ── PÁGINA: EMPRESAS Y CONTRATOS ─────────────
# ─────────────────────────────────────────────
elif pagina == "empresas":
    st.markdown("""
    <div class="section-header">
        <div>
            <h2>🏢 Empresas y Análisis de Contratos</h2>
            <p>Top empresas contratantes y duración de contratos</p>
        </div>
    </div>
    """, unsafe_allow_html=True)

    if "Razón social" not in df.columns:
        st.warning("No se encontró la columna 'Razón social' en el dataset.")
    else:
        df_emp = df[df["Razón social"].notna() & (df["Razón social"] != "")]
        total_empresas = df_emp["Razón social"].nunique()
        total_contratos = int(df.get("Tiene contrato", pd.Series([False]*len(df))).sum())

        c1, c2, c3 = st.columns(3)
        tarjeta_metrica(c1, f"{total_empresas:,}", "Empresas únicas")
        tarjeta_metrica(c2, f"{total_contratos:,}", "Total contratos")
        if total_empresas > 0:
            tarjeta_metrica(c3, f"{round(total_contratos/total_empresas,1)}", "Contratos / empresa", color="#E74C3C")

        # Top empresas
        top_n = st.slider("Top N empresas", 5, 30, 15)
        top_emp = (df_emp.groupby("Razón social").size()
                   .sort_values(ascending=False).head(top_n).reset_index(name="Contratos"))
        top_emp["Razón social"] = top_emp["Razón social"].str[:50]
        fig_emp = px.bar(top_emp, y="Razón social", x="Contratos",
                         orientation="h",
                         color="Contratos",
                         color_continuous_scale=["#FFCCCC","#E74C3C"],
                         title=f"🏆 Top {top_n} Empresas por Número de Contratos")
        fig_emp.update_layout(plot_bgcolor="white", paper_bgcolor="white",
                              font_family="Inter", height=max(400, top_n*28),
                              yaxis={"categoryorder":"total ascending"})
        st.plotly_chart(fig_emp, use_container_width=True)

        # Por nivel
        if "Nivel formación" in df.columns:
            top_emp_niv = (df_emp.groupby(["Razón social","Nivel formación"])
                           .size().reset_index(name="Contratos")
                           .sort_values("Contratos", ascending=False)
                           .groupby("Nivel formación").head(5))
            top_emp_niv["Razón social"] = top_emp_niv["Razón social"].str[:40]
            fig_emp_niv = px.bar(top_emp_niv, y="Razón social", x="Contratos",
                                 color="Nivel formación",
                                 orientation="h",
                                 barmode="stack",
                                 title="Top Empresas por Nivel de Formación",
                                 color_discrete_sequence=PALETA)
            fig_emp_niv.update_layout(plot_bgcolor="white", paper_bgcolor="white",
                                      font_family="Inter", height=500,
                                      yaxis={"categoryorder":"total ascending"})
            st.plotly_chart(fig_emp_niv, use_container_width=True)

    # Duración de contratos
    if "Duración contrato (días)" in df.columns:
        st.markdown("### ⏱️ Distribución de duración de contratos")
        df_dur = df[df["Duración contrato (días)"].notna() & (df["Duración contrato (días)"]>0)]
        fig_hist = px.histogram(df_dur, x="Duración contrato (días)",
                                nbins=40,
                                color_discrete_sequence=["#000080"],
                                title="Histograma: Duración de contratos (días)")
        fig_hist.update_layout(plot_bgcolor="white", paper_bgcolor="white",
                               font_family="Inter")
        st.plotly_chart(fig_hist, use_container_width=True)

        if "Nivel formación" in df.columns:
            fig_box = px.box(df_dur, x="Nivel formación", y="Duración contrato (días)",
                             color="Nivel formación",
                             color_discrete_sequence=PALETA,
                             title="Duración de Contratos por Nivel")
            fig_box.update_layout(plot_bgcolor="white", paper_bgcolor="white",
                                  font_family="Inter")
            st.plotly_chart(fig_box, use_container_width=True)

    # ── PREDICCIÓN: Dificultad de Contratación por Programa ────────────────
    st.markdown("---")
    st.markdown("""
    <div class="section-header">
        <div>
            <h2>🔮 Predicción: Dificultad de Contratación por Programa</h2>
            <p>Programas con mayor competencia en el mercado laboral — basado en oferta vs. contratos obtenidos</p>
        </div>
    </div>
    """, unsafe_allow_html=True)

    _esp_col_emp = None
    for _c in df.columns:
        if "especialidad" in _c.lower():
            _esp_col_emp = _c
            break

    if _esp_col_emp and "Tiene contrato" in df.columns:
        df_pred_emp = df.copy()
        df_pred_emp[_esp_col_emp] = df_pred_emp[_esp_col_emp].apply(limpiar_texto)
        df_pred_emp["Nivel formación"] = df_pred_emp[_esp_col_emp].apply(detectar_nivel)
        df_pred_emp = df_pred_emp[df_pred_emp["Nivel formación"].isin(["TÉCNICO", "TECNÓLOGO"])]

        resumen_pred_emp = (
            df_pred_emp.groupby([_esp_col_emp, "Nivel formación"])
            .agg(Total=("Tiene contrato", "count"), Contratados=("Tiene contrato", "sum"))
            .reset_index()
        )
        resumen_pred_emp["Sin_contrato"] = resumen_pred_emp["Total"] - resumen_pred_emp["Contratados"]
        resumen_pred_emp["Tasa_contratacion"] = (
            resumen_pred_emp["Contratados"] / resumen_pred_emp["Total"] * 100
        ).round(1)
        resumen_pred_emp["Score_dificultad"] = (
            (100 - resumen_pred_emp["Tasa_contratacion"]) * 0.6
            + (resumen_pred_emp["Sin_contrato"] / resumen_pred_emp["Total"].max() * 100) * 0.4
        ).round(1)
        resumen_pred_emp["Programa_corto"] = resumen_pred_emp[_esp_col_emp].apply(
            lambda x: x[:55] + "…" if len(str(x)) > 55 else x
        )

        def _clasificar_dificultad(score):
            if score >= 65: return "MUY DIFÍCIL"
            elif score >= 45: return "DIFÍCIL"
            elif score >= 25: return "MODERADO"
            else: return "FAVORABLE"

        resumen_pred_emp["Nivel_dificultad"] = resumen_pred_emp["Score_dificultad"].apply(_clasificar_dificultad)
        resumen_pred_filtrado_emp = resumen_pred_emp[resumen_pred_emp["Total"] >= 30].copy()
        top_dificiles_emp = resumen_pred_filtrado_emp.nlargest(15, "Score_dificultad")

        color_map_dif_emp = {
            "MUY DIFÍCIL": "#E74C3C",
            "DIFÍCIL":     "#F39C12",
            "MODERADO":    "#F1C40F",
            "FAVORABLE":   "#27AE60",
        }

        col_pred1, col_pred2 = st.columns([3, 2])
        with col_pred1:
            fig_pred_emp = px.bar(
                top_dificiles_emp.sort_values("Score_dificultad"),
                x="Score_dificultad", y="Programa_corto",
                color="Nivel_dificultad", color_discrete_map=color_map_dif_emp,
                orientation="h",
                title="Top 15 programas con mayor dificultad de contratación",
                text="Tasa_contratacion",
                hover_data={"Total": True, "Contratados": True, "Sin_contrato": True, "Nivel formación": True},
            )
            fig_pred_emp.update_traces(texttemplate="%{text}% contratados", textposition="outside")
            fig_pred_emp.update_layout(
                plot_bgcolor="white", paper_bgcolor="white", font_family="Inter",
                height=max(480, len(top_dificiles_emp) * 36),
                yaxis={"categoryorder": "total ascending"},
                yaxis_title="", xaxis_title="Score de dificultad (0–100)",
                legend_title="Dificultad",
                margin=dict(l=10, r=120, t=50, b=10),
            )
            st.plotly_chart(fig_pred_emp, use_container_width=True)

        with col_pred2:
            avg_tasa_emp = resumen_pred_filtrado_emp["Tasa_contratacion"].mean()
            fig_scatter_emp = px.scatter(
                resumen_pred_filtrado_emp,
                x="Total", y="Tasa_contratacion",
                color="Nivel_dificultad", color_discrete_map=color_map_dif_emp,
                size="Sin_contrato", size_max=30,
                hover_name="Programa_corto",
                hover_data={"Total": True, "Contratados": True, "Sin_contrato": True},
                title="Oferta vs. Tasa de contratación por programa",
                labels={"Total": "Total aprendices (oferta)", "Tasa_contratacion": "Tasa de contratación (%)"},
            )
            fig_scatter_emp.add_hline(
                y=avg_tasa_emp, line_dash="dash", line_color="#7F8C8D",
                annotation_text=f"Promedio {avg_tasa_emp:.1f}%",
                annotation_position="bottom right",
            )
            fig_scatter_emp.update_layout(
                plot_bgcolor="white", paper_bgcolor="white", font_family="Inter",
                height=480, legend_title="Dificultad",
            )
            st.plotly_chart(fig_scatter_emp, use_container_width=True)

        prog_mas_dificil_emp = top_dificiles_emp.iloc[-1] if not top_dificiles_emp.empty else None
        programas_sw_emp = resumen_pred_filtrado_emp[
            resumen_pred_filtrado_emp[_esp_col_emp].str.upper().str.contains(
                "SOFTWARE|SISTEMAS|PROGRAMACION|DESARROLLO", na=False
            )
        ]
        col_ins1, col_ins2 = st.columns(2)
        with col_ins1:
            if prog_mas_dificil_emp is not None:
                st.markdown(f"""
                <div class="alert-error">
                🔴 <strong>Programa más competido:</strong><br>
                <em>{prog_mas_dificil_emp['Programa_corto']}</em><br>
                Solo el <strong>{prog_mas_dificil_emp['Tasa_contratacion']}%</strong> logra contrato
                ({int(prog_mas_dificil_emp['Sin_contrato'])} de {int(prog_mas_dificil_emp['Total'])} sin contrato)
                </div>
                """, unsafe_allow_html=True)
        with col_ins2:
            if not programas_sw_emp.empty:
                avg_sw_emp = programas_sw_emp["Tasa_contratacion"].mean()
                total_sw_sin_emp = int(programas_sw_emp["Sin_contrato"].sum())
                st.markdown(f"""
                <div class="alert-warning">
                💻 <strong>Área de Software/Sistemas:</strong><br>
                Tasa promedio de contratación: <strong>{avg_sw_emp:.1f}%</strong><br>
                <strong>{total_sw_sin_emp:,}</strong> aprendices sin contrato en esta área.
                </div>
                """, unsafe_allow_html=True)

        with st.expander("📋 Ver tabla completa de predicción por programa"):
            tabla_pred_emp = (
                resumen_pred_filtrado_emp[[
                    "Programa_corto", "Nivel formación", "Total",
                    "Contratados", "Sin_contrato", "Tasa_contratacion",
                    "Score_dificultad", "Nivel_dificultad"
                ]]
                .sort_values("Score_dificultad", ascending=False)
                .rename(columns={
                    "Programa_corto":    "Programa",
                    "Nivel formación":   "Nivel",
                    "Tasa_contratacion": "Tasa (%)",
                    "Score_dificultad":  "Score dificultad",
                    "Nivel_dificultad":  "Clasificación",
                    "Sin_contrato":      "Sin contrato",
                })
            )
            st.dataframe(tabla_pred_emp, use_container_width=True, hide_index=True)
            buf_pred_emp = io.BytesIO()
            tabla_pred_emp.to_excel(buf_pred_emp, index=False, engine="openpyxl")
            st.download_button(
                "⬇️ Descargar predicción como Excel",
                data=buf_pred_emp.getvalue(),
                file_name="prediccion_dificultad_contratacion.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
    else:
        st.markdown("""
        <div class="alert-warning">
        ⚠️ Se necesita la columna <strong>Especialidad</strong> y datos de contrato para generar la predicción.
        </div>
        """, unsafe_allow_html=True)

# ─────────────────────────────────────────────
# ── PÁGINA: MAPA DE COLOMBIA ──────────────────
# ─────────────────────────────────────────────
elif pagina == "mapa":
    st.markdown("""
    <div class="section-header">
        <div>
            <h2>🗺️ Mapa de Colombia</h2>
            <p>Selecciona un departamento para ver sus aprendices, estado de contratos y programas</p>
        </div>
    </div>
    """, unsafe_allow_html=True)

    import unicodedata

    # Coordenadas de las capitales de departamento (lat, lon, nombre capital)
    # No requiere ningún archivo externo (.geojson / .shp)
    CAPITALES_CO = {
        "AMAZONAS": (-4.20, -69.94, "Leticia"),
        "ANTIOQUIA": (6.25, -75.57, "Medellín"),
        "ARAUCA": (7.09, -70.76, "Arauca"),
        "ATLANTICO": (10.96, -74.80, "Barranquilla"),
        "BOLIVAR": (10.39, -75.51, "Cartagena"),
        "BOYACA": (5.54, -73.36, "Tunja"),
        "CALDAS": (5.07, -75.52, "Manizales"),
        "CAQUETA": (1.61, -75.61, "Florencia"),
        "CASANARE": (5.33, -72.40, "Yopal"),
        "CAUCA": (2.44, -76.61, "Popayán"),
        "CESAR": (10.46, -73.25, "Valledupar"),
        "CHOCO": (5.69, -76.66, "Quibdó"),
        "CORDOBA": (8.75, -75.88, "Montería"),
        "CUNDINAMARCA": (4.71, -74.07, "Bogotá D.C."),
        "BOGOTA DC": (4.71, -74.07, "Bogotá D.C."),
        "GUAINIA": (3.88, -67.92, "Inírida"),
        "GUAVIARE": (2.57, -72.64, "San José del Guaviare"),
        "HUILA": (2.93, -75.28, "Neiva"),
        "LA GUAJIRA": (11.54, -72.91, "Riohacha"),
        "MAGDALENA": (10.41, -74.41, "Santa Marta"),
        "META": (4.15, -73.64, "Villavicencio"),
        "NARINO": (1.21, -77.28, "Pasto"),
        "NORTE DE SANTANDER": (7.89, -72.51, "Cúcuta"),
        "PUTUMAYO": (1.15, -76.65, "Mocoa"),
        "QUINDIO": (4.46, -75.67, "Armenia"),
        "RISARALDA": (4.81, -75.69, "Pereira"),
        "SANTANDER": (7.12, -73.13, "Bucaramanga"),
        "SUCRE": (9.30, -75.40, "Sincelejo"),
        "TOLIMA": (4.44, -75.23, "Ibagué"),
        "VALLE DEL CAUCA": (3.45, -76.53, "Cali"),
        "VAUPES": (1.25, -70.23, "Mitú"),
        "VICHADA": (4.42, -69.80, "Puerto Carreño"),
        "ARCHIPIELAGO DE SAN ANDRES": (12.58, -81.70, "San Andrés"),
    }

    def _normalizar_depto(nombre):
        if pd.isna(nombre) or not str(nombre).strip():
            return ""
        s = str(nombre).upper().strip()
        s = ''.join(c for c in unicodedata.normalize('NFD', s) if unicodedata.category(c) != 'Mn')
        s = re.sub(r'[.,]', '', s)
        s = re.sub(r'\s+', ' ', s).strip()
        alias = {
            "BOGOTA": "BOGOTA DC", "BOGOTA D C": "BOGOTA DC",
            "VALLE": "VALLE DEL CAUCA",
            "GUAJIRA": "LA GUAJIRA",
            "SAN ANDRES": "ARCHIPIELAGO DE SAN ANDRES",
            "SAN ANDRES Y PROVIDENCIA": "ARCHIPIELAGO DE SAN ANDRES",
            "NORTE SANTANDER": "NORTE DE SANTANDER",
        }
        return alias.get(s, s)

    def _extraer_depto_de_ciudad(ciudad):
        """La columna 'Ciudad' tiene formato 'MUNICIPIO (DEPARTAMENTO)'.
        Extraemos el departamento real entre paréntesis."""
        if pd.isna(ciudad):
            return None
        m = re.search(r'\(([^)]+)\)\s*$', str(ciudad).strip())
        return m.group(1).strip() if m else str(ciudad).strip()

    # 'Regional' suele tener un único valor fijo (ej. Regional Distrito Capital)
    # aunque los aprendices estén repartidos por todo el país (formación virtual/a distancia).
    # Por eso el departamento real se obtiene de la columna 'Ciudad'.
    if "Ciudad" not in df.columns:
        st.warning("⚠️ No se encontró la columna 'Ciudad' en el dataset.")
        st.stop()

    df_mapa = df.copy()
    df_mapa["Depto_bruto"] = df_mapa["Ciudad"].apply(_extraer_depto_de_ciudad)
    df_mapa["Depto_norm"] = df_mapa["Depto_bruto"].apply(_normalizar_depto)
    df_mapa["Coords"] = df_mapa["Depto_norm"].map(CAPITALES_CO)

    sin_cruzar = df_mapa["Coords"].isna().sum()
    if sin_cruzar > 0:
        deptos_sin_cruzar = sorted(df_mapa[df_mapa["Coords"].isna()]["Ciudad"].dropna().unique().tolist())
        with st.expander(f"⚠️ {sin_cruzar} registros no se pudieron ubicar en el mapa", expanded=False):
            st.write("Valores de 'Ciudad' que no coinciden con ningún departamento:", deptos_sin_cruzar)

    df_geo = df_mapa.dropna(subset=["Coords"]).copy()
    df_geo["lat"] = df_geo["Coords"].apply(lambda c: c[0])
    df_geo["lon"] = df_geo["Coords"].apply(lambda c: c[1])
    df_geo["Capital"] = df_geo["Coords"].apply(lambda c: c[2])

    agg_dict = {"Total_aprendices": ("Depto_norm", "count")}
    if "Tiene contrato" in df_geo.columns:
        agg_dict["Contratados"] = ("Tiene contrato", "sum")

    resumen_geo = (
        df_geo.groupby(["Depto_norm", "lat", "lon", "Capital"])
        .agg(**agg_dict)
        .reset_index()
    )
    if "Contratados" in resumen_geo.columns:
        resumen_geo["Sin_contrato"] = resumen_geo["Total_aprendices"] - resumen_geo["Contratados"]
        resumen_geo["Tasa (%)"] = (resumen_geo["Contratados"] / resumen_geo["Total_aprendices"] * 100).round(1)

    # ── Selector de departamento (única fuente de verdad) ────────────────
    deptos_disponibles = sorted(resumen_geo["Depto_norm"].unique().tolist())

    # Si se hizo clic en el mapa en el render anterior, preseleccionamos ese depto
    if "mapa_click_depto" in st.session_state and st.session_state["mapa_click_depto"] in deptos_disponibles:
        if st.session_state.get("mapa_depto_sel") != st.session_state["mapa_click_depto"]:
            st.session_state["mapa_depto_sel"] = st.session_state["mapa_click_depto"]

    col_sel, col_metric = st.columns([2, 3])
    with col_sel:
        depto_sel = st.selectbox(
            "📍 Selecciona un departamento",
            ["-- Ninguno --"] + deptos_disponibles,
            key="mapa_depto_sel",
        )

    # ── Mapa de puntos sobre Colombia (sin archivos externos) ───────────
    resumen_geo["Seleccionado"] = resumen_geo["Depto_norm"].apply(
        lambda d: "Seleccionado" if d == depto_sel else "Otros"
    )
    fig_mapa = px.scatter_geo(
        resumen_geo,
        lat="lat", lon="lon",
        size="Total_aprendices",
        color="Seleccionado" if depto_sel != "-- Ninguno --" else "Total_aprendices",
        color_discrete_map={"Seleccionado": "#E74C3C", "Otros": "#27AE60"},
        color_continuous_scale="Greens",
        hover_name="Depto_norm",
        hover_data={"Capital": True, "Total_aprendices": True, "lat": False, "lon": False},
        size_max=45,
    )
    fig_mapa.update_traces(marker_line_color="white", marker_line_width=1)
    fig_mapa.update_geos(
        lonaxis_range=[-82, -66], lataxis_range=[-5, 14],
        showcountries=True, countrycolor="lightgray",
        showland=True, landcolor="#F4F4F4",
        showocean=True, oceancolor="#EAF4FB",
        showsubunits=True,
    )
    fig_mapa.update_layout(
        title="Aprendices por departamento — clic para seleccionar (tamaño = cantidad)",
        margin=dict(l=0, r=0, t=40, b=0), height=480,
        font_family="Inter", paper_bgcolor="white",
        showlegend=(depto_sel != "-- Ninguno --"),
        legend_title="",
    )

    evento_mapa = st.plotly_chart(
        fig_mapa, use_container_width=True,
        on_select="rerun", key="mapa_colombia_click",
    )

    # Si hubo un clic nuevo en el mapa, lo guardamos y forzamos un rerun
    # para que el selectbox se actualice en la siguiente vuelta del script.
    if evento_mapa and evento_mapa.get("selection") and evento_mapa["selection"].get("points"):
        pts = evento_mapa["selection"]["points"]
        if pts:
            depto_clickeado = pts[0].get("hovertext")
            if depto_clickeado and depto_clickeado != st.session_state.get("mapa_click_depto"):
                st.session_state["mapa_click_depto"] = depto_clickeado
                st.rerun()

    depto_final = depto_sel if depto_sel != "-- Ninguno --" else None

    with col_metric:
        if depto_final:
            fila = resumen_geo[resumen_geo["Depto_norm"] == depto_final]
            if not fila.empty:
                f = fila.iloc[0]
                mc1, mc2, mc3 = st.columns(3)
                tarjeta_metrica(mc1, f"{int(f['Total_aprendices']):,}", f"Aprendices en {depto_final}")
                if "Contratados" in f.index:
                    tarjeta_metrica(mc2, f"{int(f['Contratados']):,}", "Con contrato", f"{f['Tasa (%)']}%")
                    tarjeta_metrica(mc3, f"{int(f['Sin_contrato']):,}", "Sin contrato", color="#E74C3C")

    st.markdown("---")

    if not depto_final:
        st.info("👆 Haz clic en un punto del mapa o selecciona un departamento arriba para ver el detalle.")
    else:
        df_depto = df_geo[df_geo["Depto_norm"] == depto_final].copy()
        st.markdown(f"### 📍 Detalle: {depto_final} — {len(df_depto):,} aprendices")

        col_g1, col_g2 = st.columns(2)

        with col_g1:
            if "Estado aprendiz" in df_depto.columns:
                vc_estado = df_depto["Estado aprendiz"].value_counts().reset_index()
                vc_estado.columns = ["Estado", "Cantidad"]
                vc_estado = vc_estado.head(10)
                fig_estado_dep = px.bar(
                    vc_estado.sort_values("Cantidad"),
                    x="Cantidad", y="Estado", orientation="h",
                    color="Cantidad", color_continuous_scale="Blues",
                    text="Cantidad", title="Estado de los contratos",
                )
                fig_estado_dep.update_traces(textposition="outside")
                fig_estado_dep.update_layout(
                    plot_bgcolor="white", paper_bgcolor="white", font_family="Inter",
                    yaxis_title="", xaxis_title="Aprendices", coloraxis_showscale=False,
                    height=max(320, len(vc_estado) * 40),
                )
                st.plotly_chart(fig_estado_dep, use_container_width=True)
            else:
                st.info("No hay columna 'Estado aprendiz' disponible.")

        with col_g2:
            if "Tiene contrato" in df_depto.columns:
                vc_tc = df_depto["Tiene contrato"].map({True: "Con contrato", False: "Sin contrato"}).value_counts()
                fig_tc_dep = fig_pie(
                    vc_tc.index.tolist(), vc_tc.values.tolist(),
                    "Aprendices con / sin contrato",
                    colors=["#27AE60" if l == "Con contrato" else "#E74C3C" for l in vc_tc.index],
                )
                st.plotly_chart(fig_tc_dep, use_container_width=True)
            else:
                st.info("No hay columna 'Tiene contrato' disponible.")

        esp_col_mapa = next((c for c in df_depto.columns if "especialidad" in c.lower()), None)
        if esp_col_mapa:
            vc_prog = df_depto[esp_col_mapa].value_counts().head(10).reset_index()
            vc_prog.columns = ["Programa", "Cantidad"]
            vc_prog["Programa_corto"] = vc_prog["Programa"].apply(lambda x: x[:50] + "…" if len(str(x)) > 50 else x)
            fig_prog_dep = px.bar(
                vc_prog.sort_values("Cantidad"),
                x="Cantidad", y="Programa_corto", orientation="h",
                color="Cantidad", color_continuous_scale="Oranges",
                text="Cantidad", title=f"Top 10 programas en {depto_final}",
            )
            fig_prog_dep.update_traces(textposition="outside")
            fig_prog_dep.update_layout(
                plot_bgcolor="white", paper_bgcolor="white", font_family="Inter",
                yaxis_title="", xaxis_title="Aprendices", coloraxis_showscale=False,
                height=max(350, len(vc_prog) * 38),
            )
            st.plotly_chart(fig_prog_dep, use_container_width=True)
        else:
            st.info("No hay columna de 'Especialidad' / programa disponible.")

        with st.expander(f"📋 Ver tabla completa de aprendices en {depto_final}"):
            cols_tabla_dep = [c for c in ["Nombres", "Apellidos", esp_col_mapa, "Estado Aprendiz",
                                          "Nivel formación", "Razón social"] if c and c in df_depto.columns]
            st.dataframe(df_depto[cols_tabla_dep], use_container_width=True, hide_index=True)


elif pagina == "desercion":
    st.markdown("""
    <div class="section-header">
        <div>
            <h2>⚠️ Deserción y Riesgo</h2>
            <p>Identificación de aprendices en riesgo o con contratos cancelados</p>
        </div>
    </div>
    """, unsafe_allow_html=True)

    if "Estado aprendiz" not in df.columns:
        st.warning("No se encontró la columna 'Estado aprendiz'.")
    else:
        # Categorías de riesgo
        estados_desercion = [
            "Alumno Retirado","Cancelado","Bajo Rendimiento Académico",
            "Aplazado","Inhabilitado Por Actualizacion",
            "Inhabilitado por limite de tiempo (1 año)"
        ]
        estados_riesgo = ["Aprendiz no interesado en contrato","Pendiente Por Certificar"]
        estados_activos = ["Contratado","En Proceso de Selección","Disponible","Aprendiz Aplica"]

        df_des = df[df["Estado aprendiz"].isin(estados_desercion)]
        df_riesgo = df[df["Estado aprendiz"].isin(estados_riesgo)]
        df_activos = df[df["Estado aprendiz"].isin(estados_activos)]

        c1, c2, c3, c4 = st.columns(4)
        tarjeta_metrica(c1, f"{len(df_des):,}", "Aprendices en deserción/cancelación", color="#E74C3C")
        tarjeta_metrica(c2, f"{round(len(df_des)/len(df)*100,1)}%", "Tasa de deserción general", color="#E74C3C")
        tarjeta_metrica(c3, f"{len(df_riesgo):,}", "En riesgo (sin contrato/pendiente)", color="#F39C12")
        tarjeta_metrica(c4, f"{len(df_activos):,}", "Aprendices activos/contratados", color="#27AE60")

        col1, col2 = st.columns(2)
        with col1:
            # Causas de deserción
            vc_des = df_des["Estado aprendiz"].value_counts().reset_index()
            vc_des.columns = ["Causa","Cantidad"]
            fig_des = px.bar(vc_des, x="Cantidad", y="Causa", orientation="h",
                             color="Cantidad",
                             color_continuous_scale=["#FADBD8","#E74C3C"],
                             title="Causas de deserción / cancelación")
            fig_des.update_layout(plot_bgcolor="white", paper_bgcolor="white",
                                  font_family="Inter",
                                  yaxis={"categoryorder":"total ascending"})
            st.plotly_chart(fig_des, use_container_width=True)

        with col2:
            # Deserción por nivel
            if "Nivel formación" in df_des.columns and not df_des.empty:
                vc_des_niv = df_des["Nivel formación"].value_counts().reset_index()
                vc_des_niv.columns = ["Nivel","Cantidad"]
                fig_des_niv = fig_pie(vc_des_niv["Nivel"].tolist(),
                                      vc_des_niv["Cantidad"].tolist(),
                                      "Deserción por Nivel de Formación",
                                      colors=["#E74C3C","#8E44AD","#F39C12","#2980B9"])
                st.plotly_chart(fig_des_niv, use_container_width=True)

        # Comparativo Técnico vs Tecnólogo
        st.markdown("### 🔵🟢 Comparativo Técnico vs Tecnólogo: Deserción")
        df_comp = df[df["Nivel formación"].isin(["TÉCNICO","TECNÓLOGO"])]
        if not df_comp.empty and "Estado aprendiz" in df_comp.columns:
            pivot_comp = (df_comp.groupby(["Nivel formación","Estado aprendiz"])
                          .size().reset_index(name="Cantidad")
                          .sort_values("Cantidad", ascending=False))
            fig_comp2 = px.bar(pivot_comp, x="Estado aprendiz", y="Cantidad",
                               color="Nivel formación", barmode="group",
                               title="Estado aprendiz: Técnico vs Tecnólogo",
                               color_discrete_map={
                                   "TÉCNICO": "#7F8C8D",
                                   "TECNÓLOGO": "#2980B9"
                               })
            fig_comp2.update_layout(plot_bgcolor="white", paper_bgcolor="white",
                                    font_family="Inter", xaxis_tickangle=-30)
            st.plotly_chart(fig_comp2, use_container_width=True)

        # Motivode estado
        if "Motivo estado aprendiz" in df.columns:
            motivos = df["Motivo estado aprendiz"].dropna()
            if len(motivos) > 0:
                st.markdown("### 📝 Muestra de motivos registrados")
                motivos_sample = motivos.head(20).reset_index(drop=True)
                for i, m in enumerate(motivos_sample):
                    if str(m).strip():
                        st.markdown(f"- {str(m)[:200]}")

# ─────────────────────────────────────────────
# ── PÁGINA: EXPLORAR DATOS ───────────────────
# ─────────────────────────────────────────────
elif pagina == "explorar":
    st.markdown("""
    <div class="section-header">
        <div>
            <h2>🔍 Explorador de Datos</h2>
            <p>Busca, filtra y descarga registros del dataset</p>
        </div>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("### 🔧 Filtros avanzados")
    fc1, fc2 = st.columns(2)
    fc3, fc4 = st.columns(2)
    fc5, _ = st.columns(2)

    dff = df.copy()

    if "Nivel formación" in df.columns:
        opts_niv = ["Todos"] + sorted(df["Nivel formación"].dropna().unique().tolist())
        f_niv = fc1.selectbox("Nivel de formación", opts_niv, key="exp_niv")
        if f_niv != "Todos":
            dff = dff[dff["Nivel formación"] == f_niv]

    if "Etapa formación" in df.columns:
        opts_etapa = ["Todas"] + sorted(df["Etapa formación"].dropna().unique().tolist())
        f_etapa = fc2.selectbox("Etapa de formación", opts_etapa, key="exp_etapa")
        if f_etapa != "Todas":
            dff = dff[dff["Etapa formación"] == f_etapa]

    if "Estado aprendiz" in df.columns:
        opts_ea = ["Todos"] + sorted(df["Estado aprendiz"].dropna().unique().tolist())
        f_ea = fc3.selectbox("Estado aprendiz", opts_ea, key="exp_ea")
        if f_ea != "Todos":
            dff = dff[dff["Estado aprendiz"] == f_ea]

    if "Estado Contrato" in df.columns:
        opts_ec = ["Todos"] + sorted(df["Estado Contrato"].dropna().unique().tolist())
        f_ec = fc4.selectbox("Estado contrato", opts_ec, key="exp_ec")
        if f_ec != "Todos":
            dff = dff[dff["Estado Contrato"] == f_ec]

    if "Regional" in df.columns:
        opts_reg = ["Todas"] + sorted(df["Regional"].dropna().unique().tolist())
        f_reg = fc5.selectbox("Regional", opts_reg, key="exp_reg")
        if f_reg != "Todas":
            dff = dff[dff["Regional"] == f_reg]

    # Búsqueda por texto
    buscar = st.text_input("🔍 Buscar por nombre, documento, empresa o especialidad")
    if buscar:
        mask = pd.Series([False]*len(dff), index=dff.index)
        for col in ["Nombres","Apellidos","Numero Documento","Razón social","Especialidad"]:
            if col in dff.columns:
                mask |= dff[col].astype(str).str.contains(buscar, case=False, na=False)
        dff = dff[mask]

    st.info(f"Mostrando **{len(dff):,}** registros")

    # Selección de columnas visibles
    all_cols = dff.columns.tolist()
    default_cols = [c for c in ["Apellidos","Nombres","Especialidad","Nivel formación",
                                 "Etapa formación","Trimestre actual","Estado aprendiz",
                                 "Estado Contrato","Razón social","Fecha inicio contrato",
                                 "Fecha fin contrato"] if c in all_cols]
    cols_vis = st.multiselect("Columnas a mostrar", all_cols, default=default_cols)
    if not cols_vis:
        cols_vis = all_cols[:10]

    st.dataframe(dff[cols_vis].head(500), use_container_width=True)

    # Descarga
    buf = io.BytesIO()
    dff.to_excel(buf, index=False, engine="openpyxl")
    st.download_button(
        "⬇️ Descargar filtrado como Excel",
        data=buf.getvalue(),
        file_name="sena_filtrado.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
# ─────────────────────────────────────────────
# ── PÁGINA: ACTIVIDAD ECONÓMICA ──────────────
# ─────────────────────────────────────────────
elif pagina == "actividad_economica":
    st.markdown("""
    <div class="section-header">
        <div>
            <h2>🏭 Actividad Económica de las Empresas</h2>
            <p>¿Qué sectores contratan más aprendices SENA y en qué programas?</p>
        </div>
    </div>
    """, unsafe_allow_html=True)

    # ── 1. Cargar dataset de empresas reguladas ───────────────────────────
    df_reg, reg_info = _cargar_empresas_reguladas()

    if df_reg is None:
        st.error(f"""
        ❌ No se pudo cargar el archivo de empresas reguladas.
        **Detalle:** {reg_info}

        Asegúrate de que el archivo **{_EMPRESAS_REG_FILE}** esté en la misma
        carpeta que `index.py` y vuelve a ejecutar la aplicación.
        """)
        st.stop()

    st.markdown(f"""
    <div class="alert-info">
    ✅ Dataset de empresas reguladas cargado automáticamente desde <code>{reg_info}</code>
    — <strong>{len(df_reg):,}</strong> registros · {df_reg['NIT'].nunique():,} empresas únicas
    </div>
    """, unsafe_allow_html=True)

    # ── 2. Verificar dataset principal de aprendices ──────────────────────
    if df_global is None:
        st.markdown("""
        <div class="alert-warning">
        ⚠️ <strong>Carga el dataset de aprendices (Libro1)</strong> usando el panel lateral
        para activar el análisis cruzado por ficha y programa.
        <br>Se mostrarán estadísticas del dataset de empresas mientras tanto.
        </div>
        """, unsafe_allow_html=True)

    # ── 3. Mapa de códigos CIIU → descripción completa ────────────────────
    # Construir desde el propio dataset
    ciiu_map = (
        df_reg[["Actividad económica", "Descripción actividad económica"]]
        .drop_duplicates()
        .dropna()
        .set_index("Actividad económica")["Descripción actividad económica"]
        .to_dict()
    )

    # ── 4. Métricas del dataset de empresas reguladas ─────────────────────
    total_empresas_reg = df_reg[df_reg["NIT"] != 0]["NIT"].nunique()
    total_sectores = df_reg["Actividad económica"].nunique()
    empresas_con_contratos = (
    int((df_global["Estado Contrato"] == "VIGENTE").sum())
    if df_global is not None and "Estado Contrato" in df_global.columns else 0
)

    c1, c2, c3 = st.columns(3)
    tarjeta_metrica(c1, f"{total_empresas_reg:,}", "Empresas únicas registradas")
    tarjeta_metrica(c2, f"{total_sectores}", "Cantidad de sectores económicos ", color="#2980B9")
    tarjeta_metrica(c3, f"{empresas_con_contratos:,}", "Contratos vigentes Regional", color="#F39C12")

    st.markdown("---")

    # ── 5. Análisis por categoría de actividad económica ─────────────────
    st.markdown("### 📊 Distribución de Empresas por Sector Económico")

    # Agregar descripción al dataset de empresas
    df_reg["Sector"] = df_reg["Actividad económica"].map(ciiu_map).fillna(df_reg["Actividad económica"])
    df_reg["Sector_corto"] = df_reg["Sector"].apply(lambda x: x[:55] + "…" if len(str(x)) > 55 else x)

    # Número de empresas por sector
    empresas_por_sector = (
        df_reg.groupby(["Actividad económica", "Sector_corto"])
        .agg(
            Empresas=("Razon social", "nunique"),
            Contratos_vigentes=("Contratos vigentes", "sum") if "Contratos vigentes" in df_reg.columns else ("Razon social", "count"),
            Cuota_total=("Cuota", "sum") if "Cuota" in df_reg.columns else ("Razon social", "count"),
        )
        .reset_index()
        .sort_values("Empresas", ascending=False)
    )

    top_n_sectores = st.slider("Mostrar top N sectores", 5, len(empresas_por_sector), min(15, len(empresas_por_sector)), key="slider_sectores")
    df_top_sec = empresas_por_sector.head(top_n_sectores).copy()

        # Pie chart: proporción de empresas por sector (todos los sectores)
    df_pie_full = empresas_por_sector.sort_values("Empresas", ascending=False).copy()

    # Agrupar sectores pequeños en "Otros" para que el gráfico no quede saturado
    umbral_otros = 8  # nº de sectores principales a mostrar individualmente
    if len(df_pie_full) > umbral_otros:
        top_otros = df_pie_full.head(umbral_otros)
        resto = df_pie_full.iloc[umbral_otros:]
        fila_otros = pd.DataFrame({
            "Sector_corto": ["Otros sectores"],
            "Empresas": [resto["Empresas"].sum()],
        })
        df_pie_full = pd.concat([top_otros[["Sector_corto", "Empresas"]], fila_otros], ignore_index=True)

    fig_sec_pie = go.Figure(go.Pie(
        labels=df_pie_full["Sector_corto"].tolist(),
        values=df_pie_full["Empresas"].tolist(),
        marker_colors=PALETA,
        hole=0.45,
        textinfo="percent",
        textfont_size=12,
        textposition="outside",
        pull=[0.02] * len(df_pie_full),
    ))
    fig_sec_pie.update_layout(
        title=f"Proporción de empresas por sector ({len(empresas_por_sector)})",
        font_family="Inter",
        plot_bgcolor="white", paper_bgcolor="white",
        margin=dict(t=60, b=20, l=20, r=20),
        showlegend=True,
        legend=dict(orientation="v", x=1.02, y=0.5, font=dict(size=11)),
        height=500,
    )
    st.plotly_chart(fig_sec_pie, use_container_width=True)  

    # Contratos vigentes por sector (contando "VIGENTE" en Estado Contrato de aprendices)
    if df_global is not None and "Estado Contrato" in df_global.columns:
        st.markdown("### 📋 Contratos Vigentes y Cuota de Aprendices por Sector")

        # Normalizar razón social para cruzar aprendices con empresas reguladas
        def _normalizar_razon(s):
            if pd.isna(s):
                return ""
            return re.sub(r'\s+', ' ', str(s)).strip().upper()

        df_apr_norm = df_global.copy()
        df_apr_norm["_razon_norm"] = df_apr_norm["Razón social"].apply(_normalizar_razon) \
            if "Razón social" in df_apr_norm.columns else ""

        df_reg_norm2 = df_reg.copy()
        df_reg_norm2["_razon_norm"] = df_reg_norm2["Razon social"].apply(_normalizar_razon)
        df_reg_norm2["Sector_corto"] = df_reg_norm2["Actividad económica"].map(ciiu_map).fillna(df_reg_norm2["Actividad económica"])
        df_reg_norm2["Sector_corto"] = df_reg_norm2["Sector_corto"].apply(lambda x: x[:55] + "…" if len(str(x)) > 55 else x)

        mapa_sector = df_reg_norm2.drop_duplicates("_razon_norm").set_index("_razon_norm")["Sector_corto"]
        df_apr_norm["Sector_corto"] = df_apr_norm["_razon_norm"].map(mapa_sector)

        df_vigentes_sector = (
            df_apr_norm[df_apr_norm["Estado Contrato"] == "VIGENTE"]
            .groupby("Sector_corto")
            .size()
            .reset_index(name="Contratos_vigentes")
            .sort_values("Contratos_vigentes", ascending=False)
        )

        df_cv = df_vigentes_sector.head(top_n_sectores).copy()
        df_cv = df_cv.merge(
            empresas_por_sector[["Sector_corto", "Cuota_total"]],
            on="Sector_corto", how="left"
        )

        tab_cv1, tab_cv2 = st.tabs(["📄 Contratos vigentes", "🎯 Cuota de aprendices"])

        with tab_cv1:
            fig_cv = px.bar(
                df_cv.sort_values("Contratos_vigentes"),
                x="Contratos_vigentes", y="Sector_corto", orientation="h",
                color="Contratos_vigentes",
                color_continuous_scale=["#D6EAF8", "#2980B9"],
                title=f"Contratos vigentes por sector (Estado Contrato) — Top {top_n_sectores}",
                text="Contratos_vigentes",
            )
            fig_cv.update_traces(textposition="outside")
            fig_cv.update_layout(
                plot_bgcolor="white", paper_bgcolor="white", font_family="Inter",
                height=max(400, top_n_sectores * 32),
                yaxis={"categoryorder": "total ascending"},
                coloraxis_showscale=False, yaxis_title="", xaxis_title="Contratos vigentes",
            )
            st.plotly_chart(fig_cv, use_container_width=True)

        with tab_cv2:
            if "Cuota" in df_reg.columns:
                fig_cuota = px.bar(
                    df_cv.sort_values("Cuota_total"),
                    x="Cuota_total", y="Sector_corto", orientation="h",
                    color="Cuota_total",
                    color_continuous_scale=["#F5EEF8", "#8E44AD"],
                    title=f"Cuota total de aprendices por sector — Top {top_n_sectores}",
                    text="Cuota_total",
                )
                fig_cuota.update_traces(textposition="outside")
                fig_cuota.update_layout(
                    plot_bgcolor="white", paper_bgcolor="white", font_family="Inter",
                    height=max(400, top_n_sectores * 32),
                    yaxis={"categoryorder": "total ascending"},
                    coloraxis_showscale=False, yaxis_title="", xaxis_title="Cuota aprendices",
                )
                st.plotly_chart(fig_cuota, use_container_width=True)

    # ── 6. Análisis cruzado con aprendices (Libro1) ───────────────────────
    st.markdown("---")
    st.markdown("### 🔗 Análisis Cruzado: Aprendices × Sector Económico")

    if df_global is None:
        st.markdown("""
        <div class="alert-warning">
        👈 <strong>Sube el dataset de aprendices (Libro1)</strong> en el panel lateral
        para ver qué fichas y programas corresponden a cada sector económico.
        </div>
        """, unsafe_allow_html=True)
    else:
        # Normalizar razón social en ambos datasets para el cruce
        def normalizar_razon(s):
            if pd.isna(s):
                return ""
            return re.sub(r'\s+', ' ', str(s)).strip().upper()

        df_reg_norm = df_reg.copy()
        df_reg_norm["_razon_key"] = df_reg_norm["Razon social"].apply(normalizar_razon)

        df_apr = df_global.copy()
        # Detectar columna de razón social en aprendices
        razon_col = None
        for c in df_apr.columns:
            if "raz" in c.lower() and "social" in c.lower():
                razon_col = c
                break
            elif c.lower() in ["empresa", "razon social", "razón social"]:
                razon_col = c
                break

        if razon_col is None:
            st.warning("⚠️ No se encontró la columna 'Razón social' en el dataset de aprendices. Verifica el nombre de la columna.")
            st.stop()

        df_apr["_razon_key"] = df_apr[razon_col].apply(normalizar_razon)

        # Detectar columna de NÚMERO de documento para contar aprendices ÚNICOS
        # (sin esto, se contaban filas, y un mismo aprendiz puede tener varias
        # filas en Libro1 por múltiples fichas/contratos)
        #
        # IMPORTANTE: se excluyen columnas tipo "Tipo de Documento" (CC, TI, CE...)
        # porque esa columna también contiene la palabra "documento" pero solo
        # tiene un puñado de valores posibles (NO identifica aprendices únicos).
        doc_col = None
        candidatos_doc = []
        for c in df_apr.columns:
            cl = c.lower().strip()
            if "tipo" in cl and "documento" in cl:
                continue  # excluir explícitamente "Tipo de Documento"
            if ("numero" in cl or "número" in cl) and "documento" in cl:
                candidatos_doc.append((0, c))  # prioridad máxima: "Número Documento"
            elif "documento" in cl:
                candidatos_doc.append((1, c))
            elif cl in ["cedula", "cédula", "nit", "identificacion", "identificación"]:
                candidatos_doc.append((2, c))
        if candidatos_doc:
            candidatos_doc.sort(key=lambda x: x[0])
            doc_col = candidatos_doc[0][1]

        # Tabla de mapeo: razon_key → sector
        sector_lookup = (
            df_reg_norm[["_razon_key", "Actividad económica", "Sector", "Sector_corto"]]
            .drop_duplicates(subset=["_razon_key"])
        )

        # Cruce
        df_cruce = df_apr.merge(sector_lookup, on="_razon_key", how="left")

        if doc_col:
            # Conteo de APRENDICES ÚNICOS (por Número de Documento), no de filas
            n_aprendices_total = df_cruce[doc_col].nunique()
            n_cruzados = df_cruce.loc[df_cruce["Actividad económica"].notna(), doc_col].nunique()
        else:
            # Fallback: si no se detecta columna de documento, se usa el conteo de filas
            st.warning("⚠️ No se encontró una columna de 'Número de Documento' en el dataset de aprendices; "
                        "las tarjetas muestran número de filas, no de aprendices únicos.")
            n_aprendices_total = len(df_cruce)
            n_cruzados = df_cruce["Actividad económica"].notna().sum()

        pct_cruce = round(n_cruzados / n_aprendices_total * 100, 1) if n_aprendices_total else 0.0

        if doc_col:
            st.caption(f"🔎 Columna usada para identificar aprendices únicos: **{doc_col}**")

        col_cruce1, col_cruce2, col_cruce3 = st.columns(3)
        tarjeta_metrica(col_cruce1, f"{n_aprendices_total:,}", "Aprendices únicos en dataset Libro1")
        tarjeta_metrica(col_cruce2, f"{n_cruzados:,}", "Aprendices únicos con sector identificado", color="#27AE60")
        tarjeta_metrica(col_cruce3, f"{pct_cruce}%", "Tasa de cruce exitoso",
                        color="#27AE60" if pct_cruce >= 60 else "#F39C12")

        if pct_cruce < 30:
            st.markdown("""
            <div class="alert-warning">
            ⚠️ La tasa de cruce es baja. Puede deberse a diferencias de escritura en la razón social
            entre los dos archivos. Revisa que los nombres de empresa sean consistentes.
            </div>
            """, unsafe_allow_html=True)

        df_con_sector = df_cruce[df_cruce["Actividad económica"].notna()].copy()

        # ── Filtro por sector ──────────────────────────────────────────────
        sectores_disponibles = sorted(df_con_sector["Sector_corto"].dropna().unique().tolist())
        sector_sel = st.selectbox(
            "🔍 Filtrar por sector económico específico (opcional)",
            ["Todos los sectores"] + sectores_disponibles,
            key="sel_sector_cruce"
        )

        df_vista = df_con_sector.copy()
        if sector_sel != "Todos los sectores":
            df_vista = df_vista[df_vista["Sector_corto"] == sector_sel]

        # ── 6a. Aprendices por sector ──────────────────────────────────────
        st.markdown("#### 👷 Aprendices contratados por sector económico")

        # Detectar columna de estado del aprendiz
        estado_col = None
        for c in df_vista.columns:
            if "estado" in c.lower() and "aprendiz" in c.lower():
                estado_col = c
                break

        # Contar aprendices por sector (todos los estados)
        apr_por_sector = (
            df_con_sector.groupby("Sector_corto")
            .size()
            .reset_index(name="Aprendices")
            .sort_values("Aprendices", ascending=False)
            .head(top_n_sectores)
        )

        col_apr1, col_apr2 = st.columns([3, 2])
        with col_apr1:
            fig_apr_sec = px.bar(
                apr_por_sector.sort_values("Aprendices"),
                x="Aprendices", y="Sector_corto", orientation="h",
                color="Aprendices",
                color_continuous_scale=["#ECF0F1", "#7F8C8D"],
                title=f"Aprendices por sector económico — Top {top_n_sectores}",
                text="Aprendices",
            )
            fig_apr_sec.update_traces(textposition="outside")
            fig_apr_sec.update_layout(
                plot_bgcolor="white", paper_bgcolor="white", font_family="Inter",
                height=max(400, top_n_sectores * 32),
                yaxis={"categoryorder": "total ascending"},
                coloraxis_showscale=False, yaxis_title="", xaxis_title="Aprendices",
            )
            st.plotly_chart(fig_apr_sec, use_container_width=True)

        with col_apr2:
            fig_apr_pie = fig_pie(
                apr_por_sector["Sector_corto"].tolist(),
                apr_por_sector["Aprendices"].tolist(),
                "Proporción de aprendices por sector",
            )
            st.plotly_chart(fig_apr_pie, use_container_width=True)

        # ── 6b. Aprendices por estado y sector ────────────────────────────
        if estado_col:
            st.markdown("#### 📌 Estado del aprendiz por sector")
            apr_estado_sector = (
                df_vista.groupby(["Sector_corto", estado_col])
                .size().reset_index(name="Cantidad")
            )
            if not apr_estado_sector.empty:
                fig_est_sec = px.bar(
                    apr_estado_sector,
                    x="Cantidad", y="Sector_corto", color=estado_col,
                    orientation="h", barmode="stack",
                    title="Estado de aprendices por sector económico",
                    color_discrete_sequence=PALETA,
                )
                fig_est_sec.update_layout(
                    plot_bgcolor="white", paper_bgcolor="white", font_family="Inter",
                    height=max(400, len(apr_estado_sector["Sector_corto"].unique()) * 30),
                    yaxis={"categoryorder": "total ascending"},
                    yaxis_title="", xaxis_title="Aprendices",
                )
                st.plotly_chart(fig_est_sec, use_container_width=True)

        # ── 6c. Fichas y programas por sector ─────────────────────────────
        st.markdown("#### 🎓 Fichas y Programas por Sector Económico")

        # Detectar columna de ficha y especialidad
        ficha_col = None
        esp_col = None
        for c in df_vista.columns:
            cl = c.lower().strip()
            if cl == "ficha" and ficha_col is None:
                ficha_col = c
            if "especialidad" in cl and esp_col is None:
                esp_col = c

        tab_f1, tab_f2, tab_f3 = st.tabs(["📋 Programas más contratados", "🔢 Fichas por sector", "📊 Tabla detallada"])

        with tab_f1:
            if esp_col:
                prog_sector = (
                    df_vista.groupby(["Sector_corto", esp_col])
                    .size().reset_index(name="Aprendices")
                    .sort_values("Aprendices", ascending=False)
                )

                # Top 10 programas globales
                top_progs = (
                    df_vista[esp_col].value_counts().head(12).index.tolist()
                )
                prog_top_data = prog_sector[prog_sector[esp_col].isin(top_progs)].copy()
                prog_top_data[esp_col] = prog_top_data[esp_col].apply(
                    lambda x: x[:50] + "…" if len(str(x)) > 50 else x
                )

                if not prog_top_data.empty:
                    fig_prog = px.bar(
                        prog_top_data,
                        x="Aprendices", y=esp_col, color="Sector_corto",
                        orientation="h", barmode="stack",
                        title="Top programas — distribución por sector económico",
                        color_discrete_sequence=PALETA,
                    )
                    fig_prog.update_layout(
                        plot_bgcolor="white", paper_bgcolor="white", font_family="Inter",
                        height=max(450, len(top_progs) * 36),
                        yaxis={"categoryorder": "total ascending"},
                        yaxis_title="", xaxis_title="Aprendices",
                        legend_title="Sector",
                    )
                    st.plotly_chart(fig_prog, use_container_width=True)
            else:
                st.warning("No se encontró columna 'Especialidad' en el dataset de aprendices.")

        with tab_f2:
            if ficha_col and esp_col:
                ficha_sector = (
                    df_vista.groupby(["Sector_corto", ficha_col, esp_col])
                    .size().reset_index(name="Aprendices")
                    .sort_values("Aprendices", ascending=False)
                )
                # Top N fichas por sector seleccionado
                top_fichas = (
                    ficha_sector.groupby(ficha_col)["Aprendices"]
                    .sum()
                    .sort_values(ascending=False)
                    .head(15)
                    .index.tolist()
                )
                fichas_data = ficha_sector[ficha_sector[ficha_col].isin(top_fichas)].copy()
                fichas_data["Ficha_Prog"] = fichas_data[ficha_col].astype(str) + " – " + fichas_data[esp_col].apply(
                    lambda x: x[:40] + "…" if len(str(x)) > 40 else x
                )

                if not fichas_data.empty:
                    agg_fichas = (
                        fichas_data.groupby(["Ficha_Prog", "Sector_corto"])["Aprendices"]
                        .sum().reset_index()
                    )
                    fig_fichas = px.bar(
                        agg_fichas,
                        x="Aprendices", y="Ficha_Prog", color="Sector_corto",
                        orientation="h", barmode="stack",
                        title="Top fichas — distribución por sector económico",
                        color_discrete_sequence=PALETA,
                    )
                    fig_fichas.update_layout(
                        plot_bgcolor="white", paper_bgcolor="white", font_family="Inter",
                        height=max(450, len(top_fichas) * 34),
                        yaxis={"categoryorder": "total ascending"},
                        yaxis_title="", xaxis_title="Aprendices",
                        legend_title="Sector",
                    )
                    st.plotly_chart(fig_fichas, use_container_width=True)
            elif not ficha_col:
                st.warning("No se encontró columna 'Ficha' en el dataset de aprendices.")

        with tab_f3:
            # Tabla resumen: sector × programa × ficha
            cols_tabla = ["Sector_corto"]
            if ficha_col:
                cols_tabla.append(ficha_col)
            if esp_col:
                cols_tabla.append(esp_col)
            if estado_col:
                cols_tabla.append(estado_col)
            if "Ciudad" in df_vista.columns:
                cols_tabla.append("Ciudad")

            cols_tabla_disponibles = [c for c in cols_tabla if c in df_vista.columns]
            resumen_tabla = (
                df_vista[cols_tabla_disponibles]
                .copy()
                .rename(columns={"Sector_corto": "Sector Económico"})
            )
            resumen_tabla = resumen_tabla.drop_duplicates().sort_values("Sector Económico")

            st.dataframe(resumen_tabla, use_container_width=True, hide_index=True)

            # Descarga
            buf_act = io.BytesIO()
            resumen_tabla.to_excel(buf_act, index=False, engine="openpyxl")
            st.download_button(
                "⬇️ Descargar tabla como Excel",
                data=buf_act.getvalue(),
                file_name="actividad_economica_aprendices.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

        
    # ── 7. Tabla de sectores del archivo de empresas reguladas ────────────
    st.markdown("---")
    with st.expander("📖 Catálogo completo de sectores económicos"):
        catalogo = (
            df_reg[["Actividad económica", "Descripción actividad económica"]]
            .drop_duplicates()
            .sort_values("Actividad económica")
            .rename(columns={
                "Actividad económica": "Código CIIU",
                "Descripción actividad económica": "Descripción del sector"
            })
        )
        st.dataframe(catalogo, use_container_width=True, hide_index=True)