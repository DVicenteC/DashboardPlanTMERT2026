"""
Dashboard TMERT 2026 - Gestión Integral (Programación + Análisis EP)
Autor: Diego Vicente Contreras y Claude AI - IST 2026
"""

import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
import duckdb
import re
import io
from datetime import datetime, timedelta

# ── 1. CONFIGURACIÓN DE PÁGINA ───────────────────────────────────────────────
st.set_page_config(
    page_title="Dashboard TMERT 2026 - IST",
    page_icon="🏥",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ── 2. ESTILOS ───────────────────────────────────────────────────────────────
st.markdown("""
    <style>
    .main {background-color: #F8F9FA;}
    h1 {color: #2E86AB;}
    .stMetric {background-color: white; padding: 10px; border-radius: 5px;}
    .metric-card {
        background-color: white;
        padding: 20px;
        border-radius: 10px;
        border-left: 5px solid #2E86AB;
        box-shadow: 2px 2px 5px rgba(0,0,0,0.05);
    }
    .status-ep {color: #FF4B4B; font-weight: bold;}
    .detalle-section {
        background-color: #f0f7fb;
        padding: 15px;
        border-radius: 8px;
        margin-top: 20px;
    }
    </style>
""", unsafe_allow_html=True)

# ── 3. SISTEMA DE AUTENTICACIÓN ───────────────────────────────────────────────
def check_password():
    """Retorna True si el usuario ingresó las credenciales correctas."""

    def password_entered():
        """Revisa si las credenciales son correctas."""
        if (
            st.session_state["username"] == st.secrets["credentials"]["username"]
            and st.session_state["password"] == st.secrets["credentials"]["password"]
        ):
            st.session_state["password_correct"] = True
            del st.session_state["password"]  # Eliminar contraseña de session_state
            del st.session_state["username"]
        else:
            st.session_state["password_correct"] = False

    if st.session_state.get("password_correct", False):
        return True

    # Página de login con estilo premium
    st.markdown("""
        <style>
        .login-container {
            max-width: 450px;
            margin: 50px auto;
            padding: 40px;
            background-color: white;
            border-radius: 20px;
            box-shadow: 0 15px 35px rgba(0,0,0,0.1);
            border: 1px solid #E0E0E0;
        }
        .login-header {
            text-align: center;
            margin-bottom: 30px;
        }
        .login-title {
            color: #2E86AB;
            font-size: 24px;
            font-weight: bold;
            margin-top: 10px;
        }
        </style>
    """, unsafe_allow_html=True)

    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.markdown('<div class="login-container">', unsafe_allow_html=True)
        st.markdown('<div class="login-header">', unsafe_allow_html=True)
        st.markdown('<h1>🏥</h1>', unsafe_allow_html=True)
        st.markdown('<div class="login-title">Acceso Programación TMERT 2026</div>', unsafe_allow_html=True)
        st.markdown('</div>', unsafe_allow_html=True)

        st.text_input("Usuario", key="username", placeholder="Ingrese su RUT/ID")
        st.text_input("Contraseña", type="password", key="password", placeholder="••••••••")

        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("🚀 Ingresar al Dashboard", use_container_width=True, on_click=password_entered):
            pass  # on_click maneja la lógica

        if "password_correct" in st.session_state and not st.session_state["password_correct"]:
            st.error("❌ Credenciales incorrectas. Intente nuevamente.")

        st.markdown('<div style="text-align: center; margin-top: 20px; color: #888; font-size: 12px;">© 2026 IST - Especialidades Técnicas</div>', unsafe_allow_html=True)
        st.markdown('</div>', unsafe_allow_html=True)

    return False

# Solo continuar si está autenticado
if not check_password():
    st.stop()

# ── 4. CONEXIÓN A GOOGLE SHEETS ───────────────────────────────────────────────
def construir_url_exportacion(url_sheet):
    """Construye la URL de exportación CSV a partir de la URL de Google Sheets"""
    match_id = re.search(r'/d/([a-zA-Z0-9_-]+)', url_sheet)
    if not match_id:
        st.error("❌ No se pudo extraer el ID del spreadsheet de la URL configurada en secrets.toml")
        st.stop()
    spreadsheet_id = match_id.group(1)
    match_gid = re.search(r'gid=(\d+)', url_sheet)
    gid = match_gid.group(1) if match_gid else '0'
    return f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}/export?format=csv&gid={gid}"

def normalizar_columnas_tmert(df):
    """
    Normaliza nombres de columnas del CSV de Google Sheets para que coincidan
    con los nombres esperados (con tildes y caracteres especiales).
    Google Sheets exporta CSV sin tildes ni el símbolo ° en los encabezados.
    """
    mapeo_columnas = {
        'Fecha Asistencia Tecnica TMERT 2026*':              'Fecha Asistencia Técnica TMERT 2026*',
        'Region':                                            'Región',
        'Direccion CT':                                      'Dirección CT',
        'N de trabajadores(as) a evaluar 2026 N hombres':   'N° de trabajadores(as) a evaluar 2026 N° hombres',
        'N de trabajadores(as) a evaluar 2026 N mujeres':   'N° de trabajadores(as) a evaluar 2026 N° mujeres',
    }
    return df.rename(columns=mapeo_columnas)

# Vocabulario de "centro de trabajo activo" en 'Estado Centro de Trabajo'.
# El maestro de adherentes informa 'Si' (columna Est Sucursal); versiones previas
# usaban 'Activa'. El resto de los valores (Cerrada, Pasiva, Anulada, Desafiliada)
# son NO activos. Se aceptan ambos vocabularios para no romper el filtro si la
# fuente cambia de nuevo.
CT_ACTIVO_VALORES = {'SI', 'SÍ', 'ACTIVA', 'ACTIVO', 'S'}


def es_ct_activo(serie):
    """Máscara booleana de centros de trabajo activos, tolerante al vocabulario."""
    return serie.astype(str).str.strip().str.upper().isin(CT_ACTIVO_VALORES)


def parsear_fecha_flexible(serie):
    """
    Parsea una serie de fechas que puede tener formatos mixtos:
    - DD-MM-YYYY (formato Excel original)
    - M/D/YYYY o MM/DD/YYYY (formato Google Sheets export)
    - YYYY-MM-DD (formato ISO)
    """
    # ISO primero (lo que sube el procesador a GSheets)
    resultado = pd.to_datetime(serie, format='%Y-%m-%d', errors='coerce')
    mascara = resultado.isna() & serie.notna() & (serie.astype(str).str.strip() != '')
    if mascara.any():
        resultado[mascara] = pd.to_datetime(serie[mascara], format='%d-%m-%Y', errors='coerce')
    mascara = resultado.isna() & serie.notna() & (serie.astype(str).str.strip() != '')
    if mascara.any():
        resultado[mascara] = pd.to_datetime(serie[mascara], format='%d/%m/%Y', errors='coerce')
    mascara = resultado.isna() & serie.notna() & (serie.astype(str).str.strip() != '')
    if mascara.any():
        resultado[mascara] = pd.to_datetime(serie[mascara], dayfirst=True, errors='coerce')
    return resultado

# ── 4. CARGA DE DATOS ─────────────────────────────────────────────────────────
@st.cache_data(ttl=300)
def cargar_datos_seguimiento_tmert():
    """Carga datos de seguimiento TMERT desde Google Sheets."""
    try:
        url_seg = st.secrets["gsheets"].get("seguimiento_tmert")
        if not url_seg:
            return pd.DataFrame()
        
        export_url = construir_url_exportacion(url_seg)
        df = pd.read_csv(export_url, dtype=str)
        
        if df.empty:
            return pd.DataFrame()
            
        df.columns = df.columns.str.strip()
        
        # Parsear fechas (Día Primero) - Excluyendo columnas que son booleanas o de estado
        cols_fecha = [c for c in df.columns if ('Fecha' in c or 'Prescripción' in c) 
                      and 'Pilar' not in c and 'Estado' not in c]
        for col in cols_fecha:
            df[col] = parsear_fecha_flexible(df[col])
            
        # Convertir columnas booleanas (Pilar 1-4, Meta 5, Validado MK, Es_Programado)
        cols_bool = [c for c in df.columns if 'Pilar' in c or 'Cumplida' in c or 'Validado' in c or 'Programado' in c]
        bool_map = {
            'TRUE': True, 'FALSE': False, 
            '1': True, '0': False,
            'VERDADERO': True, 'FALSO': False,
            'VERDADERO ': True, 'FALSO ': False,
            'NAN': False, 'NONE': False, 'NAT': False
        }
        for col in cols_bool:
            # Asegurar limpieza y mapeo robusto
            df[col] = df[col].astype(str).str.upper().str.strip().map(bool_map).fillna(False)

        return df
    except Exception:
        return pd.DataFrame()

def load_data():
    try:
        # Leer URL desde secrets
        url_sheet = st.secrets["gsheets"]["url"]
        export_url = construir_url_exportacion(url_sheet)

        # Descargar CSV desde Google Sheets
        df = pd.read_csv(export_url)

        # Normalizar nombres de columnas (CSV no tiene tildes ni °)
        df = normalizar_columnas_tmert(df)

        # Cargar en DuckDB en memoria para mayor velocidad
        con = duckdb.connect(':memory:')
        con.register('tmert_raw', df)
        df = con.execute("SELECT * FROM tmert_raw").fetchdf()
        con.close()

        df.columns = df.columns.str.strip()

        # Columnas EP (denuncias)
        columnas_ep = ['folios', 'ocupaciones', 'tareas', 'observaciones',
                       'cie10', 'diagnosticos', 'segmentos']
        for col in columnas_ep:
            if col not in df.columns:
                df[col] = ""
            else:
                df[col] = df[col].fillna("")

        # Un registro tiene EP si su celda de folios contiene al menos un valor no vacío
        df['Tiene EP'] = df['folios'].astype(str).str.strip().ne("")

        # Alias 'ID-CT' como duplicado de la columna larga (sin renombrar la original)
        _id_long = next((c for c in df.columns if 'Identificador' in c and 'CT' in c), None)
        if _id_long and 'ID-CT' not in df.columns:
            df['ID-CT'] = df[_id_long].astype(str).str.upper().str.strip()

        # Columnas base
        if 'Región' in df.columns:
            df['Región'] = df['Región'].fillna("S/R").astype(str).str.replace(".0", "", regex=False)
        else:
            df['Región'] = "N/A"

        df['Ergonomo'] = df['Ergonomo'].fillna("No Asignado") if 'Ergonomo' in df.columns else "No Asignado"

        for col, default in [('Gerencia - Cuenta Nacional', 'Sin Dato'), ('Holding', 'Sin Dato')]:
            if col not in df.columns:
                df[col] = default
            else:
                df[col] = df[col].fillna(default).astype(str)

        # Fecha y columnas temporales para programación
        col_fecha = 'Fecha Asistencia Técnica TMERT 2026*'
        if col_fecha in df.columns:
            df[col_fecha] = parsear_fecha_flexible(df[col_fecha])
        df['fecha'] = df[col_fecha] if col_fecha in df.columns else pd.NaT

        nombres_meses = {
            1: 'Enero', 2: 'Febrero', 3: 'Marzo', 4: 'Abril',
            5: 'Mayo', 6: 'Junio', 7: 'Julio', 8: 'Agosto',
            9: 'Septiembre', 10: 'Octubre', 11: 'Noviembre', 12: 'Diciembre'
        }
        df['mes'] = df['fecha'].dt.month
        df['mes_nombre'] = df['mes'].map(nombres_meses)

        # Comuna CT
        if 'Comuna CT2' in df.columns:
            df['Comuna CT'] = df['Comuna CT2'].astype(str)
        elif 'Comuna CT' not in df.columns:
            df['Comuna CT'] = "S/D"
        df['Comuna CT'] = df['Comuna CT'].fillna("S/D").astype(str)

        df = df.dropna(how='all')
        return df

    except KeyError:
        st.error("❌ No se encontró la URL de Google Sheets en secrets.toml")
        st.info("Configura el archivo `.streamlit/secrets.toml` con la sección [gsheets] y la clave `url`.")
        st.stop()
    except Exception as e:
        st.error(f"❌ Error al cargar datos desde Google Sheets: {str(e)}")
        st.exception(e)
        st.stop()

# ── 4. FUNCIÓN AUXILIAR RANKING (EP) ─────────────────────────────────────────
def obtener_ranking_limpio(df, columna, separador_principal="||", separador_secundario=",",
                           separadores_extra=None):
    """
    Cuenta la frecuencia de elementos atómicos en una columna multi-valor.

    Estructura de los datos:
      - Separador entre registros de folio : ' || '
      - Separador secundario interno       :
          · tareas, diagnósticos  → ','
          · ocupaciones           → ' | '
          · segmentos             → ' ' (espacio)  → pasar separadores_extra=[" "]
    """
    if columna not in df.columns:
        return pd.DataFrame(columns=['Nombre', 'Cantidad'])

    datos = df[df[columna].astype(str).str.strip() != ""][columna].astype(str)

    if datos.empty:
        return pd.DataFrame(columns=['Nombre', 'Cantidad'])

    series = datos.str.split(separador_principal, regex=False).explode().str.strip()
    series = series.str.split(separador_secundario, regex=False).explode().str.strip()

    if separadores_extra:
        for sep in separadores_extra:
            series = series.str.split(sep, regex=False).explode().str.strip()

    conteo = series[series != ""].value_counts().reset_index()
    conteo.columns = ['Nombre', 'Cantidad']
    return conteo

# ── 5. ORDEN ANATÓMICO DE SEGMENTOS ──────────────────────────────────────────
# Proximal → distal, luego Derecho → Izquierdo; resto al final (alphabético)
ORDEN_SEGMENTOS = [
    "HOMBRO_DER", "HOMBRO_IZQ",
    "CODO_DER",   "CODO_IZQ",
    "MUÑECA_DER", "MUÑECA_IZQ",
    "MANO_DER",   "MANO_IZQ",
    "DEDOS_DER",  "DEDOS_IZQ",
    "PULGAR_DER", "PULGAR_IZQ",
    "CERVICAL",
    "LUMBAR",
]
_ORDEN_IDX = {s: i for i, s in enumerate(ORDEN_SEGMENTOS)}

def ordenar_segmentos(lista: list[str]) -> list[str]:
    """Ordena segmentos según criterio anatómico proximal→distal, DER→IZQ."""
    return sorted(lista, key=lambda s: (_ORDEN_IDX.get(s.upper(), len(ORDEN_SEGMENTOS)), s))


# ── 6. CONTEO DE FOLIOS EP ───────────────────────────────────────────────────
def contar_folios_distintos(df_sub):
    """
    Cuenta el número de folios EP únicos presentes en un subconjunto de filas.
    Cada folio está separado por ' || ' dentro de la celda.
    """
    if 'folios' not in df_sub.columns:
        return 0
    series = (df_sub['folios']
              .astype(str)
              .str.split("||", regex=False)
              .explode()
              .str.strip())
    series = series[series != ""]
    return int(series.nunique())


def folios_por_empresa(df_sub):
    """
    Devuelve un DataFrame con el número de folios EP distintos por empresa,
    ordenado de mayor a menor.
    """
    if 'folios' not in df_sub.columns or 'Nombre Empleador' not in df_sub.columns:
        return pd.DataFrame(columns=['Empresa', 'Folios EP'])

    series = (df_sub[['Nombre Empleador', 'folios']]
              .assign(folios=df_sub['folios'].astype(str))
              .set_index('Nombre Empleador')['folios']
              .str.split("||", regex=False)
              .explode()
              .str.strip())

    df_exp = series.reset_index()
    df_exp.columns = ['Empresa', 'folio']
    df_exp = df_exp[df_exp['folio'] != ""]
    result = (df_exp.groupby('Empresa')['folio']
              .nunique()
              .reset_index()
              .rename(columns={'folio': 'Folios EP'})
              .sort_values('Folios EP', ascending=False))
    return result


# ── 6. FUNCIONES DE GRÁFICOS (PROGRAMACIÓN) ──────────────────────────────────
def grafico_barras_mensuales(df):
    if len(df) == 0:
        return None
    conteo = df.groupby('mes', observed=True).size().reset_index(name='cantidad')
    nombres_meses_es = {
        1: 'Enero', 2: 'Febrero', 3: 'Marzo', 4: 'Abril',
        5: 'Mayo', 6: 'Junio', 7: 'Julio', 8: 'Agosto',
        9: 'Septiembre', 10: 'Octubre', 11: 'Noviembre', 12: 'Diciembre'
    }
    conteo['mes_nombre'] = conteo['mes'].map(nombres_meses_es).astype(str)
    conteo = conteo.sort_values('mes')
    fig = px.bar(
        conteo, x='mes_nombre', y='cantidad',
        title='Carga Mensual de Asistencias Técnicas TMERT',
        labels={'mes_nombre': 'Mes', 'cantidad': 'Cantidad de Asistencias'},
        color_discrete_sequence=['#2E86AB'], height=420,
        category_orders={"mes_nombre": list(nombres_meses_es.values())}
    )
    fig.update_traces(texttemplate='%{y}', textposition='outside')
    fig.update_layout(xaxis_tickangle=-45, xaxis_title='Mes',
                      yaxis_title='Asistencias Técnicas')
    return fig

def grafico_top_regiones(df):
    if len(df) == 0:
        return None
    regiones = df['Región'].value_counts().reset_index()
    regiones.columns = ['Región', 'Cantidad']
    fig = px.bar(
        regiones, x='Cantidad', y='Región',
        title='Distribución por Región', orientation='h',
        color_discrete_sequence=['#F39C12'], height=420
    )
    fig.update_traces(texttemplate='%{x}', textposition='outside')
    fig.update_layout(yaxis={'categoryorder': 'total ascending'})
    return fig

def grafico_top_ergonomos(df):
    if len(df) == 0:
        return None
    ergonomos = df['Ergonomo'].value_counts().head(10).reset_index()
    ergonomos.columns = ['Ergónomo', 'Cantidad']
    fig = px.bar(
        ergonomos, x='Cantidad', y='Ergónomo',
        title='Top 10 Especialistas con Mayor Carga',
        orientation='h', color_discrete_sequence=['#A23B72'], height=420
    )
    fig.update_traces(texttemplate='%{x}', textposition='outside')
    fig.update_layout(yaxis={'categoryorder': 'total ascending'})
    return fig

# ── 6b. EVOLUCIÓN TEMPORAL (reconstrucción por fechas reales) ─────────────────
# Mapeo pilar -> columna de fecha real de ejecución en SIGECO.
# El Pilar 5 (Seguimiento) ahora SÍ tiene fecha: el reporte base incorporó
# 'FECHA SEGUIMIENTO PRESCRIPCION 1/2' y el procesador consolida la más temprana
# de las dos (basta cualquiera de los dos seguimientos) en la columna de abajo.
PILAR_FECHA_REAL = {
    'Pilar 1 - Difusión':            'Fecha AT Difusión (real)',
    'Pilar 2 - Capacitación':        'Fecha AT Capacitación (real)',
    'Pilar 3 - Diseño Cap Pract':    'Fecha Diseño Cap Práctica (real)',
    'Pilar 4 - Prescripción Caract': 'Fecha Prescripción Caracterización (real)',
}

COL_FECHA_SEGUIMIENTO = 'Fecha Seguimiento Prescripción Caracterización (real)'

# Etiqueta corta para leyendas/tablas
PILAR_LABEL_CORTO = {
    'Pilar 1 - Difusión':            'P1 Difusión',
    'Pilar 2 - Capacitación':        'P2 Capacitación',
    'Pilar 3 - Diseño Cap Pract':    'P3 Diseño Cap.',
    'Pilar 4 - Prescripción Caract': 'P4 Prescripción',
    'Pilar 5 - Seguimiento':         'P5 Seguimiento',
}

# Inicio y fin del horizonte del plan (24 meses)
EVO_INICIO = pd.Timestamp('2025-01-01')
EVO_FIN_PLAN = pd.Timestamp('2026-12-31')


def serie_acumulada(fechas, idx):
    """Conteo acumulado: cuántas fechas son <= a cada punto del índice idx.
    Es el 'histograma acumulado' — una curva que solo sube."""
    s = pd.to_datetime(pd.Series(fechas), errors='coerce').dropna().sort_values()
    if s.empty:
        return np.zeros(len(idx), dtype=int)
    return np.searchsorted(s.values, np.asarray(idx.values), side='right')


def cuatro_pilares_fecha(df_seg):
    """Fecha EXACTA en que cada CT completó los 4 pilares fechables = fecha del
    ÚLTIMO de los 4 pilares, exigiendo que los 4 tengan fecha. Devuelve una serie
    de fechas (NaT donde falta algún pilar).

    Es el precursor de Meta 5 y su techo: como Meta 5 además exige el Seguimiento
    (sin fecha en SIGECO), la Meta 5 real a cualquier fecha pasada es <= esta curva.
    No se aproxima nada: cuenta solo CTs con sus 4 pilares efectivamente fechados."""
    cols = [c for c in PILAR_FECHA_REAL.values() if c in df_seg.columns]
    if not cols:
        return pd.Series(pd.NaT, index=df_seg.index)
    fechas = df_seg[cols].apply(pd.to_datetime, errors='coerce')
    todas_presentes = fechas.notna().all(axis=1)
    ult = fechas.max(axis=1)  # el pilar más tardío
    return ult.where(todas_presentes, other=pd.NaT)


def meta5_fecha(df_seg):
    """Ubica en el tiempo los CTs que YA cuentan como Meta 5, sin redefinir la meta.

    El universo es exactamente la columna 'Meta 5 Cumplida' que calcula el
    procesador (Es_Programado + 5 pilares, con el Pilar 5 desde
    ESTADO SEGUIMIENTO). Las fechas SOLO se usan para saber CUÁNDO ocurrió:
    fecha = la del último de los 5 pilares. Devuelve NaT donde el CT no cumple
    Meta 5 o donde falta alguna fecha, así el total de la curva nunca supera el
    conteo oficial de Meta 5."""
    cols = [c for c in PILAR_FECHA_REAL.values() if c in df_seg.columns]
    if (COL_FECHA_SEGUIMIENTO not in df_seg.columns
            or 'Meta 5 Cumplida' not in df_seg.columns
            or len(cols) < len(PILAR_FECHA_REAL)):
        return pd.Series(pd.NaT, index=df_seg.index)
    cols = cols + [COL_FECHA_SEGUIMIENTO]
    fechas = df_seg[cols].apply(pd.to_datetime, errors='coerce')
    cumple = df_seg['Meta 5 Cumplida'].fillna(False).astype(bool)
    ult = fechas.max(axis=1)
    return ult.where(cumple & fechas.notna().all(axis=1), other=pd.NaT)


# ── 6. FUNCIÓN PARETO EP ──────────────────────────────────────────────────────
def grafico_pareto(df_ep, columna, titulo, separador_secundario=","):
    """
    Genera un gráfico de Pareto (barras + línea acumulada) para la columna indicada.
    Retorna (fig, df_pareto) con Rank, Nombre, Casos EP, %, % Acumulado, Vital.
    """
    ranking = obtener_ranking_limpio(df_ep, columna,
                                     separador_secundario=separador_secundario)
    if ranking.empty:
        return None, pd.DataFrame()

    ranking = ranking.head(30).copy()
    total = ranking['Cantidad'].sum()
    ranking['Pct']     = (ranking['Cantidad'] / total * 100).round(1)
    ranking['PctAcum'] = ranking['Pct'].cumsum().round(1)
    ranking['Rank']    = range(1, len(ranking) + 1)
    ranking['Vital']   = ranking['PctAcum'] <= 80.0

    # Colores: vital = coral IST-friendly, resto = gris azulado
    colores = ['#C0392B' if v else '#AEC6CF' for v in ranking['Vital']]
    # Etiqueta en barra: mostrar valor solo si hay espacio (top 15)
    textos = [str(int(v)) if i < 15 else '' for i, v in enumerate(ranking['Cantidad'])]

    fig = go.Figure()

    # Barras
    fig.add_trace(go.Bar(
        x=ranking['Nombre'],
        y=ranking['Cantidad'],
        name='Casos EP',
        marker=dict(color=colores, line=dict(width=0)),
        text=textos,
        textposition='outside',
        textfont=dict(size=10, color='#333333'),
        yaxis='y1',
        hovertemplate='<b>%{x}</b><br>Casos EP: %{y}<extra></extra>',
    ))

    # Línea acumulada suavizada
    fig.add_trace(go.Scatter(
        x=ranking['Nombre'],
        y=ranking['PctAcum'],
        name='% Acumulado',
        mode='lines+markers',
        line=dict(color='#2C3E50', width=2.5, shape='spline'),
        marker=dict(size=6, color='#2C3E50', symbol='circle'),
        yaxis='y2',
        hovertemplate='%{y:.1f} %<extra>% Acumulado</extra>',
    ))

    # Área sombreada bajo la curva (zona vital)
    idx_corte = int(ranking['Vital'].sum()) - 1
    if idx_corte >= 0:
        x_vital = ranking['Nombre'].iloc[:idx_corte + 1].tolist()
        y_vital = ranking['PctAcum'].iloc[:idx_corte + 1].tolist()
        fig.add_trace(go.Scatter(
            x=x_vital + x_vital[::-1],
            y=y_vital + [0] * len(y_vital),
            fill='toself',
            fillcolor='rgba(192,57,43,0.08)',
            line=dict(width=0),
            showlegend=False,
            yaxis='y2',
            hoverinfo='skip',
        ))

    # Línea de corte 80 %
    fig.add_shape(
        type='line', xref='paper', yref='y2',
        x0=0, x1=1, y0=80, y1=80,
        line=dict(color='#E67E22', width=1.5, dash='dot')
    )
    fig.add_annotation(
        xref='paper', yref='y2', x=1.01, y=80,
        text='<b>80 %</b>', showarrow=False,
        xanchor='left', font=dict(color='#E67E22', size=11)
    )

    fig.update_layout(
        title=dict(text=titulo, font=dict(size=15, color='#2C3E50'), x=0),
        plot_bgcolor='white',
        paper_bgcolor='white',
        xaxis=dict(
            tickangle=-38,
            tickfont=dict(size=10, color='#555'),
            showgrid=False,
            linecolor='#DDDDDD',
        ),
        yaxis=dict(
            title='Casos EP',
            title_font=dict(size=11, color='#C0392B'),
            tickfont=dict(size=10),
            showgrid=True,
            gridcolor='#F0F0F0',
            side='left',
        ),
        yaxis2=dict(
            title='% Acumulado',
            title_font=dict(size=11, color='#2C3E50'),
            tickfont=dict(size=10),
            side='right',
            overlaying='y',
            range=[0, 115],
            showgrid=False,
            ticksuffix=' %',
        ),
        legend=dict(
            orientation='h', yanchor='bottom', y=1.02, x=0,
            font=dict(size=11),
        ),
        height=520,
        margin=dict(l=50, r=70, t=60, b=120),
        bargap=0.25,
    )
    return fig, ranking


# ── 7. FUNCIÓN DETALLE COMPLETO ───────────────────────────────────────────────
def mostrar_resumen_detallado(df_filtrado, seccion='tab1'):
    if len(df_filtrado) == 0:
        st.info("No hay asistencias técnicas para mostrar con los filtros seleccionados.")
        return

    st.markdown("### 📋 Detalle de Asistencias Técnicas TMERT")

    col1, col2, col3 = st.columns(3)
    with col1:
        st.markdown("#### Por Región")
        region_counts = df_filtrado['Región'].value_counts().reset_index()
        region_counts.columns = ['Región', 'Cantidad']
        st.dataframe(region_counts, use_container_width=True, hide_index=True)
    with col2:
        st.markdown("#### Por Ergónomo")
        ergonom_counts = df_filtrado['Ergonomo'].value_counts().head(10).reset_index()
        ergonom_counts.columns = ['Ergónomo', 'Cantidad']
        st.dataframe(ergonom_counts, use_container_width=True, hide_index=True)
    with col3:
        st.markdown("#### Por Comuna")
        comuna_counts = df_filtrado['Comuna CT'].value_counts().head(10).reset_index()
        comuna_counts.columns = ['Comuna', 'Cantidad']
        st.dataframe(comuna_counts, use_container_width=True, hide_index=True)

    st.markdown("---")

    col_hombres = 'N° de trabajadores(as) a evaluar 2026 N° hombres'
    col_mujeres  = 'N° de trabajadores(as) a evaluar 2026 N° mujeres'

    if col_hombres in df_filtrado.columns and col_mujeres in df_filtrado.columns:
        st.markdown("#### Resumen de Trabajadores a Evaluar")
        total_h = df_filtrado[col_hombres].fillna(0).sum()
        total_m = df_filtrado[col_mujeres].fillna(0).sum()
        w1, w2, w3 = st.columns(3)
        w1.metric("Total Trabajadores", f"{int(total_h + total_m):,}")
        w2.metric("Hombres", f"{int(total_h):,}")
        w3.metric("Mujeres", f"{int(total_m):,}")
        st.markdown("---")

    st.markdown("#### Listado Completo de Asistencias Técnicas")

    columnas_detalle = ['fecha', 'Región', 'Ergonomo', 'Nombre Empleador', 'Nombre CT', 'Comuna CT']
    if col_hombres in df_filtrado.columns:
        columnas_detalle.append(col_hombres)
    if col_mujeres in df_filtrado.columns:
        columnas_detalle.append(col_mujeres)
    if 'Dirección CT' in df_filtrado.columns:
        columnas_detalle.insert(6, 'Dirección CT')

    df_tabla = df_filtrado[[c for c in columnas_detalle if c in df_filtrado.columns]].copy()
    df_tabla['fecha'] = df_tabla['fecha'].dt.strftime('%d-%m-%Y')

    if col_hombres in df_tabla.columns and col_mujeres in df_tabla.columns:
        df_tabla['Total Trabajadores'] = (
            df_tabla[col_hombres].fillna(0) + df_tabla[col_mujeres].fillna(0)
        )

    df_tabla = df_tabla.rename(columns={
        'fecha': 'Fecha', 'Ergonomo': 'Ergónomo',
        'Nombre CT': 'Centro de Trabajo', 'Comuna CT': 'Comuna',
        'Dirección CT': 'Dirección', col_hombres: 'N° Hombres', col_mujeres: 'N° Mujeres'
    }).sort_values('Fecha')

    st.dataframe(df_tabla, use_container_width=True, height=400, hide_index=True)

    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        df_tabla.to_excel(writer, index=False, sheet_name='Detalle_TMERT')
    st.download_button(
        label="📥 Descargar Detalle en Excel",
        data=buffer.getvalue(),
        file_name=f'detalle_tmert_{datetime.now().strftime("%d-%m-%Y")}.xlsx',
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        key=f'download_btn_{seccion}'
    )

# ── 6d. TAREA 3: DUEÑO DE LA ASESORÍA EN NO-PROGRAMADOS ──────────────────────
def rellenar_dueno_noprog(df):
    """En CTs no programados (Es_Programado == False) sin Ergónomo asignado,
    atribuye la asesoría al 'Último Profesional Registra (sigeco)'. Si tampoco
    hay último profesional, queda 'Sin asignar' (nadie atendió). Solo visual:
    no toca los programados (esos conservan el Ergónomo del plan)."""
    if df is None or df.empty:
        return df
    if not {'Ergonomo', 'Es_Programado'}.issubset(df.columns):
        return df
    df = df.copy()
    _vacios = ['', 'nan', 'None']
    _noprog = df['Es_Programado'] == False
    # 1. Rellenar con el último profesional que registró en SIGECO
    if 'Último Profesional Registra (sigeco)' in df.columns:
        _vacio = df['Ergonomo'].isna() | df['Ergonomo'].astype(str).str.strip().isin(_vacios)
        _mask = _vacio & _noprog
        df.loc[_mask, 'Ergonomo'] = df.loc[_mask, 'Último Profesional Registra (sigeco)']
    # 2. Lo que quede vacío en no programados → 'Sin asignar'
    _vacio2 = df['Ergonomo'].isna() | df['Ergonomo'].astype(str).str.strip().isin(_vacios)
    df.loc[_vacio2 & _noprog, 'Ergonomo'] = 'Sin asignar'
    return df


# ── 7. INTERFAZ PRINCIPAL ─────────────────────────────────────────────────────
df_raw = load_data()
df_seg_raw = cargar_datos_seguimiento_tmert()

if df_raw is not None:

    # ── SIDEBAR — Filtros Bidireccionales Robustos (Cross-filtering) ──────────
    st.sidebar.image("https://www.ist.cl/wp-content/themes/ist/img/logo-ist.png", width=100)
    st.sidebar.title("🔍 Filtros de Gestión")

    solo_ep = st.sidebar.toggle("🚨 Ver solo centros con denuncias de EP", value=False)
    solo_activos = st.sidebar.toggle("🟢 Ver solo centros de trabajo activos", value=False)

    # IDs de CT activos (ver es_ct_activo / CT_ACTIVO_VALORES) tomados del seguimiento,
    # para poder filtrar también la programación: df_raw NO trae esa columna, así que
    # se cruza por ID-CT (match verificado 5.500/5.500, upper+strip sin normalización extra).
    _activos_ids = set()
    if solo_activos and not df_seg_raw.empty \
            and 'Estado Centro de Trabajo' in df_seg_raw.columns and 'ID-CT' in df_seg_raw.columns:
        _activos_ids = set(
            df_seg_raw.loc[
                es_ct_activo(df_seg_raw['Estado Centro de Trabajo']), 'ID-CT'
            ].astype(str).str.upper().str.strip()
        )
        if not _activos_ids:
            st.sidebar.warning(
                "⚠️ El filtro de CT activos no encontró coincidencias en "
                "'Estado Centro de Trabajo'. Valores presentes: "
                + ", ".join(sorted(
                    df_seg_raw['Estado Centro de Trabajo'].dropna().astype(str).unique()
                )[:6])
            )

    # Base: dataset de referencia (los toggles EP y "activos" actúan como pre-filtros)
    _base_t = df_raw.copy()
    if solo_ep and 'Tiene EP' in _base_t.columns:
        _base_t = _base_t[_base_t['Tiene EP'] == True]
    if solo_activos and _activos_ids and 'ID-CT' in _base_t.columns:
        _base_t = _base_t[
            _base_t['ID-CT'].astype(str).str.upper().str.strip().isin(_activos_ids)
        ]
    _base_t = _base_t.copy()

    # ── 1. Inicialización de session_state ────────────────────────────────────
    _defaults_t = {
        'tmert_ergo':     'Todos',
        'tmert_gerencia': 'Todas',
        'tmert_holding':  'Todos',
        'tmert_emplea':   'Todos',
        'tmert_region':   'Todas',
    }
    for k, v in _defaults_t.items():
        if k not in st.session_state:
            st.session_state[k] = v

    # Reset pendiente (del botón)
    if st.session_state.get('_tmert_reset', False):
        for k, v in _defaults_t.items():
            st.session_state[k] = v
        st.session_state['_tmert_reset'] = False

    # ── 2. Definición de filtros ──────────────────────────────────────────────
    # (key_session, columna_df, valor_todos)
    _defs_t = [
        ('tmert_ergo',     'Ergonomo',                  'Todos'),
        ('tmert_gerencia', 'Gerencia - Cuenta Nacional', 'Todas'),
        ('tmert_holding',  'Holding',                   'Todos'),
        ('tmert_emplea',   'Nombre Empleador',           'Todos'),
        ('tmert_region',   'Región',                    'Todas'),
    ]

    # ── 3. Funciones de cálculo core ──────────────────────────────────────────
    def _sin_t(excluir_key):
        """Filtra _base_t aplicando todos los filtros EXCEPTO el indicado."""
        dff = _base_t
        for key, col, all_val in _defs_t:
            if key == excluir_key:
                continue
            val = st.session_state.get(key, all_val)
            if val != all_val and col in dff.columns:
                dff = dff[dff[col] == val]
        return dff

    def _opts_t(key, col):
        dff = _sin_t(key)
        if col not in dff.columns:
            return []
        return sorted(dff[col].dropna().astype(str).unique().tolist())

    # ── 4. Pase de validación iterativa (Elimina el 'doble click') ─────────────
    # Garantiza que los valores en session_state sean válidos antes de crear widgets.
    for _ in range(len(_defs_t)):
        changed = False
        for key, col, all_val in _defs_t:
            val = st.session_state.get(key, all_val)
            if val == all_val:
                continue
            available = _opts_t(key, col)
            if val not in available:
                st.session_state[key] = all_val
                changed = True
        if not changed:
            break

    # ── 5. Renderizado de Widgets con key= ────────────────────────────────────
    # Ergónomo
    filtro_ergo = st.sidebar.selectbox(
        "Especialista / Ergónomo",
        ["Todos"] + _opts_t('tmert_ergo', 'Ergonomo'),
        key='tmert_ergo'
    )

    # Gerencia
    filtro_gerencia = st.sidebar.selectbox(
        "Gerencia - Cuenta Nacional",
        ["Todas"] + _opts_t('tmert_gerencia', 'Gerencia - Cuenta Nacional'),
        key='tmert_gerencia'
    )

    # Holding
    filtro_holding = st.sidebar.selectbox(
        "Holding",
        ["Todos"] + _opts_t('tmert_holding', 'Holding'),
        key='tmert_holding'
    )

    # Empleador
    filtro_empleador = st.sidebar.selectbox(
        "Nombre Empleador",
        ["Todos"] + _opts_t('tmert_emplea', 'Nombre Empleador'),
        key='tmert_emplea'
    )

    # Región
    filtro_reg = st.sidebar.selectbox(
        "Región",
        ["Todas"] + _opts_t('tmert_region', 'Región'),
        key='tmert_region'
    )

    # Mes: El mes es independiente de la base bidireccional porque solo aplica a df_prog
    meses_espanol = ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio',
                     'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre']
    filtro_mes = st.sidebar.selectbox("Mes (Programación)", ["Todos"] + meses_espanol)

    # ── 6. Resultado Final ────────────────────────────────────────────────────
    df_f = _base_t.copy()
    for key, col, all_val in _defs_t:
        val = st.session_state.get(key, all_val)
        if val != all_val and col in df_f.columns:
            df_f = df_f[df_f[col] == val]

    # Contador y Reseteo
    st.sidebar.markdown("---")
    st.sidebar.caption(f"🔎 **{len(df_f):,}** registros con los filtros actuales")

    if st.sidebar.button("🔄 Resetear Filtros"):
        st.session_state['_tmert_reset'] = True
        st.rerun()

    # ── APLICAR FILTROS ───────────────────────────────────────────────────────
    # df viene del cross-filtering del sidebar
    df = df_f.copy()

    # Aplicar los mismos filtros sobre df_seg (fuente separada)
    df_seg = df_seg_raw.copy() if not df_seg_raw.empty else pd.DataFrame()
    # Tarea 3: dueño de la asesoría en no programados (Ergónomo vacío → Último Profesional)
    df_seg = rellenar_dueno_noprog(df_seg)
    # Filtro EP: el seguimiento NO tiene 'folios' (esa columna vive sólo en df_raw/plan).
    # Por eso cruzamos por ID-CT contra los CTs marcados como Tiene EP en el plan.
    if solo_ep and not df_seg.empty and 'ID-CT' in df_seg.columns \
            and 'ID-CT' in df_raw.columns and 'Tiene EP' in df_raw.columns:
        _ep_ids_seg = set(
            df_raw.loc[df_raw['Tiene EP'] == True, 'ID-CT']
                  .astype(str).str.upper().str.strip()
        )
        df_seg = df_seg[
            df_seg['ID-CT'].astype(str).str.upper().str.strip().isin(_ep_ids_seg)
        ]
    if not df_seg.empty:
        if filtro_ergo != "Todos" and 'Ergonomo' in df_seg.columns:
            df_seg = df_seg[df_seg['Ergonomo'] == filtro_ergo]
        if filtro_gerencia != "Todas" and 'Gerencia - Cuenta Nacional' in df_seg.columns:
            df_seg = df_seg[df_seg['Gerencia - Cuenta Nacional'] == filtro_gerencia]
        if filtro_holding != "Todos" and 'Holding' in df_seg.columns:
            df_seg = df_seg[df_seg['Holding'] == filtro_holding]
        if filtro_empleador != "Todos" and 'Nombre Empleador' in df_seg.columns:
            df_seg = df_seg[df_seg['Nombre Empleador'] == filtro_empleador]
        if filtro_reg != "Todas" and 'Región' in df_seg.columns:
            df_seg = df_seg[df_seg['Región'] == filtro_reg]

    # Tarea 1: filtro "solo centros de trabajo activos" (ver es_ct_activo)
    if solo_activos and not df_seg.empty and 'Estado Centro de Trabajo' in df_seg.columns:
        df_seg = df_seg[es_ct_activo(df_seg['Estado Centro de Trabajo'])]

    # df_prog: registros con fecha programada (para tab Programación)
    df_prog = df[df['fecha'].notna()].copy()
    if filtro_mes != "Todos":
        meses_es_a_num = {m: i+1 for i, m in enumerate(meses_espanol)}
        df_prog = df_prog[df_prog['mes'] == meses_es_a_num[filtro_mes]]

    # ── TÍTULO ────────────────────────────────────────────────────────────────
    st.title("🏥 Dashboard TMERT 2026 - Gestión Integral")
    # Fecha de corte real: estampada por el procesador en columna _FechaCorte
    _fecha_corte = ""
    if df_seg_raw is not None and "_FechaCorte" in df_seg_raw.columns:
        _fc = df_seg_raw["_FechaCorte"].iloc[0]
        try:
            _fecha_corte = pd.to_datetime(_fc).strftime("%d/%m/%Y")
        except Exception:
            _fecha_corte = str(_fc)
    st.markdown(
        f"**IST · Especialidades Técnicas** | "
        f"Datos al: **{_fecha_corte}**" if _fecha_corte else
        "**IST · Especialidades Técnicas**"
    )

    # ── MÉTRICAS GLOBALES ─────────────────────────────────────────────────────
    m1, m2, m3 = st.columns(3)
    with m1:
        st.metric("AT Programadas", f"{len(df_prog):,}")
    with m2:
        if not df_seg.empty and 'Meta 5 Cumplida' in df_seg.columns:
            _meta5 = df_seg['Meta 5 Cumplida'].sum()
            _porc5 = (_meta5 / len(df_prog) * 100) if len(df_prog) > 0 else 0
            st.metric("Meta 5 Completa", f"{_meta5:,}", f"{_porc5:.1f}% del plan")
        else:
            st.metric("Meta 5 Completa", "S/D")
    with m3:
        n_ep = contar_folios_distintos(df)
        n_emp_ep = df[df['Tiene EP']]['Nombre Empleador'].nunique() if n_ep > 0 else 0
        st.metric("Folios EP", n_ep, f"{n_emp_ep} empresa(s)", delta_color="inverse")

    st.markdown("---")

    # ── TABS ──────────────────────────────────────────────────────────────────
    tabs = st.tabs([
        "📊 Programación",
        "🔍 Análisis de Denuncias EP",
        "✅ Estado Seguimiento",
        "👨‍⚕️ Evolución e Indicadores por Profesional",
        "🩺 Vigilancia de la Salud",
    ])
    tab1, tab2, tab_seg, tab_ind, tab_vs = tabs

    # ── TAB 1: PROGRAMACIÓN ───────────────────────────────────────────────────
    with tab1:
        if len(df_prog) > 0:
            col_c1, col_c2 = st.columns(2)
            with col_c1:
                fig_barras = grafico_barras_mensuales(df_prog)
                if fig_barras:
                    st.plotly_chart(fig_barras, use_container_width=True)
            with col_c2:
                fig_reg = grafico_top_regiones(df_prog)
                if fig_reg:
                    st.plotly_chart(fig_reg, use_container_width=True)

            st.divider()

            fig_ergo = grafico_top_ergonomos(df_prog)
            if fig_ergo:
                st.plotly_chart(fig_ergo, use_container_width=True)

            with st.expander("📋 Ver Detalle de Asistencias Técnicas", expanded=False):
                mostrar_resumen_detallado(df_prog, seccion='tab1')
        else:
            st.warning("⚠️ No hay registros con fecha programada para los filtros seleccionados.")

    # ── TAB: ESTADO SEGUIMIENTO ───────────────────────────────────────────────
    with tab_seg:
        if df_seg.empty:
            st.info("ℹ️ Sube un archivo de seguimiento para activar esta vista.")
        else:
            st.subheader("🎯 Cumplimiento Meta 5 (4 Pilares + Seguimiento 1)")

            total_plan = len(df_prog)
            cumplen_meta = df_seg['Meta 5 Cumplida'].sum() if 'Meta 5 Cumplida' in df_seg.columns else 0
            avance_meta = (cumplen_meta / total_plan) if total_plan > 0 else 0

            st.progress(avance_meta, text=f"Progreso hacia Meta Anual: {avance_meta*100:.1f}% ({cumplen_meta}/{total_plan} CTs)")

            # Métricas de avance
            c1, c2, c3 = st.columns(3)
            atrasadas = (df_seg['Estado AT'] == 'Pendiente atrasada').sum() if 'Estado AT' in df_seg.columns else 0

            c1.metric("Universo Plan", f"{total_plan:,}")
            c2.metric("Meta 5 Completa", f"{cumplen_meta:,}", "4 Pilares + Seg. 1")
            c3.metric("Pendientes Atrasadas", f"{atrasadas:,}", delta_color="inverse")

            st.divider()

            # Desglose por Pilares (incluye Pilar 5 - Seguimiento)
            st.markdown("#### 🧱 Desglose por Pilares de Cumplimiento")
            cp1, cp2, cp3, cp4, cp5 = st.columns(5)
            for col_m, label, p_col in zip(
                [cp1, cp2, cp3, cp4, cp5],
                ["P1: Difusión", "P2: Capacitación", "P3: Diseño Cap", "P4: Prescripción", "P5: Seguimiento"],
                ["Pilar 1 - Difusión", "Pilar 2 - Capacitación", "Pilar 3 - Diseño Cap Pract",
                 "Pilar 4 - Prescripción Caract", "Pilar 5 - Seguimiento"]
            ):
                if p_col in df_seg.columns:
                    val = int(df_seg[p_col].sum())
                    col_m.metric(label, f"{val:,}", f"{(val/total_plan*100):.0f}%" if total_plan > 0 else "0%")

            st.divider()

            # Resumen Meta 5 por Región
            st.markdown("**Resumen Meta 5 por Región**")
            if 'Meta 5 Cumplida' in df_seg.columns and 'Región' in df_seg.columns:
                res_reg = df_seg.groupby('Región')['Meta 5 Cumplida'].value_counts().unstack(fill_value=0)
                st.dataframe(res_reg, use_container_width=True)

            st.divider()

            # Tabla detallada con estado de avance por CT
            st.markdown("#### 📋 Detalle de Avance por Centro de Trabajo")
            col_h = 'N° de trabajadores(as) a evaluar 2026 N° hombres'
            col_m = 'N° de trabajadores(as) a evaluar 2026 N° mujeres'

            cols_s = [
                'Región', 'Ergonomo', 'Nombre Empleador', 'ID-CT', 'Nombre CT', 'Dirección CT', 'Comuna CT',
                'Estado Centro de Trabajo',
                'Es_Programado',
                'Estado AT',
                'Meta 5 Cumplida',
                'Cuantas AT Tiene',
                'Pilar 1 - Difusión',
                'Pilar 2 - Capacitación',
                'Pilar 3 - Diseño Cap Pract',
                'Pilar 4 - Prescripción Caract',
                'Pilar 5 - Seguimiento',
                'Estado Seguimiento Prescripción Caracterización (sigeco)',
                'Último Profesional Registra (sigeco)',
                'Fecha AT Difusión (real)',
                'Fecha AT Capacitación (real)',
                'Fecha Prescripción Caracterización (real)',
                'Fecha Diseño Cap Práctica (real)',
                # Vigilancia ambiental (Identificación y Evaluación)
                'Fecha Últ. Identificación Inicial (istprod)',
                'Fecha Identificación Avanzada (real)',
                'Condición Identificación Avanzada (istprod)',
                'Prescripción Evaluación Inicial (sigeco)',
                'Estado Seguimiento Eval Inicial (sigeco)',
                'Fecha Prescripción Eval Avanzada (sigeco)',
            ]
            cols_s = [c for c in cols_s if c in df_seg.columns]

            df_det = df_seg[cols_s].copy()
            for c in df_det.columns:
                if 'Fecha' in c or pd.api.types.is_datetime64_any_dtype(df_det[c]):
                    df_det[c] = pd.to_datetime(df_det[c], errors='coerce').dt.strftime('%d-%m-%Y').fillna('')
            if 'Cuantas AT Tiene' in df_det.columns:
                df_det['Cuantas AT Tiene'] = pd.to_numeric(
                    df_det['Cuantas AT Tiene'], errors='coerce').fillna(0).astype(int)
            # Programado / No Programado (legible)
            if 'Es_Programado' in df_det.columns:
                df_det['Es_Programado'] = df_det['Es_Programado'].map(
                    {True: 'Programado', False: 'No Programado'}).fillna('')
            # Condición de identificación avanzada → etiqueta corta
            _cond_map_tabla = {
                'Identificación Avanzada Condición Aceptable':  'Aceptable',
                'Identificación Avanzada Condición Critica':    'Crítica',
                'Identificación Avanzada Condición No Critica': 'No Crítica',
            }
            if 'Condición Identificación Avanzada (istprod)' in df_det.columns:
                df_det['Condición Identificación Avanzada (istprod)'] = (
                    df_det['Condición Identificación Avanzada (istprod)'].map(_cond_map_tabla).fillna(''))
            # Encabezados cortos para las columnas nuevas
            df_det = df_det.rename(columns={
                'Es_Programado': 'Programado',
                'Fecha Últ. Identificación Inicial (istprod)': 'Fecha Ident. Inicial',
                'Fecha Identificación Avanzada (real)': 'Fecha Ident. Avanzada',
                'Condición Identificación Avanzada (istprod)': 'Condición Ident. Avanzada',
                'Prescripción Evaluación Inicial (sigeco)': 'Presc. Eval. Inicial',
                'Estado Seguimiento Eval Inicial (sigeco)': 'Estado Seg. Eval. Inicial',
                'Fecha Prescripción Eval Avanzada (sigeco)': 'Fecha Presc. Eval. Avanzada',
            })

            st.dataframe(df_det, use_container_width=True, hide_index=True)

            buffer_seg = io.BytesIO()
            with pd.ExcelWriter(buffer_seg, engine='openpyxl') as writer:
                df_det.to_excel(writer, index=False, sheet_name='Estado_Seguimiento')
            st.download_button(
                label="📥 Descargar Estado de Seguimiento en Excel",
                data=buffer_seg.getvalue(),
                file_name=f'estado_seguimiento_tmert_{datetime.now().strftime("%d-%m-%Y")}.xlsx',
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                key='download_seg'
            )

            # ── 🌡️ VIGILANCIA AMBIENTAL: IDENTIFICACIÓN Y EVALUACIÓN ──────────
            st.divider()
            st.markdown("#### 🌡️ Vigilancia Ambiental: Identificación y Evaluación")
            st.caption(
                "Flujo de identificación del riesgo TMERT y su evaluación cualitativa/cuantitativa. "
                f"Respeta los filtros del panel lateral · universo actual: **{len(df_seg):,}** CTs."
            )

            _COND_COL = 'Condición Identificación Avanzada (istprod)'

            def _cond_corta(_s):
                """Etiqueta corta de la condición (ojo: 'No Critica' contiene 'Critica')."""
                _s = str(_s)
                if 'No Critica' in _s: return 'No Crítica'
                if 'Critica'   in _s: return 'Crítica'
                if 'Aceptable' in _s: return 'Aceptable'
                return None

            def _nn_seg(_col):
                """Conteo de CTs con la columna no vacía dentro del universo filtrado."""
                if _col not in df_seg.columns:
                    return 0
                _s = df_seg[_col].astype(str).str.strip()
                return int((df_seg[_col].notna() & ~_s.isin(['', 'nan', 'None', 'NaT'])).sum())

            _N_vig = len(df_seg)
            _etapas = [
                ('Identificación Inicial',     'Fecha Últ. Identificación Inicial (istprod)'),
                ('Identificación Avanzada',    'Fecha Identificación Avanzada (real)'),
                ('Presc. Evaluación Inicial',  'Prescripción Evaluación Inicial (sigeco)'),
                ('Presc. Evaluación Avanzada', 'Fecha Prescripción Eval Avanzada (sigeco)'),
            ]
            _cov = pd.DataFrame([(lbl, _nn_seg(c)) for lbl, c in _etapas],
                                columns=['Etapa', 'CTs'])

            _cvz1, _cvz2 = st.columns([3, 2])
            with _cvz1:
                st.markdown("**Cobertura por etapa**")
                _fig_cov = px.bar(_cov, x='CTs', y='Etapa', orientation='h', text='CTs',
                                  color_discrete_sequence=['#2E86AB'], height=300)
                _fig_cov.update_layout(
                    yaxis={'categoryorder': 'array',
                           'categoryarray': _cov['Etapa'].tolist()[::-1]},
                    margin=dict(l=0, t=10, b=0), xaxis_title='CTs', yaxis_title=''
                )
                _fig_cov.update_traces(textposition='outside')
                st.plotly_chart(_fig_cov, use_container_width=True)
                st.caption(" · ".join(f"{lbl}: **{n}** de {_N_vig:,}"
                                      for lbl, n in zip(_cov['Etapa'], _cov['CTs'])))
            with _cvz2:
                st.markdown("**Condición (Ident. Avanzada)**")
                if _COND_COL in df_seg.columns:
                    _cc = df_seg[_COND_COL].map(_cond_corta).dropna()
                    if not _cc.empty:
                        _cnt = (_cc.value_counts()
                                   .rename_axis('Condición').reset_index(name='CTs'))
                        _fig_pie = px.pie(
                            _cnt, names='Condición', values='CTs', hole=0.45,
                            color='Condición',
                            color_discrete_map={'Aceptable': '#1A936F',
                                                'No Crítica': '#E67E22',
                                                'Crítica': '#C0392B'},
                            height=300)
                        _fig_pie.update_traces(textinfo='value+percent')
                        _fig_pie.update_layout(margin=dict(t=10, b=0, l=0, r=0),
                                               legend=dict(orientation='h', y=-0.15))
                        st.plotly_chart(_fig_pie, use_container_width=True)
                    else:
                        st.info("Sin datos de condición en el filtro actual.")
                else:
                    st.info("Columna de condición no disponible.")

            # ── Alertas de plazo (Protocolo TMERT) ──
            if _COND_COL in df_seg.columns:
                _hoy_vig = pd.Timestamp.today().normalize()
                _dfc = df_seg.copy()
                _dfc['_cond'] = _dfc[_COND_COL].map(_cond_corta)
                _fa = pd.to_datetime(_dfc.get('Fecha Identificación Avanzada (real)'), errors='coerce')
                _fi = pd.to_datetime(_dfc.get('Fecha Últ. Identificación Inicial (istprod)'), errors='coerce')

                # Críticas: máx 90 días desde la evaluación cualitativa (Ident. Avanzada)
                _cri = _dfc[_dfc['_cond'] == 'Crítica'].copy()
                _dias = (_hoy_vig - _fa.reindex(_cri.index)).dt.days
                _cri['Días desde eval. cualitativa'] = _dias
                _venc = int((_dias > 90).sum())
                _enpz = int(((_dias >= 0) & (_dias <= 90)).sum())
                _pend = int(_dias.isna().sum())

                st.markdown(
                    "**🔴 Críticas — intervención en máx. 90 días desde la evaluación "
                    "cualitativa (Identificación Avanzada) · Protocolo TMERT**"
                )
                _m1, _m2, _m3 = st.columns(3)
                _m1.metric("Críticas", f"{len(_cri):,}")
                _m2.metric("Vencidas (>90 días)", f"{_venc:,}", delta_color="inverse")
                _m3.metric("En plazo / pendientes", f"{_enpz} / {_pend}")

                if not _cri.empty:
                    _presc_ok = (
                        pd.to_datetime(_dfc.get('Prescripción Evaluación Inicial (sigeco)'), errors='coerce').notna()
                        | pd.to_datetime(_dfc.get('Prescripción Evaluación Inicial (istprod)'), errors='coerce').notna()
                        | pd.to_datetime(_dfc.get('Fecha Prescripción Eval Avanzada (sigeco)'), errors='coerce').notna()
                    )
                    _cri['Estado plazo'] = _cri['Días desde eval. cualitativa'].map(
                        lambda d: 'Pendiente eval.' if pd.isna(d)
                        else ('Vencida' if d > 90 else 'En plazo'))
                    _cri['¿Ya con Presc. Evaluación?'] = _presc_ok.reindex(_cri.index).map(
                        {True: 'Sí', False: 'No'})
                    _cri_disp = _cri.copy()
                    _cri_disp['Fecha Ident. Avanzada'] = _fa.reindex(_cri.index).dt.strftime('%d-%m-%Y').fillna('')
                    _cols_cri = [c for c in [
                        'Ergonomo', 'Nombre Empleador', 'Nombre CT', 'Región',
                        'Fecha Ident. Avanzada', 'Días desde eval. cualitativa',
                        'Estado plazo', '¿Ya con Presc. Evaluación?'
                    ] if c in _cri_disp.columns]
                    _cri_disp = _cri_disp[_cols_cri].sort_values(
                        'Días desde eval. cualitativa', ascending=False, na_position='first')

                    def _hl_venc(row):
                        _v = row.get('Estado plazo') == 'Vencida'
                        return ['background-color: #fdecea' if _v else ''] * len(row)

                    st.dataframe(_cri_disp.style.apply(_hl_venc, axis=1),
                                 use_container_width=True, hide_index=True)
                    _buf_cri = io.BytesIO()
                    with pd.ExcelWriter(_buf_cri, engine='openpyxl') as _w:
                        _cri_disp.to_excel(_w, index=False, sheet_name='Criticas_TMERT')
                    st.download_button(
                        "📥 Descargar Críticas en Excel", data=_buf_cri.getvalue(),
                        file_name=f'criticas_tmert_{datetime.now().strftime("%d-%m-%Y")}.xlsx',
                        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                        key='download_criticas')

                # Aceptables: vigencia 36 meses desde la Identificación Inicial
                _ace = _dfc[_dfc['_cond'] == 'Aceptable']
                _meses_a = (_hoy_vig - _fi.reindex(_ace.index)).dt.days / 30.44
                _vig_venc = int((_meses_a > 36).sum())
                _vig_con = int(_fi.reindex(_ace.index).notna().sum())
                st.markdown(
                    "**🟢 Aceptables — vigencia de la identificación: 36 meses desde la "
                    "Identificación Inicial**"
                )
                _a1, _a2, _a3 = st.columns(3)
                _a1.metric("Aceptables", f"{len(_ace):,}")
                _a2.metric("Vencidas (>36 meses)", f"{_vig_venc:,}", delta_color="inverse")
                _a3.metric("Con fecha de identificación", f"{_vig_con:,}")
                if _vig_con > 0 and _vig_venc == 0:
                    st.caption("Ninguna identificación «Aceptable» supera los 36 meses "
                               "(datos 2025–2026; todas vigentes a la fecha).")

    # ── TAB: VIGILANCIA DE LA SALUD ───────────────────────────────────────────
    with tab_vs:
        if df_seg.empty:
            st.info("ℹ️ Se requiere data de seguimiento para calcular la vigilancia de la salud.")
        else:
            st.subheader("🩺 Vigilancia de la Salud — Protocolo TMERT")
            st.caption(
                "Vigilancia de las **personas**: qué centros de trabajo con condición Crítica (C) o "
                "No Crítica (M) tienen trabajadores pendientes de ingresar a vigilancia de la salud. "
                "La vigilancia del **ambiente** (identificación, evaluación y prescripción de medidas) "
                "vive en el tab «✅ Estado Seguimiento». Respeta los filtros del panel lateral · "
                f"universo actual: **{len(df_seg):,}** CTs."
            )

            # Nombres EXACTOS de la hoja de seguimiento (no pasa por normalizar_columnas_tmert).
            _VS_COND  = 'Condición o Nivel de Riesgo (C,M,A)'
            _VS_DH    = 'N° de hombres que deben ingresar a vigilancia de salud 2026'
            _VS_DM    = 'N° de mujeres que deben ingresar a vigilancia de salud 2026'
            _VS_EH    = 'N° de hombres evaluados 2026'
            _VS_EM    = 'N° de mujeres evaluadas 2026'
            _VS_FEXAM = 'Fecha última Vigilancia de Salud 2026'
            _VS_FEVAL = 'Fecha evaluación Vigilancia de Salud 2026'
            _VS_FIA   = 'Fecha Identificación Avanzada (real)'
            # Plazo de ingreso a vigilancia de la salud desde la Identificación Avanzada:
            # 30 días (criterio confirmado por jefatura, 26-08-2026). NO confundir con los
            # 90 días de intervención de las Críticas que usa la vigilancia ambiental.
            _VS_PLAZO_DIAS = 30

            _vs_falta = [c for c in (_VS_COND, _VS_DH, _VS_DM, _VS_EH, _VS_EM)
                         if c not in df_seg.columns]
            if _vs_falta:
                st.warning(
                    "⚠️ La hoja de seguimiento no trae las columnas de vigilancia de la salud: "
                    + " · ".join(f"«{c}»" for c in _vs_falta)
                    + ". Vuelve a correr el procesador y a subir la hoja para activar esta vista."
                )
            else:
                _vs = df_seg.copy()

                # Los conteos llegan como texto (la hoja se lee con dtype=str).
                # El NaN se DEJA como NaN: "sin nómina VS" y "nómina de 0 personas" son
                # cosas distintas, y esa diferencia es la que separa al pendiente que
                # nunca ingresó del que ingresó incompleto.
                for _c in (_VS_DH, _VS_DM, _VS_EH, _VS_EM):
                    _vs[_c] = pd.to_numeric(_vs[_c], errors='coerce')
                _vs['_deben'] = _vs[[_VS_DH, _VS_DM]].sum(axis=1, min_count=1)
                _vs['_evalu'] = _vs[[_VS_EH, _VS_EM]].sum(axis=1, min_count=1)

                # Condición C/M/A normalizada; los vacíos quedan como '' (no NaN) para que
                # los filtros por negación no se traguen los nulos.
                _vs['_cond'] = (_vs[_VS_COND].astype(str).str.strip().str.upper()
                                .replace({'NAN': '', 'NONE': '', 'NAT': '', '<NA>': ''}))
                _vs.loc[_vs[_VS_COND].isna(), '_cond'] = ''

                _vs_es_cm    = _vs['_cond'].isin(['C', 'M'])
                _vs_con_nom  = _vs['_deben'].notna()
                _vs_completa = _vs_con_nom & (_vs['_evalu'].fillna(0) >= _vs['_deben'])

                _vs['_estado_vs'] = np.where(
                    ~_vs_con_nom, 'Sin ingresar',
                    np.where(_vs_completa, 'Completa', 'Incompleta'))
                _vs['_brecha'] = (_vs['_deben'] - _vs['_evalu'].fillna(0)).clip(lower=0)

                # Antigüedad desde la Identificación Avanzada que originó la condición.
                _vs_hoy = pd.Timestamp.today().normalize()
                _vs_fia = (pd.to_datetime(_vs[_VS_FIA], errors='coerce')
                           if _VS_FIA in _vs.columns
                           else pd.Series(pd.NaT, index=_vs.index))
                _vs['_dias_ia'] = (_vs_hoy - _vs_fia).dt.days

                _VS_LBL_FUERA = f'Fuera de plazo (>{_VS_PLAZO_DIAS} días)'
                _vs['_estado_plazo'] = np.select(
                    [_vs['_estado_vs'] == 'Completa',
                     _vs['_dias_ia'].isna(),
                     _vs['_dias_ia'] > _VS_PLAZO_DIAS],
                    ['Completa', 'Sin fecha Ident. Avanzada', _VS_LBL_FUERA],
                    default='En plazo')

                _cm = _vs[_vs_es_cm].copy()

                # ── a) SEMÁFORO DE PENDIENTES ──────────────────────────────────
                st.markdown("#### 🚦 Semáforo de pendientes")
                _n_c   = int((_cm['_cond'] == 'C').sum())
                _n_m   = int((_cm['_cond'] == 'M').sum())
                _n_sin = int((_cm['_estado_vs'] == 'Sin ingresar').sum())
                _n_inc = int((_cm['_estado_vs'] == 'Incompleta').sum())
                _n_com = int((_cm['_estado_vs'] == 'Completa').sum())

                _s1, _s2, _s3, _s4 = st.columns(4)
                _s1.metric("CT que requieren VS (C o M)", f"{len(_cm):,}",
                           f"{_n_c} Críticas · {_n_m} No Críticas", delta_color="off")
                _s2.metric("🔴 Sin ingresar a vigilancia", f"{_n_sin:,}",
                           "sin nómina VS", delta_color="inverse")
                _s3.metric("🟠 Vigilancia incompleta", f"{_n_inc:,}",
                           "evaluados < deben", delta_color="inverse")
                _s4.metric("🟢 Vigilancia completa", f"{_n_com:,}",
                           "evaluados ≥ deben", delta_color="off")

                st.markdown(
                    f"**⏱️ Plazo de ingreso a vigilancia: {_VS_PLAZO_DIAS} días desde la "
                    "Identificación Avanzada** · se mide sobre los CT C/M que aún no están completos"
                )
                _pend_cm    = _cm[_cm['_estado_vs'] != 'Completa']
                _n_fuera    = int((_pend_cm['_dias_ia'] > _VS_PLAZO_DIAS).sum())
                _n_dentro   = int(((_pend_cm['_dias_ia'] >= 0) &
                                   (_pend_cm['_dias_ia'] <= _VS_PLAZO_DIAS)).sum())
                _n_sinfecha = int(_pend_cm['_dias_ia'].isna().sum())
                _p1, _p2, _p3 = st.columns(3)
                _p1.metric(f"Fuera de plazo (>{_VS_PLAZO_DIAS} días)", f"{_n_fuera:,}",
                           delta_color="inverse")
                _p2.metric(f"En plazo (≤{_VS_PLAZO_DIAS} días)", f"{_n_dentro:,}",
                           delta_color="off")
                _p3.metric("Sin fecha de Ident. Avanzada", f"{_n_sinfecha:,}", delta_color="off")

                st.markdown("**👥 Trabajadores en CT con condición C o M**")
                _t_deben   = int(_cm['_deben'].sum()) if _cm['_deben'].notna().any() else 0
                _t_evalu   = int(_cm['_evalu'].sum()) if _cm['_evalu'].notna().any() else 0
                _cobertura = (_t_evalu / _t_deben * 100) if _t_deben > 0 else 0.0
                _w1, _w2, _w3 = st.columns(3)
                _w1.metric("Deben ingresar a VS", f"{_t_deben:,}")
                _w2.metric("Evaluados", f"{_t_evalu:,}")
                _w3.metric("Cobertura", f"{_cobertura:.1f}%",
                           f"brecha: {_t_deben - _t_evalu:,} trabajadores",
                           delta_color="inverse")
                st.caption(
                    "⚠️ Los conteos de vigilancia de la salud cubren la ventana **2025-2026** "
                    "(el nombre de la columna dice 2026 sólo por compatibilidad con la hoja) y "
                    "cuentan **RUT distintos en toda la ventana**: no son cifras anuales y no "
                    "deben sumarse año a año."
                )

                # ── b) TABLA ACCIONABLE DE PENDIENTES ──────────────────────────
                st.divider()
                st.markdown("#### 📋 Pendientes de vigilancia de la salud")
                st.caption(
                    "CT con condición C o M que no están completos. Orden: primero Críticas, "
                    "luego No Críticas; dentro de cada grupo, mayor antigüedad desde la "
                    "Identificación Avanzada arriba. **Los CT sin fecha de Identificación "
                    "Avanzada quedan al final** (sin fecha no se puede juzgar la antigüedad). "
                    "En rojo, las Críticas que todavía no ingresan a vigilancia."
                )

                if _pend_cm.empty:
                    st.success("✅ No hay CT con condición C o M pendientes de vigilancia de la "
                               "salud en el filtro actual.")
                else:
                    _tab_vs = _pend_cm.copy()
                    _tab_vs['_ord_cond'] = _tab_vs['_cond'].map({'C': 0, 'M': 1}).fillna(9)
                    _tab_vs = _tab_vs.sort_values(
                        ['_ord_cond', '_dias_ia'], ascending=[True, False], na_position='last')

                    _tab_vs['Condición'] = _tab_vs['_cond'].map(
                        {'C': 'C · Crítica', 'M': 'M · No Crítica'}).fillna(_tab_vs['_cond'])
                    _tab_vs['Fecha Ident. Avanzada'] = (
                        _vs_fia.reindex(_tab_vs.index).dt.strftime('%d-%m-%Y').fillna(''))
                    _tab_vs['Días desde Ident. Avanzada'] = _tab_vs['_dias_ia']
                    _tab_vs['Estado plazo'] = _tab_vs['_estado_plazo']
                    _tab_vs['Estado VS']    = _tab_vs['_estado_vs']
                    _tab_vs['Brecha']       = _tab_vs['_brecha']
                    for _orig, _dest in ((_VS_DH, 'Deben H'), (_VS_DM, 'Deben M'),
                                         (_VS_EH, 'Evaluados H'), (_VS_EM, 'Evaluados M')):
                        _tab_vs[_dest] = _tab_vs[_orig]
                    for _fcol, _dest in ((_VS_FEXAM, 'Fecha última VS (examen)'),
                                         (_VS_FEVAL, 'Fecha últ. evaluación VS')):
                        if _fcol in _tab_vs.columns:
                            _tab_vs[_dest] = (pd.to_datetime(_tab_vs[_fcol], errors='coerce')
                                              .dt.strftime('%d-%m-%Y').fillna(''))

                    _cols_vs = [c for c in [
                        'Región', 'Ergonomo', 'Nombre Empleador', 'ID-CT', 'Nombre CT',
                        'Comuna CT', 'Estado Centro de Trabajo', 'Condición',
                        'Fecha Ident. Avanzada', 'Días desde Ident. Avanzada',
                        'Estado plazo', 'Estado VS',
                        'Deben H', 'Deben M', 'Evaluados H', 'Evaluados M', 'Brecha',
                        'Fecha última VS (examen)', 'Fecha últ. evaluación VS',
                    ] if c in _tab_vs.columns]
                    _tab_vs_disp = _tab_vs[_cols_vs].copy()
                    # Int64 nullable: enteros sin ".0" y conservando el NaN de
                    # los CT sin nómina VS (que no es lo mismo que un 0).
                    for _icol in ('Días desde Ident. Avanzada', 'Deben H', 'Deben M',
                                  'Evaluados H', 'Evaluados M', 'Brecha'):
                        if _icol in _tab_vs_disp.columns:
                            _tab_vs_disp[_icol] = _tab_vs_disp[_icol].astype('Int64')

                    def _hl_vs(row):
                        """Rojo suave: Crítica que todavía no ingresa a vigilancia."""
                        _rojo = (str(row.get('Condición', '')).startswith('C')
                                 and row.get('Estado VS') == 'Sin ingresar')
                        return ['background-color: #fdecea' if _rojo else ''] * len(row)

                    st.dataframe(_tab_vs_disp.style.apply(_hl_vs, axis=1),
                                 use_container_width=True, hide_index=True)

                    _buf_vs = io.BytesIO()
                    with pd.ExcelWriter(_buf_vs, engine='openpyxl') as _w:
                        _tab_vs_disp.to_excel(_w, index=False, sheet_name='Pendientes_VS')
                    st.download_button(
                        "📥 Descargar pendientes de VS en Excel", data=_buf_vs.getvalue(),
                        file_name=f'pendientes_vigilancia_salud_{datetime.now().strftime("%d-%m-%Y")}.xlsx',
                        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                        key='download_vs_pendientes')

                if _VS_FEXAM not in _vs.columns:
                    st.caption(
                        f"ℹ️ La hoja no trae «{_VS_FEXAM}» (último examen de vigilancia); se muestra "
                        "sólo la fecha de evaluación médica. Aparecerá al volver a correr el "
                        "procesador y resubir la hoja de seguimiento."
                    )

                # ── c) BRECHA INVERSA ──────────────────────────────────────────
                st.divider()
                st.markdown("#### 🔎 Brecha inversa — vigilancia sin condición registrada")
                _vs_inv       = _vs[_vs_con_nom & (_vs['_cond'] == '')]
                _vs_nom_total = int(_vs_con_nom.sum())
                _vs_nom_ace   = int((_vs_con_nom & (_vs['_cond'] == 'A')).sum())
                _b1, _b2 = st.columns(2)
                _b1.metric("CT con nómina VS", f"{_vs_nom_total:,}")
                _b2.metric("… sin condición registrada", f"{len(_vs_inv):,}",
                           "sin Ident. Avanzada en ISTProd", delta_color="inverse")
                st.caption(
                    "Trabajadores en vigilancia de la salud cuyo CT no tiene Identificación "
                    "Avanzada registrada en ISTProd. Es un hallazgo de calidad de datos: o falta "
                    "cargar la identificación, o la nómina está asociada al CT equivocado."
                    + (f" Además, **{_vs_nom_ace}** CT con nómina VS tienen condición Aceptable (A)."
                       if _vs_nom_ace else "")
                )
                if not _vs_inv.empty:
                    with st.expander(f"📄 Ver los {len(_vs_inv):,} CT con nómina VS sin condición",
                                     expanded=False):
                        _inv_disp = _vs_inv.copy()
                        _inv_disp['Deben H']     = _inv_disp[_VS_DH]
                        _inv_disp['Deben M']     = _inv_disp[_VS_DM]
                        _inv_disp['Evaluados H'] = _inv_disp[_VS_EH]
                        _inv_disp['Evaluados M'] = _inv_disp[_VS_EM]
                        for _fcol, _dest in ((_VS_FEXAM, 'Fecha última VS (examen)'),
                                             (_VS_FEVAL, 'Fecha últ. evaluación VS')):
                            if _fcol in _inv_disp.columns:
                                _inv_disp[_dest] = (pd.to_datetime(_inv_disp[_fcol], errors='coerce')
                                                    .dt.strftime('%d-%m-%Y').fillna(''))
                        _cols_inv = [c for c in [
                            'Región', 'Ergonomo', 'Nombre Empleador', 'ID-CT', 'Nombre CT',
                            'Comuna CT', 'Estado Centro de Trabajo',
                            'Deben H', 'Deben M', 'Evaluados H', 'Evaluados M',
                            'Fecha última VS (examen)', 'Fecha últ. evaluación VS',
                        ] if c in _inv_disp.columns]
                        _inv_disp = _inv_disp[_cols_inv].copy()
                        for _icol in ('Deben H', 'Deben M', 'Evaluados H', 'Evaluados M'):
                            if _icol in _inv_disp.columns:
                                _inv_disp[_icol] = _inv_disp[_icol].astype('Int64')
                        st.dataframe(_inv_disp, use_container_width=True, hide_index=True)
                        _buf_inv = io.BytesIO()
                        with pd.ExcelWriter(_buf_inv, engine='openpyxl') as _w:
                            _inv_disp.to_excel(_w, index=False, sheet_name='VS_sin_condicion')
                        st.download_button(
                            "📥 Descargar brecha inversa en Excel", data=_buf_inv.getvalue(),
                            file_name=f'vs_sin_condicion_{datetime.now().strftime("%d-%m-%Y")}.xlsx',
                            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                            key='download_vs_inversa')

                # ── d) COBERTURA POR REGIÓN Y POR ERGÓNOMO ─────────────────────
                st.divider()
                st.markdown("#### 🗺️ Dónde está la gestión pendiente")
                st.caption("CT con condición C o M según estado de vigilancia de la salud. "
                           "Ordenado por cantidad de pendientes (sin ingresar + incompletos).")
                if _cm.empty:
                    st.info("No hay CT con condición C o M en el filtro actual.")
                else:
                    _MAPA_VS = {'Sin ingresar': '#C0392B', 'Incompleta': '#E67E22',
                                'Completa': '#1A936F'}
                    _g1, _g2 = st.columns(2)
                    for _box, _dim, _tit in ((_g1, 'Región', 'Por Región'),
                                             (_g2, 'Ergonomo', 'Por Ergónomo')):
                        with _box:
                            st.markdown(f"**{_tit}**")
                            if _dim not in _cm.columns:
                                st.info(f"Columna «{_dim}» no disponible.")
                                continue
                            _gdf = (_cm.assign(**{_dim: _cm[_dim].fillna('(sin dato)')})
                                       .groupby([_dim, '_estado_vs']).size()
                                       .reset_index(name='CTs'))
                            if _gdf.empty:
                                st.info("Sin datos en el filtro actual.")
                                continue
                            _orden_vs = (_gdf[_gdf['_estado_vs'] != 'Completa']
                                         .groupby(_dim)['CTs'].sum()
                                         .reindex(_gdf[_dim].unique(), fill_value=0)
                                         .fillna(0).sort_values(ascending=True))
                            _fig_vs = px.bar(
                                _gdf, x='CTs', y=_dim, color='_estado_vs', orientation='h',
                                height=max(320, 26 * len(_orden_vs) + 140),
                                color_discrete_map=_MAPA_VS,
                                category_orders={'_estado_vs': ['Sin ingresar', 'Incompleta',
                                                                'Completa']})
                            _fig_vs.update_layout(
                                barmode='stack',
                                yaxis={'categoryorder': 'array',
                                       'categoryarray': _orden_vs.index.tolist()},
                                margin=dict(l=0, t=10, b=0, r=0),
                                xaxis_title='CTs con condición C o M', yaxis_title='',
                                legend=dict(orientation='h', y=-0.12, title=''))
                            st.plotly_chart(_fig_vs, use_container_width=True)

                # ── e) TODO LO REALIZADO EN VIGILANCIA DE LA SALUD ─────────────
                st.divider()
                st.markdown("#### 🗂️ Todo lo realizado en vigilancia de la salud")

                _fexam_s = (pd.to_datetime(_vs[_VS_FEXAM], errors='coerce')
                            if _VS_FEXAM in _vs.columns else pd.Series(pd.NaT, index=_vs.index))
                _feval_s = (pd.to_datetime(_vs[_VS_FEVAL], errors='coerce')
                            if _VS_FEVAL in _vs.columns else pd.Series(pd.NaT, index=_vs.index))
                # "Realizado" = el CT está en vigilancia de la salud: tiene nómina VS o
                # alguna fecha de vigilancia informada. Incluye programados y NO
                # programados, y las cuatro condiciones (C, M, A y sin condición).
                _vs_hecho = _vs_con_nom | _fexam_s.notna() | _feval_s.notna()
                _real = _vs[_vs_hecho].copy()

                if _real.empty:
                    st.info("ℹ️ No hay centros de trabajo con actividad de vigilancia de la "
                            "salud en el filtro actual.")
                else:
                    if 'Es_Programado' in _real.columns:
                        _real['Programado'] = _real['Es_Programado'].map(
                            {True: 'Programado', False: 'No Programado'}).fillna('')
                        _n_prog   = int((_real['Es_Programado'] == True).sum())
                        _n_noprog = int((_real['Es_Programado'] == False).sum())
                    else:
                        _real['Programado'] = ''
                        _n_prog = _n_noprog = 0

                    _r_deben = int(_real['_deben'].sum()) if _real['_deben'].notna().any() else 0
                    _r_evalu = int(_real['_evalu'].sum()) if _real['_evalu'].notna().any() else 0
                    _r_cob   = (_r_evalu / _r_deben * 100) if _r_deben > 0 else 0.0

                    _r1, _r2, _r3, _r4 = st.columns(4)
                    _r1.metric("CT en vigilancia de la salud", f"{len(_real):,}")
                    _r2.metric("Del plan", f"{_n_prog:,}", "programados", delta_color="off")
                    _r3.metric("Fuera del plan", f"{_n_noprog:,}", "no programados",
                               delta_color="off")
                    _r4.metric("Trabajadores evaluados", f"{_r_evalu:,}",
                               f"de {_r_deben:,} · {_r_cob:.1f}%", delta_color="off")

                    _cnt_cond = _real['_cond'].value_counts()
                    st.caption(
                        "Inventario completo de la vigilancia de la salud: **programados y no "
                        "programados en una sola tabla**, con todas las condiciones. Ordenado por "
                        "la vigilancia más reciente arriba (los CT sin fecha quedan al final). "
                        "Condición: "
                        + " · ".join(
                            f"**{int(_cnt_cond.get(_k, 0))}** {_lbl}"
                            for _k, _lbl in (('C', 'Críticas'), ('M', 'No Críticas'),
                                             ('A', 'Aceptables'), ('', 'sin condición registrada'))
                        )
                    )

                    _real['Condición'] = _real['_cond'].map(
                        {'C': 'C · Crítica', 'M': 'M · No Crítica', 'A': 'A · Aceptable'}
                    ).fillna('(sin condición)')
                    _real['Estado VS'] = _real['_estado_vs']
                    _real['Total deben']     = _real['_deben']
                    _real['Total evaluados'] = _real['_evalu']
                    _real['Brecha']          = _real['_brecha']
                    _real['Cobertura %'] = ((_real['_evalu'].fillna(0) / _real['_deben'] * 100)
                                            .where(_real['_deben'] > 0).round(1))
                    for _orig, _dest in ((_VS_DH, 'Deben H'), (_VS_DM, 'Deben M'),
                                         (_VS_EH, 'Evaluados H'), (_VS_EM, 'Evaluados M')):
                        _real[_dest] = _real[_orig]
                    _real['Fecha Ident. Avanzada'] = (
                        _vs_fia.reindex(_real.index).dt.strftime('%d-%m-%Y').fillna(''))
                    _real['Fecha última VS (examen)'] = (
                        _fexam_s.reindex(_real.index).dt.strftime('%d-%m-%Y').fillna(''))
                    _real['Fecha últ. evaluación VS'] = (
                        _feval_s.reindex(_real.index).dt.strftime('%d-%m-%Y').fillna(''))

                    # Orden: la vigilancia más reciente (examen o evaluación) primero.
                    _real['_ult_vs'] = pd.concat(
                        [_fexam_s.reindex(_real.index), _feval_s.reindex(_real.index)],
                        axis=1).max(axis=1)
                    _real = _real.sort_values('_ult_vs', ascending=False, na_position='last')

                    _cols_real = [c for c in [
                        'Programado', 'Región', 'Ergonomo', 'Nombre Empleador', 'ID-CT',
                        'Nombre CT', 'Comuna CT', 'Estado Centro de Trabajo',
                        'Condición', 'Fecha Ident. Avanzada', 'Estado VS',
                        'Deben H', 'Deben M', 'Total deben',
                        'Evaluados H', 'Evaluados M', 'Total evaluados',
                        'Brecha', 'Cobertura %',
                        'Fecha última VS (examen)', 'Fecha últ. evaluación VS',
                    ] if c in _real.columns]
                    _real_disp = _real[_cols_real].copy()
                    for _icol in ('Deben H', 'Deben M', 'Total deben', 'Evaluados H',
                                  'Evaluados M', 'Total evaluados', 'Brecha'):
                        if _icol in _real_disp.columns:
                            _real_disp[_icol] = _real_disp[_icol].astype('Int64')

                    def _hl_real(row):
                        """Rojo suave: vigilancia incompleta en un CT Crítico."""
                        _rojo = (str(row.get('Condición', '')).startswith('C')
                                 and row.get('Estado VS') == 'Incompleta')
                        return ['background-color: #fdecea' if _rojo else ''] * len(row)

                    st.dataframe(_real_disp.style.apply(_hl_real, axis=1),
                                 use_container_width=True, hide_index=True)

                    _buf_real = io.BytesIO()
                    with pd.ExcelWriter(_buf_real, engine='openpyxl') as _w:
                        _real_disp.to_excel(_w, index=False, sheet_name='VS_Realizado')
                    st.download_button(
                        "📥 Descargar todo lo realizado en VS (Excel)", data=_buf_real.getvalue(),
                        file_name=f'vigilancia_salud_realizado_{datetime.now().strftime("%d-%m-%Y")}.xlsx',
                        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                        key='download_vs_realizado')

    # ── TAB: EVOLUCIÓN E INDICADORES POR PROFESIONAL ──────────────────────────
    with tab_ind:
        if df_seg.empty:
            st.info("ℹ️ Se requiere data de seguimiento para calcular indicadores por profesional.")
        else:
            st.subheader("👨‍⚕️ Evolución e Indicadores por Profesional — TMERT 2026")

            COLS_PILAR = [
                'Pilar 1 - Difusión', 'Pilar 2 - Capacitación',
                'Pilar 3 - Diseño Cap Pract', 'Pilar 4 - Prescripción Caract',
                'Pilar 5 - Seguimiento'
            ]

            def _build_ind(df_src, modo='programado'):
                """Construye tabla de indicadores agrupada por ergónomo.
                modo: 'programado' (Es_Programado=True), 'no_programado' (=False),
                       o 'todos' (sin filtro)."""
                if 'Es_Programado' in df_src.columns and modo == 'programado':
                    _df_plan = df_src[df_src['Es_Programado'] == True].copy()
                elif 'Es_Programado' in df_src.columns and modo == 'no_programado':
                    _df_plan = df_src[df_src['Es_Programado'] == False].copy()
                else:
                    _df_plan = df_src.copy()

                if _df_plan.empty:
                    return pd.DataFrame()

                # Rellenar Ergonomo vacío (típico en no programados) sin tocar el dato origen
                _df_plan = _df_plan.copy()
                _df_plan['Ergonomo'] = (
                    _df_plan['Ergonomo'].astype(str).str.strip()
                    .replace({'': 'Sin asignar', 'nan': 'Sin asignar', 'None': 'Sin asignar'})
                    .fillna('Sin asignar')
                )

                cols_p = [c for c in COLS_PILAR if c in _df_plan.columns]
                ind = pd.DataFrame()
                ind['CTs Asignados'] = _df_plan.groupby('Ergonomo').size()
                if cols_p:
                    _any_at = _df_plan[cols_p].any(axis=1)
                    ind['Con alguna AT'] = _any_at.groupby(_df_plan['Ergonomo']).sum()
                else:
                    ind['Con alguna AT'] = 0
                if 'Meta 5 Cumplida' in _df_plan.columns:
                    ind['Meta 5'] = _df_plan.groupby('Ergonomo')['Meta 5 Cumplida'].sum()
                else:
                    ind['Meta 5'] = 0
                ind = ind.fillna(0)
                ind = ind.reset_index().rename(columns={'Ergonomo': 'Ergónomo'})
                ind['% Inicio'] = (ind['Con alguna AT'] / ind['CTs Asignados'] * 100).round(1)
                ind['% Meta 5'] = (ind['Meta 5'] / ind['CTs Asignados'] * 100).round(1)
                ind['Eficiencia'] = (
                    ind['Meta 5'] / ind['Con alguna AT'].replace(0, float('nan')) * 100
                ).round(1).fillna(0)
                # Aporte %: contribución del profesional sobre total de actividades del equipo
                _tot_at = int(ind['Con alguna AT'].sum())
                if _tot_at > 0:
                    ind['Aporte %'] = (ind['Con alguna AT'] / _tot_at * 100).round(1)
                else:
                    ind['Aporte %'] = 0.0
                # Aporte Meta 5 %: contribución del profesional sobre el total de Meta 5 del equipo
                _tot_meta5 = int(ind['Meta 5'].sum())
                if _tot_meta5 > 0:
                    ind['Aporte Meta 5 %'] = (ind['Meta 5'] / _tot_meta5 * 100).round(1)
                else:
                    ind['Aporte Meta 5 %'] = 0.0
                return ind

            # IDs con denuncia EP: desde df_raw (plan base, siempre tiene folios).
            _ep_ct_ids = set()
            if 'ID-CT' in df_raw.columns and 'Tiene EP' in df_raw.columns:
                _ep_ct_ids = set(
                    df_raw.loc[df_raw['Tiene EP'] == True, 'ID-CT']
                          .astype(str).str.upper().str.strip()
                )

            # Base para Tab 4: filtros del sidebar + filtro EP cuando el toggle está activo.
            _df_ind_total = df_seg_raw.copy() if not df_seg_raw.empty else pd.DataFrame()
            if not _df_ind_total.empty and 'ID-CT' in _df_ind_total.columns:
                # Normalizar ID-CT del seguimiento una vez (upper+strip) para match consistente
                _df_ind_total['_ID_NORM'] = (
                    _df_ind_total['ID-CT'].astype(str).str.upper().str.strip()
                )
            # Tarea 3: dueño de la asesoría en no programados (antes de filtrar/agrupar)
            _df_ind_total = rellenar_dueno_noprog(_df_ind_total)
            if not _df_ind_total.empty:
                for _k, _col_f, _all_v in _defs_t:
                    _v = st.session_state.get(_k, _all_v)
                    if _v != _all_v and _col_f in _df_ind_total.columns:
                        _df_ind_total = _df_ind_total[_df_ind_total[_col_f] == _v]
                # Aplicar filtro EP si toggle activo
                if solo_ep and _ep_ct_ids and '_ID_NORM' in _df_ind_total.columns:
                    _df_ind_total = _df_ind_total[
                        _df_ind_total['_ID_NORM'].isin(_ep_ct_ids)
                    ].copy()
                # Tarea 1: filtro "solo centros de trabajo activos"
                if solo_activos and 'Estado Centro de Trabajo' in _df_ind_total.columns:
                    _df_ind_total = _df_ind_total[
                        es_ct_activo(_df_ind_total['Estado Centro de Trabajo'])
                    ].copy()

            # Foco EP como subconjunto (solo si solo_ep=False)
            if not solo_ep and _ep_ct_ids and not _df_ind_total.empty and '_ID_NORM' in _df_ind_total.columns:
                _df_ind_ep = _df_ind_total[
                    _df_ind_total['_ID_NORM'].isin(_ep_ct_ids)
                ].copy()
            else:
                _df_ind_ep = pd.DataFrame()

            # Diagnóstico (visible solo si toggle EP está activo) para ayudar a depurar match
            if solo_ep:
                _seg_ids = set(
                    df_seg_raw['ID-CT'].astype(str).str.upper().str.strip()
                ) if 'ID-CT' in df_seg_raw.columns else set()
                _matches = len(_ep_ct_ids & _seg_ids)
                st.caption(
                    f"🔎 Diagnóstico EP — IDs EP en plan: {len(_ep_ct_ids)} · "
                    f"IDs en seguimiento: {len(_seg_ids)} · "
                    f"Coincidencias: {_matches} · "
                    f"Filas tras filtro EP: {len(_df_ind_total)}"
                )

            # Promedio del equipo sobre el total programado (sin filtro EP).
            # Refleja el toggle "solo CT activos" para usar el mismo universo que las
            # grillas; NO aplica los selectores de foco (ergónomo, región…) para que el
            # baseline siga siendo "del equipo" y vs. Promedio tenga sentido.
            _df_prom = df_seg_raw.copy() if not df_seg_raw.empty else pd.DataFrame()
            if solo_activos and not _df_prom.empty and 'Estado Centro de Trabajo' in _df_prom.columns:
                _df_prom = _df_prom[es_ct_activo(_df_prom['Estado Centro de Trabajo'])]
            ind_todos  = _build_ind(_df_prom, modo='programado') if not _df_prom.empty else pd.DataFrame()
            prom_meta5 = ind_todos['% Meta 5'].mean() if not ind_todos.empty else 0

            # Pace: distribuir el total de CTs en 24 meses (Ene 2025 – Dic 2026)
            _inicio_plan    = pd.Timestamp('2025-01-01')
            _hoy_calc       = pd.Timestamp.today().normalize()
            _months_elapsed = min(
                max((_hoy_calc.year - _inicio_plan.year) * 12
                    + (_hoy_calc.month - _inicio_plan.month) + 1, 1),
                24
            )

            def _build_full_ind(df_base, modo):
                """Arma indicadores con vs. Promedio, Esperado, vs. Pace y EP."""
                _ind = _build_ind(df_base, modo=modo)
                if _ind.empty:
                    return _ind
                _ind['vs. Promedio (pp)'] = (_ind['% Meta 5'] - prom_meta5).round(1)
                if modo == 'programado':
                    _ind['Esperado'] = (_ind['CTs Asignados'] * _months_elapsed / 24).round(1)
                    _ind['vs. Pace'] = (_ind['Con alguna AT'] - _ind['Esperado']).round(1)
                # Columnas EP (solo si _df_ind_ep tiene datos y solo_ep=False)
                if not _df_ind_ep.empty:
                    _ind_ep = _build_ind(_df_ind_ep, modo=modo)
                    if not _ind_ep.empty:
                        _ep_sub = _ind_ep[
                            ['Ergónomo'] + [c for c in ['CTs Asignados', 'Meta 5', 'Con alguna AT']
                                             if c in _ind_ep.columns]
                        ].rename(columns={
                            'CTs Asignados': 'CTs EP',
                            'Meta 5':        'Meta5 EP',
                            'Con alguna AT': 'Con AT EP',
                        })
                        _ep_sub['% Meta5 EP'] = (
                            _ep_sub['Meta5 EP'] / _ep_sub['CTs EP'].replace(0, float('nan')) * 100
                        ).round(1).fillna(0)
                        _ind = _ind.merge(_ep_sub, on='Ergónomo', how='left')
                        for _ec in ['CTs EP', 'Meta5 EP', 'Con AT EP', '% Meta5 EP']:
                            if _ec in _ind.columns:
                                _ind[_ec] = _ind[_ec].fillna(0)
                return _ind

            ind        = _build_full_ind(_df_ind_total, modo='programado')
            ind_noprog = _build_full_ind(_df_ind_total, modo='no_programado')

            # ── NIVEL 1: Tabla resumen ─────────────────────────────────────
            st.markdown("#### 📋 Resumen por Profesional")
            st.caption(
                f"Promedio del equipo: **{prom_meta5:.1f}%** Meta 5"
                f"{' · solo CT activos' if solo_activos else ''} | "
                f"Pace = distribución lineal en 24 meses (Ene 2025–Dic 2026), "
                f"mes actual = **{_months_elapsed}** de 24"
            )

            def _color_delta(val):
                if val > 0:
                    return 'color: #2e7d32; font-weight: bold'
                elif val < 0:
                    return 'color: #c62828; font-weight: bold'
                return ''

            st.info(
                "**Columnas TOTAL**: sobre todos los CTs asignados del profesional.  \n"
                "**Columnas EP**: solo sobre CTs con denuncias de EP.  \n"
                "**Aporte Meta 5 %**: contribución del profesional sobre el total de Meta 5 cumplidas del equipo.  \n"
                "**Aporte %**: contribución del profesional sobre el total de actividades realizadas (Con alguna AT) del equipo.  \n"
                "**Esperado**: CTs que debería tener con AT según pace lineal.  \n"
                "**vs. Pace**: realizados menos esperados (positivo = adelantado).  \n"
                "**Eficiencia**: de los que iniciaron, cuántos llegaron a Meta 5."
            )

            _fmt = {
                '% Inicio':          '{:.1f}%',
                '% Meta 5':          '{:.1f}%',
                'Aporte Meta 5 %':   '{:.1f}%',
                '% Meta5 EP':        '{:.1f}%',
                'vs. Promedio (pp)': '{:+.1f}',
                'Eficiencia':        '{:.1f}%',
                'Esperado':          '{:.1f}',
                'vs. Pace':          '{:+.1f}',
                'Aporte %':          '{:.1f}%',
            }

            def _render_grilla(_ind_df, _incluye_pace=True):
                if _ind_df.empty:
                    st.caption("Sin datos.")
                    return
                _tiene_ep = 'CTs EP' in _ind_df.columns
                _cols = [
                    'Ergónomo',
                    'CTs Asignados', 'Meta 5', '% Meta 5', 'Aporte Meta 5 %', 'Con alguna AT', '% Inicio',
                    'Aporte %', 'Eficiencia',
                ]
                if _tiene_ep:
                    _cols += ['CTs EP', 'Meta5 EP', '% Meta5 EP', 'Con AT EP']
                if _incluye_pace:
                    _cols += ['Esperado', 'vs. Pace']
                _cols += ['vs. Promedio (pp)']
                _disp = _ind_df[[c for c in _cols if c in _ind_df.columns]] \
                            .sort_values('% Meta 5', ascending=False)
                _styled = [c for c in ['vs. Pace', 'vs. Promedio (pp)'] if c in _disp.columns]
                st.dataframe(
                    _disp.style.map(_color_delta, subset=_styled).format(_fmt),
                    use_container_width=True, hide_index=True
                )

            _n_prog   = int((_df_ind_total['Es_Programado'] == True).sum()) \
                            if 'Es_Programado' in _df_ind_total.columns else len(_df_ind_total)
            _n_noprog = int((_df_ind_total['Es_Programado'] == False).sum()) \
                            if 'Es_Programado' in _df_ind_total.columns else 0

            st.markdown(f"##### 1) Programadas (n={_n_prog:,})")
            _render_grilla(ind, _incluye_pace=True)

            st.markdown(f"##### 2) No Programadas (n={_n_noprog:,})")
            st.caption("Actividades realizadas fuera del plan que igualmente suman a la meta.")
            _render_grilla(ind_noprog, _incluye_pace=False)

            st.divider()

            # ── CURVA DE AVANCE ACUMULADO vs. PACE ESPERADO ───────────────
            st.markdown("#### 📈 Curva de Avance Acumulado vs. Pace Esperado")
            st.caption(
                "Reconstruida con las **fechas reales de ejecución** de SIGECO (resolución diaria, "
                "desde ene-2025). Pilares 1–5 y «4 Pilares completos» son **exactos**. El **P5 "
                "Seguimiento** usa las fechas de seguimiento de prescripción (basta **cualquiera "
                "de los dos** seguimientos; se toma la más temprana). La curva **Meta 5 cumplida** "
                "no redefine la meta: son los mismos CTs de «Meta 5 Cumplida», solo ubicados en "
                "el tiempo, y es un **piso** porque a algunos les falta alguna fecha. "
                "Respeta los filtros del panel lateral."
            )

            _cols_fp = {p: c for p, c in PILAR_FECHA_REAL.items() if c in df_seg.columns}
            _total_evo = len(df_seg)

            if not _cols_fp or _total_evo == 0:
                st.info("No hay columnas de fecha real de pilares para construir la curva.")
            else:
                _4p_fecha = cuatro_pilares_fecha(df_seg)
                _m5_fecha  = meta5_fecha(df_seg)
                _meta5_real = int(df_seg['Meta 5 Cumplida'].sum()) \
                    if 'Meta 5 Cumplida' in df_seg.columns else 0
                _hoy_evo = pd.Timestamp.today().normalize()

                _ver_pct = st.toggle("Ver en % del plan", value=False, key="evo_pct")
                _idx = pd.date_range(EVO_INICIO, _hoy_evo, freq='D')
                _div = (_total_evo / 100.0) if (_ver_pct and _total_evo) else 1.0

                _colores_p = {
                    'Pilar 1 - Difusión':            '#7B2D8B',
                    'Pilar 2 - Capacitación':        '#E67E22',
                    'Pilar 3 - Diseño Cap Pract':    '#1A936F',
                    'Pilar 4 - Prescripción Caract': '#2E86AB',
                }

                _fig_evo = go.Figure()
                for _p, _c in _cols_fp.items():
                    _y = serie_acumulada(df_seg[_c], _idx) / _div
                    _fig_evo.add_trace(go.Scatter(
                        x=_idx, y=_y, name=PILAR_LABEL_CORTO.get(_p, _p),
                        mode='lines', line=dict(width=1.8, color=_colores_p.get(_p)),
                        hovertemplate='%{x|%d-%m-%Y}<br>%{y:.0f}<extra>' + PILAR_LABEL_CORTO.get(_p, _p) + '</extra>',
                    ))

                # P5 Seguimiento (fechas de seguimiento de prescripción)
                if COL_FECHA_SEGUIMIENTO in df_seg.columns:
                    _y_p5 = serie_acumulada(df_seg[COL_FECHA_SEGUIMIENTO], _idx) / _div
                    _fig_evo.add_trace(go.Scatter(
                        x=_idx, y=_y_p5, name='P5 Seguimiento',
                        mode='lines', line=dict(width=1.8, color='#C0392B'),
                        hovertemplate='%{x|%d-%m-%Y}<br>%{y:.0f}<extra>P5 Seguimiento</extra>',
                    ))

                _y_4p = serie_acumulada(_4p_fecha, _idx) / _div
                _fig_evo.add_trace(go.Scatter(
                    x=_idx, y=_y_4p, name='4 Pilares completos',
                    mode='lines', line=dict(width=3.2, color='#4F0B7B'),
                    hovertemplate='%{x|%d-%m-%Y}<br>%{y:.0f}<extra>4 Pilares completos</extra>',
                ))

                _n_m5_fechable = int(_m5_fecha.notna().sum())
                if _n_m5_fechable:
                    _y_m5 = serie_acumulada(_m5_fecha, _idx) / _div
                    _fig_evo.add_trace(go.Scatter(
                        x=_idx, y=_y_m5, name='Meta 5 cumplida',
                        mode='lines', line=dict(width=3.2, color='#B8860B'),
                        hovertemplate='%{x|%d-%m-%Y}<br>%{y:.0f}<extra>Meta 5 cumplida</extra>',
                    ))

                _secs = np.asarray((_idx - EVO_INICIO).total_seconds())
                _frac = np.clip(_secs / (EVO_FIN_PLAN - EVO_INICIO).total_seconds(), 0, 1)
                _y_pace = (_total_evo * _frac) / _div
                _fig_evo.add_trace(go.Scatter(
                    x=_idx, y=_y_pace, name='Pace esperado',
                    mode='lines', line=dict(width=2, color='#E41395', dash='dash'),
                    hovertemplate='%{x|%d-%m-%Y}<br>%{y:.0f}<extra>Pace esperado</extra>',
                ))

                # Línea vertical en HOY
                _fig_evo.add_shape(
                    type='line', xref='x', yref='paper',
                    x0=_hoy_evo, x1=_hoy_evo, y0=0, y1=1,
                    line=dict(color='gray', width=1.5, dash='dot'),
                )
                _fig_evo.add_annotation(
                    x=_hoy_evo, y=1, xref='x', yref='paper',
                    text='Hoy', showarrow=False,
                    xanchor='left', yanchor='bottom', font=dict(color='gray', size=11),
                )

                _fig_evo.update_layout(
                    xaxis_title='Fecha',
                    yaxis_title='% del plan' if _ver_pct else 'CTs acumulados',
                    plot_bgcolor='white', paper_bgcolor='white',
                    legend=dict(orientation='h', yanchor='bottom', y=1.02, x=0),
                    height=460, margin=dict(l=10, r=20, t=30, b=40),
                    hovermode='x unified',
                )
                st.plotly_chart(_fig_evo, use_container_width=True)
                # El conteo oficial de Meta 5 no cambia: es la columna 'Meta 5 Cumplida'.
                # La curva es un subconjunto suyo (los que además tienen los 5 pilares
                # fechados), nunca un universo distinto.
                _sin_fecha_m5 = max(0, _meta5_real - _n_m5_fechable)
                st.caption(
                    f"📌 **Meta 5 hoy (real, con Seguimiento): {_meta5_real:,} CTs** "
                    f"({_meta5_real / _total_evo * 100:.1f}% del plan). "
                    f"La curva «Meta 5 cumplida» ubica en el tiempo **{_n_m5_fechable:,}** de "
                    f"ellos; los **{_sin_fecha_m5:,}** restantes cumplen la meta pero no tienen "
                    f"alguna de las 5 fechas informada, por eso no aparecen en la curva."
                )

            st.divider()

            # ── RITMO ACUMULADO: realizado vs pace esperado ────────────────
            st.markdown("#### 📈 Ritmo Acumulado — Realizado vs. Pace Esperado (2025–2026)")
            st.caption(
                "La línea punteada muestra el ritmo ideal distribuyendo los CTs en 24 meses. "
                "La línea sólida es el acumulado real de CTs con al menos una AT registrada."
            )

            _df_pace = _df_ind_total.copy()
            _n_pace = int((_df_pace['Es_Programado'] == True).sum()) if 'Es_Programado' in _df_pace.columns else len(_df_pace)

            if _n_pace > 0 and 'Fecha real AT' in _df_pace.columns:
                _meses = pd.date_range(start='2025-01-01', end='2026-12-01', freq='MS')
                _pace_mensual = _n_pace / 24

                _fechas_real = pd.to_datetime(_df_pace['Fecha real AT'], errors='coerce').dropna()
                _fechas_real = _fechas_real[
                    (_fechas_real >= '2025-01-01') & (_fechas_real <= '2026-12-31')
                ]

                _acum_real = []
                for _m in _meses:
                    _fin = _m + pd.offsets.MonthEnd(0)
                    _acum_real.append(int((_fechas_real <= _fin).sum()))

                _df_chart = pd.DataFrame({
                    'Mes':      _meses,
                    'Esperado': [(i + 1) * _pace_mensual for i in range(24)],
                    'Realizado': _acum_real,
                })
                _df_chart['Mes_str'] = _df_chart['Mes'].dt.strftime('%b %Y')

                _fig_pace = go.Figure()
                _fig_pace.add_trace(go.Scatter(
                    x=_df_chart['Mes_str'], y=_df_chart['Esperado'],
                    name='Pace esperado', mode='lines',
                    line=dict(color='#E41395', width=2, dash='dash'),
                ))
                _fig_pace.add_trace(go.Scatter(
                    x=_df_chart['Mes_str'], y=_df_chart['Realizado'],
                    name='Acumulado real', mode='lines+markers',
                    line=dict(color='#4F0B7B', width=2.5),
                    marker=dict(size=6),
                ))

                # Línea vertical en el mes actual (add_shape evita el bug de Plotly
                # con ejes categóricos que falla al usar add_vline + annotation)
                _hoy_str = _hoy_calc.strftime('%b %Y')
                if _hoy_str in _df_chart['Mes_str'].values:
                    _fig_pace.add_shape(
                        type='line', xref='x', yref='paper',
                        x0=_hoy_str, x1=_hoy_str, y0=0, y1=1,
                        line=dict(color='gray', width=1.5, dash='dot')
                    )
                    _fig_pace.add_annotation(
                        x=_hoy_str, y=1, xref='x', yref='paper',
                        text='Hoy', showarrow=False,
                        xanchor='left', yanchor='bottom',
                        font=dict(color='gray', size=11)
                    )

                _fig_pace.update_layout(
                    xaxis_title='Mes', yaxis_title='CTs acumulados',
                    plot_bgcolor='white', paper_bgcolor='white',
                    legend=dict(orientation='h', yanchor='bottom', y=1.02),
                    height=420, xaxis=dict(tickangle=-45),
                    margin=dict(l=10, r=20, t=30, b=80),
                )
                st.plotly_chart(_fig_pace, use_container_width=True)

                # Métricas de pace al mes actual
                _idx_hoy = next(
                    (i for i, m in enumerate(_meses)
                     if m.year == _hoy_calc.year and m.month == _hoy_calc.month),
                    None
                )
                if _idx_hoy is not None:
                    _esp_hoy = _df_chart['Esperado'].iloc[_idx_hoy]
                    _real_hoy = _df_chart['Realizado'].iloc[_idx_hoy]
                    _delta = int(_real_hoy - _esp_hoy)
                    _pp1, _pp2, _pp3 = st.columns(3)
                    _pp1.metric("CTs con AT a la fecha", f"{int(_real_hoy):,}")
                    _pp2.metric("Pace esperado a la fecha", f"{int(_esp_hoy):,}")
                    _pp3.metric("Diferencia vs. pace", f"{_delta:+,}",
                                delta_color="normal" if _delta >= 0 else "inverse")
            else:
                st.info("No hay datos de AT real para calcular el ritmo acumulado.")

            st.divider()

            # ── NIVEL 2: Desglose por pilar (separado: Programado vs. No Programado) ──
            st.markdown("#### 🧱 Avance por Pilar")
            cols_p = [c for c in COLS_PILAR if c in df_seg.columns]

            def _pilar_lbl(col):
                """Etiqueta corta del pilar para encabezado de columna."""
                return (col.replace('Pilar ', 'P').replace(' - ', ': ')
                           .replace(' Cap Pract', ' Cap').replace(' Caract', ''))

            def _tabla_pilar(df_subset, cols_p, valor='n'):
                """Pivot NUMÉRICO por ergónomo (ordenable correctamente al hacer clic).
                valor='n'   -> N° de actividades por pilar
                valor='pct' -> % sobre el total de CTs del profesional"""
                filas = []
                for ergo, grp in df_subset.groupby('Ergonomo'):
                    total = len(grp)
                    fila = {'Ergónomo': ergo, 'CTs': total}
                    for col in cols_p:
                        n = int(grp[col].sum()) if col in grp.columns else 0
                        fila[_pilar_lbl(col)] = (
                            round(n / total * 100, 1) if (valor == 'pct' and total > 0)
                            else (0.0 if valor == 'pct' else n)
                        )
                    filas.append(fila)
                return pd.DataFrame(filas) if filas else pd.DataFrame()

            if cols_p and not df_seg.empty:
                _modo_val = st.radio(
                    "Mostrar:",
                    ["N° de actividades", "% del total del profesional"],
                    horizontal=True, key="pilar_modo",
                )
                _valor = 'pct' if _modo_val.startswith('%') else 'n'
                _pilar_lbls = [_pilar_lbl(c) for c in cols_p]
                _fmt_col = "%.1f%%" if _valor == 'pct' else "%d"
                _colcfg = {
                    c: st.column_config.NumberColumn(c, format=_fmt_col) for c in _pilar_lbls
                }

                def _render_pilar(df_subset, titulo):
                    st.markdown(titulo)
                    _t = _tabla_pilar(df_subset, cols_p, valor=_valor)
                    if _t.empty:
                        st.caption("Sin datos.")
                        return
                    # Orden por defecto: mayor a menor según el primer pilar (P1: Difusión)
                    _sort_col = _pilar_lbls[0] if _pilar_lbls else 'CTs'
                    _t = _t.sort_values(_sort_col, ascending=False)
                    st.dataframe(
                        _t, use_container_width=True, hide_index=True,
                        column_config=_colcfg,
                    )

                if 'Es_Programado' in df_seg.columns:
                    df_prog_only = df_seg[df_seg['Es_Programado'] == True]
                    df_nopr_only = df_seg[df_seg['Es_Programado'] == False]
                else:
                    df_prog_only, df_nopr_only = df_seg, df_seg.iloc[0:0]

                st.caption(
                    "Ordenado de mayor a menor por P1: Difusión. Haz clic en cualquier "
                    "encabezado para reordenar (ahora ordena por número, no por texto)."
                )
                _render_pilar(df_prog_only, f"**📋 Programado** — {len(df_prog_only):,} CTs")
                _render_pilar(df_nopr_only, f"**📋 No Programado** — {len(df_nopr_only):,} CTs")
            else:
                st.info("No hay columnas de pilares disponibles en los datos.")

    # ── TAB 2: ANÁLISIS DE DENUNCIAS EP ───────────────────────────────────────
    with tab2:
        st.header("🔍 Análisis de Denuncias EP")
        df_ep = df[df['Tiene EP']]

        if len(df_ep) > 0:

            # Precomputar rankings
            rank_seg_gen  = obtener_ranking_limpio(df_ep, 'segmentos', separadores_extra=[" "])
            rank_diag_gen = obtener_ranking_limpio(df_ep, 'diagnosticos')
            rank_emp_gen  = folios_por_empresa(df_ep)

            # ── Métricas resumen ──────────────────────────────────────────────
            n_folios_total = contar_folios_distintos(df_ep)
            n_emp_total    = df_ep['Nombre Empleador'].nunique()
            n_vis_total    = len(df_ep)
            ms1, ms2, ms3 = st.columns(3)
            ms1.metric("Folios EP distintos", f"{n_folios_total:,}")
            ms2.metric("Empresas con EP", f"{n_emp_total:,}")
            ms3.metric("Visitas con EP", f"{n_vis_total:,}")

            st.divider()

            # ── Top Diagnósticos + Top Empresas ──────────────────────────────
            col_g1, col_g2 = st.columns(2)
            with col_g1:
                if not rank_diag_gen.empty:
                    fig_diag = px.bar(
                        rank_diag_gen.head(10), x='Cantidad', y='Nombre', orientation='h',
                        color_discrete_sequence=['#E67E22'],
                        title="Top Diagnósticos en Denuncias EP"
                    )
                    fig_diag.update_layout(yaxis={'categoryorder': 'total ascending'}, margin=dict(l=0))
                    st.plotly_chart(fig_diag, use_container_width=True)
            with col_g2:
                if not rank_emp_gen.empty:
                    fig_emp = px.bar(
                        rank_emp_gen.head(10), x='Folios EP', y='Empresa', orientation='h',
                        color_discrete_sequence=['#1A936F'],
                        title="Top Empresas con Denuncias EP (por folios)"
                    )
                    fig_emp.update_layout(yaxis={'categoryorder': 'total ascending'}, margin=dict(l=0))
                    st.plotly_chart(fig_emp, use_container_width=True)

            st.divider()

            # ── Ranking EP: Treemap + Tabla ───────────────────────────────────
            st.subheader("📊 Ranking EP — Puestos de Trabajo y Tareas")
            st.caption(
                "El **Treemap** muestra el peso relativo de cada categoría por área. "
                "Las celdas **rojas** concentran el 80 % de los casos (principio vital)."
            )

            df_ep_pareto = df_ep.copy()
            rank_seg_p = obtener_ranking_limpio(df_ep_pareto, 'segmentos', separadores_extra=[" "])
            opciones_seg_p = ["Todos"] + (ordenar_segmentos(rank_seg_p['Nombre'].tolist())
                                          if not rank_seg_p.empty else [])
            cf1, cf2 = st.columns([2, 2])
            with cf1:
                seg_pareto = st.selectbox("Filtrar por Segmento Corporal:", opciones_seg_p, key="pareto_seg")
            with cf2:
                dimension = st.radio("Analizar por:",
                                     ["Ocupaciones (Puestos de Trabajo)", "Tareas"],
                                     horizontal=True, key="pareto_dim")

            if seg_pareto != "Todos":
                df_ep_pareto = df_ep_pareto[
                    df_ep_pareto['segmentos'].str.contains(seg_pareto, case=False, na=False, regex=False)
                ]

            if len(df_ep_pareto) > 0:
                col_p = 'ocupaciones' if "Ocupaciones" in dimension else 'tareas'
                sep_p = ' | '          if "Ocupaciones" in dimension else ','
                fig_p, df_p = grafico_pareto(df_ep_pareto, col_p, "Ranking EP",
                                         separador_secundario=sep_p)

                if fig_p and not df_p.empty:
                    # Métricas de concentración
                    pm1, pm2, pm3 = st.columns(3)
                    n_vital   = int(df_p['Vital'].sum())
                    n_total_p = len(df_p)
                    casos_vital = int(df_p[df_p['Vital']]['Cantidad'].sum())
                    pct_casos = round(casos_vital / df_p['Cantidad'].sum() * 100, 1) if df_p['Cantidad'].sum() > 0 else 0
                    pm1.metric("Categorías vitales 🔴", f"{n_vital} de {n_total_p}")
                    pm2.metric("% de categorías vitales", f"{round(n_vital/n_total_p*100,1)} %" if n_total_p > 0 else "0 %")
                    pm3.metric("Casos EP que concentran", f"{pct_casos} %")

                    st.plotly_chart(fig_p, use_container_width=True)

                    # Tabla detallada del ranking
                    st.markdown("#### 📋 Detalle del Ranking")
                    df_display = df_p[['Rank', 'Nombre', 'Cantidad', 'Pct', 'PctAcum', 'Vital']].copy()
                    df_display.columns = ['Rank', 'Nombre', 'Casos EP', '% del Total', '% Acumulado', 'Vital 🔴']
                    st.dataframe(
                        df_display,
                        use_container_width=True,
                        hide_index=True,
                        column_config={'Vital 🔴': st.column_config.CheckboxColumn("Vital 🔴")}
                    )
                else:
                    st.info("No hay datos suficientes para construir el ranking.")
            else:
                st.info(f"No hay registros EP con segmento «{seg_pareto}» para los filtros actuales.")

            st.divider()

            # ── EXPLORADOR: De lo General a lo Particular ─────────────────────
            st.subheader("🔍 Explorador por Segmento o Diagnóstico")
            st.caption(
                "Selecciona un segmento corporal o diagnóstico para ver en qué empresas, "
                "puestos de trabajo y tareas se concentra ese riesgo."
            )

            modo = st.radio(
                "Explorar por:", ["Segmento Corporal", "Diagnóstico"],
                horizontal=True, key="modo_explor"
            )
            col_explor = 'segmentos' if modo == "Segmento Corporal" else 'diagnosticos'
            rank_base   = rank_seg_gen if modo == "Segmento Corporal" else rank_diag_gen

            if not rank_base.empty:
                if modo == "Segmento Corporal":
                    _opts = ["Todos"] + ordenar_segmentos(rank_base['Nombre'].tolist())
                else:
                    _opts = ["Todos"] + sorted(rank_base['Nombre'].tolist())
                seleccion = st.selectbox(
                    f"Selecciona un {'segmento corporal' if modo == 'Segmento Corporal' else 'diagnóstico'}:",
                    _opts,
                    key="explorador_selector"
                )

                df_drill = df_ep.copy() if seleccion == "Todos" else df_ep[
                    df_ep[col_explor].str.contains(seleccion, case=False, na=False, regex=False)
                ]

                if len(df_drill) > 0:
                    n_folios_drill = contar_folios_distintos(df_drill)
                    n_emp_drill = df_drill['Nombre Empleador'].nunique()
                    label_sel = "todos los registros EP" if seleccion == "Todos" else f"«{seleccion}»"
                    st.markdown(
                        f"**{n_folios_drill} folio(s)** para {label_sel} "
                        f"— en **{n_emp_drill} empresa(s)** · {len(df_drill)} visita(s)"
                    )

                    col_d1, col_d2, col_d3 = st.columns(3)
                    with col_d1:
                        st.markdown("**🏢 Empresas**")
                        st.dataframe(folios_por_empresa(df_drill), use_container_width=True, hide_index=True)
                    with col_d2:
                        st.markdown("**💼 Puestos de Trabajo**")
                        r_ocup = obtener_ranking_limpio(
                            df_drill, 'ocupaciones', separador_secundario=' | '
                        ).rename(columns={'Nombre': 'Puesto', 'Cantidad': 'Casos'})
                        if not r_ocup.empty:
                            st.dataframe(r_ocup, use_container_width=True, hide_index=True)
                        else:
                            st.info("Sin datos de puestos.")
                    with col_d3:
                        st.markdown("**🛠️ Tareas Asociadas**")
                        r_tar = obtener_ranking_limpio(df_drill, 'tareas').rename(
                            columns={'Nombre': 'Tarea', 'Cantidad': 'Casos'}
                        )
                        if not r_tar.empty:
                            st.dataframe(r_tar, use_container_width=True, hide_index=True)
                        else:
                            st.info("Sin datos de tareas.")

                    st.markdown("#### 📋 Ver registros individuales")
                    cols_det = ['Nombre Empleador', 'Nombre CT', 'Región', 'Ergonomo',
                                'segmentos', 'ocupaciones', 'tareas', 'diagnosticos', 'folios',
                                'observaciones']
                    cols_det = [c for c in cols_det if c in df_drill.columns]
                    st.dataframe(df_drill[cols_det], use_container_width=True, hide_index=True)
                else:
                    st.info(f"No hay registros con «{seleccion}» para los filtros actuales.")
            else:
                st.info("No hay datos suficientes para el explorador.")

        else:
            st.warning("⚠️ No se encontraron registros con Denuncias de EP en el filtro actual.")

    # ── TAB PARETO EP (integrado en tab2) ────────────────────────────────────
    if False:
        st.header("📈 Análisis de Pareto — Intervención Preventiva")
        st.caption(
            "Identifica los puestos de trabajo y tareas que concentran el mayor número "
            "de denuncias EP (principio 80/20). Las barras **rojas** son las categorías "
            "vitales que acumulan hasta el 80 % de los casos."
        )

        df_ep_pareto = df[df['Tiene EP']]

        if len(df_ep_pareto) > 0:
            # ── Sub-filtro por segmento corporal ──────────────────────────────
            # separadores_extra=[" "] para obtener valores atómicos (ej. HOMBRO_DER, no combinaciones)
            rank_seg_p = obtener_ranking_limpio(df_ep_pareto, 'segmentos', separadores_extra=[" "])
            opciones_seg_p = ["Todos"] + (ordenar_segmentos(rank_seg_p['Nombre'].tolist())
                                          if not rank_seg_p.empty else [])

            cf1, cf2 = st.columns([2, 2])
            with cf1:
                seg_pareto = st.selectbox(
                    "Filtrar por Segmento Corporal:",
                    opciones_seg_p, key="pareto_seg"
                )
            with cf2:
                dimension = st.radio(
                    "Analizar por:",
                    ["Ocupaciones (Puestos de Trabajo)", "Tareas"],
                    horizontal=True, key="pareto_dim"
                )

            # Aplicar sub-filtro de segmento
            if seg_pareto != "Todos":
                df_ep_pareto = df_ep_pareto[
                    df_ep_pareto['segmentos'].str.contains(
                        seg_pareto, case=False, na=False, regex=False
                    )
                ]

            if len(df_ep_pareto) > 0:
                if "Ocupaciones" in dimension:
                    col_p, sep_p = 'ocupaciones', ' | '
                    titulo_p = "Pareto · Puestos de Trabajo con EP"
                else:
                    col_p, sep_p = 'tareas', ','
                    titulo_p = "Pareto · Tareas con EP"

                if seg_pareto != "Todos":
                    titulo_p += f" — {seg_pareto}"

                fig_p, df_p = grafico_pareto(df_ep_pareto, col_p, titulo_p,
                                             separador_secundario=sep_p)

                if fig_p:
                    st.plotly_chart(fig_p, use_container_width=True)

                    # Métricas de concentración
                    n_vital = int(df_p['Vital'].sum())
                    n_total = len(df_p)
                    pct_cat = round(n_vital / n_total * 100, 1) if n_total > 0 else 0
                    casos_vital = int(df_p[df_p['Vital']]['Cantidad'].sum())
                    pct_casos = round(casos_vital / df_p['Cantidad'].sum() * 100, 1) if df_p['Cantidad'].sum() > 0 else 0

                    pm1, pm2, pm3 = st.columns(3)
                    pm1.metric("Categorías vitales (🔴)", f"{n_vital} de {n_total}")
                    pm2.metric("% de categorías vitales", f"{pct_cat} %")
                    pm3.metric("Casos EP que concentran", f"{pct_casos} %")

                    # Tabla detallada
                    st.markdown("#### 📋 Detalle del Ranking")
                    df_display = df_p[['Rank', 'Nombre', 'Cantidad', 'Pct', 'PctAcum', 'Vital']].copy()
                    df_display.columns = ['Rank', 'Nombre', 'Casos EP', '% del Total', '% Acumulado', 'Vital 🔴']
                    st.dataframe(
                        df_display,
                        use_container_width=True,
                        hide_index=True,
                        column_config={
                            'Vital 🔴': st.column_config.CheckboxColumn("Vital 🔴")
                        }
                    )

                    # ── EXPLORADOR: De la prioridad al detalle ────────────────
                    st.divider()
                    st.subheader("🔍 Explorador: Investiga una Categoría del Pareto")
                    st.caption(
                        "Selecciona un segmento corporal o diagnóstico para ver en qué empresas, "
                        "puestos y tareas se concentra ese riesgo — usando el mismo subconjunto "
                        "de registros que el Pareto de arriba."
                    )

                    rank_diag_p = obtener_ranking_limpio(df_ep_pareto, 'diagnosticos')
                    modo_p = st.radio(
                        "Explorar por:", ["Segmento Corporal", "Diagnóstico"],
                        horizontal=True, key="modo_explor_pareto"
                    )
                    col_explor_p = 'segmentos' if modo_p == "Segmento Corporal" else 'diagnosticos'
                    rank_base_p  = rank_seg_p if modo_p == "Segmento Corporal" else rank_diag_p

                    if not rank_base_p.empty:
                        # Orden anatómico para segmentos; alfabético para diagnósticos
                        if modo_p == "Segmento Corporal":
                            _opts_p = ["Todos"] + ordenar_segmentos(rank_base_p['Nombre'].tolist())
                        else:
                            _opts_p = ["Todos"] + sorted(rank_base_p['Nombre'].tolist())
                        sel_p = st.selectbox(
                            f"Selecciona un {'segmento corporal' if modo_p == 'Segmento Corporal' else 'diagnóstico'}:",
                            _opts_p,
                            key="explorador_pareto_selector"
                        )

                        df_drill_p = df_ep_pareto.copy() if sel_p == "Todos" else df_ep_pareto[
                            df_ep_pareto[col_explor_p].str.contains(sel_p, case=False, na=False, regex=False)
                        ]

                        if len(df_drill_p) > 0:
                            n_f_p = contar_folios_distintos(df_drill_p)
                            n_e_p = df_drill_p['Nombre Empleador'].nunique()
                            label_p = "todos los registros EP" if sel_p == "Todos" else f"«{sel_p}»"
                            st.markdown(
                                f"**{n_f_p} folio(s)** para {label_p} "
                                f"— en **{n_e_p} empresa(s)** · {len(df_drill_p)} visita(s)"
                            )

                            pd1, pd2, pd3 = st.columns(3)
                            with pd1:
                                st.markdown("**🏢 Empresas**")
                                st.dataframe(folios_por_empresa(df_drill_p),
                                             use_container_width=True, hide_index=True)
                            with pd2:
                                st.markdown("**💼 Puestos de Trabajo**")
                                r_oc_p = obtener_ranking_limpio(
                                    df_drill_p, 'ocupaciones', separador_secundario=' | '
                                ).rename(columns={'Nombre': 'Puesto', 'Cantidad': 'Casos'})
                                if not r_oc_p.empty:
                                    st.dataframe(r_oc_p, use_container_width=True, hide_index=True)
                                else:
                                    st.info("Sin datos de puestos.")
                            with pd3:
                                st.markdown("**🛠️ Tareas Asociadas**")
                                r_ta_p = obtener_ranking_limpio(df_drill_p, 'tareas').rename(
                                    columns={'Nombre': 'Tarea', 'Cantidad': 'Casos'}
                                )
                                if not r_ta_p.empty:
                                    st.dataframe(r_ta_p, use_container_width=True, hide_index=True)
                                else:
                                    st.info("Sin datos de tareas.")

                            with st.expander("📋 Ver registros individuales"):
                                cols_d = ['Nombre Empleador', 'Nombre CT', 'Región', 'Ergonomo',
                                          'segmentos', 'ocupaciones', 'tareas', 'diagnosticos', 'folios']
                                cols_d = [c for c in cols_d if c in df_drill_p.columns]
                                st.dataframe(df_drill_p[cols_d], use_container_width=True, hide_index=True)
                        else:
                            st.info(f"No hay registros con «{sel_p}» en el subconjunto actual.")
                    else:
                        st.info("No hay datos suficientes para el explorador.")

                else:
                    st.info("No hay datos suficientes para construir el Pareto.")
            else:
                st.info(f"No hay registros EP con segmento «{seg_pareto}» para los filtros actuales.")
        else:
            st.warning("⚠️ No se encontraron registros con Denuncias de EP en el filtro actual.")

    # ── TAB PLANILLA DETALLADA (ELIMINADO) ───────────────────────────────────
    if False:
        st.subheader("Planificación Detallada 2026")
        col_fecha_display = 'Fecha Asistencia Técnica TMERT 2026*'
        cols_finales = [col_fecha_display, 'Región', 'Ergonomo', 'Nombre Empleador',
                        'Nombre CT', 'Tiene EP', 'folios', 'segmentos', 'tareas']
        cols_finales = [c for c in cols_finales if c in df.columns]

        df_tab3 = df[cols_finales].sort_values(col_fecha_display)

        st.dataframe(
            df_tab3,
            column_config={
                "Tiene EP": st.column_config.CheckboxColumn("🚨 EP"),
                col_fecha_display: st.column_config.DateColumn("Fecha"),
                "folios": st.column_config.TextColumn("Folios"),
                "segmentos": st.column_config.TextColumn("Segmentos"),
                "tareas": st.column_config.TextColumn("Tareas")
            },
            use_container_width=True,
            hide_index=True
        )

        buffer_tab3 = io.BytesIO()
        with pd.ExcelWriter(buffer_tab3, engine='openpyxl') as writer:
            df_tab3.to_excel(writer, index=False, sheet_name='Planilla_TMERT')
        st.download_button(
            label="📥 Descargar Planilla en Excel",
            data=buffer_tab3.getvalue(),
            file_name=f'planilla_tmert_{datetime.now().strftime("%d-%m-%Y")}.xlsx',
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            key='download_tab3'
        )

# ── 8. FOOTER ─────────────────────────────────────────────────────────────────
st.markdown("---")
st.caption("Preparado por Diego Vicente Contreras y Claude AI - IST 2026")
