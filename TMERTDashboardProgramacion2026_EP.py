"""
Dashboard TMERT 2026 - Gestión Integral (Programación + Análisis EP)
Autor: Diego Vicente Contreras y Claude AI - IST 2026
"""

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import duckdb
import re
import io
from datetime import datetime

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

def parsear_fecha_flexible(serie):
    """
    Parsea una serie de fechas que puede tener formatos mixtos:
    - DD-MM-YYYY (formato Excel original)
    - M/D/YYYY o MM/DD/YYYY (formato Google Sheets export)
    - YYYY-MM-DD (formato ISO)
    """
    resultado = pd.to_datetime(serie, format='%d-%m-%Y', errors='coerce')
    mascara = resultado.isna() & serie.notna() & (serie.astype(str).str.strip() != '')
    if mascara.any():
        resultado[mascara] = pd.to_datetime(serie[mascara], format='%m/%d/%Y', errors='coerce')
    mascara2 = resultado.isna() & serie.notna() & (serie.astype(str).str.strip() != '')
    if mascara2.any():
        resultado[mascara2] = pd.to_datetime(serie[mascara2], dayfirst=True, errors='coerce')
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
            
        # Convertir columnas booleanas (Pilar 1, 2, 3, 4, Meta 5) a booleanos reales
        cols_bool = [c for c in df.columns if 'Pilar' in c or 'Cumplida' in c or 'Validado' in c]
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
    ranking['Pct']      = (ranking['Cantidad'] / total * 100).round(1)
    ranking['PctAcum']  = ranking['Pct'].cumsum().round(1)
    ranking['Rank']     = range(1, len(ranking) + 1)
    ranking['Vital']    = ranking['PctAcum'] <= 80.0

    colores = ['#E74C3C' if v else '#85C1E9' for v in ranking['Vital']]

    fig = go.Figure()

    # Barras
    fig.add_trace(go.Bar(
        x=ranking['Nombre'], y=ranking['Cantidad'],
        name='Casos EP', marker_color=colores, yaxis='y1'
    ))

    # Línea acumulada
    fig.add_trace(go.Scatter(
        x=ranking['Nombre'], y=ranking['PctAcum'],
        name='% Acumulado', mode='lines+markers',
        line=dict(color='#2E86AB', width=2),
        marker=dict(size=5), yaxis='y2'
    ))

    # Línea de corte 80 %
    fig.add_shape(
        type='line', xref='paper', yref='y2',
        x0=0, x1=1, y0=80, y1=80,
        line=dict(color='orange', width=2, dash='dash')
    )
    fig.add_annotation(
        xref='paper', yref='y2', x=1.01, y=80,
        text='80 %', showarrow=False,
        xanchor='left', font=dict(color='orange', size=11)
    )

    fig.update_layout(
        title=titulo,
        xaxis=dict(tickangle=-40),
        yaxis=dict(title='Casos EP', side='left'),
        yaxis2=dict(title='% Acumulado', side='right', overlaying='y',
                    range=[0, 112], showgrid=False),
        legend=dict(orientation='h', yanchor='bottom', y=1.02, x=0),
        height=500, margin=dict(r=60)
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
        file_name=f'detalle_tmert_{datetime.now().strftime("%Y%m%d")}.xlsx',
        mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        key=f'download_btn_{seccion}'
    )

# ── 7. INTERFAZ PRINCIPAL ─────────────────────────────────────────────────────
df_raw = load_data()
df_seg_raw = cargar_datos_seguimiento_tmert()

if df_raw is not None:

    # ── SIDEBAR — Filtros Bidireccionales Robustos (Cross-filtering) ──────────
    st.sidebar.image("https://www.ist.cl/wp-content/themes/ist/img/logo-ist.png", width=100)
    st.sidebar.title("🔍 Filtros de Gestión")

    solo_ep = st.sidebar.toggle("🚨 Ver solo centros con denuncias de EP", value=False)

    # Base: dataset de referencia (el toggle EP actúa como pre-filtro crítico)
    _base_t = df_raw[df_raw['Tiene EP'] == True].copy() if solo_ep else df_raw.copy()

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

    # df_prog: registros con fecha programada (para tab Programación)
    df_prog = df[df['fecha'].notna()].copy()
    if filtro_mes != "Todos":
        meses_es_a_num = {m: i+1 for i, m in enumerate(meses_espanol)}
        df_prog = df_prog[df_prog['mes'] == meses_es_a_num[filtro_mes]]

    # ── TÍTULO ────────────────────────────────────────────────────────────────
    st.title("🏥 Dashboard TMERT 2026 - Gestión Integral")
    st.markdown(
        f"**IST · Especialidades Técnicas** | "
        f"Datos actualizados: {datetime.now().strftime('%d/%m/%Y %H:%M')}"
    )

    # ── MÉTRICAS GLOBALES ─────────────────────────────────────────────────────
    m1, m2, m3 = st.columns(3)
    with m1:
        st.metric("AT Programadas", f"{len(df_prog):,}")
    with m2:
        if not df_seg.empty and 'Fecha real AT' in df_seg.columns:
            realizadas = df_seg['Fecha real AT'].notna().sum()
            porc = (realizadas / len(df_prog) * 100) if len(df_prog) > 0 else 0
            st.metric("AT Realizadas", f"{realizadas:,}", f"{porc:.1f}% del programa")
        else:
            st.metric("AT Realizadas", "S/D")
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
        "👨‍⚕️ Indicadores por Profesional",
    ])
    tab1, tab2, tab_seg, tab_ind = tabs

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
            c1, c2, c3, c4 = st.columns(4)
            real_at = df_seg['Fecha real AT'].notna().sum() if 'Fecha real AT' in df_seg.columns else 0
            atrasadas = (df_seg['Estado AT'] == 'Pendiente atrasada').sum() if 'Estado AT' in df_seg.columns else 0

            c1.metric("Universo Plan", f"{total_plan:,}")
            c2.metric("AT con Registro", f"{real_at:,}", f"{(real_at/total_plan*100):.1f}%" if total_plan > 0 else "0%")
            c3.metric("Meta 5 Completa", f"{cumplen_meta:,}", "4 Pilares + Seg. 1")
            c4.metric("Pendientes Atrasadas", f"{atrasadas:,}", delta_color="inverse")

            st.divider()

            # Desglose por Pilares
            st.markdown("#### 🧱 Desglose por Pilares de Cumplimiento")
            cp1, cp2, cp3, cp4 = st.columns(4)
            for col_m, label, p_col in zip(
                [cp1, cp2, cp3, cp4],
                ["P1: Difusión", "P2: Capacitación", "P3: Diseño Cap", "P4: Prescripción Caract"],
                ["Pilar 1 - Difusión", "Pilar 2 - Capacitación", "Pilar 3 - Diseño Cap Pract", "Pilar 4 - Prescripción Caract"]
            ):
                if p_col in df_seg.columns:
                    val = df_seg[p_col].sum()
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
                'Región', 'Ergonomo', 'Nombre Empleador', 'ID-CT', 'Nombre CT',
                'Estado AT',
                'Meta 5 Cumplida',
                'Pilar 1 - Difusión',
                'Pilar 2 - Capacitación',
                'Pilar 3 - Diseño Cap Pract',
                'Pilar 4 - Prescripción Caract',
                'Estado Seguimiento Prescripción Caracterización (sigeco)',
                'Fecha AT Difusión (real)',
                'Fecha AT Capacitación (real)',
                'Fecha Prescripción Caracterización (real)',
                'Fecha Diseño Cap Práctica (real)',
            ]
            cols_s = [c for c in cols_s if c in df_seg.columns]

            df_det = df_seg[cols_s].copy()
            for c in df_det.columns:
                if 'Fecha' in c:
                    df_det[c] = pd.to_datetime(df_det[c], errors='coerce').dt.strftime('%d-%m-%Y').fillna('')

            st.dataframe(df_det, use_container_width=True, hide_index=True)

            buffer_seg = io.BytesIO()
            with pd.ExcelWriter(buffer_seg, engine='openpyxl') as writer:
                df_det.to_excel(writer, index=False, sheet_name='Estado_Seguimiento')
            st.download_button(
                label="📥 Descargar Estado de Seguimiento en Excel",
                data=buffer_seg.getvalue(),
                file_name=f'estado_seguimiento_tmert_{datetime.now().strftime("%Y%m%d")}.xlsx',
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                key='download_seg'
            )

    # ── TAB: INDICADORES POR PROFESIONAL ─────────────────────────────────────
    with tab_ind:
        if df_seg.empty:
            st.info("ℹ️ Se requiere data de seguimiento para calcular indicadores por profesional.")
        else:
            st.subheader("👨‍⚕️ Indicadores de Gestión Mensual")
            
            # Agrupar por Profesional (Ergónomo)
            df_is = df_seg.copy()
            
            # Definir columnas de métricas
            # Nro. de actividades en SIGECO = Fecha real AT
            # Nro. de actividades en GOPIST = Identificaciones Iniciales + Avanzadas (istprod)
            # Nro. de registros en MK = Prescripciones
            
            agg_dict = {
                'Fecha real AT': 'count',
                'Fecha Últ. Identificación Inicial (istprod)': 'count',
                'Fecha Identificación Avanzada (real)': 'count',
                'Prescripción Evaluación Inicial (sigeco)': 'count',
                'Fecha Prescripción Eval Avanzada (sigeco)': 'count',
            }
            if 'Meta 5 Cumplida' in df_is.columns:
                agg_dict['Meta 5 Cumplida'] = 'sum'
            ind = df_is.groupby('Ergonomo').agg(agg_dict).reset_index()

            col_names = ['Ergónomo', 'SIGECO (ATs)', 'GOPIST (Ini)', 'GOPIST (Avz)',
                         'MK (Ini)', 'MK (Avz)']
            if 'Meta 5 Cumplida' in df_is.columns:
                col_names.append('Meta 5 (Cumple)')
            ind.columns = col_names

            # Totales
            ind['Total GOPIST'] = ind['GOPIST (Ini)'] + ind['GOPIST (Avz)']
            ind['Total MK'] = ind['MK (Ini)'] + ind['MK (Avz)']

            # Reordenar
            keep_cols = ['Ergónomo', 'SIGECO (ATs)', 'Total GOPIST', 'Total MK']
            if 'Meta 5 (Cumple)' in ind.columns:
                keep_cols.append('Meta 5 (Cumple)')
            ind = ind[keep_cols]
            
            # Mostrar Tabla de Ranking
            st.dataframe(ind.sort_values('SIGECO (ATs)', ascending=False), use_container_width=True, hide_index=True)
            
            # Gráficos Comparativos
            st.divider()
            col_i1, col_i2 = st.columns(2)
            with col_i1:
                y_cols_i1 = ['SIGECO (ATs)']
                if 'Meta 5 (Cumple)' in ind.columns:
                    y_cols_i1.append('Meta 5 (Cumple)')
                fig_i1 = px.bar(ind, x='Ergónomo', y=y_cols_i1,
                               barmode='group', title='SIGECO vs Meta 5')
                st.plotly_chart(fig_i1, use_container_width=True)
            with col_i2:
                fig_i2 = px.bar(ind, x='Ergónomo', y=['Total GOPIST', 'Total MK'], 
                               barmode='group', title='GOPIST vs MK')
                st.plotly_chart(fig_i2, use_container_width=True)

    # ── TAB 2: ANÁLISIS DE DENUNCIAS EP ───────────────────────────────────────
    with tab2:
        st.header("🔍 Análisis de Denuncias EP")
        df_ep = df[df['Tiene EP']]

        if len(df_ep) > 0:

            # NIVEL 1: VISTA GENERAL
            rank_seg_gen  = obtener_ranking_limpio(df_ep, 'segmentos', separadores_extra=[" "])
            rank_diag_gen = obtener_ranking_limpio(df_ep, 'diagnosticos')
            rank_emp_gen  = folios_por_empresa(df_ep)

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

            # ── PARETO EP (Detalle del Ranking) ───────────────────────────────
            st.subheader("📈 Ranking EP — Análisis de Pareto")
            st.caption(
                "Las barras **rojas** son las categorías vitales que acumulan hasta el 80 % de los casos."
            )

            df_ep_pareto = df_ep.copy()
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

                    pm1, pm2, pm3 = st.columns(3)
                    n_vital = int(df_p['Vital'].sum())
                    n_total = len(df_p)
                    pct_cat = round(n_vital / n_total * 100, 1) if n_total > 0 else 0
                    casos_vital = int(df_p[df_p['Vital']]['Cantidad'].sum())
                    pct_casos = round(casos_vital / df_p['Cantidad'].sum() * 100, 1) if df_p['Cantidad'].sum() > 0 else 0
                    pm1.metric("Categorías vitales (🔴)", f"{n_vital} de {n_total}")
                    pm2.metric("% de categorías vitales", f"{pct_cat} %")
                    pm3.metric("Casos EP que concentran", f"{pct_casos} %")

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
                    st.info("No hay datos suficientes para construir el Pareto.")
            else:
                st.info(f"No hay registros EP con segmento «{seg_pareto}» para los filtros actuales.")

            st.divider()

            # ── EXPLORADOR: De lo General a lo Particular ─────────────────────
            st.subheader("🔍 Explorador: De lo General a lo Particular")
            st.caption(
                "Elige un segmento corporal o un diagnóstico para ver en qué empresas, "
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
                                'segmentos', 'ocupaciones', 'tareas', 'diagnosticos', 'folios']
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
            file_name=f'planilla_tmert_{datetime.now().strftime("%Y%m%d")}.xlsx',
            mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            key='download_tab3'
        )

# ── 8. FOOTER ─────────────────────────────────────────────────────────────────
st.markdown("---")
st.caption("Preparado por Diego Vicente Contreras y Claude AI - IST 2026")
