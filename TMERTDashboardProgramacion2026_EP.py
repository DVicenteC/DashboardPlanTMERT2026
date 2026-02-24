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
def obtener_ranking_limpio(df, columna, separador_principal="||", separador_secundario=","):
    """
    Cuenta la frecuencia de elementos atómicos en una columna multi-valor.

    Estructura de los datos (crear_resumen_estructurado sin prefijos):
      - Separador entre registros de folio : ' || '
      - Separador secundario interno       :
          · segmentos, tareas, diagnósticos → ','
          · ocupaciones                     → ' | '
    """
    if columna not in df.columns:
        return pd.DataFrame(columns=['Nombre', 'Cantidad'])

    datos = df[df[columna].astype(str).str.strip() != ""][columna].astype(str)

    if datos.empty:
        return pd.DataFrame(columns=['Nombre', 'Cantidad'])

    series = datos.str.split(separador_principal, regex=False).explode().str.strip()
    series = series.str.split(separador_secundario, regex=False).explode().str.strip()

    conteo = series[series != ""].value_counts().reset_index()
    conteo.columns = ['Nombre', 'Cantidad']
    return conteo

# ── 5. CONTEO DE FOLIOS EP ───────────────────────────────────────────────────
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

# ── 6. FUNCIÓN DETALLE COMPLETO ───────────────────────────────────────────────
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

if df_raw is not None:

    # ── SIDEBAR ───────────────────────────────────────────────────────────────
    st.sidebar.image("https://www.ist.cl/wp-content/themes/ist/img/logo-ist.png", width=100)
    st.sidebar.title("🔍 Filtros de Gestión")

    solo_ep = st.sidebar.toggle("🚨 Ver solo Centros con EP", value=False)

    filtro_ergo = st.sidebar.selectbox(
        "Especialista", ["Todos"] + sorted(df_raw['Ergonomo'].unique().tolist())
    )
    filtro_gerencia = st.sidebar.selectbox(
        "Gerencia - Cuenta Nacional",
        ["Todas"] + sorted(df_raw['Gerencia - Cuenta Nacional'].unique().tolist())
    )
    filtro_holding = st.sidebar.selectbox(
        "Holding", ["Todos"] + sorted(df_raw['Holding'].unique().tolist())
    )

    filtro_empleador = st.sidebar.selectbox(
        "Nombre Empleador",
        ["Todos"] + sorted(df_raw['Nombre Empleador'].dropna().astype(str).unique().tolist())
    )

    filtro_reg = st.sidebar.selectbox(
        "Región", ["Todas"] + sorted(df_raw['Región'].unique().tolist())
    )

    meses_espanol = ['Enero', 'Febrero', 'Marzo', 'Abril', 'Mayo', 'Junio',
                     'Julio', 'Agosto', 'Septiembre', 'Octubre', 'Noviembre', 'Diciembre']
    filtro_mes = st.sidebar.selectbox("Mes (Programación)", ["Todos"] + meses_espanol)

    if st.sidebar.button("🔄 Resetear Filtros"):
        st.rerun()

    # ── APLICAR FILTROS ───────────────────────────────────────────────────────
    df = df_raw.copy()
    if solo_ep:
        df = df[df['Tiene EP'] == True]
    if filtro_ergo != "Todos":
        df = df[df['Ergonomo'] == filtro_ergo]
    if filtro_gerencia != "Todas":
        df = df[df['Gerencia - Cuenta Nacional'] == filtro_gerencia]
    if filtro_holding != "Todos":
        df = df[df['Holding'] == filtro_holding]
    if filtro_empleador != "Todos":
        df = df[df['Nombre Empleador'] == filtro_empleador]
    if filtro_reg != "Todas":
        df = df[df['Región'] == filtro_reg]

    # df_prog: registros con fecha programada (para tab Programación)
    df_prog = df[df['fecha'].notna()].copy()
    if filtro_mes != "Todos":
        meses_es_a_num = {m: i+1 for i, m in enumerate(meses_espanol)}
        df_prog = df_prog[df_prog['mes'] == meses_es_a_num[filtro_mes]]

    # ── TÍTULO ────────────────────────────────────────────────────────────────
    st.title("🏥 Dashboard TMERT 2026 - Gestión Integral")
    st.markdown(
        f"**IST · Circular SUSESO 3900** | "
        f"Datos actualizados: {datetime.now().strftime('%d/%m/%Y %H:%M')}"
    )

    # ── MÉTRICAS GLOBALES ─────────────────────────────────────────────────────
    m1, m2 = st.columns(2)
    with m1:
        st.metric("AT Programadas", f"{len(df_prog):,}")
    with m2:
        n_ep = contar_folios_distintos(df)
        n_emp_ep = df[df['Tiene EP']]['Nombre Empleador'].nunique() if n_ep > 0 else 0
        st.metric("Folios EP", n_ep, f"{n_emp_ep} empresa(s)", delta_color="inverse")

    st.markdown("---")

    # ── TABS ──────────────────────────────────────────────────────────────────
    tab1, tab2, tab3 = st.tabs([
        "📊 Programación y Carga",
        "🔍 Análisis de Denuncias EP",
        "📋 Planilla Detallada"
    ])

    # ── TAB 1: PROGRAMACIÓN Y CARGA ───────────────────────────────────────────
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

    # ── TAB 2: ANÁLISIS DE DENUNCIAS EP ───────────────────────────────────────
    with tab2:
        st.header("Análisis de Salud Laboral (Ranking de EP)")
        df_ep = df[df['Tiene EP']]

        if len(df_ep) > 0:

            # NIVEL 1: VISTA GENERAL
            rank_seg_gen  = obtener_ranking_limpio(df_ep, 'segmentos')
            rank_diag_gen = obtener_ranking_limpio(df_ep, 'diagnosticos')
            rank_emp_gen  = folios_por_empresa(df_ep)

            

            col_g1, col_g2 = st.columns(2)
            with col_g1:
                if not rank_seg_gen.empty:
                    fig_seg = px.bar(
                        rank_seg_gen.head(10), x='Cantidad', y='Nombre', orientation='h',
                        color_discrete_sequence=['#E67E22'],
                        title="Top Segmentos Corporal Afectados"
                    )
                    fig_seg.update_layout(yaxis={'categoryorder': 'total ascending'}, margin=dict(l=0))
                    st.plotly_chart(fig_seg, use_container_width=True)
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

            # NIVEL 2 & 3: EXPLORADOR
            st.subheader("🔍 Explorador: De lo General a lo Particular")
            st.caption(
                "Elige un segmento Corporal o un diagnóstico para ver en qué empresas, "
                "puestos de trabajo y tareas se concentra ese riesgo."
            )

            modo = st.radio(
                "Explorar por:", ["Segmento Corporal", "Diagnóstico"],
                horizontal=True, key="modo_explor"
            )
            col_explor = 'segmentos' if modo == "Segmento Corporal" else 'diagnosticos'
            rank_base   = rank_seg_gen if modo == "Segmento Corporal" else rank_diag_gen

            if not rank_base.empty:
                seleccion = st.selectbox(
                    f"Selecciona un {'segmento' if modo == 'Segmento Corporal' else 'diagnóstico'}:",
                    ["Todos"] + rank_base['Nombre'].tolist(),
                    key="explorador_selector"
                )

                if seleccion == "Todos":
                    df_drill = df_ep.copy()
                else:
                    df_drill = df_ep[
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
                        r_emp = folios_por_empresa(df_drill)
                        st.dataframe(r_emp, use_container_width=True, hide_index=True)
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

                    with st.expander("📋 Ver registros individuales"):
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

    # ── TAB 3: PLANILLA DETALLADA ─────────────────────────────────────────────
    with tab3:
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
