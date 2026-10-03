import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime
import requests
import io
import time

# ---------------------------------------------------------
# 1. CONFIGURACIÓN GENERAL Y ESTILOS EJECUTIVOS BI
# ---------------------------------------------------------
st.set_page_config(
    page_title="Monitor de Incidentes - WFM & ANS", 
    page_icon="⚡", 
    layout="wide", 
    initial_sidebar_state="expanded"
)

# Estilos CSS de Grado BI Enterprise
st.markdown('''
<style>
    .main { background-color: #f8fafc; font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif; }
    
    .hero-banner {
        background: linear-gradient(135deg, #0f172a 0%, #1e293b 60%, #0f172a 100%);
        padding: 1.5rem 2rem;
        border-radius: 16px;
        color: #ffffff;
        box-shadow: 0 10px 25px -5px rgba(15, 23, 42, 0.25);
        margin-bottom: 1.5rem;
        border-left: 6px solid #0284c7;
        display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap;
    }
    .hero-title { color: #ffffff !important; font-weight: 800 !important; font-size: 1.7rem !important; margin: 0 !important; }
    .hero-subtitle { color: #94a3b8 !important; font-size: 0.85rem !important; margin-top: 0.3rem !important; margin-bottom: 0 !important; }
    
    .status-pill {
        background: rgba(2, 132, 199, 0.15); border: 1px solid rgba(56, 189, 248, 0.4);
        color: #38bdf8; padding: 0.4rem 0.9rem; border-radius: 9999px; font-size: 0.75rem; font-weight: 700;
    }

    .kpi-card {
        background-color: #ffffff; padding: 1.1rem; border-radius: 14px;
        border: 1px solid #e2e8f0; text-align: center; box-shadow: 0 2px 4px rgba(0,0,0,0.02);
    }
    .kpi-label { font-size: 0.72rem; font-weight: 800; text-transform: uppercase; letter-spacing: 0.05em; margin-bottom: 0.3rem; }
    .kpi-val { font-size: 2.1rem; font-weight: 900; line-height: 1; }
    .kpi-sub { font-size: 0.72rem; color: #64748b; margin-top: 0.3rem; font-weight: 500; }

    /* Semaforización ANS */
    .kpi-vencido { border-left: 5px solid #dc2626; } .kpi-vencido .kpi-label { color: #dc2626; } .kpi-vencido .kpi-val { color: #991b1b; }
    .kpi-limite { border-left: 5px solid #d97706; } .kpi-limite .kpi-label { color: #d97706; } .kpi-limite .kpi-val { color: #92400e; }
    .kpi-tiempo { border-left: 5px solid #16a34a; } .kpi-tiempo .kpi-label { color: #16a34a; } .kpi-tiempo .kpi-val { color: #166534; }
    .kpi-clientes { border-left: 5px solid #0284c7; } .kpi-clientes .kpi-label { color: #0284c7; } .kpi-clientes .kpi-val { color: #075985; }

    .stTabs [data-baseweb="tab-list"] { gap: 8px; background-color: #f1f5f9; padding: 6px; border-radius: 12px; }
    .stTabs [data-baseweb="tab"] { height: 42px; border-radius: 8px; font-weight: 700; font-size: 0.85rem; }
    .stTabs [aria-selected="true"] { background-color: #ffffff !important; box-shadow: 0 2px 4px rgba(0,0,0,0.06); }
</style>
''', unsafe_allow_html=True)

# Mapeo Inteligente de Circuitos a Sectores Operativos (Centro / Norte)
CIRCUITOS_SECTOR_MAP = {
    "LIBANO": "NORTE", "LERIDA": "NORTE", "MARIQUITA": "NORTE", "FRESNO": "NORTE", "VENADILLO": "NORTE", 
    "HONDA": "NORTE", "GUAYABAL": "NORTE", "PALOCABILDO": "NORTE", "HERVEO": "NORTE", "JUNIN": "NORTE", 
    "ANZOATEGUI": "NORTE", "SANTA ISABEL": "NORTE", "EL TABLAZO": "NORTE", "PADUA": "NORTE", "ARRIEROS": "NORTE",
    "CAJAMARCA": "CENTRO", "CHAPETON": "CENTRO", "SALDAÑA": "CENTRO", "SALADO": "CENTRO", "VERGEL": "CENTRO",
    "ROVIRA": "CENTRO", "PAYANDE": "CENTRO", "PIEDRAS": "CENTRO", "ALVARADO": "CENTRO", "PAPAYO": "CENTRO",
    "SAN JORGE": "CENTRO", "BRISAS": "CENTRO", "MIROLINDO": "CENTRO", "SAN JUAN": "CENTRO", "PASTALES": "CENTRO",
    "LA VEGA": "CENTRO", "LA MIEL": "CENTRO", "SIMON BOLIVAR": "CENTRO", "PICALEÑA": "CENTRO"
}

def obtener_sector_operativo(circuito, subestacion):
    texto = f"{str(circuito)} {str(subestacion)}".upper()
    for kw, sec in CIRCUITOS_SECTOR_MAP.items():
        if kw in texto:
            return sec
    return "CENTRO"

# ---------------------------------------------------------
# 2. ENCABEZADO Y BARRA LATERAL (CONFIGURACIÓN GITHUB)
# ---------------------------------------------------------
st.markdown('''
<div class="hero-banner">
    <div>
        <h1 class="hero-title">⚡ Monitor de Incidentes Operativos & ANS WFM</h1>
        <p class="hero-subtitle">Control de Tiempos de Atención Inmediata (Sin DP / ASIG) &nbsp;|&nbsp; <b>SLA Urbano:</b> ≤ 1 día &nbsp;•&nbsp; <b>SLA Rural:</b> ≤ 3 días</p>
    </div>
    <div><span class="status-pill">● MONITOREO ACTIVO</span></div>
</div>
''', unsafe_allow_html=True)

# URL RAW de GitHub (Reemplaza con el enlace Raw de tu repositorio si deseas carga automática)
URL_RAW_GITHUB = "https://raw.githubusercontent.com/TU_USUARIO/TU_REPOSITORIO/main/Navegador%20de%20incidentes.xlsx"

with st.sidebar:
    st.image("https://img.icons8.com/color/96/electricity.png", width=55)
    st.title("Centro de Control")
    st.markdown("---")
    
    # Opción A: Cargar archivo manualmente (Actualización Instantánea 0s)
    archivo_subido = st.file_uploader("📂 Cargar Reporte WFM (.xlsx)", type=["xlsx", "xls"])
    
    st.markdown("---")
    
    # Filtro Estricto de Instrucción
    solo_vacios = st.toggle("🚫 Excluir DP / DP BT / ASIG (Solo Instrucción en Blanco)", value=True)
    
    st.markdown("---")
    st.caption("<b>SLA Mantenimiento:</b><br>• Urbano: ≤ 1 día<br>• Rural: ≤ 3 días", unsafe_allow_html=True)

# Cargar desde GitHub Raw sin problemas de Caché
@st.cache_data(ttl=60)
def cargar_desde_github_raw(url):
    try:
        url_dinamica = f"{url}?t={int(time.time())}"
        res = requests.get(url_dinamica, timeout=5)
        if res.status_code == 200:
            return io.BytesIO(res.content)
    except:
        pass
    return None

archivo_a_usar = None
if archivo_subido is not None:
    archivo_a_usar = archivo_subido
elif os.path.exists("Navegador de incidentes_01_10_2026 06_15_57.320.xlsx"):
    archivo_a_usar = "Navegador de incidentes_01_10_2026 06_15_57.320.xlsx"
elif os.path.exists("Base_DP.xlsx"):
    archivo_a_usar = "Base_DP.xlsx"
else:
    archivo_a_usar = cargar_desde_github_raw(URL_RAW_GITHUB)

if archivo_a_usar is None:
    st.info("👋 Por favor, carga el archivo Excel exportado de WFM en la barra lateral.")
    st.stop()

# ---------------------------------------------------------
# 3. PROCESAMIENTO Y FILTRADO DE DATOS
# ---------------------------------------------------------
try:
    if hasattr(archivo_a_usar, 'seek'):
        archivo_a_usar.seek(0)
        
    df_temp = pd.read_excel(archivo_a_usar, header=None, nrows=15)
    header_idx = 5
    for i in range(len(df_temp)):
        fila = [str(c).lower() for c in df_temp.iloc[i].tolist()]
        if any('identificaci' in c for c in fila) and any('instrucci' in c for c in fila):
            header_idx = i
            break

    if hasattr(archivo_a_usar, 'seek'):
        archivo_a_usar.seek(0)

    df = pd.read_excel(archivo_a_usar, header=header_idx)
    df.columns = df.columns.astype(str).str.strip()
    df = df.loc[:, ~df.columns.duplicated()].copy()

    # FILTRO SOLICITADO: EXCLUIR DP, DP BT, ASIG (SOLO CAMPOS VACÍOS EN INSTRUCCIÓN)
    def es_instruccion_vacia(val):
        if pd.isna(val) or val is None:
            return True
        s = str(val).strip().upper()
        return s in ['', 'NAN', 'NONE', 'NULL', 'UNDEFINED']

    if solo_vacios and 'Instrucción' in df.columns:
        df = df[df['Instrucción'].apply(es_instruccion_vacia)].copy().reset_index(drop=True)

    if df.empty:
        st.warning("⚠️ No se encontraron incidentes con el filtro de Instrucción en blanco.")
        st.stop()

    # Normalización de Columnas Clave
    sub_col = [c for c in df.columns if 'subestaci' in c.lower()]
    df['Subestación'] = df[sub_col[0]].fillna('SIN SUBESTACIÓN').astype(str) if sub_col else 'SIN SUBESTACIÓN'

    circ_col = [c for c in df.columns if 'circuito normal' in c.lower() or 'circuito actual' in c.lower() or 'circuito' in c.lower()]
    df['Circuito'] = df[circ_col[0]].fillna('SIN CIRCUITO').astype(str) if circ_col else 'SIN CIRCUITO'

    dir_col = [c for c in df.columns if 'direcci' in c.lower()]
    df['Dirección'] = df[dir_col[0]].fillna('SIN DIRECCIÓN').astype(str) if dir_col else 'SIN DIRECCIÓN'

    cuad_col = [c for c in df.columns if 'cuadrilla' in c.lower()]
    df['Cuadrilla'] = df[cuad_col[0]].fillna('SIN ASIGNAR').astype(str) if cuad_col else 'SIN ASIGNAR'

    if 'Afectados' in df.columns:
        df['Clientes Sin Servicio'] = pd.to_numeric(df['Afectados'], errors='coerce').fillna(0).astype(int)
    elif 'Clientes no restaurados' in df.columns:
        df['Clientes Sin Servicio'] = pd.to_numeric(df['Clientes no restaurados'], errors='coerce').fillna(0).astype(int)
    else:
        df['Clientes Sin Servicio'] = 0

    # Sector Operativo (Centro vs Norte)
    df['Sector Operativo'] = df.apply(lambda r: obtener_sector_operativo(r['Circuito'], r['Subestación']), axis=1)

    # Clasificación Urbano vs Rural
    def clasificar_urbano_rural(direccion):
        dir_upper = str(direccion).upper()
        rurales = ['VDA', 'VEREDA', 'FCA', 'FINCA', 'CORREGIMIENTO']
        return 'RURAL' if any(kw in dir_upper for kw in rurales) else 'URBANO'

    df['Tipo de Sector'] = df['Dirección'].apply(clasificar_urbano_rural)

    # Cálculo de Fechas, Días y Horas Transcurridas
    if 'Fecha de creación' in df.columns:
        df['Fecha_Creacion_DT'] = pd.to_datetime(df['Fecha de creación'], errors='coerce')
    else:
        f_col = [c for c in df.columns if 'fecha' in c.lower() and 'creaci' in c.lower()]
        df['Fecha_Creacion_DT'] = pd.to_datetime(df[f_col[0]], errors='coerce') if f_col else pd.NaT

    ahora = pd.Timestamp.now()
    df['Horas_Transcurridas'] = ((ahora - df['Fecha_Creacion_DT']).dt.total_seconds() / 3600.0).fillna(0).round(1)
    df['Días Sin Servicio'] = (df['Horas_Transcurridas'] / 24.0).round(1)

    def formato_hhmmss(horas_num):
        secs = int(horas_num * 3600)
        h = secs // 3600
        m = (secs % 3600) // 60
        s = secs % 60
        return f"{h:02d}h {m:02d}m {s:02d}s"

    df['Tiempo_Transcurrido_Str'] = df['Horas_Transcurridas'].apply(formato_hhmmss)

    # Evaluación SLA / ANS
    def estado_sla(row):
        limite_dias = 3 if row['Tipo de Sector'] == 'RURAL' else 1
        if row['Días Sin Servicio'] > limite_dias:
            return 'Vencido'
        elif row['Días Sin Servicio'] == limite_dias:
            return 'Al Límite'
        else:
            return 'A Tiempo'

    df['Estado SLA'] = df.apply(estado_sla, axis=1)

    # ---------------------------------------------------------
    # 4. PANEL DE FILTROS (SECTOR OPERATIVO: CENTRO / NORTE)
    # ---------------------------------------------------------
    st.markdown('<div style="background:#fff; padding:1.2rem; border-radius:12px; border:1px solid #e2e8f0; margin-bottom:1.2rem;">', unsafe_allow_html=True)
    f1, f2, f3, f4 = st.columns([2, 2.5, 2.5, 1])

    with f1:
        sectores_op_disponibles = ['CENTRO', 'NORTE']
        sectores_op_sel = st.multiselect("🗺️ Sector Operativo:", options=sectores_op_disponibles, default=sectores_op_disponibles)

    with f2:
        subs_disponibles = sorted(df[df['Sector Operativo'].isin(sectores_op_sel)]['Subestación'].unique())
        subs_sel = st.multiselect("📍 Subestación:", options=subs_disponibles, default=subs_disponibles)

    with f3:
        circs_disponibles = sorted(df[(df['Sector Operativo'].isin(sectores_op_sel)) & (df['Subestación'].isin(subs_sel))]['Circuito'].unique())
        circs_sel = st.multiselect("🔌 Circuito:", options=circs_disponibles, default=circs_disponibles)

    with f4:
        st.markdown("<div style='height: 28px;'></div>", unsafe_allow_html=True)
        if st.button("🔄 Restablecer", use_container_width=True):
            sectores_op_sel = sectores_op_disponibles
            subs_sel = subs_disponibles
            circs_sel = circs_disponibles

    st.markdown('</div>', unsafe_allow_html=True)

    # Aplicación de filtros
    df_filtrado = df[
        (df['Sector Operativo'].isin(sectores_op_sel)) &
        (df['Subestación'].isin(subs_sel)) &
        (df['Circuito'].isin(circs_sel))
    ].copy().reset_index(drop=True)

    if df_filtrado.empty:
        st.warning("⚠️ No hay incidentes que coincidan con la combinación de filtros seleccionada.")
        st.stop()

    # ---------------------------------------------------------
    # 5. APARTADO RÁPIDO: MÉTRICAS Y RESUMEN EJECUTIVO (KPIS)
    # ---------------------------------------------------------
    cnt_total_avisos = len(df_filtrado)
    cnt_centro = len(df_filtrado[df_filtrado['Sector Operativo'] == 'CENTRO'])
    cnt_norte = len(df_filtrado[df_filtrado['Sector Operativo'] == 'NORTE'])
    cnt_circuitos = df_filtrado['Circuito'].nunique()
    cnt_afectados_totales = int(df_filtrado['Clientes Sin Servicio'].sum())

    cnt_vencidos = int((df_filtrado['Estado SLA'] == 'Vencido').sum())
    cnt_limite = int((df_filtrado['Estado SLA'] == 'Al Límite').sum())
    cnt_tiempo = int((df_filtrado['Estado SLA'] == 'A Tiempo').sum())

    k1, k2, k3, k4, k5 = st.columns(5)
    with k1:
        st.markdown(f'''<div class="kpi-card"><div class="kpi-label">📋 Contador Avisos</div><div class="kpi-val" style="color:#0284c7;">{cnt_total_avisos}</div><div class="kpi-sub">Atención Inmediata</div></div>''', unsafe_allow_html=True)
    with k2:
        st.markdown(f'''<div class="kpi-card"><div class="kpi-label">🗺️ Avisos por Zona</div><div class="kpi-val" style="color:#334155;">{cnt_centro} <span style="font-size:1rem; color:#64748b;">C</span> | {cnt_norte} <span style="font-size:1rem; color:#64748b;">N</span></div><div class="kpi-sub">Centro vs Norte</div></div>''', unsafe_allow_html=True)
    with k3:
        st.markdown(f'''<div class="kpi-card"><div class="kpi-label">🔌 Circuitos Afectados</div><div class="kpi-val" style="color:#475569;">{cnt_circuitos}</div><div class="kpi-sub">Circuitos en servicio</div></div>''', unsafe_allow_html=True)
    with k4:
        st.markdown(f'''<div class="kpi-card kpi-clientes"><div class="kpi-label">👥 Clientes Afectados</div><div class="kpi-val">{cnt_afectados_totales:,}</div><div class="kpi-sub">Clientes sin energía</div></div>''', unsafe_allow_html=True)
    with k5:
        st.markdown(f'''<div class="kpi-card kpi-vencido"><div class="kpi-label">🔴 Vencidos SLA</div><div class="kpi-val">{cnt_vencidos}</div><div class="kpi-sub">Exceden tiempo ANS</div></div>''', unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # ---------------------------------------------------------
    # 6. PESTAÑAS DETALLADAS SOLICITADAS
    # ---------------------------------------------------------
    tab_tiempos, tab_avisos, tab_circuitos, tab_sla = st.tabs([
        "⏱️ Contador de Horas (Mayor a Menor)", 
        "👥 Control Avisos (Clientes Afectados)", 
        "🔌 Cantidad de Incidentes por Circuito",
        "🚨 Semaforización & Cumplimiento ANS"
    ])

    # PESTAÑA 1: CONTADOR DE HORAS ORGANIZADO DE MAYOR A MENOR TIEMPO
    with tab_tiempos:
        st.subheader("⏱️ Contador de Tiempo Transcurrido (Organizado de Mayor a Menor)")
        st.caption("Visualiza los incidentes priorizados de mayor a menor tiempo transcurrido desde su creación.")

        df_tiempos = df_filtrado.sort_values(by='Horas_Transcurridas', ascending=False).reset_index(drop=True)

        cols_tiempos = [
            'Identificación', 'Sector Operativo', 'Subestación', 'Circuito', 
            'Dirección', 'Tipo de Sector', 'Clientes Sin Servicio', 
            'Tiempo_Transcurrido_Str', 'Días Sin Servicio', 'Estado SLA'
        ]
        
        cols_presentes = [c for c in cols_tiempos if c in df_tiempos.columns]

        def colorear_sla(val):
            if val == 'Vencido': return 'background-color: #fee2e2; color: #991b1b; font-weight: bold;'
            if val == 'Al Límite': return 'background-color: #fef3c7; color: #92400e; font-weight: bold;'
            if val == 'A Tiempo': return 'background-color: #dcfce7; color: #166534; font-weight: bold;'
            return ''

        st.dataframe(
            df_tiempos[cols_presentes].style.map(colorear_sla, subset=['Estado SLA']),
            use_container_width=True,
            height=450
        )

    # PESTAÑA 2: CONTROL AVISOS (CLIENTES AFECTADOS: ID, CUADRILLA, AFECTADOS)
    with tab_avisos:
        st.subheader("👥 Control Avisos — Clientes Afectados por Incidente y Cuadrilla")
        st.caption("Detalle específico de Identificación del incidente, Cuadrilla asignada y Clientes Afectados.")

        df_avisos = df_filtrado.sort_values(by='Clientes Sin Servicio', ascending=False).reset_index(drop=True)

        cols_avisos = ['Identificación', 'Sector Operativo', 'Subestación', 'Circuito', 'Cuadrilla', 'Clientes Sin Servicio', 'Estado SLA']
        cols_av_presentes = [c for c in cols_avisos if c in df_avisos.columns]

        a1, a2 = st.columns([3, 1.5])
        with a1:
            st.dataframe(
                df_avisos[cols_av_presentes].style.map(colorear_sla, subset=['Estado SLA']),
                use_container_width=True,
                height=420
            )
        with a2:
            st.markdown("##### 📊 Top Cuadrillas con Más Clientes")
            cuad_summary = df_filtrado.groupby('Cuadrilla', as_index=False)['Clientes Sin Servicio'].sum().sort_values(by='Clientes Sin Servicio', ascending=False)
            st.dataframe(cuad_summary.reset_index(drop=True), use_container_width=True, height=360)

    # PESTAÑA 3: CANTIDAD DE INCIDENTES POR CIRCUITO
    with tab_circuitos:
        st.subheader("🔌 Cantidad de Incidentes Discriminados por Circuito")
        
        circ_df = df_filtrado.groupby(['Sector Operativo', 'Subestación', 'Circuito'], as_index=False).agg(
            Cantidad_Incidentes=('Identificación', 'count'),
            Total_Clientes_Afectados=('Clientes Sin Servicio', 'sum')
        ).sort_values(by='Cantidad_Incidentes', ascending=False).reset_index(drop=True)

        c1, c2 = st.columns([3, 2])
        with c1:
            fig_circ = px.bar(
                circ_df.head(15), x='Cantidad_Incidentes', y='Circuito',
                color='Sector Operativo',
                color_discrete_map={'CENTRO': '#0284c7', 'NORTE': '#10b981'},
                orientation='h',
                title="Top Circuitos con Mayor Cantidad de Incidentes",
                text='Cantidad_Incidentes'
            )
            fig_circ.update_layout(yaxis={'categoryorder': 'total ascending'}, height=400)
            st.plotly_chart(fig_circ, use_container_width=True)

        with c2:
            st.markdown("##### 📋 Resumen por Circuito")
            st.dataframe(circ_df, use_container_width=True, height=380)

    # PESTAÑA 4: SEMAFORIZACIÓN ANS & CUMPLIMIENTO
    with tab_sla:
        st.subheader("🚨 Semaforización & Cumplimiento ANS de Tiempos")
        colores_map = {'A Tiempo': '#00B050', 'Al Límite': '#FFC000', 'Vencido': '#C00000'}

        s1, s2 = st.columns([2, 3])
        with s1:
            conteo_sla = df_filtrado['Estado SLA'].value_counts().reset_index()
            conteo_sla.columns = ['Estado SLA', 'Cantidad']
            
            fig_pie = px.pie(
                conteo_sla, names='Estado SLA', values='Cantidad',
                color='Estado SLA', color_discrete_map=colores_map,
                hole=0.5
            )
            fig_pie.update_traces(textposition='inside', textinfo='percent+value')
            fig_pie.update_layout(height=350, margin=dict(t=10, b=10, l=10, r=10))
            st.plotly_chart(fig_pie, use_container_width=True)

        with s2:
            fig_top = px.bar(
                df_filtrado.sort_values(by='Días Sin Servicio', ascending=False).head(10),
                x='Días Sin Servicio', y='Identificación',
                color='Estado SLA', color_discrete_map=colores_map,
                orientation='h', title="Top 10 Incidentes con Más Días Sin Servicio",
                hover_data=['Subestación', 'Circuito', 'Clientes Sin Servicio', 'Tipo de Sector']
            )
            fig_top.update_layout(yaxis={'categoryorder': 'total ascending'}, height=350)
            st.plotly_chart(fig_top, use_container_width=True)

except Exception as e:
    st.error(f"❌ Error al procesar el archivo Excel: {e}")
