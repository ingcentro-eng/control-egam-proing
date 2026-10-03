import os
import io
import time
import requests
import pandas as pd
import streamlit as st
import plotly.express as px

# ---------------------------------------------------------
# 1. CONFIGURACIÓN GENERAL Y ESTILOS BI ENTERPRISE
# ---------------------------------------------------------
st.set_page_config(
    page_title="Control EGAM PROING - Contrato 013-26", 
    page_icon="⚡", 
    layout="wide", 
    initial_sidebar_state="expanded"
)

# Estilos CSS de Grado Ejecutivo
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

    .stTabs [data-baseweb="tab-list"] { gap: 8px; background-color: #f1f5f9; padding: 6px; border-radius: 12px; }
    .stTabs [data-baseweb="tab"] { height: 42px; border-radius: 8px; font-weight: 700; font-size: 0.85rem; }
    .stTabs [aria-selected="true"] { background-color: #ffffff !important; box-shadow: 0 2px 4px rgba(0,0,0,0.06); }
</style>
''', unsafe_allow_html=True)

# Mapeo de Circuitos a Sectores Operativos (Centro / Norte)
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

# Evaluar metas del KPI 3 según tabla oficial
def evaluar_kpi3(afectados, horas_transcurridas):
    if afectados <= 20:
        rango, meta = "1 - 20", 20.0
    elif afectados <= 100:
        rango, meta = "21 - 100", 9.0
    elif afectados <= 500:
        rango, meta = "101 - 500", 6.0
    elif afectados <= 5000:
        rango, meta = "501 - 5000", 3.0
    else:
        rango, meta = "> 5000", 1.0
        
    cumple = horas_transcurridas <= meta
    return rango, meta, cumple

# ---------------------------------------------------------
# 2. ENCABEZADO Y BARRA LATERAL (CARGA DE DATOS)
# ---------------------------------------------------------
st.markdown('''
<div class="hero-banner">
    <div>
        <h1 class="hero-title">⚡ CONTROL EGAM PROING CONTRATO 013-26</h1>
        <p class="hero-subtitle">Seguimiento WFM Tiempos & Control KPI 3 (Rangos de Clientes) — Centro y Norte</p>
    </div>
    <div><span class="status-pill">● EGAM ACTIVO</span></div>
</div>
''', unsafe_allow_html=True)

# CONFIGURACIÓN DE URL RAW DE GITHUB
# Reemplaza esta URL por la dirección Raw oficial de tu archivo en GitHub:
DEFAULT_RAW_URL = "https://raw.githubusercontent.com/TU_USUARIO/TU_REPOSITORIO/main/Navegador%20de%20incidentes.xlsx"

with st.sidebar:
    st.image("https://img.icons8.com/color/96/electricity.png", width=55)
    st.title("Centro de Control EGAM")
    st.markdown("---")
    
    st.markdown("### 1. Carga Manual")
    archivo_subido = st.file_uploader("📂 Cargar Reporte WFM (.xlsx, .csv)", type=["xlsx", "xls", "csv"])
    
    st.markdown("---")
    st.markdown("### 2. Carga Automática (GitHub)")
    url_github = st.text_input("🔗 URL Raw del Excel en GitHub:", value=DEFAULT_RAW_URL)
    
    st.markdown("---")
    excluir_dp_asig = st.toggle("🚫 Excluir DP / DP BT / ASIG (Solo WFM Atención Inmediata)", value=True)
    st.markdown("---")
    st.caption("<b>Metas KPI 3 (EGAM):</b><br>• 1-20: ≤ 20 h<br>• 21-100: ≤ 9 h<br>• 101-500: ≤ 6 h<br>• 501-5000: ≤ 3 h<br>• > 5000: ≤ 1 h")

@st.cache_data(ttl=60)
def cargar_desde_github_raw(url):
    if not url or "TU_USUARIO" in url:
        return None
    try:
        url_dinamica = f"{url}?t={int(time.time())}"
        res = requests.get(url_dinamica, timeout=8)
        if res.status_code == 200:
            return io.BytesIO(res.content)
    except Exception:
        pass
    return None

# Determinar origen del archivo
archivo_a_usar = None
origen_datos = ""

if archivo_subido is not None:
    archivo_a_usar = archivo_subido
    origen_datos = "📂 Carga Manual (Archivo Local)"
else:
    bytes_github = cargar_desde_github_raw(url_github)
    if bytes_github is not None:
        archivo_a_usar = bytes_github
        origen_datos = "🔄 Sincronización Automática (GitHub)"
    else:
        # Intento de lectura de archivos locales en la raíz
        archivos_locales = [f for f in os.listdir('.') if f.endswith(('.xlsx', '.csv')) and not f.startswith('.')]
        if archivos_locales:
            archivo_a_usar = archivos_locales[0]
            origen_datos = f"📄 Archivo Local en Servidor ({archivos_locales[0]})"

if archivo_a_usar is None:
    st.info("👋 Por favor, carga el reporte diario exportado de WFM en la barra lateral o verifica la URL Raw de GitHub.")
    st.stop()

st.sidebar.success(f"Origen de Datos: **{origen_datos}**")

# ---------------------------------------------------------
# 3. PROCESAMIENTO Y TRANSFORMACIÓN
# ---------------------------------------------------------
try:
    if hasattr(archivo_a_usar, 'seek'):
        archivo_a_usar.seek(0)

    # Detectar si es CSV o Excel
    es_csv = False
    if isinstance(archivo_a_usar, str) and archivo_a_usar.endswith('.csv'):
        es_csv = True
    elif hasattr(archivo_a_usar, 'name') and archivo_a_usar.name.endswith('.csv'):
        es_csv = True

    if es_csv:
        # 1. Intentar determinar separador (';' o ',')
        df_raw = pd.read_csv(archivo_a_usar, nrows=15, header=None)
        
        # Buscar la fila que contiene las columnas reales
        header_idx = 5
        sep_usado = ';'
        
        for i in range(min(15, len(df_raw))):
            fila_str = " ".join([str(x) for x in df_raw.iloc[i].values]).lower()
            if 'identificaci' in fila_str and ('afectados' in fila_str or 'circuito' in fila_str):
                header_idx = i
                if ';' in fila_str:
                    sep_usado = ';'
                elif ',' in fila_str:
                    sep_usado = ','
                break
                
        if hasattr(archivo_a_usar, 'seek'):
            archivo_a_usar.seek(0)
            
        df = pd.read_csv(archivo_a_usar, header=header_idx, sep=sep_usado)
    else:
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

    # Limpieza de nombres de columnas
    df.columns = df.columns.astype(str).str.strip()
    df = df.loc[:, ~df.columns.duplicated()].copy()

    # Identificación / ID Incidente
    id_col = [c for c in df.columns if 'identificaci' in c.lower() or 'incidente' in c.lower() or 'aviso' in c.lower()]
    df['Identificación'] = df[id_col[0]].astype(str) if id_col else df.index.astype(str)

    # Excluir DP y ASIG para la vista de WFM si está activo el toggle
    def es_atencion_inmediata(val):
        if pd.isna(val) or val is None:
            return True
        s = str(val).strip().upper()
        return s in ['', 'NAN', 'NONE', 'NULL', 'UNDEFINED']

    col_instruccion = [c for c in df.columns if 'instrucci' in c.lower()]
    if excluir_dp_asig and col_instruccion:
        df_wfm = df[df[col_instruccion[0]].apply(es_atencion_inmediata)].copy().reset_index(drop=True)
        if df_wfm.empty:
            df_wfm = df.copy().reset_index(drop=True)
    else:
        df_wfm = df.copy().reset_index(drop=True)

    # Nombres de columnas homogeneizados
    sub_col = [c for c in df_wfm.columns if 'subestaci' in c.lower()]
    df_wfm['Subestación'] = df_wfm[sub_col[0]].fillna('SIN SUBESTACIÓN').astype(str) if sub_col else 'SIN SUBESTACIÓN'

    circ_col = [c for c in df_wfm.columns if 'circuito' in c.lower()]
    df_wfm['Circuito'] = df_wfm[circ_col[0]].fillna('SIN CIRCUITO').astype(str) if circ_col else 'SIN CIRCUITO'

    cuad_col = [c for c in df_wfm.columns if 'cuadrilla' in c.lower()]
    df_wfm['Cuadrilla'] = df_wfm[cuad_col[0]].fillna('SIN ASIGNAR').astype(str) if cuad_col else 'SIN ASIGNAR'

    # --- LÓGICA OPTIMIZADA DE AFECTADOS ---
    # Busca prioritariamente 'Afectados' (Afectación total del evento)
    afect_col = [c for c in df_wfm.columns if c.strip().lower() == 'afectados']
    if not afect_col:
        # Si no la encuentra exacta, busca coincidencia parcial
        afect_col = [c for c in df_wfm.columns if 'afectado' in c.lower() and 'critico' not in c.lower()]
    if not afect_col:
        # Fallback a 'Clientes no restaurados'
        afect_col = [c for c in df_wfm.columns if 'clientes no restaurados' in c.lower()]

    if afect_col:
        df_wfm['Clientes Sin Servicio'] = pd.to_numeric(df_wfm[afect_col[0]], errors='coerce').fillna(0).astype(int)
    else:
        df_wfm['Clientes Sin Servicio'] = 0

    df_wfm['Sector Operativo'] = df_wfm.apply(lambda r: obtener_sector_operativo(r['Circuito'], r['Subestación']), axis=1)

    # Parseo de Fechas y Cálculo de Tiempos
    f_col = [c for c in df_wfm.columns if 'fecha' in c.lower() and ('creaci' in c.lower() or 'inicio' in c.lower() or 'reporte' in c.lower())]
    h_col = [c for c in df_wfm.columns if 'hora' in c.lower() and 'creaci' in c.lower()]

    if f_col and h_col:
        df_wfm['Fecha_Creacion_DT'] = pd.to_datetime(df_wfm[f_col[0]].astype(str) + ' ' + df_wfm[h_col[0]].astype(str), errors='coerce')
    elif f_col:
        df_wfm['Fecha_Creacion_DT'] = pd.to_datetime(df_wfm[f_col[0]], errors='coerce')
    else:
        df_wfm['Fecha_Creacion_DT'] = pd.NaT

    ahora = pd.Timestamp.now()
    df_wfm['Horas_Transcurridas'] = ((ahora - df_wfm['Fecha_Creacion_DT']).dt.total_seconds() / 3600.0).fillna(0).round(2)

    def formato_hhmmss(horas_num):
        secs = int(max(0, horas_num) * 3600)
        h = secs // 3600
        m = (secs % 3600) // 60
        s = secs % 60
        return f"{h:02d}:{m:02d}:{s:02d}"

    df_wfm['Contador Horas'] = df_wfm['Horas_Transcurridas'].apply(formato_hhmmss)

    # Evaluación KPI 3
    kpi_evals = df_wfm.apply(lambda r: evaluar_kpi3(r['Clientes Sin Servicio'], r['Horas_Transcurridas']), axis=1)
    df_wfm['Rango Clientes'] = [k[0] for k in kpi_evals]
    df_wfm['Meta Horas KPI 3'] = [k[1] for k in kpi_evals]
    df_wfm['Cumple KPI 3'] = [k[2] for k in kpi_evals]
    df_wfm['Estado KPI 3'] = df_wfm['Cumple KPI 3'].apply(lambda c: 'Cumple 🟢' if c else 'Incumple 🔴')

    # ---------------------------------------------------------
    # 4. FILTROS DE ZONA Y CIRCUITO
    # ---------------------------------------------------------
    st.markdown('<div style="background:#fff; padding:1rem 1.2rem; border-radius:12px; border:1px solid #e2e8f0; margin-bottom:1.2rem;">', unsafe_allow_html=True)
    f1, f2, f3 = st.columns([2, 3, 1])

    with f1:
        sectores_op_disponibles = ['CENTRO', 'NORTE']
        sectores_op_sel = st.multiselect("🗺️ Zona / Sector Operativo:", options=sectores_op_disponibles, default=sectores_op_disponibles)

    with f2:
        circs_disponibles = sorted(df_wfm[df_wfm['Sector Operativo'].isin(sectores_op_sel)]['Circuito'].unique())
        circs_sel = st.multiselect("🔌 Circuito Actual:", options=circs_disponibles, default=circs_disponibles)

    with f3:
        st.markdown("<div style='height: 28px;'></div>", unsafe_allow_html=True)
        if st.button("🔄 Restablecer", use_container_width=True):
            sectores_op_sel = sectores_op_disponibles
            circs_sel = circs_disponibles

    st.markdown('</div>', unsafe_allow_html=True)

    df_filtrado = df_wfm[
        (df_wfm['Sector Operativo'].isin(sectores_op_sel)) &
        (df_wfm['Circuito'].isin(circs_sel))
    ].copy().reset_index(drop=True)

    if df_filtrado.empty:
        st.warning("⚠️ No se encontraron incidentes para la combinación de filtros seleccionada.")
        st.stop()

    # ---------------------------------------------------------
    # 5. PESTAÑAS DEL DASHBOARD EGAM
    # ---------------------------------------------------------
    tab_compacto, tab_kpi3 = st.tabs([
        "⏱️ Pestaña 1: Seguimiento WFM (Atención Inmediata)", 
        "📊 Pestaña 2: Control KPI 3 (Metas Rango de Clientes)"
    ])

    # =========================================================
    # PESTAÑA 1: VISTA COMPACTA TIPO IMAGEN
    # =========================================================
    with tab_compacto:
        cnt_avisos = len(df_filtrado)
        cnt_circuitos = df_filtrado['Circuito'].nunique()
        cnt_afectados = int(df_filtrado['Clientes Sin Servicio'].sum())

        # Lógica dinámica para la subetiqueta del KPI "Avisos por Zona"
        if len(sectores_op_sel) == len(sectores_op_disponibles):
            subetiqueta_zona = "Centro y Norte"
        elif len(sectores_op_sel) == 1:
            subetiqueta_zona = f"Sector {sectores_op_sel[0].title()}"
        else:
            subetiqueta_zona = "Sin Selección"

        k1, k2, k3, k4 = st.columns(4)
        with k1:
            st.markdown(f'''<div class="kpi-card"><div class="kpi-label">Contador Avisos</div><div class="kpi-val" style="color:#0284c7;">{cnt_avisos}</div><div class="kpi-sub">Avisos WFM Activos</div></div>''', unsafe_allow_html=True)
        with k2:
            st.markdown(f'''<div class="kpi-card"><div class="kpi-label">Circuitos Afectados</div><div class="kpi-val" style="color:#0f172a;">{cnt_circuitos}</div><div class="kpi-sub">Circuitos en seguimiento</div></div>''', unsafe_allow_html=True)
        with k3:
            st.markdown(f'''<div class="kpi-card"><div class="kpi-label">Avisos por Zona</div><div class="kpi-val" style="color:#0369a1;">{cnt_avisos}</div><div class="kpi-sub">📍 {subetiqueta_zona}</div></div>''', unsafe_allow_html=True)
        with k4:
            st.markdown(f'''<div class="kpi-card"><div class="kpi-label">Total Clientes Afectados</div><div class="kpi-val" style="color:#d97706;">{cnt_afectados:,}</div><div class="kpi-sub">Afectación total real</div></div>''', unsafe_allow_html=True)
        st.markdown("<br>", unsafe_allow_html=True)

        col1, col2, col3 = st.columns([1.1, 1.2, 1])

        # COLUMNA 1: CANTIDAD DE INCIDENTES POR CIRCUITO
        with col1:
            st.markdown("##### Cantidad de Incidentes por Circuitos")
            df_circ_summary = df_filtrado.groupby('Circuito', as_index=False)['Identificación'].count()
            df_circ_summary.columns = ['Circuito actual', 'Identificación']
            df_circ_summary = df_circ_summary.sort_values(by='Identificación', ascending=False).reset_index(drop=True)

            st.dataframe(df_circ_summary, use_container_width=True, height=380)
            st.markdown(f"**Total:** `{cnt_avisos}`")

        # COLUMNA 2: CONTROL AVISOS (CLIENTES AFECTADOS) + GRÁFICA DE DONA
        with col2:
            st.markdown("##### Control Avisos (Clientes Afectados)")
            df_avisos_summary = df_filtrado[['Identificación', 'Cuadrilla', 'Clientes Sin Servicio']].sort_values(by='Clientes Sin Servicio', ascending=False).reset_index(drop=True)
            df_avisos_summary.columns = ['Identificación', 'Cuadrillas', 'Afectados']

            st.dataframe(df_avisos_summary, use_container_width=True, height=220)
            st.markdown(f"**Total Afectados:** `{cnt_afectados:,}`")

            # Gráfica de Dona por Estado
            col_estado = [c for c in df_filtrado.columns if 'estado' in c.lower()]
            if col_estado:
                df_estado = df_filtrado[col_estado[0]].value_counts().reset_index()
                df_estado.columns = ['ESTADO', 'Cantidad']
                fig_dona = px.pie(df_estado, names='ESTADO', values='Cantidad', hole=0.55, color_discrete_sequence=['#0284c7', '#38bdf8', '#93c5fd'])
                fig_dona.update_traces(textposition='inside', textinfo='percent')
                fig_dona.update_layout(height=180, margin=dict(t=5, b=5, l=5, r=5), showlegend=True)
                st.plotly_chart(fig_dona, use_container_width=True)

        # COLUMNA 3: CONTADOR HORAS REPORTES
        with col3:
            st.markdown("##### Contador Horas Reportes")
            df_horas_summary = df_filtrado.sort_values(by='Horas_Transcurridas', ascending=False)[['Identificación', 'Contador Horas']].reset_index(drop=True)

            st.dataframe(df_horas_summary, use_container_width=True, height=380)
            st.markdown("**Estado:** `● WFM En Monitoreo`")

    # =========================================================
    # PESTAÑA 2: CONTROL KPI 3 (RANGOS DE CLIENTES)
    # =========================================================
    with tab_kpi3:
        st.subheader("📊 Control KPI 3 EGAM — Evaluado por Rango de Clientes y Horas")
        st.caption("Verificación estricta de cumplimiento según la tabla oficial de tiempos de atención.")

        cnt_cumple = int((df_filtrado['Cumple KPI 3'] == True).sum())
        cnt_incumple = int((df_filtrado['Cumple KPI 3'] == False).sum())
        pct_cumple = round((cnt_cumple / len(df_filtrado)) * 100, 1) if len(df_filtrado) > 0 else 0

        m1, m2, m3 = st.columns(3)
        with m1:
            st.markdown(f'''<div class="kpi-card"><div class="kpi-label">Cumplimiento General KPI 3</div><div class="kpi-val" style="color:#16a34a;">{pct_cumple}%</div><div class="kpi-sub">Incidentes a tiempo</div></div>''', unsafe_allow_html=True)
        with m2:
            st.markdown(f'''<div class="kpi-card"><div class="kpi-label">🟢 Cumplen Meta</div><div class="kpi-val" style="color:#15803d;">{cnt_cumple}</div><div class="kpi-sub">Dentro de la meta por horas</div></div>''', unsafe_allow_html=True)
        with m3:
            st.markdown(f'''<div class="kpi-card"><div class="kpi-label">🔴 Incumplen Meta</div><div class="kpi-val" style="color:#dc2626;">{cnt_incumple}</div><div class="kpi-sub">Superan la meta permitida</div></div>''', unsafe_allow_html=True)

        st.markdown("<br>", unsafe_allow_html=True)

        df_kpi_tabla = df_filtrado[[
            'Identificación', 'Sector Operativo', 'Subestación', 'Circuito', 
            'Clientes Sin Servicio', 'Rango Clientes', 'Contador Horas', 
            'Horas_Transcurridas', 'Meta Horas KPI 3', 'Estado KPI 3'
        ]].sort_values(by='Horas_Transcurridas', ascending=False).reset_index(drop=True)

        def colorear_kpi3(val):
            if 'Cumple 🟢' in str(val): return 'background-color: #dcfce7; color: #166534; font-weight: bold;'
            if 'Incumple 🔴' in str(val): return 'background-color: #fee2e2; color: #991b1b; font-weight: bold;'
            return ''

        st.dataframe(
            df_kpi_tabla.style.map(colorear_kpi3, subset=['Estado KPI 3']),
            use_container_width=True,
            height=450
        )

except Exception as e:
    st.error(f"❌ Error al procesar el reporte de WFM: {e}")
