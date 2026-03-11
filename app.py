import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime

st.set_page_config(page_title="Planillas KPI CM", page_icon="📊", layout="centered")

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=DM+Sans:wght@300;400;500;600&family=DM+Mono&display=swap');

* { font-family: 'DM Sans', sans-serif; }

.stApp {
    background: #0f1117;
    color: #e8e8e8;
}

h1 {
    font-size: 1.6rem !important;
    font-weight: 600 !important;
    color: #ffffff !important;
    letter-spacing: -0.5px;
}

.subtitle {
    color: #6b7280;
    font-size: 0.9rem;
    margin-top: -12px;
    margin-bottom: 32px;
}

.upload-label {
    font-size: 0.8rem;
    font-weight: 500;
    color: #9ca3af;
    text-transform: uppercase;
    letter-spacing: 1px;
    margin-bottom: 8px;
}

.file-ok {
    color: #34d399;
    font-size: 0.85rem;
    display: flex;
    align-items: center;
    gap: 6px;
}

section[data-testid="stFileUploadDropzone"] {
    background: #13151f !important;
    border: 1.5px dashed #2a2d3a !important;
    border-radius: 8px !important;
}

.stButton > button {
    background: #4f6ef7 !important;
    color: white !important;
    border: none !important;
    border-radius: 8px !important;
    padding: 12px 32px !important;
    font-size: 0.95rem !important;
    font-weight: 500 !important;
    width: 100%;
    transition: all 0.2s !important;
}

.stButton > button:hover {
    background: #3d5ce0 !important;
    transform: translateY(-1px);
}

.stButton > button:disabled {
    background: #2a2d3a !important;
    color: #4b5563 !important;
}

.status-bar {
    background: #1a1d27;
    border: 1px solid #2a2d3a;
    border-radius: 8px;
    padding: 14px 18px;
    font-size: 0.85rem;
    color: #9ca3af;
    margin-top: 16px;
    font-family: 'DM Mono', monospace;
}

.divider {
    border: none;
    border-top: 1px solid #1e2130;
    margin: 24px 0;
}

.stDownloadButton > button {
    background: #1a2a1a !important;
    color: #34d399 !important;
    border: 1.5px solid #34d399 !important;
    border-radius: 8px !important;
    width: 100%;
    font-weight: 500 !important;
}

.metric-card {
    background: #1a1d27;
    border: 1px solid #2a2d3a;
    border-radius: 12px;
    padding: 20px 24px;
    text-align: center;
}

.metric-value {
    font-size: 2.2rem;
    font-weight: 600;
    color: #4f6ef7;
    font-family: 'DM Mono', monospace;
    line-height: 1;
}

.metric-label {
    font-size: 0.8rem;
    color: #6b7280;
    text-transform: uppercase;
    letter-spacing: 1px;
    margin-top: 6px;
}

div[data-testid="stSelectbox"] > div {
    background: #13151f !important;
    border: 1px solid #2a2d3a !important;
    border-radius: 8px !important;
    color: #e8e8e8 !important;
}
</style>
""", unsafe_allow_html=True)

# ── Helpers ───────────────────────────────────────────────────────────────────

MESES = {
    1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril",
    5: "Mayo", 6: "Junio", 7: "Julio", 8: "Agosto",
    9: "Septiembre", 10: "Octubre", 11: "Noviembre", 12: "Diciembre"
}

def encontrar_col(df, opciones):
    for op in opciones:
        for col in df.columns:
            if col.strip().lower() == op.strip().lower():
                return col
    return None

def filtrar_por_mes(df, campo_fecha, mes, anio):
    col = encontrar_col(df, [campo_fecha])
    if col is None:
        return pd.DataFrame()
    serie = pd.to_datetime(df[col], dayfirst=True, errors='coerce')
    mask = (serie.dt.month == mes) & (serie.dt.year == anio)
    return df[mask].copy()

def formatear_fecha(val):
    if pd.isna(val) or val is None or val == '':
        return ''
    try:
        d = pd.to_datetime(val, dayfirst=True, errors='coerce')
        if pd.isna(d):
            return str(val)
        return d.strftime('%d/%m/%Y')
    except:
        return str(val)

def generar_cm_presentados(df_pres, df_ofic, mes, anio):
    fp = filtrar_por_mes(df_pres, 'Fecha de Presentacion', mes, anio)
    fo = filtrar_por_mes(df_ofic, 'Fecha de presentacion', mes, anio)
    filas = []
    for df in [fp, fo]:
        for _, row in df.iterrows():
            def get(opts, _row=row, _df=df):
                c = encontrar_col(_df, opts)
                return _row[c] if c else ''
            filas.append({
                'Operación':  get(['REFERENCIA', 'Referencia']),
                'Facturas':   get(['FACTURAS', 'Facturas']),
                'Expediente': get(['Expediente', 'EXPEDIENTE']),
                'Ult evento': formatear_fecha(get(['ULTIMO EVENTO', 'ultimo evento'])),
                'TAD SUBIDO': formatear_fecha(get(['Fecha de Presentacion', 'Fecha de presentacion'])),
            })
    return pd.DataFrame(filas)

def generar_cm_aprobados(df_pres, df_ofic, mes, anio):
    fp = filtrar_por_mes(df_pres, 'FECHA DE APROBACION', mes, anio)
    fo = filtrar_por_mes(df_ofic, 'Fecha de aprobacion', mes, anio)
    filas = []
    for df in [fp, fo]:
        for _, row in df.iterrows():
            def get(opts, _row=row, _df=df):
                c = encontrar_col(_df, opts)
                return _row[c] if c else ''
            filas.append({
                'Referencia':        get(['REFERENCIA', 'Referencia']),
                'Facturas':          get(['FACTURAS', 'Facturas']),
                'Expediente':        get(['Expediente', 'EXPEDIENTE']),
                'Fecha':             formatear_fecha(get(['Fecha de Presentacion', 'Fecha de presentacion'])),
                'Fechadeaprobacion': formatear_fecha(get(['Fecha de aprobacion', 'FECHA DE APROBACION'])),
            })
    return pd.DataFrame(filas)

def df_a_excel(df):
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine='openpyxl') as writer:
        df.to_excel(writer, index=False)
    return buf.getvalue()

# ── UI ────────────────────────────────────────────────────────────────────────

st.markdown("# Planillas KPI CM")
st.markdown('<p class="subtitle">Generador de CM Presentados · CM Aprobados</p>', unsafe_allow_html=True)
st.markdown('<hr class="divider">', unsafe_allow_html=True)

# 1. Archivo
st.markdown('<div class="upload-label">📂 Planilla madre</div>', unsafe_allow_html=True)
archivo = st.file_uploader("Planilla madre", type=["xlsx", "xls"], label_visibility="collapsed")

if archivo:
    st.markdown(f'<div class="file-ok">✓ {archivo.name}</div>', unsafe_allow_html=True)

    try:
        xls = pd.ExcelFile(archivo)
        if 'PRESENTACIONES' not in xls.sheet_names or 'OFICIALIZADOS' not in xls.sheet_names:
            st.error(f"No se encontraron las solapas requeridas. Encontradas: {', '.join(xls.sheet_names)}")
            st.stop()
        df_pres = pd.read_excel(xls, sheet_name='PRESENTACIONES')
        df_ofic = pd.read_excel(xls, sheet_name='OFICIALIZADOS')
    except Exception as e:
        st.error(f"Error al leer el archivo: {e}")
        st.stop()

    st.markdown('<hr class="divider">', unsafe_allow_html=True)

    # 2. Período
    st.markdown('<div class="upload-label">📅 Período a analizar</div>', unsafe_allow_html=True)
    col1, col2 = st.columns(2)
    with col1:
        mes = st.selectbox("Mes", options=list(MESES.keys()), format_func=lambda x: MESES[x], index=datetime.now().month - 1)
    with col2:
        anio = st.selectbox("Año", options=list(range(datetime.now().year, 2019, -1)))

    st.markdown('<hr class="divider">', unsafe_allow_html=True)

    # 3. Vista previa
    df_presentados = generar_cm_presentados(df_pres, df_ofic, mes, anio)
    df_aprobados   = generar_cm_aprobados(df_pres, df_ofic, mes, anio)

    c1, c2 = st.columns(2)
    with c1:
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-value">{len(df_presentados)}</div>
            <div class="metric-label">CM Presentados</div>
        </div>""", unsafe_allow_html=True)
    with c2:
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-value">{len(df_aprobados)}</div>
            <div class="metric-label">CM Aprobados</div>
        </div>""", unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    if len(df_presentados) == 0 and len(df_aprobados) == 0:
        st.markdown(f'<div class="status-bar">⚠️ Sin registros para {MESES[mes]} {anio}</div>', unsafe_allow_html=True)
        st.stop()

    # 4. Descargas
    periodo = f"{str(mes).zfill(2)}-{anio}"
    col3, col4 = st.columns(2)
    with col3:
        st.download_button(
            label="⬇ CM PRESENTADOS",
            data=df_a_excel(df_presentados),
            file_name=f"CM_PRESENTADOS_{periodo}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
    with col4:
        st.download_button(
            label="⬇ CM APROBADOS",
            data=df_a_excel(df_aprobados),
            file_name=f"CM_APROBADOS_{periodo}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
else:
    st.markdown('<div class="status-bar">⏳ Esperando planilla madre...</div>', unsafe_allow_html=True)
