import streamlit as st
import pandas as pd
from io import BytesIO
from datetime import datetime

st.set_page_config(page_title="Planillas KPI CM", page_icon="📊", layout="centered")

st.title("📊 Planillas KPI CM")
st.caption("Generador de CM Presentados y CM Aprobados")

# ── Helpers ──────────────────────────────────────────────────────────────────

MESES = {
    1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril",
    5: "Mayo", 6: "Junio", 7: "Julio", 8: "Agosto",
    9: "Septiembre", 10: "Octubre", 11: "Noviembre", 12: "Diciembre"
}

def encontrar_col(df, opciones):
    """Busca una columna en el df de forma case-insensitive."""
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
            get = lambda opts: row.get(encontrar_col(df, opts) or '', '')
            filas.append({
                'Operación':   get(['REFERENCIA', 'Referencia']),
                'Facturas':    get(['FACTURAS', 'Facturas']),
                'Expediente':  get(['Expediente', 'EXPEDIENTE']),
                'Ult evento':  formatear_fecha(get(['ULTIMO EVENTO', 'ultimo evento'])),
                'TAD SUBIDO':  formatear_fecha(get(['Fecha de Presentacion', 'Fecha de presentacion'])),
            })
    return pd.DataFrame(filas)

def generar_cm_aprobados(df_pres, df_ofic, mes, anio):
    fp = filtrar_por_mes(df_pres, 'FECHA DE APROBACION', mes, anio)
    fo = filtrar_por_mes(df_ofic, 'Fecha de aprobacion', mes, anio)
    filas = []
    for df in [fp, fo]:
        for _, row in df.iterrows():
            get = lambda opts: row.get(encontrar_col(df, opts) or '', '')
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

# ── UI ───────────────────────────────────────────────────────────────────────

# 1. Subir archivo
st.subheader("1 · Planilla madre")
archivo = st.file_uploader("Subí el archivo Excel", type=["xlsx", "xls"])

if archivo:
    try:
        xls = pd.ExcelFile(archivo)
        if 'PRESENTACIONES' not in xls.sheet_names or 'OFICIALIZADOS' not in xls.sheet_names:
            st.error(f"No se encontraron las solapas requeridas. Solapas encontradas: {', '.join(xls.sheet_names)}")
            st.stop()

        df_pres = pd.read_excel(xls, sheet_name='PRESENTACIONES')
        df_ofic = pd.read_excel(xls, sheet_name='OFICIALIZADOS')
        total = len(df_pres) + len(df_ofic)
        st.success(f"✅ {archivo.name} — {total:,} filas totales")

    except Exception as e:
        st.error(f"Error al leer el archivo: {e}")
        st.stop()

    # 2. Período
    st.subheader("2 · Período a analizar")
    col1, col2 = st.columns(2)
    with col1:
        mes = st.selectbox("Mes", options=list(MESES.keys()), format_func=lambda x: MESES[x], index=datetime.now().month - 1)
    with col2:
        anio = st.selectbox("Año", options=list(range(datetime.now().year, 2019, -1)))

    # 3. Vista previa
    st.subheader("3 · Vista previa")
    df_presentados = generar_cm_presentados(df_pres, df_ofic, mes, anio)
    df_aprobados = generar_cm_aprobados(df_pres, df_ofic, mes, anio)

    c1, c2 = st.columns(2)
    c1.metric("CM Presentados", len(df_presentados), help="Filtrado por Fecha de Presentacion")
    c2.metric("CM Aprobados", len(df_aprobados), help="Filtrado por Fecha de Aprobacion")

    if len(df_presentados) == 0 and len(df_aprobados) == 0:
        st.warning(f"No se encontraron registros para {MESES[mes]} {anio}")
        st.stop()

    # 4. Generar y descargar
    st.subheader("4 · Descargar reportes")
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
