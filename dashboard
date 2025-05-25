import streamlit as st
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from io import BytesIO
import plotly.express as px

# Configurar el acceso a Google Sheets
scope = [
    "https://spreadsheets.google.com/feeds",
    "https://www.googleapis.com/auth/drive"
]

# Cargar las credenciales
creds = ServiceAccountCredentials.from_json_keyfile_name("credenciales.json", scope)
cliente = gspread.authorize(creds)

# Abrir la hoja de Google Sheets
spreadsheet = cliente.open("OrdenesMaura")
hoja = spreadsheet.sheet1

# Cargar los datos
datos = hoja.get_all_records()
df = pd.DataFrame(datos)

# Convertir columna de fecha si existe
if 'FECHA' in df.columns:
    df['FECHA'] = pd.to_datetime(df['FECHA'], errors='coerce')

# Interfaz de Streamlit
st.set_page_config(layout="wide")
st.title("Dashboard OrdenesMaura")

# Filtros
st.sidebar.header("Filtros")
clientes = st.sidebar.multiselect("Filtrar por Cliente (HNAME):", options=df['HNAME'].unique())
productos = st.sidebar.multiselect("Filtrar por Producto (LPROD):", options=df['LPROD'].unique())
responsables = st.sidebar.multiselect("Filtrar por Responsable:", options=df['RESPONSABLE'].unique() if 'RESPONSABLE' in df.columns else [])
orden_id = st.sidebar.text_input("Buscar por código LORD:")

# Filtro de rango de fechas
if 'FECHA' in df.columns:
    st.sidebar.markdown("---")
    fecha_min = df['FECHA'].min()
    fecha_max = df['FECHA'].max()
    rango_fecha = st.sidebar.date_input("Filtrar por rango de fechas:", value=(fecha_min, fecha_max))
    if isinstance(rango_fecha, tuple) and len(rango_fecha) == 2:
        df = df[(df['FECHA'] >= pd.to_datetime(rango_fecha[0])) & (df['FECHA'] <= pd.to_datetime(rango_fecha[1]))]

# Aplicar filtros
if clientes:
    df = df[df['HNAME'].isin(clientes)]
if productos:
    df = df[df['LPROD'].isin(productos)]
if responsables:
    df = df[df['RESPONSABLE'].isin(responsables)]
if orden_id:
    df = df[df['LORD'].astype(str).str.contains(orden_id, case=False)]

# Mostrar tabla completa con scroll
st.subheader("Vista general de datos")
st.dataframe(df, use_container_width=True)

# Métricas rápidas
st.subheader("Resumen de Solicitudes")
col1, col2, col3 = st.columns(3)
col1.metric("Total Solicitudes", len(df))
col2.metric("Productos únicos (LPROD)", df['LPROD'].nunique())
col3.metric("Clientes únicos (HNAME)", df['HNAME'].nunique())

# Gráfico: solicitudes por producto
st.subheader("Solicitudes por Producto (LPROD)")
st.bar_chart(df['LPROD'].value_counts())

# Gráfico: solicitudes por cliente
st.subheader("Solicitudes por Cliente (HNAME)")
st.bar_chart(df['HNAME'].value_counts())

# Gráfico: solicitudes por Responsable (si existe)
if 'RESPONSABLE' in df.columns:
    st.subheader("Solicitudes por Responsable")
    st.bar_chart(df['RESPONSABLE'].value_counts())

# Gráfico de líneas por fecha (si FECHA existe)
if 'FECHA' in df.columns:
    st.subheader("Solicitudes en el tiempo")
    df_fecha = df.dropna(subset=['FECHA'])
    solicitudes_por_dia = df_fecha.groupby('FECHA').size().reset_index(name='Solicitudes')
    fig_linea = px.line(solicitudes_por_dia, x='FECHA', y='Solicitudes', title="Solicitudes por Fecha")
    st.plotly_chart(fig_linea, use_container_width=True)

# Tabla dinámica por Cliente y Responsable
if 'RESPONSABLE' in df.columns:
    st.subheader("Resumen por Cliente y Responsable")
    pivot_df = df.pivot_table(index=['HNAME', 'RESPONSABLE'], aggfunc='size').reset_index(name='Total')
    st.dataframe(pivot_df, use_container_width=True)

# Botón para exportar a CSV y Excel
st.subheader("Exportar Datos Filtrados")
@st.cache_data
def convertir_csv(dataframe):
    return dataframe.to_csv(index=False).encode('utf-8')

@st.cache_data
def convertir_excel(dataframe):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        dataframe.to_excel(writer, index=False, sheet_name='Ordenes')
    return output.getvalue()

csv = convertir_csv(df)
excel = convertir_excel(df)

col_csv, col_excel = st.columns(2)
col_csv.download_button(
    label="Descargar como CSV",
    data=csv,
    file_name='ordenes_filtradas.csv',
    mime='text/csv'
)

col_excel.download_button(
    label="Descargar como Excel",
    data=excel,
    file_name='ordenes_filtradas.xlsx',
    mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
)
