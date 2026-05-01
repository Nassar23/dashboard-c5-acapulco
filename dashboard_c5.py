import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
import io
from datetime import datetime, timedelta

# Configuración de la página
st.set_page_config(page_title="Dashboard C5 Acapulco", page_icon="📹", layout="wide")

@st.cache_resource
def get_drive_service():
    """Conectar con Google Drive usando credenciales de Streamlit Secrets"""
    credentials = service_account.Credentials.from_service_account_info(
        st.secrets["gcp_service_account"],
        scopes=['https://www.googleapis.com/auth/drive.readonly']
    )
    return build('drive', 'v3', credentials=credentials)

@st.cache_data(ttl=300)
def listar_archivos_drive():
    """Listar todos los archivos Excel en la carpeta de Drive"""
    try:
        service = get_drive_service()
        folder_id = st.secrets["folder_id"]
        
        # Buscar TODOS los archivos
        query = f"'{folder_id}' in parents and trashed=false"
        
        results = service.files().list(
            q=query,
            fields="files(id, name, createdTime, modifiedTime, mimeType)",
            orderBy='createdTime desc'
        ).execute()
        
        archivos = results.get('files', [])
        
        # Filtrar solo Excel
        archivos_excel = [a for a in archivos if a['name'].endswith(('.xlsx', '.xlsm', '.xls'))]
        
        return archivos_excel
        
    except Exception as e:
        st.error(f"❌ Error al listar archivos: {str(e)}")
        return []

def descargar_archivo_drive(file_id):
    """Descargar archivo desde Google Drive"""
    service = get_drive_service()
    request = service.files().get_media(fileId=file_id)
    file_data = io.BytesIO()
    downloader = MediaIoBaseDownload(file_data, request)
    
    done = False
    while not done:
        status, done = downloader.next_chunk()
    
    file_data.seek(0)
    return file_data

@st.cache_data(ttl=300)
def cargar_datos():
    """Cargar datos desde Google Drive"""
    archivos = listar_archivos_drive()
    
    if not archivos:
        return None, [], []
    
    dataframes = []
    archivos_info = []
    errores = []
    
    for archivo in archivos:
        nombre_archivo = archivo['name']
        
        try:
            file_data = descargar_archivo_drive(archivo['id'])
            df = pd.read_excel(file_data, sheet_name='BARRIDO_ACTIVO', engine='openpyxl', header=3)
            
            # Limpiar nombres de columnas
            df.columns = df.columns.str.strip().str.replace('\n', ' ').str.upper()
            
            # Buscar columna de ID (puede tener diferentes nombres)
            col_id = None
            for col in df.columns:
                if 'CÁMARA' in col or 'CAMARA' in col or 'ID' in col:
                    col_id = col
                    break
            
            if col_id is None:
                errores.append(f"{nombre_archivo}: No se encontró columna de ID de cámara")
                continue
            
            # Renombrar columna para estandarizar
            df.rename(columns={col_id: 'ID_CAMARA'}, inplace=True)
            
            # Buscar columna de ESTADO
            col_estado = None
            for col in df.columns:
                if 'ESTADO' in col:
                    col_estado = col
                    break
            
            if col_estado:
                df.rename(columns={col_estado: 'ESTADO'}, inplace=True)
            else:
                df['ESTADO'] = 'DESCONOCIDO'
            
            # Limpiar datos
            df = df.dropna(subset=['ID_CAMARA'])
            df = df[df['ID_CAMARA'].astype(str).str.strip() != '']
            
            # Extraer fecha del nombre del archivo
            try:
                partes = nombre_archivo.split('_')
                fecha_parte = partes[-1].replace('.xlsm', '').replace('.xlsx', '').replace('.xls', '')
                
                if len(fecha_parte) == 6 and fecha_parte.isdigit():
                    dia = fecha_parte[0:2]
                    mes = fecha_parte[2:4]
                    anio = '20' + fecha_parte[4:6]
                    fecha_dt = pd.to_datetime(f"{anio}-{mes}-{dia}")
                    fecha_str = fecha_dt.strftime('%d/%m/%Y')
                elif len(fecha_parte) == 8 and fecha_parte.isdigit():
                    dia = fecha_parte[0:2]
                    mes = fecha_parte[2:4]
                    anio = fecha_parte[4:8]
                    fecha_dt = pd.to_datetime(f"{anio}-{mes}-{dia}")
                    fecha_str = fecha_dt.strftime('%d/%m/%Y')
                else:
                    fecha_dt = pd.to_datetime('today')
                    fecha_str = fecha_dt.strftime('%d/%m/%Y')
            except:
                fecha_dt = pd.to_datetime('today')
                fecha_str = fecha_dt.strftime('%d/%m/%Y')
            
            df['fecha_barrido'] = fecha_dt
            df['fecha_str'] = fecha_str
            df['archivo_origen'] = nombre_archivo
            
            dataframes.append(df)
            archivos_info.append({'nombre': nombre_archivo, 'fecha': fecha_str, 'registros': len(df)})
            
        except Exception as e:
            errores.append(f"{nombre_archivo}: {str(e)}")
    
    if dataframes:
        df_completo = pd.concat(dataframes, ignore_index=True)
        return df_completo, archivos_info, errores
    
    return None, [], errores

def crear_grafico_estado(df):
    """Crear gráfico de barras del estado de cámaras"""
    if 'ESTADO' not in df.columns:
        st.warning("⚠️ No se encontró la columna ESTADO")
        return None
    
    conteo = df['ESTADO'].value_counts().reset_index()
    conteo.columns = ['Estado', 'Cantidad']
    
    fig = px.bar(conteo, x='Estado', y='Cantidad', 
                 title='Estado General de Cámaras',
                 color='Estado',
                 color_discrete_map={'OPERANDO': '#00CC96', 'FUERA DE SERVICIO': '#EF553B', 'EN MANTENIMIENTO': '#FFA15A'})
    
    fig.update_layout(showlegend=False, height=400)
    return fig

def crear_grafico_tendencia(df):
    """Crear gráfico de tendencia temporal"""
    if 'fecha_barrido' not in df.columns or 'ESTADO' not in df.columns:
        return None
    
    tendencia = df.groupby(['fecha_barrido', 'ESTADO']).size().reset_index(name='Cantidad')
    
    fig = px.line(tendencia, x='fecha_barrido', y='Cantidad', color='ESTADO',
                  title='Tendencia de Estado de Cámaras',
                  labels={'fecha_barrido': 'Fecha', 'Cantidad': 'Número de Cámaras'})
    
    fig.update_layout(height=400)
    return fig

def crear_diagrama_gantt(df):
    """Crear diagrama de Gantt para visualizar disponibilidad de cámaras"""
    if 'ID_CAMARA' not in df.columns or 'fecha_barrido' not in df.columns:
        return None
    
    df_gantt = df[['ID_CAMARA', 'fecha_barrido', 'ESTADO']].copy()
    df_gantt = df_gantt.sort_values(['ID_CAMARA', 'fecha_barrido'])
    
    # Limitar a las primeras 50 cámaras
    camaras_top = df_gantt['ID_CAMARA'].unique()[:50]
    df_gantt = df_gantt[df_gantt['ID_CAMARA'].isin(camaras_top)]
    
    fig = px.timeline(df_gantt, x_start='fecha_barrido', x_end='fecha_barrido',
                      y='ID_CAMARA', color='ESTADO',
                      title='Historial de Estado por Cámara (Top 50)',
                      color_discrete_map={'OPERANDO': '#00CC96', 'FUERA DE SERVICIO': '#EF553B'})
    
    fig.update_layout(height=600, xaxis_title='Fecha', yaxis_title='Cámara')
    return fig

# Interfaz principal
st.title("📹 Dashboard de Monitoreo de Cámaras C5 - Acapulco")

# Sidebar
with st.sidebar:
    st.header("⚙️ Configuración")
    
    if st.button("🔄 Actualizar", use_container_width=True):
        st.cache_data.clear()
        st.rerun()
    
    st.info("📁 Conectado a Google Drive")

# Cargar datos
df, archivos_info, errores = cargar_datos()

if df is not None and len(df) > 0:
    
    # Mostrar columnas detectadas (debug)
    with st.expander("🔍 Columnas detectadas en los datos"):
        st.write(list(df.columns))
    
    # Métricas principales
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        total_camaras = len(df['ID_CAMARA'].unique())
        st.metric("📹 Total Cámaras", total_camaras)
    
    with col2:
        if 'ESTADO' in df.columns:
            operando = len(df[df['ESTADO'].str.contains('OPERANDO', case=False, na=False)])
            st.metric("✅ Operando", operando)
        else:
            st.metric("✅ Operando", "N/A")
    
    with col3:
        if 'ESTADO' in df.columns:
            fuera = len(df[df['ESTADO'].str.contains('FUERA', case=False, na=False)])
            st.metric("❌ Fuera de Servicio", fuera)
        else:
            st.metric("❌ Fuera de Servicio", "N/A")
    
    with col4:
        if 'ESTADO' in df.columns:
            operando = len(df[df['ESTADO'].str.contains('OPERANDO', case=False, na=False)])
            fuera = len(df[df['ESTADO'].str.contains('FUERA', case=False, na=False)])
            if operando + fuera > 0:
                disponibilidad = (operando / (operando + fuera)) * 100
                st.metric("📊 Disponibilidad", f"{disponibilidad:.1f}%")
            else:
                st.metric("📊 Disponibilidad", "N/A")
        else:
            st.metric("📊 Disponibilidad", "N/A")
    
    # Gráficos
    st.markdown("---")
    
    col1, col2 = st.columns(2)
    
    with col1:
        fig_estado = crear_grafico_estado(df)
        if fig_estado:
            st.plotly_chart(fig_estado, use_container_width=True)
    
    with col2:
        fig_tendencia = crear_grafico_tendencia(df)
        if fig_tendencia:
            st.plotly_chart(fig_tendencia, use_container_width=True)
    
    # Diagrama de Gantt
    st.markdown("---")
    fig_gantt = crear_diagrama_gantt(df)
    if fig_gantt:
        st.plotly_chart(fig_gantt, use_container_width=True)
    
    # Tabla de datos
    st.markdown("---")
    st.subheader("📋 Datos Detallados")
    
    columnas_mostrar = ['ID_CAMARA', 'ESTADO', 'fecha_str', 'archivo_origen']
    columnas_disponibles = [col for col in columnas_mostrar if col in df.columns]
    
    if columnas_disponibles:
        df_display = df[columnas_disponibles].copy()
        nombres_columnas = {'ID_CAMARA': 'Cámara', 'ESTADO': 'Estado', 'fecha_str': 'Fecha', 'archivo_origen': 'Archivo'}
        df_display.columns = [nombres_columnas.get(col, col) for col in df_display.columns]
        st.dataframe(df_display, use_container_width=True, height=400)
    
    # Información de archivos procesados
    with st.expander("📂 Archivos Procesados"):
        for info in archivos_info:
            st.success(f"✅ {info['nombre']} - {info['fecha']} ({info['registros']} registros)")
    
    # Errores
    if errores:
        with st.expander("⚠️ Errores Encontrados"):
            for error in errores:
                st.warning(error)

else:
    st.error("❌ No se encontraron datos en Google Drive")
    st.info("Verifica que hayas subido archivos Excel con el formato: C5_Acapulco_DDMMAA.xlsx")
    
    if errores:
        with st.expander("⚠️ Errores Encontrados"):
            for error in errores:
                st.warning(error)
