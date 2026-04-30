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
        
        st.sidebar.info(f"🔍 Carpeta ID: {folder_id}")
        
        # Buscar TODOS los archivos (sin filtro de tipo)
        query = f"'{folder_id}' in parents and trashed=false"
        
        results = service.files().list(
            q=query,
            fields="files(id, name, createdTime, modifiedTime, mimeType)",
            orderBy='createdTime desc'
        ).execute()
        
        archivos = results.get('files', [])
        
        st.sidebar.success(f"✅ Total archivos: {len(archivos)}")
        
        # Filtrar solo Excel
        archivos_excel = [a for a in archivos if a['name'].endswith(('.xlsx', '.xlsm', '.xls'))]
        
        st.sidebar.info(f"📊 Archivos Excel: {len(archivos_excel)}")
        
        if archivos_excel:
            st.sidebar.write("**Primeros 5 archivos:**")
            for i, arch in enumerate(archivos_excel[:5]):
                st.sidebar.text(f"{i+1}. {arch['name']}")
        else:
            st.sidebar.warning("⚠️ No hay archivos Excel")
            if archivos:
                st.sidebar.write("**Archivos encontrados (otros tipos):**")
                for i, arch in enumerate(archivos[:5]):
                    st.sidebar.text(f"{i+1}. {arch['name']} ({arch.get('mimeType', 'unknown')})")
        
        return archivos_excel
        
    except Exception as e:
        st.sidebar.error(f"❌ Error: {str(e)}")
        import traceback
        st.sidebar.code(traceback.format_exc())
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
    
    # Debug
    st.sidebar.info(f"📂 Procesando {len(archivos)} archivos...")
    
    if not archivos:
        st.sidebar.warning("⚠️ Lista de archivos vacía")
        return None, [], []
    
    dataframes = []
    archivos_info = []
    errores = []
    
    for archivo in archivos:
        nombre_archivo = archivo['name']
        
        # Debug
        st.sidebar.text(f"📄 Procesando: {nombre_archivo}")
        
        try:
            file_data = descargar_archivo_drive(archivo['id'])
            df = pd.read_excel(file_data, sheet_name='BARRIDO_ACTIVO', engine='openpyxl', header=3)
            df.columns = df.columns.str.strip().str.replace('\n', ' ')
            
            if 'ID CÁMARA' not in df.columns:
                errores.append(f"{nombre_archivo}: Sin columna ID CÁMARA")
                continue
            
            df = df.dropna(subset=['ID CÁMARA'])
            df = df[df['ID CÁMARA'].astype(str).str.strip() != '']
            
            try:
                partes = nombre_archivo.split('_')
                fecha_parte = partes[-1].replace('.xlsm', '').replace('.xlsx', '').replace('.xls', '')
                
                if len(fecha_parte) == 6 and fecha_parte.isdigit():
                    dia = fecha_parte[0:2]
                    mes = fecha_parte[2:4]
                    anio = '20' + fecha_parte[4:6]
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
    if 'ID CÁMARA' not in df.columns or 'fecha_barrido' not in df.columns:
        return None
    
    df_gantt = df[['ID CÁMARA', 'fecha_barrido', 'ESTADO']].copy()
    df_gantt = df_gantt.sort_values(['ID CÁMARA', 'fecha_barrido'])
    
    # Limitar a las primeras 50 cámaras para mejor visualización
    camaras_top = df_gantt['ID CÁMARA'].unique()[:50]
    df_gantt = df_gantt[df_gantt['ID CÁMARA'].isin(camaras_top)]
    
    fig = px.timeline(df_gantt, x_start='fecha_barrido', x_end='fecha_barrido',
                      y='ID CÁMARA', color='ESTADO',
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
    # Métricas principales
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.metric("📹 Total Cámaras", len(df['ID CÁMARA'].unique()))
    
    with col2:
        operando = len(df[df['ESTADO'] == 'OPERANDO'])
        st.metric("✅ Operando", operando)
    
    with col3:
        fuera = len(df[df['ESTADO'] == 'FUERA DE SERVICIO'])
        st.metric("❌ Fuera de Servicio", fuera)
    
    with col4:
        if operando + fuera > 0:
            disponibilidad = (operando / (operando + fuera)) * 100
            st.metric("📊 Disponibilidad", f"{disponibilidad:.1f}%")
    
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
    
    if 'ID CÁMARA' in df.columns:
        df_display = df[['ID CÁMARA', 'ESTADO', 'fecha_str', 'archivo_origen']].copy()
        df_display.columns = ['Cámara', 'Estado', 'Fecha', 'Archivo']
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
