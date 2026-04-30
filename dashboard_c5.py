import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime
import io
from google.oauth2 import service_account
from googleapiclient.discovery import build

st.set_page_config(page_title="Monitor de Cámaras C5", page_icon="📹", layout="wide")

# ========== CONEXIÓN A GOOGLE DRIVE ==========
@st.cache_resource
def get_drive_service():
    """Conectar con Google Drive usando credenciales del secrets"""
    credentials = service_account.Credentials.from_service_account_info(
        st.secrets["gcp_service_account"],
        scopes=['https://www.googleapis.com/auth/drive.readonly']
    )
    return build('drive', 'v3', credentials=credentials)

@st.cache_data(ttl=300)
def listar_archivos_drive():
    """Listar todos los archivos Excel en la carpeta de Drive"""
    service = get_drive_service()
    folder_id = st.secrets["folder_id"]
    
    query = f"'{folder_id}' in parents and (mimeType='application/vnd.ms-excel' or mimeType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet') and trashed=false"
    
    results = service.files().list(
        q=query,
        fields="files(id, name, createdTime, modifiedTime)",
        orderBy='createdTime desc'
    ).execute()
    
    return results.get('files', [])

@st.cache_data(ttl=300)
def descargar_archivo_drive(file_id):
    """Descargar archivo Excel desde Drive"""
    service = get_drive_service()
    request = service.files().get_media(fileId=file_id)
    
    file_data = io.BytesIO()
    downloader = request.execute()
    file_data.write(downloader)
    file_data.seek(0)
    
    return file_data

# ========== FUNCIONES DE CARGA (ADAPTADAS) ==========
@st.cache_data(ttl=300)
def cargar_catalogo_fallas():
    """Carga el catálogo de fallas desde el primer archivo Excel encontrado"""
    archivos = listar_archivos_drive()
    
    if not archivos:
        return None
    
    try:
        file_data = descargar_archivo_drive(archivos[0]['id'])
        df_catalogo = pd.read_excel(file_data, sheet_name='CATÁLOGO_FALLAS', engine='openpyxl', header=1)
        df_catalogo.columns = df_catalogo.columns.str.strip()
        
        col_id = None
        col_desc = None
        
        for col in df_catalogo.columns:
            if 'ID' in col.upper() and 'FALLA' in col.upper():
                col_id = col
            if 'DESCRIP' in col.upper():
                col_desc = col
        
        if col_id and col_desc:
            df_catalogo = df_catalogo.dropna(subset=[col_id])
            df_catalogo = df_catalogo[df_catalogo[col_id].astype(str).str.strip() != '']
            
            catalogo_dict = dict(zip(
                df_catalogo[col_id].astype(str).str.strip(), 
                df_catalogo[col_desc].astype(str).str.strip()
            ))
            return catalogo_dict
            
    except Exception as e:
        st.sidebar.error(f"Error al cargar catálogo: {str(e)}")
    
    return None

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
        
        if '(' in nombre_archivo:
            continue
            
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
                fecha_parte = partes[-1].replace('.xlsm', '').replace('.xlsx', '')
                
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

def clasificar_estatus(estatus):
    estatus_str = str(estatus).upper()
    if any(palabra in estatus_str for palabra in ['OK', 'OPERATIV', 'FUNCIONAL', 'ACTIV']):
        return 'OPERATIVA'
    elif any(palabra in estatus_str for palabra in ['FALLA', 'ERROR', 'INACTIV', 'OFFLINE', 'FUERA', 'SIN']):
        return 'CON FALLA'
    else:
        return 'OTRO'

# ========== INTERFAZ PRINCIPAL ==========
st.title("📹 Dashboard de Monitoreo de Cámaras C5 - Acapulco")
st.markdown("---")

with st.sidebar:
    st.header("⚙️ Configuración")
    st.info("📁 Conectado a Google Drive")
    if st.button("🔄 Actualizar", use_container_width=True):
        st.cache_data.clear()
        st.rerun()
    st.markdown("---")

with st.spinner('🔄 Cargando datos desde Google Drive...'):
    df_completo, archivos_info, errores = cargar_datos()
    catalogo_fallas = cargar_catalogo_fallas()

if errores:
    with st.expander("⚠️ Advertencias"):
        for error in errores:
            st.warning(error)

if df_completo is None or len(df_completo) == 0:
    st.error("❌ No se encontraron datos en Google Drive")
    st.info("Verifica que hayas subido archivos Excel con el formato: C5_Acapulco_DDMMAA.xlsx")
    st.stop()

if 'ID CÁMARA' not in df_completo.columns:
    st.error("❌ Falta columna ID CÁMARA")
    st.stop()

for col in ['ZONA', 'SISTEMA', 'ESTATUS', 'IP', 'ID FALLA']:
    if col not in df_completo.columns:
        df_completo[col] = 'N/A'

df_completo['ESTATUS'] = df_completo['ESTATUS'].fillna('SIN ESTATUS').astype(str)
df_completo['ZONA'] = df_completo['ZONA'].fillna('SIN ZONA').astype(str)
df_completo['SISTEMA'] = df_completo['SISTEMA'].fillna('SIN SISTEMA').astype(str)
df_completo['ID FALLA'] = df_completo['ID FALLA'].fillna('N/A').astype(str)

df_completo['estatus_clasificado'] = df_completo['ESTATUS'].apply(clasificar_estatus)

if catalogo_fallas:
    df_completo['descripcion_falla'] = df_completo['ID FALLA'].map(catalogo_fallas).fillna(df_completo['ID FALLA'])
else:
    df_completo['descripcion_falla'] = df_completo['ID FALLA']

with st.sidebar:
    st.success(f"✅ {len(archivos_info)} archivos cargados")
    st.info(f"📊 {len(df_completo)} registros totales")
    st.info(f"📅 Periodo: {df_completo['fecha_str'].min()} al {df_completo['fecha_str'].max()}")
    if catalogo_fallas:
        st.success(f"📖 Catálogo: {len(catalogo_fallas)} tipos de falla")
    with st.expander("📄 Archivos procesados"):
        for info in archivos_info:
            st.text(f"📅 {info['fecha']}")
            st.text(f"   {info['registros']} registros")
            st.markdown("---")

df_ultimo = df_completo.sort_values('fecha_barrido', ascending=False).groupby('ID CÁMARA').first().reset_index()

st.subheader("📊 Resumen General")

col1, col2, col3, col4, col5 = st.columns(5)

total_camaras = df_completo['ID CÁMARA'].nunique()
total_barridos = len(archivos_info)
camaras_con_fallas = df_completo[df_completo['estatus_clasificado'] == 'CON FALLA']['ID CÁMARA'].nunique()
camaras_operativas = total_camaras - camaras_con_fallas
tasa_disponibilidad = (camaras_operativas / total_camaras * 100) if total_camaras > 0 else 0

col1.metric("🎥 Total Cámaras", total_camaras)
col2.metric("📋 Barridos Realizados", total_barridos)
col3.metric("✅ Siempre Operativas", camaras_operativas, delta=f"{tasa_disponibilidad:.1f}%")
col4.metric("❌ Con Fallas", camaras_con_fallas)
col5.metric("📊 Tasa de Fallas", f"{(camaras_con_fallas/total_camaras*100):.1f}%")

st.markdown("---")

tab1, tab2, tab3, tab4, tab5 = st.tabs(["📊 Estado Actual", "📈 Fallas por Día", "🔄 Fallas Recurrentes", "🗺️ Por Zona", "📋 Datos Completos"])

with tab1:
    st.subheader("Estado Actual (Último Barrido)")
    
    col_a, col_b = st.columns(2)
    
    with col_a:
        st.markdown("### Distribución por Estatus")
        ec = df_ultimo['estatus_clasificado'].value_counts()
        colores = {'OPERATIVA': '#00CC66', 'CON FALLA': '#FF4444', 'OTRO': '#FFA500'}
        fig1 = px.pie(
            values=ec.values, 
            names=ec.index, 
            hole=0.4,
            color=ec.index,
            color_discrete_map=colores
        )
        st.plotly_chart(fig1, use_container_width=True)
    
    with col_b:
        st.markdown("### Cámaras por Sistema")
        sc = df_ultimo['SISTEMA'].value_counts().head(10)
        fig2 = px.bar(x=sc.index, y=sc.values, color=sc.values, color_continuous_scale='Blues')
        fig2.update_layout(showlegend=False)
        st.plotly_chart(fig2, use_container_width=True)
    
    st.markdown("### 🚨 Cámaras con Problemas (Último Estado)")
    df_prob = df_ultimo[df_ultimo['estatus_clasificado'] == 'CON FALLA']
    
    if len(df_prob) > 0:
        st.dataframe(
            df_prob[['ID CÁMARA', 'ZONA', 'SISTEMA', 'ESTATUS', 'ID FALLA', 'descripcion_falla', 'fecha_str']].sort_values('ZONA'),
            use_container_width=True,
            height=300
        )
    else:
        st.success("✅ Todas las cámaras operativas en el último barrido")

with tab2:
    st.subheader("📈 Evolución de Fallas por Día")
    
    df_fallas_dia = df_completo[df_completo['estatus_clasificado'] == 'CON FALLA'].copy()
    
    if len(df_fallas_dia) > 0:
        df_agrupado = df_fallas_dia.groupby(['fecha_str', 'descripcion_falla']).size().reset_index(name='cantidad')
        df_agrupado['fecha_dt'] = pd.to_datetime(df_agrupado['fecha_str'], format='%d/%m/%Y')
        df_agrupado = df_agrupado.sort_values('fecha_dt')
        
        fig_fallas = px.bar(
            df_agrupado,
            x='fecha_str',
            y='cantidad',
            color='descripcion_falla',
            title='Fallas Detectadas por Día y Tipo',
            labels={'fecha_str': 'Fecha', 'cantidad': 'Número de Fallas', 'descripcion_falla': 'Tipo de Falla'}
        )
        fig_fallas.update_layout(xaxis_tickangle=-45, height=500)
        st.plotly_chart(fig_fallas, use_container_width=True)
        
        st.markdown("---")
        st.markdown("### 📊 Resumen de Fallas por Tipo")
        
        col_c, col_d = st.columns(2)
        
        with col_c:
            fallas_tipo = df_fallas_dia['descripcion_falla'].value_counts().head(10)
            fig_tipo = px.pie(
                values=fallas_tipo.values,
                names=fallas_tipo.index,
                title='Top 10 Tipos de Falla'
            )
            st.plotly_chart(fig_tipo, use_container_width=True)
        
        with col_d:
            st.markdown("#### Detalle de Fallas")
            fallas_detalle = df_fallas_dia.groupby(['ID FALLA', 'descripcion_falla']).size().reset_index(name='Total')
            fallas_detalle = fallas_detalle.sort_values('Total', ascending=False)
            st.dataframe(
                fallas_detalle.head(10),
                use_container_width=True,
                height=350
            )
    else:
        st.info("No se detectaron fallas en el periodo analizado")

with tab3:
    st.subheader("🔄 Análisis de Fallas Recurrentes")
    
    df_fallas_rec = df_completo[df_completo['estatus_clasificado'] == 'CON FALLA'].copy()
    
    if len(df_fallas_rec) > 0:
        fallas_por_camara = df_fallas_rec.groupby('ID CÁMARA').agg({
            'fecha_barrido': 'count',
            'fecha_str': lambda x: list(sorted(set(x))),
            'ZONA': 'first',
            'SISTEMA': 'first',
            'ID FALLA': lambda x: ', '.join(sorted(set(x))),
            'descripcion_falla': lambda x: ' | '.join(sorted(set(x)))
        }).reset_index()
        
        fallas_por_camara.columns = ['ID CÁMARA', 'Días con Falla', 'Fechas', 'ZONA', 'SISTEMA', 'Códigos Falla', 'Descripción Fallas']
        fallas_por_camara = fallas_por_camara.sort_values('Días con Falla', ascending=False)
        
        st.markdown("### 🔴 Top 30 Cámaras con Más Fallas")
        
        fig_recurrentes = go.Figure()
        
        top30 = fallas_por_camara.head(30)
        max_fallas = top30['Días con Falla'].max()
        
        fig_recurrentes.add_trace(go.Bar(
            x=top30['ID CÁMARA'],
            y=top30['Días con Falla'],
            marker=dict(
                color=top30['Días con Falla'],
                colorscale='Reds',
                cmin=0,
                cmax=max_fallas,
                colorbar=dict(title="Días<br>con Falla")
            ),
            text=top30['Días con Falla'],
            textposition='outside',
            hovertemplate='<b>%{x}</b><br>Días con falla: %{y}<br><extra></extra>'
        ))
        
        fig_recurrentes.update_layout(
            title='Top 30 Cámaras con Mayor Número de Fallas',
            xaxis_tickangle=-45,
            height=500,
            yaxis_title="Días con Falla",
            xaxis_title="ID Cámara",
            showlegend=False
        )
        st.plotly_chart(fig_recurrentes, use_container_width=True)
        
        st.markdown("---")
        st.markdown("### 📅 Línea de Tiempo de Fallas (Diagrama de Gantt)")
        
        col_gantt1, col_gantt2 = st.columns([1, 3])
        
        with col_gantt1:
            umbral_gantt = st.slider(
                "Días mínimos de falla:", 
                1, 
                int(fallas_por_camara['Días con Falla'].max()), 
                min(5, int(fallas_por_camara['Días con Falla'].max())),
                key="gantt_slider"
            )
            num_camaras = st.slider(
                "Número de cámaras:", 
                5, 
                30, 
                15, 
                key="num_camaras_gantt"
            )
            
            st.markdown("**Opciones de vista:**")
            agrupar_por_zona = st.checkbox("Agrupar por zona", value=False, key="agrupar_zona")
        
        with col_gantt2:
            st.info(f"📊 Mostrando las {num_camaras} cámaras con más fallas (≥{umbral_gantt} días)")
        
        df_gantt = fallas_por_camara[fallas_por_camara['Días con Falla'] >= umbral_gantt].head(num_camaras).copy()
        
        if len(df_gantt) > 0:
            gantt_data = []
            for idx, row in df_gantt.iterrows():
                fechas = row['Fechas']
                for fecha in fechas:
                    gantt_data.append({
                        'Cámara': row['ID CÁMARA'],
                        'Fecha': fecha,
                        'Zona': row['ZONA'],
                        'Sistema': row['SISTEMA'],
                        'Total_Fallas': row['Días con Falla']
                    })
            
            df_gantt_plot = pd.DataFrame(gantt_data)
            df_gantt_plot['Fecha_dt'] = pd.to_datetime(df_gantt_plot['Fecha'], format='%d/%m/%Y')
            
            if agrupar_por_zona:
                orden_camaras = df_gantt.sort_values(['ZONA', 'Días con Falla'], ascending=[True, False])['ID CÁMARA'].tolist()
            else:
                orden_camaras = df_gantt.sort_values('Días con Falla', ascending=False)['ID CÁMARA'].tolist()
            
            fig_gantt = go.Figure()
            
            colores_zona_map = {
                'ZONA 1': '#FF4444', 'ZONA 2': '#4444FF', 'ZONA 3': '#44FF44',
                'ZONA 4': '#FFAA00', 'ZONA 5': '#FF44FF', 'ZONA 6': '#00FFFF',
                'ZONA 7': '#FFFF44', 'ZONA 8': '#AA44FF', 'ZONA 9': '#44FFAA',
                'ZONA 10': '#FF8844'
            }
            
            zonas_unicas = sorted(df_gantt_plot['Zona'].unique())
            colores_extra = ['#FF6B6B', '#4ECDC4', '#45B7D1', '#FFA07A', '#98D8C8', '#F7DC6F']
            
            for i, zona in enumerate(zonas_unicas):
                if zona not in colores_zona_map:
                    colores_zona_map[zona] = colores_extra[i % len(colores_extra)]
            
            zonas_mostradas = set()
            
            for camara in orden_camaras:
                df_cam = df_gantt_plot[df_gantt_plot['Cámara'] == camara]
                zona = df_cam['Zona'].iloc[0]
                total_fallas = df_cam['Total_Fallas'].iloc[0]
                
                tamano_marker = 12 + (total_fallas / df_gantt['Días con Falla'].max() * 8)
                
                fig_gantt.add_trace(go.Scatter(
                    x=df_cam['Fecha_dt'],
                    y=[camara] * len(df_cam),
                    mode='markers',
                    marker=dict(
                        size=tamano_marker,
                        symbol='square',
                        color=colores_zona_map.get(zona, '#888888'),
                        line=dict(width=2, color='white'),
                        opacity=0.9
                    ),
                    name=zona,
                    legendgroup=zona,
                    showlegend=zona not in zonas_mostradas,
                    hovertemplate=f'<b>{camara}</b><br>Fecha: %{{x|%d/%m/%Y}}<br>Zona: {zona}<br>Total fallas: {total_fallas}<extra></extra>'
                ))
                zonas_mostradas.add(zona)
            
            fig_gantt.update_layout(
                title={
                    'text': 'Línea de Tiempo de Fallas por Cámara',
                    'font': {'size': 18, 'color': '#FFFFFF'}
                },
                xaxis_title="Fecha del Barrido",
                yaxis_title="ID Cámara",
                height=max(400, len(orden_camaras) * 30),
                hovermode='closest',
                showlegend=True,
                legend=dict(
                    title=dict(text="Zona", font=dict(size=14)),
                    orientation="v",
                    yanchor="top",
                    y=1,
                    xanchor="left",
                    x=1.02,
                    bgcolor='rgba(0,0,0,0.5)',
                    bordercolor='white',
                    borderwidth=1
                ),
                yaxis=dict(
                    categoryorder='array',
                    categoryarray=orden_camaras[::-1],
                    tickfont=dict(size=10)
                ),
                xaxis=dict(
                    tickformat='%d/%m',
                    dtick=86400000.0,
                    gridcolor='rgba(128,128,128,0.2)',
                    showgrid=True
                ),
                plot_bgcolor='rgba(0,0,0,0.1)',
                paper_bgcolor='rgba(0,0,0,0)'
            )
            
            st.plotly_chart(fig_gantt, use_container_width=True)
            
            col_exp1, col_exp2 = st.columns(2)
            
            with col_exp1:
                st.info("""
                **📖 Cómo leer el diagrama:**
                - **Eje Vertical (Y):** Cámaras ordenadas por fallas
                - **Eje Horizontal (X):** Fechas de barridos
                - **Cuadrados:** Día con falla detectada
                - **Tamaño:** Más grande = más fallas totales
                - **Color:** Zona geográfica
                """)
            
            with col_exp2:
                st.success("""
                **🔍 Patrones a identificar:**
                - **Línea continua:** Falla persistente
                - **Puntos aislados:** Fallas intermitentes
                - **Agrupación:** Problema en zona específica
                - **Inicio/fin:** Momento de aparición/solución
                """)
            
        else:
            st.warning(f"No hay cámaras con {umbral_gantt}+ días de falla")
        
        st.markdown("---")
        st.markdown("### 📋 Tabla Detallada de Fallas Recurrentes")
        
        col_tabla1, col_tabla2 = st.columns([1, 3])
        
        with col_tabla1:
            umbral_tabla = st.slider("Filtrar por días de falla:", 1, int(fallas_por_camara['Días con Falla'].max()), 2, key="tabla_slider")
            ordenar_por = st.selectbox("Ordenar por:", ["Días con Falla", "Zona", "ID Cámara"], key="ordenar_tabla")
        
        df_filtrado_recurrente = fallas_por_camara[fallas_por_camara['Días con Falla'] >= umbral_tabla].copy()
        df_filtrado_recurrente['Fechas_str'] = df_filtrado_recurrente['Fechas'].apply(lambda x: ', '.join(x))
        df_filtrado_recurrente['% Fallas'] = (df_filtrado_recurrente['Días con Falla'] / len(archivos_info) * 100).round(1)
        
        if ordenar_por == "Días con Falla":
            df_filtrado_recurrente = df_filtrado_recurrente.sort_values('Días con Falla', ascending=False)
        elif ordenar_por == "Zona":
            df_filtrado_recurrente = df_filtrado_recurrente.sort_values(['ZONA', 'Días con Falla'], ascending=[True, False])
        else:
            df_filtrado_recurrente = df_filtrado_recurrente.sort_values('ID CÁMARA')
        
        with col_tabla2:
            st.info(f"📊 {len(df_filtrado_recurrente)} cámaras con {umbral_tabla}+ días de falla")
        
        st.dataframe(
            df_filtrado_recurrente[['ID CÁMARA', 'Días con Falla', '% Fallas', 'ZONA', 'SISTEMA', 'Códigos Falla', 'Descripción Fallas', 'Fechas_str']],
            use_container_width=True,
            height=400
        )
        
        csv_recurrentes = df_filtrado_recurrente.to_csv(index=False, encoding='utf-8-sig').encode('utf-8-sig')
        st.download_button(
            "📥 Descargar Fallas Recurrentes",
            csv_recurrentes,
            f"fallas_recurrentes_{datetime.now().strftime('%Y%m%d')}.csv",
            "text/csv",
            key="download_recurrentes"
        )
    else:
        st.success("✅ No hay fallas recurrentes en el periodo analizado")

with tab4:
    st.subheader("🗺️ Análisis por Zona")
    
    col_e, col_f = st.columns(2)
    
    with col_e:
        st.markdown("### Cámaras por Zona")
        zc = df_ultimo['ZONA'].value_counts()
        fig_zonas = px.bar(
            x=zc.index,
            y=zc.values,
            color=zc.values,
            color_continuous_scale='Viridis',
            labels={'x': 'Zona', 'y': 'Cantidad'},
            text=zc.values
        )
        fig_zonas.update_traces(textposition='outside')
        st.plotly_chart(fig_zonas, use_container_width=True)
    
    with col_f:
        st.markdown("### Cámaras con Fallas por Zona")
        df_fallas_zona = df_completo[df_completo['estatus_clasificado'] == 'CON FALLA']
        if len(df_fallas_zona) > 0:
            fz = df_fallas_zona.groupby('ZONA')['ID CÁMARA'].nunique().reset_index(name='Cámaras con Fallas')
            fz = fz.sort_values('Cámaras con Fallas', ascending=False)
            fig_fz = px.bar(
                fz,
                x='ZONA',
                y='Cámaras con Fallas',
                color='Cámaras con Fallas',
                color_continuous_scale='Reds',
                text='Cámaras con Fallas'
            )
            fig_fz.update_traces(textposition='outside')
            st.plotly_chart(fig_fz, use_container_width=True)
        else:
            st.info("Sin fallas registradas")
    
    st.markdown("---")
    zona_sel = st.selectbox("Filtrar por zona:", ['Todas'] + sorted(df_ultimo['ZONA'].unique().tolist()), key="zona_tab4")
    df_z = df_completo if zona_sel == 'Todas' else df_completo[df_completo['ZONA'] == zona_sel]
    
    st.info(f"📊 {len(df_z)} registros en zona: {zona_sel}")
    st.dataframe(
        df_z[['ID CÁMARA', 'ZONA', 'SISTEMA', 'ESTATUS', 'estatus_clasificado', 'ID FALLA', 'descripcion_falla', 'fecha_str']].sort_values('fecha_str', ascending=False),
        use_container_width=True,
        height=400
    )

with tab5:
    st.subheader("🔍 Explorador de Datos Completo")
    
    col_g, col_h, col_i, col_j = st.columns(4)
    
    est_list = sorted([str(x) for x in df_completo['estatus_clasificado'].unique() if pd.notna(x)])
    zon_list = sorted([str(x) for x in df_completo['ZONA'].unique() if pd.notna(x)])
    fec_list = sorted([str(x) for x in df_completo['fecha_str'].unique() if pd.notna(x)], reverse=True)
    falla_list = sorted([str(x) for x in df_completo['ID FALLA'].unique() if pd.notna(x) and x != 'N/A'])
    
    with col_g:
        est_fil = st.selectbox("Estatus:", ['Todos'] + est_list, key="estatus_tab5")
    
    with col_h:
        zon_fil = st.selectbox("Zona:", ['Todas'] + zon_list, key="zona_tab5")
    
    with col_i:
        fec_fil = st.selectbox("Fecha:", ['Todas'] + fec_list, key="fecha_tab5")
    
    with col_j:
        falla_fil = st.selectbox("Tipo Falla:", ['Todas'] + falla_list, key="falla_tab5")
    
    buscar = st.text_input("🔍 Buscar cámara:", placeholder="ID...", key="buscar_tab5")
    
    df_fil = df_completo.copy()
    
    if est_fil != 'Todos':
        df_fil = df_fil[df_fil['estatus_clasificado'] == est_fil]
    
    if zon_fil != 'Todas':
        df_fil = df_fil[df_fil['ZONA'] == zon_fil]
    
    if fec_fil != 'Todas':
        df_fil = df_fil[df_fil['fecha_str'] == fec_fil]
    
    if falla_fil != 'Todas':
        df_fil = df_fil[df_fil['ID FALLA'].str.contains(falla_fil, case=False, na=False)]
    
    if buscar:
        df_fil = df_fil[df_fil['ID CÁMARA'].astype(str).str.contains(buscar, case=False, na=False)]
    
    st.info(f"📊 Mostrando {len(df_fil)} de {len(df_completo)} registros")
    
    st.dataframe(
        df_fil[['ID CÁMARA', 'ZONA', 'SISTEMA', 'ESTATUS', 'estatus_clasificado', 'ID FALLA', 'descripcion_falla', 'fecha_str']].sort_values('fecha_str', ascending=False),
        use_container_width=True,
        height=400
    )
    
    csv = df_fil.to_csv(index=False, encoding='utf-8-sig').encode('utf-8-sig')
    st.download_button(
        "📥 Descargar Datos Filtrados",
        csv,
        f"camaras_filtrado_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
        "text/csv",
        key="download_tab5"
    )

st.markdown("---")
st.caption(f"🕐 Actualizado: {datetime.now().strftime('%d/%m/%Y %H:%M')} | 📊 {len(df_completo)} registros | 📁 {len(archivos_info)} archivos | 📅 Periodo: {df_completo['fecha_str'].min()} - {df_completo['fecha_str'].max()}")
st.caption("Dashboard C5 Acapulco - Sistema de Monitoreo de Cámaras")
