"""
Dashboard C5 Acapulco - Monitoreo de Cámaras
Versión 4.0 - Lee el catálogo PMI maestro + barridos diarios desde Google Drive

Estructura de carpeta en Drive:
  - Catalogo_PMI_Acapulco_DDMMAA.xlsx  → catálogo maestro (toma el más reciente)
  - C5_Acapulco_DDMMAA.xlsx            → barridos diarios (todos)
"""
import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload
import io
import re
from datetime import datetime
from collections import defaultdict

# ============================================================================
# CONFIGURACIÓN GENERAL
# ============================================================================
st.set_page_config(
    page_title="Monitor C5 Acapulco",
    page_icon="📹",
    layout="wide",
    initial_sidebar_state="auto"
)

# CSS responsivo
st.markdown("""
<style>
    .main .block-container {
        padding-top: 1rem;
        padding-bottom: 1rem;
        padding-left: 1rem;
        padding-right: 1rem;
        max-width: 100%;
    }
    [data-testid="stMetricValue"] { font-size: clamp(1.2rem, 3.5vw, 2rem); }
    [data-testid="stMetricLabel"] { font-size: clamp(0.75rem, 2vw, 0.95rem); }
    .stTabs [data-baseweb="tab-list"] { flex-wrap: wrap; gap: 4px; }
    .stTabs [data-baseweb="tab"] {
        font-size: clamp(0.7rem, 1.8vw, 0.95rem);
        padding: 6px 10px;
    }
    .dataframe { font-size: clamp(0.7rem, 1.8vw, 0.9rem); }
    @media (max-width: 768px) {
        .main .block-container { padding-left: 0.4rem; padding-right: 0.4rem; }
        h1 { font-size: 1.4rem !important; }
        h2 { font-size: 1.15rem !important; }
        h3 { font-size: 1rem !important; }
    }
    .stAlert { padding: 0.5rem 1rem; }
</style>
""", unsafe_allow_html=True)

# ============================================================================
# CONEXIÓN A GOOGLE DRIVE
# ============================================================================

@st.cache_resource
def get_drive_service():
    creds = service_account.Credentials.from_service_account_info(
        st.secrets["gcp_service_account"],
        scopes=['https://www.googleapis.com/auth/drive.readonly']
    )
    return build('drive', 'v3', credentials=creds)


@st.cache_data(ttl=300)
def listar_archivos_drive():
    """Lista todos los Excel de la carpeta y los clasifica en catálogo / barridos"""
    try:
        service = get_drive_service()
        folder_id = st.secrets["folder_id"]
        query = f"'{folder_id}' in parents and trashed=false"
        results = service.files().list(
            q=query,
            fields="files(id, name, modifiedTime)",
            orderBy='modifiedTime desc',
            pageSize=1000
        ).execute()
        archivos = results.get('files', [])
        excels = [a for a in archivos if a['name'].lower().endswith(('.xlsx', '.xlsm', '.xls'))]
        # Saltar duplicados con paréntesis y archivos temporales
        excels = [a for a in excels if '(' not in a['name'] and not a['name'].startswith('~$')]
        
        # Detección flexible del catálogo: cualquier archivo cuyo nombre contenga
        # 'catalogo', 'catálogo', 'pmi' o 'maestro' (case-insensitive)
        # y que NO empiece con 'C5_Acapulco' (esos son barridos)
        def es_catalogo(nombre):
            n = nombre.lower()
            if n.startswith('c5_acapulco'):
                return False
            return any(kw in n for kw in ['catalogo', 'catálogo', 'pmi', 'maestro'])
        
        catalogos = [a for a in excels if es_catalogo(a['name'])]
        barridos = [a for a in excels if a not in catalogos]
        return catalogos, barridos
    except Exception as e:
        st.error(f"❌ Error al listar archivos: {str(e)}")
        return [], []


def descargar_archivo_drive(file_id):
    service = get_drive_service()
    request = service.files().get_media(fileId=file_id)
    file_data = io.BytesIO()
    downloader = MediaIoBaseDownload(file_data, request)
    done = False
    while not done:
        _, done = downloader.next_chunk()
    file_data.seek(0)
    return file_data


# ============================================================================
# PROCESAMIENTO DEL CATÁLOGO PMI MAESTRO
# ============================================================================

def _extraer_ips_principales(texto):
    if pd.isna(texto): return []
    s = re.sub(r'\[.*?\]', '', str(texto))
    return re.findall(r'\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}', s)


@st.cache_data(ttl=300)
def cargar_catalogo_pmi():
    """Carga el catálogo PMI más reciente y construye el universo de dispositivos."""
    catalogos, _ = listar_archivos_drive()
    if not catalogos:
        return None, None, None
    
    # Tomar el más reciente
    cat_archivo = catalogos[0]
    file_data = descargar_archivo_drive(cat_archivo['id'])
    
    # Hoja PMI
    df_pmi = pd.read_excel(file_data, sheet_name='PMI', engine='openpyxl', header=0)
    
    # Construir universo de dispositivos por IP
    registros = []
    for _, row in df_pmi.iterrows():
        pmi = row['Nomenclatura PMI']
        zona_num = row['Zona']
        zona_str = f"ZONA {int(zona_num)}" if pd.notna(zona_num) else None
        info_base = {
            'PMI': pmi, 'ZONA': zona_str, 'MUNICIPIO': row['Municipio'],
            'LOCALIDAD': row['Localidad'], 'COLONIA': row['Colonia'],
            'DIRECCION': row['Dirección completa'],
            'LATITUD': row['Latitud'], 'LONGITUD': row['Longitud'],
        }
        ips_lpr = set(_extraer_ips_principales(row['IPs c/LPR (con IP Hikvision alt.)']))
        ips_facial = set(_extraer_ips_principales(row['IPs c/Facial']))
        
        for ip in _extraer_ips_principales(row['IPs PTZ (con IP Hikvision alt.)']):
            registros.append({
                'ID_DISPOSITIVO': f"CAM-{ip}", 'IP': ip, 'TIPO': 'PTZ',
                'TIENE_LPR': 'Sí' if ip in ips_lpr else 'No',
                'TIENE_FACIAL': 'Sí' if ip in ips_facial else 'No',
                **info_base,
            })
        for ip in _extraer_ips_principales(row['IPs FIJA']):
            registros.append({
                'ID_DISPOSITIVO': f"CAM-{ip}", 'IP': ip, 'TIPO': 'FIJA',
                'TIENE_LPR': 'No', 'TIENE_FACIAL': 'No',
                **info_base,
            })
        for ip in _extraer_ips_principales(row['IPs Sin Clasif']):
            registros.append({
                'ID_DISPOSITIVO': f"CAM-{ip}", 'IP': ip, 'TIPO': 'Sin_Clasif',
                'TIENE_LPR': 'No', 'TIENE_FACIAL': 'No',
                **info_base,
            })
        # Botones del PMI
        if pd.notna(row.get('ID Botón')):
            try:
                bot_num = int(row['ID Botón'])
                ip_ref = (_extraer_ips_principales(row['IPs PTZ (con IP Hikvision alt.)']) +
                         _extraer_ips_principales(row['IPs FIJA']))
                ip_ref = ip_ref[0] if ip_ref else None
                registros.append({
                    'ID_DISPOSITIVO': f"BOT-{bot_num:03d}", 'IP': ip_ref, 'TIPO': 'BOTÓN',
                    'TIENE_LPR': 'No', 'TIENE_FACIAL': 'No',
                    **info_base,
                })
            except (ValueError, TypeError):
                pass
    
    df_universo = pd.DataFrame(registros).drop_duplicates(subset=['ID_DISPOSITIVO']).reset_index(drop=True)
    
    # Catálogo de fallas
    try:
        file_data.seek(0)
        df_fallas = pd.read_excel(file_data, sheet_name='CATÁLOGO_FALLAS', engine='openpyxl', header=1)
        df_fallas.columns = df_fallas.columns.str.strip()
        df_fallas = df_fallas.dropna(subset=['ID_FALLA'])
    except Exception:
        df_fallas = None
    
    return df_universo, df_fallas, cat_archivo['name']


# ============================================================================
# PROCESAMIENTO DE BARRIDOS
# ============================================================================

def _detectar_columnas_barrido(df):
    """Detecta columnas según versión de plantilla (vieja o nueva)."""
    cols = {c.upper().strip(): c for c in df.columns}
    
    # Plantilla nueva (v3): ID_DISPOSITIVO, ESTATUS, ID_FALLA, MARCA, etc.
    if 'ID_DISPOSITIVO' in cols:
        return {
            'id': cols['ID_DISPOSITIVO'],
            'ip': cols.get('IP'),
            'pmi': cols.get('PMI'),
            'tipo': cols.get('TIPO'),
            'marca': cols.get('MARCA'),
            'zona': cols.get('ZONA'),
            'municipio': cols.get('MUNICIPIO'),
            'localidad': cols.get('LOCALIDAD'),
            'tiene_lpr': cols.get('TIENE_LPR'),
            'tiene_facial': cols.get('TIENE_FACIAL'),
            'verificar': cols.get('VERIFICAR_EXIST') or cols.get('VERIFICAR_EXISTENCIA'),
            'estatus': cols.get('ESTATUS'),
            'id_falla': cols.get('ID_FALLA'),
            'descripcion_falla': cols.get('DESCRIPCIÓN_FALLA') or cols.get('DESCRIPCION_FALLA'),
            'severidad': cols.get('SEVERIDAD'),
            'existe_fisica': cols.get('EXISTE_FISICAMENTE'),
            'pmi_real': cols.get('PMI_REAL'),
            'obs': cols.get('OBS_VISUAL'),
            'version': 'v3'
        }
    # Plantilla vieja: ID CÁMARA con header en fila 4
    if 'ID CÁMARA' in cols:
        return {
            'id': cols['ID CÁMARA'],
            'estatus': cols.get('ESTATUS'),
            'id_falla': cols.get('ID FALLA'),
            'zona': cols.get('ZONA'),
            'tipo': cols.get('TIPO'),
            'version': 'vieja'
        }
    return None


@st.cache_data(ttl=300)
def cargar_barridos(_universo_hash):
    """Carga TODOS los barridos disponibles. Detecta automáticamente versión nueva o vieja."""
    _, barridos = listar_archivos_drive()
    if not barridos:
        return None, [], []
    
    dataframes = []
    archivos_info = []
    errores = []
    
    for archivo in barridos:
        nombre = archivo['name']
        nombre_sin_ext = nombre.rsplit('.', 1)[0]
        try:
            file_data = descargar_archivo_drive(archivo['id'])
            
            # Intentar leer con header=2 (nueva v3) primero
            df = None
            for hdr in [2, 3]:
                try:
                    file_data.seek(0)
                    df_test = pd.read_excel(file_data, sheet_name='BARRIDO', engine='openpyxl', header=hdr) \
                        if hdr == 2 else pd.read_excel(file_data, sheet_name='BARRIDO_ACTIVO', engine='openpyxl', header=hdr)
                    df_test.columns = [str(c).strip().replace('\n', ' ') for c in df_test.columns]
                    cols_map = _detectar_columnas_barrido(df_test)
                    if cols_map:
                        df = df_test
                        break
                except Exception:
                    continue
            
            if df is None:
                errores.append(f"{nombre}: No se pudo identificar la estructura del barrido")
                continue
            
            cols_map = _detectar_columnas_barrido(df)
            
            # Renombrar columnas a nombres estándar
            rename_dict = {}
            for std_name, real_name in cols_map.items():
                if std_name == 'version': continue
                if real_name and real_name in df.columns:
                    rename_dict[real_name] = std_name.upper()
            df = df.rename(columns=rename_dict)
            
            # Limpiar
            id_col = 'ID' if cols_map['version'] == 'v3' else 'ID'  # ya renombrado
            df = df.dropna(subset=[id_col])
            df = df[df[id_col].astype(str).str.strip() != '']
            df = df[~df[id_col].astype(str).str.contains('REPETIDA', na=False)]
            
            # Extraer fecha del nombre
            partes = nombre_sin_ext.split('_')
            fecha_str_parte = partes[-1]
            try:
                if len(fecha_str_parte) == 6 and fecha_str_parte.isdigit():
                    dia, mes, anio = fecha_str_parte[:2], fecha_str_parte[2:4], '20' + fecha_str_parte[4:6]
                    fecha_dt = pd.to_datetime(f"{anio}-{mes}-{dia}")
                elif len(fecha_str_parte) == 8 and fecha_str_parte.isdigit():
                    dia, mes, anio = fecha_str_parte[:2], fecha_str_parte[2:4], fecha_str_parte[4:8]
                    fecha_dt = pd.to_datetime(f"{anio}-{mes}-{dia}")
                else:
                    fecha_dt = pd.to_datetime(archivo['modifiedTime']).tz_localize(None)
            except Exception:
                fecha_dt = pd.to_datetime(archivo['modifiedTime']).tz_localize(None)
            
            df['FECHA_BARRIDO'] = fecha_dt
            df['FECHA_STR'] = fecha_dt.strftime('%d/%m/%Y')
            df['ARCHIVO_ORIGEN'] = nombre_sin_ext
            df['VERSION_PLANTILLA'] = cols_map['version']
            
            dataframes.append(df)
            archivos_info.append({
                'nombre': nombre_sin_ext, 'fecha': fecha_dt.strftime('%d/%m/%Y'),
                'registros': len(df), 'version': cols_map['version']
            })
        except Exception as e:
            errores.append(f"{nombre}: {str(e)}")
    
    if not dataframes:
        return None, archivos_info, errores
    
    df_completo = pd.concat(dataframes, ignore_index=True)
    return df_completo, archivos_info, errores


# ============================================================================
# CRUCE: UNIVERSO + BARRIDO → ESTADO COMPLETO
# ============================================================================

def construir_estado_completo(df_universo, df_barrido):
    """
    Hace OUTER JOIN del universo (catálogo PMI) con el barrido más reciente.
    - Cámaras del catálogo NO presentes en barrido → SIN DATO
    - Cámaras del barrido NO presentes en catálogo → marcadas FUENTE='Solo en barrido'
      (sucede con BOT-XXX y LPR-XXX que no están en el catálogo PMI)
    """
    if df_universo is None or len(df_universo) == 0:
        return None
    
    if df_barrido is None or len(df_barrido) == 0:
        df = df_universo.copy()
        df['ESTATUS'] = 'SIN DATO'
        df['MARCA'] = 'POR VERIFICAR'
        df['ID_FALLA'] = None
        df['DESCRIPCION_FALLA'] = None
        df['SEVERIDAD'] = None
        df['VERIFICAR'] = None
        df['OBS'] = None
        df['FECHA_BARRIDO'] = pd.NaT
        df['FECHA_STR'] = ''
        df['FUENTE'] = 'Solo en catálogo'
        return df
    
    # Tomar el último barrido por dispositivo
    df_barr_ult = df_barrido.sort_values('FECHA_BARRIDO', ascending=False).drop_duplicates(subset=['ID']).copy()
    df_barr_ult = df_barr_ult.rename(columns={'ID': 'ID_DISPOSITIVO'})
    
    # Columnas del barrido a traer (incluyendo metadatos del catálogo embebido en la plantilla v3)
    cols_barr = ['ID_DISPOSITIVO', 'ESTATUS', 'ID_FALLA', 'DESCRIPCION_FALLA',
                 'SEVERIDAD', 'MARCA', 'VERIFICAR', 'EXISTE_FISICA', 'PMI_REAL',
                 'OBS', 'FECHA_BARRIDO', 'FECHA_STR',
                 # Metadatos del catálogo embebido (para los huérfanos)
                 'IP', 'PMI', 'TIPO', 'ZONA', 'MUNICIPIO', 'LOCALIDAD',
                 'TIENE_LPR', 'TIENE_FACIAL']
    cols_disponibles = [c for c in cols_barr if c in df_barr_ult.columns]
    
    df_b_subset = df_barr_ult[cols_disponibles].copy()
    
    # Marcar IDs del barrido que NO están en el catálogo
    ids_universo = set(df_universo['ID_DISPOSITIVO'])
    ids_barrido = set(df_b_subset['ID_DISPOSITIVO'].astype(str))
    huerfanos = ids_barrido - ids_universo
    
    # 1) LEFT JOIN: cámaras del catálogo (con datos del barrido si existen)
    cols_join = [c for c in cols_disponibles if c not in 
                 ['IP','PMI','TIPO','ZONA','MUNICIPIO','LOCALIDAD','TIENE_LPR','TIENE_FACIAL']]
    df_left = df_universo.merge(df_b_subset[cols_join], on='ID_DISPOSITIVO', how='left')
    df_left['FUENTE'] = 'Catálogo PMI'
    
    # 2) Agregar huérfanos del barrido (BOT-XXX y LPR-XXX sin PMI)
    if huerfanos:
        df_huerf = df_b_subset[df_b_subset['ID_DISPOSITIVO'].astype(str).isin(huerfanos)].copy()
        # Asegurar columnas del universo
        for col in ['PMI', 'TIPO', 'ZONA', 'MUNICIPIO', 'LOCALIDAD', 'COLONIA',
                    'DIRECCION', 'LATITUD', 'LONGITUD', 'TIENE_LPR', 'TIENE_FACIAL', 'IP']:
            if col not in df_huerf.columns:
                df_huerf[col] = None
        df_huerf['FUENTE'] = 'Solo en barrido'
        # Asegurar todas las columnas del df_left para concatenar limpiamente
        for col in df_left.columns:
            if col not in df_huerf.columns:
                df_huerf[col] = None
        df_huerf = df_huerf[df_left.columns]
        df = pd.concat([df_left, df_huerf], ignore_index=True)
    else:
        df = df_left
    
    df['ESTATUS'] = df['ESTATUS'].fillna('SIN DATO')
    if 'MARCA' in df.columns:
        df['MARCA'] = df['MARCA'].fillna('POR VERIFICAR')
    else:
        df['MARCA'] = 'POR VERIFICAR'
    return df


def clasificar_estatus(s):
    s = str(s).upper().strip()
    if 'OK' in s and 'SIN DATO' not in s: return 'OPERATIVA'
    if 'SIN SEÑAL' in s or 'SIN SENAL' in s: return 'CON FALLA'
    if 'OTRO' in s: return 'CON FALLA'
    if 'FALLA' in s: return 'CON FALLA'
    if 'SIN DATO' in s: return 'SIN DATO'
    return 'SIN DATO'


# ============================================================================
# INTERFAZ PRINCIPAL
# ============================================================================

st.title("📹 Dashboard C5 Acapulco — Monitoreo de Cámaras")

# Sidebar
with st.sidebar:
    st.header("⚙️ Configuración")
    st.info("📁 Conectado a Google Drive")
    if st.button("🔄 Actualizar Datos", use_container_width=True):
        st.cache_data.clear()
        st.cache_resource.clear()
        st.rerun()
    st.markdown("---")

# Cargar
with st.spinner('Cargando catálogo PMI desde Google Drive...'):
    df_universo, df_fallas, nombre_catalogo = cargar_catalogo_pmi()

if df_universo is None:
    st.error("❌ No se encontró el catálogo PMI en la carpeta de Drive.")
    st.info("El catálogo debe llamarse: `Catalogo_PMI_Acapulco_DDMMAA.xlsx`")
    
    # Diagnóstico: mostrar qué archivos SÍ se ven en la carpeta
    with st.expander("🔍 Diagnóstico — archivos detectados en Drive", expanded=True):
        try:
            cats_diag, barr_diag = listar_archivos_drive()
            st.write(f"**Archivos catálogo detectados:** {len(cats_diag)}")
            if cats_diag:
                for a in cats_diag:
                    st.text(f"   📚 {a['name']}")
            else:
                st.warning("   No se detectó ningún archivo cuyo nombre contenga 'catalogo' o 'catálogo'")
            
            st.write(f"**Archivos barrido detectados:** {len(barr_diag)}")
            if barr_diag:
                for a in barr_diag:
                    st.text(f"   📋 {a['name']}")
            else:
                st.warning("   No se detectó ningún archivo de barrido")
            
            if not cats_diag and not barr_diag:
                st.error("⚠️ No se ve NINGÚN archivo en la carpeta. Posibles causas:")
                st.markdown("""
                - La cuenta de servicio (`xxx@xxx.iam.gserviceaccount.com`) no tiene acceso a la carpeta. Verifica en Drive: clic derecho en la carpeta → Compartir → agregar el email de la cuenta de servicio como Lector.
                - El `folder_id` en `secrets.toml` apunta a otra carpeta.
                - La carpeta está vacía.
                """)
            elif not cats_diag and barr_diag:
                st.warning("⚠️ Falta subir el catálogo PMI. Sube `Catalogo_PMI_Acapulco_270426.xlsx` a la misma carpeta donde están los barridos.")
        except Exception as e:
            st.error(f"Error al diagnosticar: {str(e)}")
    
    st.stop()

with st.spinner('Cargando barridos diarios...'):
    df_barrido, archivos_info, errores = cargar_barridos(len(df_universo))

# Construir estado completo
df_estado = construir_estado_completo(df_universo, df_barrido)
df_estado['ESTATUS_CLASIF'] = df_estado['ESTATUS'].apply(clasificar_estatus)

# Sidebar: info de archivos
with st.sidebar:
    st.success(f"📚 Catálogo: `{nombre_catalogo}`")
    st.info(f"🎯 Universo: {len(df_universo)} dispositivos")
    if archivos_info:
        st.success(f"📋 Barridos cargados: {len(archivos_info)}")
        with st.expander("📄 Ver archivos"):
            for info in archivos_info:
                st.text(f"📅 {info['fecha']} ({info['registros']} reg.)")
    else:
        st.warning("⚠️ No hay archivos de barrido aún")
    
    if errores:
        with st.expander(f"⚠️ {len(errores)} advertencias"):
            for e in errores: st.warning(e)

# ============================================================================
# PESTAÑAS
# ============================================================================
tab1, tab2, tab3, tab4, tab5, tab6, tab7, tab8 = st.tabs([
    "📊 Resumen",
    "📋 Cobertura",
    "⚙️ Estado",
    "🎯 Analíticos",
    "🔁 Histórico",
    "🗺️ Geográfico",
    "🔍 Explorador",
    "🛠️ Fallas Atendidas"
])

# ----------------------------------------------------------------------------
# TAB 1: RESUMEN EJECUTIVO
# ----------------------------------------------------------------------------
with tab1:
    st.subheader("🏛️ Universo Instalado (Catálogo PMI)")
    
    n_pmi = df_universo['PMI'].nunique()
    n_cam = (df_universo['TIPO'].isin(['PTZ', 'FIJA', 'Sin_Clasif'])).sum()
    n_bot = (df_universo['TIPO'] == 'BOTÓN').sum()
    n_lpr = (df_universo['TIENE_LPR'] == 'Sí').sum()
    n_facial = (df_universo['TIENE_FACIAL'] == 'Sí').sum()
    
    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("📍 PMIs (Postes)", n_pmi)
    c2.metric("🎥 Cámaras", n_cam)
    c3.metric("🔘 Botones", n_bot)
    c4.metric("🚗 c/ LPR", n_lpr)
    c5.metric("👤 c/ Facial", n_facial)
    
    # Por tipo de cámara
    n_ptz = (df_universo['TIPO'] == 'PTZ').sum()
    n_fija = (df_universo['TIPO'] == 'FIJA').sum()
    n_sc = (df_universo['TIPO'] == 'Sin_Clasif').sum()
    st.caption(f"Cámaras desglosadas: **{n_ptz}** PTZ • **{n_fija}** Fijas • **{n_sc}** Sin clasificar")
    
    st.markdown("---")
    st.subheader("📋 Estado del Último Barrido")
    
    if df_barrido is not None and len(df_barrido) > 0:
        ultima_fecha = df_estado['FECHA_STR'].dropna().mode()[0] if df_estado['FECHA_STR'].notna().any() else 'N/D'
        st.caption(f"Datos del barrido: **{ultima_fecha}**")
        
        n_op = (df_estado['ESTATUS_CLASIF'] == 'OPERATIVA').sum()
        n_falla = (df_estado['ESTATUS_CLASIF'] == 'CON FALLA').sum()
        n_sd = (df_estado['ESTATUS_CLASIF'] == 'SIN DATO').sum()
        total = len(df_estado)
        cobertura = ((n_op + n_falla) / total * 100) if total else 0
        disponibilidad = (n_op / (n_op + n_falla) * 100) if (n_op + n_falla) else 0
        
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("✅ Operativas", n_op, f"{n_op/total*100:.1f}%")
        c2.metric("❌ Con Falla", n_falla, f"{n_falla/total*100:.1f}%", delta_color="inverse")
        c3.metric("⚪ Sin Dato", n_sd, f"{n_sd/total*100:.1f}%", delta_color="off")
        c4.metric("📊 Disponibilidad", f"{disponibilidad:.1f}%", f"Cobertura: {cobertura:.1f}%")
        
        # Gráficas
        col_a, col_b = st.columns(2)
        
        with col_a:
            st.markdown("##### Estado General")
            datos_pie = pd.DataFrame({
                'Estado': ['Operativas', 'Con Falla', 'Sin Dato'],
                'Cantidad': [n_op, n_falla, n_sd]
            })
            datos_pie = datos_pie[datos_pie['Cantidad'] > 0]
            fig = px.pie(datos_pie, values='Cantidad', names='Estado', hole=0.45,
                        color='Estado',
                        color_discrete_map={'Operativas': '#00CC66', 'Con Falla': '#FF4444', 'Sin Dato': '#A0A0A0'})
            fig.update_layout(height=320, margin=dict(l=10, r=10, t=20, b=10),
                              legend=dict(orientation='h', y=-0.15))
            st.plotly_chart(fig, use_container_width=True)
        
        with col_b:
            st.markdown("##### Distribución por Marca")
            if 'MARCA' in df_estado.columns:
                marca_counts = df_estado['MARCA'].value_counts().reset_index()
                marca_counts.columns = ['Marca', 'Cantidad']
                fig = px.bar(marca_counts, x='Marca', y='Cantidad', text='Cantidad',
                            color='Marca',
                            color_discrete_map={
                                'PANASONIC': '#1F4E78', 'HIKVISION': '#666666',
                                'GENETEC': '#548235', 'POR VERIFICAR': '#FFC000',
                                'OTRA': '#9C5700'
                            })
                fig.update_traces(textposition='outside')
                fig.update_layout(height=320, margin=dict(l=10, r=10, t=20, b=10), showlegend=False)
                st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("📭 Aún no hay barridos cargados. Sube uno a Google Drive y refresca.")

# ----------------------------------------------------------------------------
# TAB 2: COBERTURA
# ----------------------------------------------------------------------------
with tab2:
    st.subheader("📋 Cobertura del Barrido")
    st.caption("¿Cuántas cámaras del catálogo se revisaron vs cuántas faltaron?")
    
    if df_barrido is None or len(df_barrido) == 0:
        st.info("📭 No hay datos de barrido para analizar cobertura.")
    else:
        # Cobertura general
        n_op = (df_estado['ESTATUS_CLASIF'] == 'OPERATIVA').sum()
        n_falla = (df_estado['ESTATUS_CLASIF'] == 'CON FALLA').sum()
        n_sd = (df_estado['ESTATUS_CLASIF'] == 'SIN DATO').sum()
        total = len(df_estado)
        revisadas = n_op + n_falla
        
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("🎯 Universo Total", total)
        c2.metric("✅ Revisadas", revisadas, f"{revisadas/total*100:.1f}%")
        c3.metric("⚪ Pendientes", n_sd, f"{n_sd/total*100:.1f}%", delta_color="inverse")
        if 'VERIFICAR' in df_estado.columns:
            n_verif = (df_estado['VERIFICAR'].astype(str).str.upper() == 'SÍ').sum()
            c4.metric("⚠️ Por Verificar", n_verif, "Existencia física")
        
        st.markdown("---")
        col_a, col_b = st.columns(2)
        
        with col_a:
            st.markdown("##### Cobertura por Zona")
            cob_zona = df_estado.groupby('ZONA').agg(
                Total=('ID_DISPOSITIVO', 'count'),
                Revisadas=('ESTATUS_CLASIF', lambda x: (x != 'SIN DATO').sum())
            ).reset_index()
            cob_zona['% Cobertura'] = (cob_zona['Revisadas'] / cob_zona['Total'] * 100).round(1)
            cob_zona = cob_zona.sort_values('ZONA')
            fig = px.bar(cob_zona, x='ZONA', y='% Cobertura', text='% Cobertura',
                        color='% Cobertura', color_continuous_scale='RdYlGn',
                        range_color=[0, 100])
            fig.update_traces(textposition='outside')
            fig.update_layout(height=350, margin=dict(l=10, r=10, t=20, b=10),
                              yaxis_range=[0, 110])
            st.plotly_chart(fig, use_container_width=True)
        
        with col_b:
            st.markdown("##### Cobertura por Municipio")
            cob_muni = df_estado.groupby('MUNICIPIO').agg(
                Total=('ID_DISPOSITIVO', 'count'),
                Revisadas=('ESTATUS_CLASIF', lambda x: (x != 'SIN DATO').sum())
            ).reset_index()
            cob_muni['% Cobertura'] = (cob_muni['Revisadas'] / cob_muni['Total'] * 100).round(1)
            fig = px.bar(cob_muni, x='MUNICIPIO', y='% Cobertura', text='% Cobertura',
                        color='% Cobertura', color_continuous_scale='RdYlGn',
                        range_color=[0, 100])
            fig.update_traces(textposition='outside')
            fig.update_layout(height=350, margin=dict(l=10, r=10, t=20, b=10),
                              yaxis_range=[0, 110])
            st.plotly_chart(fig, use_container_width=True)
        
        st.markdown("---")
        st.markdown("##### 🔍 Cámaras Pendientes de Revisar (SIN DATO)")
        df_pend = df_estado[df_estado['ESTATUS_CLASIF'] == 'SIN DATO'].copy()
        if len(df_pend) > 0:
            cols_show = ['ID_DISPOSITIVO', 'PMI', 'TIPO', 'MARCA', 'ZONA', 'MUNICIPIO', 'LOCALIDAD']
            cols_disp = [c for c in cols_show if c in df_pend.columns]
            st.dataframe(df_pend[cols_disp].sort_values(['ZONA','PMI']),
                        use_container_width=True, height=320)
            csv = df_pend[cols_disp].to_csv(index=False, encoding='utf-8-sig').encode('utf-8-sig')
            st.download_button("📥 Descargar Pendientes", csv,
                              f"pendientes_{datetime.now().strftime('%Y%m%d')}.csv",
                              "text/csv", key="dwn_pend")
        else:
            st.success("✅ Todas las cámaras del catálogo fueron revisadas")

# ----------------------------------------------------------------------------
# TAB 3: ESTADO OPERATIVO
# ----------------------------------------------------------------------------
with tab3:
    st.subheader("⚙️ Estado Operativo Detallado")
    
    if df_barrido is None:
        st.info("📭 Sin datos de barrido.")
    else:
        col_a, col_b = st.columns(2)
        
        with col_a:
            st.markdown("##### Operatividad por Tipo")
            df_solo_rev = df_estado[df_estado['ESTATUS_CLASIF'] != 'SIN DATO']
            if len(df_solo_rev) > 0:
                op_tipo = df_solo_rev.groupby('TIPO').agg(
                    Total=('ID_DISPOSITIVO', 'count'),
                    Operativas=('ESTATUS_CLASIF', lambda x: (x == 'OPERATIVA').sum())
                ).reset_index()
                op_tipo['% Op'] = (op_tipo['Operativas'] / op_tipo['Total'] * 100).round(1)
                fig = px.bar(op_tipo, x='TIPO', y='% Op', text='% Op',
                            color='% Op', color_continuous_scale='RdYlGn',
                            range_color=[0, 100])
                fig.update_traces(textposition='outside')
                fig.update_layout(height=350, margin=dict(l=10, r=10, t=20, b=10),
                                  yaxis_range=[0, 110])
                st.plotly_chart(fig, use_container_width=True)
        
        with col_b:
            st.markdown("##### Operatividad por Marca")
            if 'MARCA' in df_solo_rev.columns and len(df_solo_rev) > 0:
                op_marca = df_solo_rev.groupby('MARCA').agg(
                    Total=('ID_DISPOSITIVO', 'count'),
                    Operativas=('ESTATUS_CLASIF', lambda x: (x == 'OPERATIVA').sum())
                ).reset_index()
                op_marca['% Op'] = (op_marca['Operativas'] / op_marca['Total'] * 100).round(1)
                fig = px.bar(op_marca, x='MARCA', y='% Op', text='% Op',
                            color='% Op', color_continuous_scale='RdYlGn',
                            range_color=[0, 100])
                fig.update_traces(textposition='outside')
                fig.update_layout(height=350, margin=dict(l=10, r=10, t=20, b=10),
                                  yaxis_range=[0, 110])
                st.plotly_chart(fig, use_container_width=True)
        
        st.markdown("---")
        st.markdown("##### 🚨 Distribución de Fallas por Severidad")
        if 'SEVERIDAD' in df_estado.columns:
            df_con_falla = df_estado[df_estado['ESTATUS_CLASIF'] == 'CON FALLA']
            if len(df_con_falla) > 0 and df_con_falla['SEVERIDAD'].notna().any():
                sev_count = df_con_falla['SEVERIDAD'].fillna('Sin clasificar').value_counts().reset_index()
                sev_count.columns = ['Severidad', 'Cantidad']
                fig = px.bar(sev_count, x='Severidad', y='Cantidad', text='Cantidad',
                            color='Severidad',
                            color_discrete_map={'CRÍTICA':'#C00000','ALTA':'#FF8C00',
                                              'MEDIA':'#FFC000','BAJA':'#92D050'})
                fig.update_traces(textposition='outside')
                fig.update_layout(height=300, margin=dict(l=10, r=10, t=20, b=10), showlegend=False)
                st.plotly_chart(fig, use_container_width=True)
        
        st.markdown("---")
        st.markdown("##### 🔴 Cámaras con Falla (Último Barrido)")
        df_falla_ult = df_estado[df_estado['ESTATUS_CLASIF'] == 'CON FALLA'].copy()
        if len(df_falla_ult) > 0:
            cols_show = ['ID_DISPOSITIVO', 'IP', 'PMI', 'TIPO', 'MARCA', 'ZONA',
                        'MUNICIPIO', 'ESTATUS', 'ID_FALLA', 'DESCRIPCION_FALLA',
                        'SEVERIDAD', 'OBS']
            cols_disp = [c for c in cols_show if c in df_falla_ult.columns]
            st.dataframe(df_falla_ult[cols_disp].sort_values(['SEVERIDAD','ZONA'], na_position='last'),
                        use_container_width=True, height=350)
        else:
            st.success("✅ Sin fallas en el último barrido")

# ----------------------------------------------------------------------------
# TAB 4: ANALÍTICOS
# ----------------------------------------------------------------------------
with tab4:
    st.subheader("🎯 Cámaras con Analítico (LPR / Facial)")
    
    df_lpr = df_estado[df_estado['TIENE_LPR'] == 'Sí'].copy()
    df_facial = df_estado[df_estado['TIENE_FACIAL'] == 'Sí'].copy()
    
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("🚗 Total LPR", len(df_lpr))
    if len(df_lpr) > 0:
        op_lpr = (df_lpr['ESTATUS_CLASIF'] == 'OPERATIVA').sum()
        c2.metric("✅ LPR Operativas", op_lpr, f"{op_lpr/len(df_lpr)*100:.1f}%")
    c3.metric("👤 Total Facial", len(df_facial))
    if len(df_facial) > 0:
        op_f = (df_facial['ESTATUS_CLASIF'] == 'OPERATIVA').sum()
        c4.metric("✅ Facial Operativas", op_f, f"{op_f/len(df_facial)*100:.1f}%")
    
    st.markdown("---")
    
    col_a, col_b = st.columns(2)
    with col_a:
        st.markdown("##### Estado de Cámaras LPR")
        if len(df_lpr) > 0:
            d = df_lpr['ESTATUS_CLASIF'].value_counts().reset_index()
            d.columns = ['Estado', 'Cantidad']
            fig = px.pie(d, values='Cantidad', names='Estado', hole=0.4,
                        color='Estado',
                        color_discrete_map={'OPERATIVA':'#00CC66','CON FALLA':'#FF4444','SIN DATO':'#A0A0A0'})
            fig.update_layout(height=300, margin=dict(l=10, r=10, t=20, b=10))
            st.plotly_chart(fig, use_container_width=True)
    
    with col_b:
        st.markdown("##### Estado de Cámaras Facial")
        if len(df_facial) > 0:
            d = df_facial['ESTATUS_CLASIF'].value_counts().reset_index()
            d.columns = ['Estado', 'Cantidad']
            fig = px.pie(d, values='Cantidad', names='Estado', hole=0.4,
                        color='Estado',
                        color_discrete_map={'OPERATIVA':'#00CC66','CON FALLA':'#FF4444','SIN DATO':'#A0A0A0'})
            fig.update_layout(height=300, margin=dict(l=10, r=10, t=20, b=10))
            st.plotly_chart(fig, use_container_width=True)
    
    st.markdown("---")
    st.markdown("##### 📋 Detalle de Cámaras con Analítico")
    df_analiticas = df_estado[(df_estado['TIENE_LPR']=='Sí') | (df_estado['TIENE_FACIAL']=='Sí')].copy()
    if len(df_analiticas) > 0:
        cols_show = ['ID_DISPOSITIVO', 'IP', 'PMI', 'TIPO', 'MARCA',
                    'TIENE_LPR', 'TIENE_FACIAL', 'ZONA', 'MUNICIPIO',
                    'ESTATUS', 'ID_FALLA']
        cols_disp = [c for c in cols_show if c in df_analiticas.columns]
        st.dataframe(df_analiticas[cols_disp].sort_values(['MUNICIPIO','ZONA']),
                    use_container_width=True, height=350)

# ----------------------------------------------------------------------------
# TAB 5: HISTÓRICO
# ----------------------------------------------------------------------------
with tab5:
    st.subheader("🔁 Histórico y Fallas Recurrentes")
    
    if df_barrido is None or len(archivos_info) < 2:
        st.info(f"📭 Se necesitan al menos 2 barridos para análisis histórico. Actualmente hay {len(archivos_info) if archivos_info else 0}.")
    else:
        df_hist = df_barrido.copy()
        df_hist['ESTATUS_CLASIF'] = df_hist['ESTATUS'].apply(clasificar_estatus)
        df_falla_hist = df_hist[df_hist['ESTATUS_CLASIF'] == 'CON FALLA'].copy()
        
        if len(df_falla_hist) == 0:
            st.success("✅ No se han registrado fallas en el periodo")
        else:
            # Top cámaras con más días de falla
            cols_pmi = ['ID', 'PMI'] if 'PMI' in df_falla_hist.columns else ['ID']
            falla_por_cam = df_falla_hist.groupby(cols_pmi).agg(
                Dias_con_Falla=('FECHA_BARRIDO', 'nunique'),
                Fallas_unicas=('ID_FALLA', lambda x: ', '.join(sorted(set(str(v) for v in x.dropna()))))
            ).reset_index().sort_values('Dias_con_Falla', ascending=False)
            
            top30 = falla_por_cam.head(30)
            
            st.markdown("##### Top 30 Cámaras con Más Fallas")
            fig = go.Figure()
            fig.add_trace(go.Bar(
                x=top30['ID'], y=top30['Dias_con_Falla'],
                marker=dict(color=top30['Dias_con_Falla'], colorscale='Reds'),
                text=top30['Dias_con_Falla'], textposition='outside'
            ))
            fig.update_layout(
                xaxis_tickangle=-45, height=450,
                yaxis_title="Días con falla", xaxis_title="Cámara",
                margin=dict(l=10, r=10, t=20, b=10)
            )
            st.plotly_chart(fig, use_container_width=True)
            
            st.markdown("---")
            st.markdown("##### Tendencia Diaria de Fallas")
            tend = df_hist.groupby(['FECHA_STR', 'ESTATUS_CLASIF']).size().reset_index(name='Cantidad')
            tend['FECHA_DT'] = pd.to_datetime(tend['FECHA_STR'], format='%d/%m/%Y')
            tend = tend.sort_values('FECHA_DT')
            fig = px.line(tend, x='FECHA_DT', y='Cantidad', color='ESTATUS_CLASIF',
                         markers=True,
                         color_discrete_map={'OPERATIVA':'#00CC66','CON FALLA':'#FF4444'})
            fig.update_layout(height=350, margin=dict(l=10, r=10, t=20, b=10))
            st.plotly_chart(fig, use_container_width=True)
            
            st.markdown("---")
            st.markdown("##### 📋 Detalle de Fallas Recurrentes")
            umbral = st.slider("Mostrar cámaras con al menos N días de falla:",
                              1, int(falla_por_cam['Dias_con_Falla'].max()), 2)
            fpc_filt = falla_por_cam[falla_por_cam['Dias_con_Falla'] >= umbral]
            st.dataframe(fpc_filt, use_container_width=True, height=350)

# ----------------------------------------------------------------------------
# TAB 6: GEOGRÁFICO
# ----------------------------------------------------------------------------
with tab6:
    st.subheader("🗺️ Vista Geográfica")
    
    df_geo = df_estado.dropna(subset=['LATITUD', 'LONGITUD']).copy()
    if len(df_geo) == 0:
        st.warning("⚠️ El catálogo no contiene coordenadas válidas")
    else:
        # Color por estatus
        color_map = {'OPERATIVA': '#00CC66', 'CON FALLA': '#FF4444', 'SIN DATO': '#A0A0A0'}
        df_geo['color'] = df_geo['ESTATUS_CLASIF'].map(color_map)
        
        c1, c2, c3 = st.columns(3)
        c1.metric("📍 Puntos en mapa", len(df_geo))
        c2.metric("🏘️ Colonias", df_geo['COLONIA'].nunique())
        c3.metric("🏛️ Municipios", df_geo['MUNICIPIO'].nunique())
        
        st.markdown("##### Mapa de Cámaras (color = estatus)")
        fig = px.scatter_mapbox(
            df_geo, lat='LATITUD', lon='LONGITUD',
            color='ESTATUS_CLASIF',
            color_discrete_map=color_map,
            hover_name='ID_DISPOSITIVO',
            hover_data={'PMI': True, 'TIPO': True, 'MARCA': True, 'ZONA': True,
                       'COLONIA': True, 'LATITUD': False, 'LONGITUD': False},
            zoom=11, height=600
        )
        fig.update_layout(mapbox_style='open-street-map',
                         margin=dict(l=0, r=0, t=0, b=0))
        st.plotly_chart(fig, use_container_width=True)
        
        st.markdown("---")
        col_a, col_b = st.columns(2)
        with col_a:
            st.markdown("##### Top 15 Colonias con más Cámaras")
            top_col = df_geo['COLONIA'].value_counts().head(15).reset_index()
            top_col.columns = ['Colonia', 'Cantidad']
            fig = px.bar(top_col, x='Cantidad', y='Colonia', orientation='h',
                        text='Cantidad', color='Cantidad', color_continuous_scale='Blues')
            fig.update_traces(textposition='outside')
            fig.update_layout(height=400, margin=dict(l=10, r=10, t=20, b=10),
                              yaxis={'categoryorder':'total ascending'})
            st.plotly_chart(fig, use_container_width=True)
        
        with col_b:
            st.markdown("##### Top 15 Colonias con más Fallas")
            df_falla_geo = df_geo[df_geo['ESTATUS_CLASIF'] == 'CON FALLA']
            if len(df_falla_geo) > 0:
                top_falla = df_falla_geo['COLONIA'].value_counts().head(15).reset_index()
                top_falla.columns = ['Colonia', 'Fallas']
                fig = px.bar(top_falla, x='Fallas', y='Colonia', orientation='h',
                            text='Fallas', color='Fallas', color_continuous_scale='Reds')
                fig.update_traces(textposition='outside')
                fig.update_layout(height=400, margin=dict(l=10, r=10, t=20, b=10),
                                  yaxis={'categoryorder':'total ascending'})
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.info("Sin fallas registradas")

# ----------------------------------------------------------------------------
# TAB 7: EXPLORADOR
# ----------------------------------------------------------------------------
with tab7:
    st.subheader("🔍 Explorador de Datos")
    
    col_a, col_b, col_c, col_d = st.columns(4)
    with col_a:
        marcas_op = ['Todas'] + sorted([str(x) for x in df_estado['MARCA'].dropna().unique()])
        f_marca = st.selectbox("Marca", marcas_op, key="exp_marca")
    with col_b:
        muni_op = ['Todos'] + sorted([str(x) for x in df_estado['MUNICIPIO'].dropna().unique()])
        f_muni = st.selectbox("Municipio", muni_op, key="exp_muni")
    with col_c:
        zona_op = ['Todas'] + sorted([str(x) for x in df_estado['ZONA'].dropna().unique()])
        f_zona = st.selectbox("Zona", zona_op, key="exp_zona")
    with col_d:
        tipo_op = ['Todos'] + sorted([str(x) for x in df_estado['TIPO'].dropna().unique()])
        f_tipo = st.selectbox("Tipo", tipo_op, key="exp_tipo")
    
    col_e, col_f, col_g = st.columns(3)
    with col_e:
        est_op = ['Todos', 'OPERATIVA', 'CON FALLA', 'SIN DATO']
        f_est = st.selectbox("Estatus", est_op, key="exp_est")
    with col_f:
        f_lpr = st.selectbox("Tiene LPR", ['Todos', 'Sí', 'No'], key="exp_lpr")
    with col_g:
        f_buscar = st.text_input("🔍 Buscar (ID/IP/PMI)", key="exp_busc")
    
    df_f = df_estado.copy()
    if f_marca != 'Todas': df_f = df_f[df_f['MARCA'] == f_marca]
    if f_muni != 'Todos': df_f = df_f[df_f['MUNICIPIO'] == f_muni]
    if f_zona != 'Todas': df_f = df_f[df_f['ZONA'] == f_zona]
    if f_tipo != 'Todos': df_f = df_f[df_f['TIPO'] == f_tipo]
    if f_est != 'Todos': df_f = df_f[df_f['ESTATUS_CLASIF'] == f_est]
    if f_lpr != 'Todos': df_f = df_f[df_f['TIENE_LPR'] == f_lpr]
    if f_buscar:
        mask = (df_f['ID_DISPOSITIVO'].astype(str).str.contains(f_buscar, case=False, na=False) |
                df_f['IP'].astype(str).str.contains(f_buscar, case=False, na=False) |
                df_f['PMI'].astype(str).str.contains(f_buscar, case=False, na=False))
        df_f = df_f[mask]
    
    st.info(f"📊 {len(df_f)} de {len(df_estado)} dispositivos")
    
    cols_export = ['ID_DISPOSITIVO', 'IP', 'PMI', 'TIPO', 'MARCA', 'ZONA',
                  'MUNICIPIO', 'LOCALIDAD', 'COLONIA', 'TIENE_LPR', 'TIENE_FACIAL',
                  'ESTATUS', 'ID_FALLA', 'SEVERIDAD', 'FECHA_STR']
    cols_disp = [c for c in cols_export if c in df_f.columns]
    
    st.dataframe(df_f[cols_disp], use_container_width=True, height=400)
    
    csv = df_f[cols_disp].to_csv(index=False, encoding='utf-8-sig').encode('utf-8-sig')
    st.download_button("📥 Descargar Datos Filtrados", csv,
                      f"camaras_{datetime.now().strftime('%Y%m%d_%H%M')}.csv",
                      "text/csv", key="dwn_exp")

# ----------------------------------------------------------------------------
# TAB 8: FALLAS ATENDIDAS (placeholder)
# ----------------------------------------------------------------------------
with tab8:
    st.subheader("🛠️ Reporte de Fallas Atendidas")
    st.info("""
    📋 **Esta pestaña está reservada para el reporte de fallas atendidas.**
    
    Cuando compartas el formato del archivo de fallas atendidas, se programará aquí:
    - Total de fallas atendidas vs pendientes
    - Tiempo promedio de atención por severidad
    - Tickets cerrados por zona/municipio
    - Histórico de atención
    """)
    st.markdown("---")
    st.markdown("**Por ahora, esto es lo que sabemos del estado actual:**")
    
    if df_estado is not None:
        n_falla = (df_estado['ESTATUS_CLASIF'] == 'CON FALLA').sum()
        n_critica = 0
        n_alta = 0
        if 'SEVERIDAD' in df_estado.columns:
            df_f = df_estado[df_estado['ESTATUS_CLASIF'] == 'CON FALLA']
            n_critica = (df_f['SEVERIDAD'] == 'CRÍTICA').sum()
            n_alta = (df_f['SEVERIDAD'] == 'ALTA').sum()
        
        c1, c2, c3 = st.columns(3)
        c1.metric("❌ Fallas Activas", n_falla, "Por atender")
        c2.metric("🔴 Críticas", n_critica, "Prioridad máxima")
        c3.metric("🟠 Altas", n_alta, "Prioridad alta")

# ============================================================================
# FOOTER
# ============================================================================
st.markdown("---")
st.caption(
    f"🕐 Actualizado: {datetime.now().strftime('%d/%m/%Y %H:%M')} | "
    f"📚 Catálogo: {nombre_catalogo} | "
    f"📋 Barridos: {len(archivos_info) if archivos_info else 0} | "
    f"🎯 Universo: {len(df_universo)} dispositivos"
)
