import streamlit as st
import google.generativeai as genai
import msal
import openpyxl
import io
import json

# 1. Configuración de la interfaz visual de Carla
st.set_page_config(page_title="Asistente de Laboratorio IQ", page_icon="🧪")
st.title("🧪 Puente de Calificaciones para Carla")
st.write("Conéctate a OneDrive, sube los exámenes y califica con IA de forma precisa.")

# Cargar las llaves secretas guardadas de forma segura en la nube
try:
    GENAI_KEY = st.secrets["CONEXION_IA"]
    CLIENT_ID = st.secrets["CLIENT_ID"]
    TENANT_ID = st.secrets["TENANT_ID"]
    CLIENT_SECRET = st.secrets["SECRETO_CLIENTE"]
except Exception:
    st.error("Falta configurar las llaves secretas en Streamlit Cloud.")
    st.stop()

# Configurar la IA de Google
genai.configure(api_key=GENAI_KEY)

# 2. Manejo de la conexión segura con Microsoft OneDrive
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPE = ["Files.ReadWrite.All", "User.Read"]

msal_app = msal.ConfidentialClientApplication(
    CLIENT_ID, authority=AUTHORITY, client_credential=CLIENT_SECRET
)

# Botón de inicio de sesión para Carla
if "auth_token" not in st.session_state:
    st.session_state.auth_token = None

if not st.session_state.auth_token:
    st.subheader("Paso 1: Autorizar acceso")
    st.info("Para que la app pueda buscar tus archivos de notas, primero debes conectarte.")
    
    if st.button("🔌 Conectar OneDrive"):
        result = msal_app.acquire_token_for_client(scopes=["https://graph.microsoft.com/.default"])
        if "access_token" in result:
            st.session_state.auth_token = result["access_token"]
            st.success("¡Conectado exitosamente a OneDrive!")
            st.rerun()
        else:
            st.error("No se pudo conectar. Verifica tus llaves secretas de Microsoft.")
            st.stop()
else:
    st.sidebar.success("🟢 OneDrive Conectado")

    # 3. Configuración del criterio de Carla (Modelo de respuestas)
    st.subheader("Paso 2: Define el criterio de evaluación de hoy")
    modelo_respuestas = st.text_area(
        "Pega aquí la solución del coloquio/informe y las penalizaciones exactas que deseas aplicar:",
        placeholder="Ejemplo: La respuesta matemática correcta es 125 W/m²K. Si confunde unidades resta 2 puntos. Si el balance no cierra, resta 3 puntos.",
        height=150
    )

    # 4. Carga de archivos por calificar
    st.subheader("Paso 3: Sube los documentos de los estudiantes")
    archivos_subidos = st.file_uploader(
        "Puedes arrastrar fotos de coloquios a mano (JPG/PNG) o informes digitales (PDF):", 
        type=["png", "jpg", "jpeg", "pdf"], 
        accept_multiple_files=True
    )

    # 5. Selección del Excel de notas de OneDrive
    st.subheader("Paso 4: Selecciona el archivo de Excel de notas")
    excel_file = st.file_uploader("Arrastra o selecciona el archivo Excel de notas donde se guardarán los resultados:", type=["xlsx"])

    if archivos_subidos and excel_file and modelo_respuestas:
        if st.button("🚀 Iniciar Calificación Automatizada"):
            progreso = st.progress(0)
            status_text = st.empty()
            
            # Cargar el Excel en memoria para modificarlo
            wb = openpyxl.load_workbook(io.BytesIO(excel_file.read()))
            sheet = wb.active
            
            # Asignar registros buscando por el nombre de las columnas en la primera fila
            headers = [str(cell.value).lower().strip() for cell in sheet[1]]
            col_nombre = headers.index("nombre") + 1 if "nombre" in headers else 1
            col_nota = headers.index("nota") + 1 if "nota" in headers else 2
            
            # Buscamos la columna de observaciones. Priorizamos la columna Y si existiera, o mapeamos dinámicamente
            if "observacion" in headers:
                col_feedback = headers.index("observacion") + 1
            elif "feedback" in headers:
                col_feedback = headers.index("feedback") + 1
            else:
                col_feedback = 3
            
            total_archivos = len(archivos_subidos)
