import streamlit as st
import google.generativeai as genai
import msal
import openpyxl
import io
import json
import requests

# 1. Configuración de la interfaz visual de Carla
st.set_page_config(page_title="Asistente de Laboratorio IQ", page_icon="🧪")
st.title("🧪Calificador")
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
    # Crear enlace de autenticación
    auth_url = msal_app.get_authorization_request_url(SCOPE, redirect_uri="http://localhost:8501")
    st.subheader("Paso 1: Autorizar acceso")
    st.info("Para que la app pueda buscar tus archivos de notas, primero debes conectarte.")
    
    # En un entorno real en la nube, capturamos el token por código para simplificar al usuario
    token_input = st.text_input("Pega aquí el código de autorización de Microsoft o inicia sesión directamente:")
    if st.button("🔌 Conectar OneDrive"):
        # Intento de conexión simulada o directa según las credenciales del inquilino
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
            
            # Buscar dónde poner las cosas inteligentemente en el formato de Carla
            headers = [str(cell.value).lower().strip() for cell in sheet[1]]
            col_nombre = headers.index("nombre") + 1 if "nombre" in headers else 1
            col_nota = headers.index("nota") + 1 if "nota" in headers else 2
            col_feedback = headers.index("observacion") + 1 if "observacion" in headers else (headers.index("feedback") + 1 if "feedback" in headers else 3)
            
            total_archivos = len(archivos_subidos)
            
            for index, archivo in enumerate(archivos_subidos):
                status_text.text(f"Analizando archivo {index+1} de {total_archivos}: {archivo.name}")
                
                # Preparar datos para enviar a la IA
                bytes_data = archivo.read()
                
                model = genai.GenerativeModel('gemini-1.5-flash')
                
                prompt = f"""
                Actúa como un calificador experto de laboratorio de Ingeniería Química para la EPN.
                Evalúa este documento de forma rigurosa basándote estrictamente en este Modelo de Respuestas y Criterios del profesor:
                {modelo_respuestas}
                
                Busca el nombre del estudiante en el documento.
                Entrega el resultado exclusivamente en este formato JSON limpio, no agregues texto extra ni marcas de bloque:
                {{
                  "nombre": "Nombre y Apellido detectado",
                  "nota": 8.5,
                  "justificacion": "Breve explicación técnica de por qué obtuvo esa nota basándote en la penalización."
                }}
                """
                
                # Enviar archivo a Gemini
                response = model.generate_content([
                    {"mime_type": archivo.type, "data": bytes_data},
                    prompt
                ])
                
                try:
                    # Limpiar y parsear JSON devuelto por la IA
                    texto_ia = response.text.replace("```json", "").replace("
```", "").strip()
                    datos = json.loads(texto_ia)
                    
                    nombre_estudiante = datos.get("nombre", "Desconocido")
                    nota_final = datos.get("nota", 0)
                    feedback_final = datos.get("justificacion", "")
                    
                    # Buscar al alumno en el Excel de Carla para no dañar sus filas
                    alumno_encontrado = False
                    for row in range(2, sheet.max_row + 1):
                        celda_nombre = str(sheet.cell(row=row, column=col_nombre).value).lower().strip()
                        if nombre_estudiante.lower().strip() in celda_nombre or celda_nombre in nombre_estudiante.lower().strip():
                            sheet.cell(row=row, column=col_nota).value = float(nota_final)
                            sheet.cell(row=row, column=col_feedback).value = feedback_final
                            alumno_encontrado = True
                            break
                    
                    # Si el alumno no estaba en la lista, añadirlo al final de manera segura
                    if not alumno_encontrado:
                        sheet.append([nombre_estudiante, float(nota_final), feedback_final])
                        
                except Exception as e:
                    st.warning(f"Error procesando {archivo.name}: {str(e)}")
                
                progreso.progress((index + 1) / total_archivos)
            
            status_text.text("¡Proceso terminado con éxito!")
            st.success("Se analizaron todos los exámenes basados en tu criterio.")
            
            # Devolverle a Carla su Excel modificado listo para guardar en su OneDrive
            output = io.BytesIO()
            wb.save(output)
            output.seek(0)
            
            st.download_button(
                label="📥 Descargar Excel con Notas Actualizadas",
                data=output,
                file_name=f"Notas_Actualizadas_{excel_file.name}",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
