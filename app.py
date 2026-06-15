import streamlit as st
import google.generativeai as genai
import openpyxl
import io
import json

# 1. Configuración de la interfaz visual de Carla
st.set_page_config(page_title="Asistente de Laboratorio IQ", page_icon="🧪")
st.title("🧪 Calificador Automatizado de Laboratorio (SAEw)")
st.write("Sube la rúbrica, los informes de los alumnos y el acta del SAEw para calificar al instante.")

# Cargar la llave de la IA de forma segura
try:
    GENAI_KEY = st.secrets["CONEXION_IA"]
except Exception:
    st.error("Falta configurar la llave secretas 'CONEXION_IA' en Streamlit Cloud.")
    st.stop()

# Configurar la IA de Google
genai.configure(api_key=GENAI_KEY)

# 2. Configuración del criterio de Carla (Modelo de respuestas)
st.subheader("Paso 1: Define el criterio de evaluación de hoy")
modelo_respuestas = st.text_area(
    "Pega aquí la solución de la práctica o coloquio y las penalizaciones exactas:",
    placeholder="Ejemplo: La respuesta matemática correcta es 125 W/m²K. Si confunde unidades resta 2 puntos.",
    height=120
)

# 3. Carga de archivos por calificar
st.subheader("Paso 2: Sube los documentos de los estudiantes")
archivos_subidos = st.file_uploader(
    "Puedes arrastrar fotos de coloquios (JPG/PNG) o informes grupales en PDF:", 
    type=["png", "jpg", "jpeg", "pdf"], 
    accept_multiple_files=True
)

# 4. Selección y parsing dinámico del Excel del SAEw
st.subheader("Paso 3: Archivo de Calificaciones SAEw")
excel_file = st.file_uploader("Arrastra aquí el archivo Excel original exportado del SAEw:", type=["xlsx"])

practica_seleccionada = None
header_row_idx = None
practice_row_idx = None

if excel_file:
    try:
        bytes_in_mem = excel_file.read()
        wb_reader = openpyxl.load_workbook(io.BytesIO(bytes_in_mem))
        sheet_reader = wb_reader.active
        
        # Buscar dinámicamente la fila de cabecera principal (donde está "Estudiante")
        for r in range(1, 15):
            row_vals = [str(sheet_reader.cell(row=r, column=c).value).lower() for c in range(1, 20)]
            if any("estudiante" in val or "código" in val for val in row_vals):
                header_row_idx = r
                break
        
        if header_row_idx and header_row_idx > 1:
            practice_row_idx = header_row_idx - 1
            listado_practicas = []
            for c in range(1, sheet_reader.max_column + 1):
                val = sheet_reader.cell(row=practice_row_idx, column=c).value
                if val and str(val).strip() and str(val).strip() not in listado_practicas:
                    if "escuela" not in str(val).lower() and "fuente" not in str(val).lower():
                        listado_practicas.append(str(val).strip())
            
            if listado_practicas:
                practica_seleccionada = st.selectbox("🎯 ¿Qué práctica vas a calificar hoy?", listado_practicas)
            else:
                st.error("No se detectaron nombres de prácticas en la cabecera del archivo.")
        else:
            st.error("No se reconoció la estructura estándar del archivo SAEw (Falta columna 'Estudiante').")
            
        excel_file.seek(0) # Resetear puntero para el proceso de ejecución
    except Exception as e:
        st.error(f"Error analizando el archivo Excel: {e}")

# 5. Ejecución del Proceso de Calificación
if archivos_subidos and excel_file and modelo_respuestas and practica_seleccionada:
    if st.button("🚀 Iniciar Calificación Automatizada"):
        progreso = st.progress(0)
        status_text = st.empty()
        
        wb = openpyxl.load_workbook(io.BytesIO(excel_file.read()))
        sheet = wb.active
        
        # Localizar columna de estudiantes
        col_nombre = 4 # Fallback estándar
        for c in range(1, 20):
            val_h = str(sheet.cell(row=header_row_idx, column=c).value).lower()
            if "estudiante" in val_h:
                col_nombre = c
                break
        
        # Localizar la columna de la práctica seleccionada
        start_col = None
        for c in range(1, sheet.max_column + 1):
            val_p = sheet.cell(row=practice_row_idx, column=c).value
            if val_p and str(val_p).strip().lower() == practica_seleccionada.lower():
                start_col = c
                break
        
        # Buscar la columna "I" (Informe) dentro de esa práctica
        col_nota = None
        if start_col:
            for c in range(start_col, start_col + 5):
                val_sub = str(sheet.cell(row=header_row_idx, column=c).value).lower().strip()
                if val_sub == 'i' or val_sub == 'i2' or 'inf' in val_sub:
                    col_nota = c
                    break
        
        if not col_nota:
            col_nota = start_col + 2 if start_col else 10 # Fallback de posición
        
        # Crear columna de comentarios al extremo derecho para no alterar las fórmulas del SAEw
        col_feedback = sheet.max_column + 1
        sheet.cell(row=header_row_idx, column=col_feedback).value = f"Observaciones - {practica_seleccionada}"
        
        total_archivos = len(archivos_subidos)
        
        for index, archivo in enumerate(archivos_subidos):
            status_text.text(f"Analizando informe {index+1} de {total_archivos}: {archivo.name}")
            
            bytes_data = archivo.read()
            model = genai.GenerativeModel('gemini-1.5-flash')
            
            prompt = f"""
            Actúa como un calificador experto de laboratorio de Ingeniería Química para la EPN.
            Evalúa este documento basándote estrictamente en este Modelo de Respuestas:
            {modelo_respuestas}
            
            Identifica a todos los integrantes del grupo que firman el informe.
            Entrega el resultado exclusivamente en este formato JSON limpio, sin marcas de bloque:
            {{
              "nombres": ["Nombre Integrante 1", "Nombre Integrante 2"],
              "nota": 8.5,
              "justificacion": "Breve explicación técnica de la calificación."
            }}
            """
            
            response = model.generate_content([
                {"mime_type": archivo.type, "data": bytes_data},
                prompt
            ])
            
            try:
                texto_ia = response.text.replace("```json", "").replace("```", "").strip()
                datos = json.loads(texto_ia)
                
                lista_nombres = datos.get("nombres", [])
                nota_final = datos.get("nota", 0)
                feedback_final = datos.get("justificacion", "")
                
                # Asignar la nota a cada integrante del grupo en el Excel
                for nombre_estudiante in lista_nombres:
                    for row in range(header_row_idx + 1, sheet.max_row + 1):
                        celda_nombre = str(sheet.cell(row=row, column=col_nombre).value).lower().strip()
                        if nombre_estudiante.lower().strip() in celda_nombre or celda_nombre in nombre_estudiante.lower().strip():
                            sheet.cell(row=row, column=col_nota).value = float(nota_final)
                            sheet.cell(row=row, column=col_feedback).value = feedback_final
                            break
                            
            except Exception as e:
                st.warning(f"Error procesando {archivo.name}: {str(e)}")
            
            progreso.progress((index + 1) / total_archivos)
        
        status_text.text("¡Proceso terminado con éxito!")
        st.success(f"Notas aplicadas correctamente para la práctica: {practica_seleccionada}")
        
        output = io.BytesIO()
        wb.save(output)
        output.seek(0)
        
        st.download_button(
            label="📥 Descargar Excel con Notas Actualizadas",
            data=output,
            file_name=f"Notas_Actualizadas_{excel_file.name}",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
