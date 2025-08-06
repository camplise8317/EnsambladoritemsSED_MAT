import streamlit as st
import pandas as pd
from docxtpl import DocxTemplate
import google.generativeai as genai
import os
import re
import time
import zipfile
from io import BytesIO

# --- CONFIGURACIÓN DE LA PÁGINA DE STREAMLIT ---
st.set_page_config(
    page_title="Ensamblador de Fichas Técnicas con IA",
    page_icon="🤖",
    layout="wide"
)

# --- FUNCIONES DE LÓGICA ---

# Función para limpiar HTML (de tu código original)
def limpiar_html(texto_html):
    if not isinstance(texto_html, str):
        return texto_html
    cleanr = re.compile('<.*?>')
    texto_limpio = re.sub(cleanr, '', texto_html)
    return texto_limpio

# Función para configurar el modelo Gemini
def setup_model(api_key):
    try:
        genai.configure(api_key=api_key)
        generation_config = {
            "temperature": 0.6, "top_p": 1, "top_k": 1, "max_output_tokens": 8192
        }
        safety_settings = [
            {"category": "HARM_CATEGORY_HARASSMENT", "threshold": "BLOCK_MEDIUM_AND_ABOVE"},
            {"category": "HARM_CATEGORY_HATE_SPEECH", "threshold": "BLOCK_MEDIUM_AND_ABOVE"},
            {"category": "HARM_CATEGORY_SEXUALLY_EXPLICIT", "threshold": "BLOCK_MEDIUM_AND_ABOVE"},
            {"category": "HARM_CATEGORY_DANGEROUS_CONTENT", "threshold": "BLOCK_MEDIUM_AND_ABOVE"},
        ]
        model = genai.GenerativeModel(
            model_name="gemini-1.5-pro-latest",
            generation_config=generation_config,
            safety_settings=safety_settings
        )
        return model
    except Exception as e:
        st.error(f"Error al configurar la API de Google: {e}")
        return None

# Funciones para construir prompts (adaptadas de tu código)
def construir_prompt_analisis(fila):
    fila = fila.fillna('')
    descripcion_item = (
        f"Enunciado: {fila.get('Enunciado', '')}\n"
        f"A. {fila.get('OpcionA', '')}\n" # Asegúrate de que el Excel tenga 'OpcionA', 'OpcionB', etc.
        f"B. {fila.get('OpcionB', '')}\n"
        f"C. {fila.get('OpcionC', '')}\n"
        f"D. {fila.get('OpcionD', '')}\n"
        f"Respuesta correcta: {fila.get('AlternativaClave', '')}"
    )
    return f"""
🎯 ROL DEL SISTEMA
Eres un experto en evaluación educativa con un profundo conocimiento de la pedagogía urbana, especializado en enseñanza de las matematicas y procesos cognitivos en el contexto de Bogotá. Tu misión es analizar un ítem de evaluación para proporcionar un análisis tripartito: un resumen de lo que evalúa, la ruta cognitiva detallada para la respuesta correcta, y un análisis de los errores asociados a las opciones incorrectas.

🧠 INSUMOS DE ENTRADA
- Texto/Fragmento: {fila.get('ItemContexto', 'No aplica')}
- Descripción del Ítem: {fila.get('Pregunta', 'No aplica')}
- Imagen asociada al ítem {fila.get('Imagen_pregunta', 'No aplica')}
- Componente: {fila.get('ComponenteNombre', 'No aplica')}
- Competencia: {fila.get('CompetenciaNombre', '')}
- Aprendizaje Priorizado: {fila.get('AfirmacionNombre', '')}
- Evidencia de Aprendizaje: {fila.get('EvidenciaNombre', '')}
- Grado Escolar: {fila.get('ItemGradoId', '')}
- Opción A. {fila.get('OpcionA', '')}\n"
- Opción B. {fila.get('OpcionB', '')}\n"
- Opción C. {fila.get('OpcionC', '')}\n"
- Opción D. {fila.get('OpcionD', '')}\n"
- Respuesta correcta: {fila.get('AlternativaClave', '')}"


📝 INSTRUCCIONES PARA EL ANÁLISIS DEL ÍTEM
Genera el análisis del ítem siguiendo estas reglas y en el orden exacto solicitado:

### 1. Qué Evalúa
Basándote en la Competencia, el Aprendizaje Priorizado y la Evidencia, redacta una frase concisa y clara (1-2 renglones) que resuma la habilidad específica que el ítem está evaluando. Debes comenzar la frase obligatoriamente con "Este ítem evalúa la capacidad del estudiante para...".

### 2. Ruta Cognitiva Correcta
Describe el paso a paso lógico y cognitivo que un estudiante debe seguir para llegar a la respuesta correcta. La explicación debe ser clara y basarse en los verbos del `CRITERIO COGNITIVO` que se define más abajo.

### 3. Análisis de Opciones No Válidas
Para cada una de las TRES opciones incorrectas, explica el posible razonamiento erróneo del estudiante. Describe la confusión o el error conceptual que lo llevaría a elegir esa opción y luego clarifica por qué es incorrecta.

📘 CRITERIO COGNITIVO PARA MATEMÁTICAS
Identifica la competencia principal del ítem y selecciona los verbos cognitivos adecuados de las siguientes listas. 

1. Interpretación y Comunicación (Comprender y representar información)
Verbos de menor complejidad (FORTALECER): identificar, leer (datos, gráficos), reconocer, nombrar, contar, localizar, señalar.
Verbos de mayor complejidad (AVANZAR): representar (en gráficos, tablas), describir, comparar, clasificar, organizar, traducir (de lenguaje verbal a matemático).

2. Formulación y Solución de Problemas (Aplicar procedimientos y estrategias)
Verbos de menor complejidad (FORTALECER): calcular, medir, aplicar (una fórmula), resolver (operaciones directas), completar (secuencias), usar (un algoritmo).
Verbos de mayor complejidad (AVANZAR): formular (un plan o ecuación), plantear, modelar, diseñar (una estrategia), optimizar, descomponer (un problema).

3. Argumentación (Justificar y validar procesos y resultados)
Verbos de menor complejidad (FORTALECER): verificar, explicar (los pasos), mostrar, relacionar, ejemplificar.
Verbos de mayor complejidad (AVANZAR): justificar (un método), validar (un resultado), probar, generalizar, demostrar, evaluar (la pertinencia de una solución).

✍️ FORMATO DE SALIDA DEL ANÁLISIS
**REGLA CRÍTICA:** Responde únicamente con el texto solicitado y en la estructura definida a continuación. Es crucial que los tres títulos aparezcan en la respuesta, en el orden correcto. No agregues introducciones, conclusiones ni frases de cierre.

Qué Evalúa:
[Frase concisa de 1-2 renglones.]

Ruta Cognitiva Correcta:
Descripción concisa y paso a paso del proceso cognitivo. Debe estar escrita como un parrafo continuo y no como una lista

Análisis de Opciones No Válidas:
- El estudiante podría escoger la [OpcionX] porque [razonamiento erróneo]. Sin embargo, esto es incorrecto porque [razón].
"""

def construir_prompt_recomendaciones(fila):
    fila = fila.fillna('')
    return f"""
🎯 ROL DEL SISTEMA
Eres un experto en evaluación educativa especializado en enseñanza de las matematicas con un profundo conocimiento de la pedagogía urbana. Tu misión es generar dos recomendaciones pedagógicas personalizadas a partir de cada ítem de evaluación formativa: una para Fortalecer y otra para Avanzar en el aprendizaje. Deberás identificar de manera endógena los verbos clave de los procesos cognitivos implicados, basándote en la competencia, el aprendizaje priorizado, la evidencia de aprendizaje, el grado escolar, la edad (para gado 3 niños de 9 a 11 años, grado 6 de 11 a 13 años, grado noveno de 13 a 15 años) y el nivel educativo general de los estudiantes. Luego, integrarás estos verbos de forma fluida en la redacción de las recomendaciones. Considerarás las características cognitivas y pedagógicas del ítem.

🧠 INSUMOS DE ENTRADA
- Texto/Fragmento: {fila.get('ItemContexto', 'No aplica')}
- Descripción del Ítem: {fila.get('Pregunta', 'No aplica')}
- Imagen asociada al ítem {fila.get('Imagen_pregunta', 'No aplica')}
- Componente: {fila.get('ComponenteNombre', 'No aplica')}
- Competencia: {fila.get('CompetenciaNombre', '')}
- Aprendizaje Priorizado: {fila.get('AfirmacionNombre', '')}
- Evidencia de Aprendizaje: {fila.get('EvidenciaNombre', '')}
- Tipología Textual (Solo para Lectura Crítica): {fila.get('Tipologia Textual', 'No aplica')}
- Grado Escolar: {fila.get('ItemGradoId', '')}
- Respuesta correcta: {fila.get('AlternativaClave', 'No aplica')}
- Opción A: {fila.get('OpcionA', 'No aplica')}
- Opción B: {fila.get('OpcionB', 'No aplica')}
- Opción C: {fila.get('OpcionC', 'No aplica')}
- Opción D: {fila.get('OpcionD', 'No aplica')}

📝 INSTRUCCIONES PARA GENERAR LAS RECOMENDACIONES
Para cada ítem, genera dos recomendaciones claras y accionables, siguiendo los siguientes criterios:

### Reglas Generales Clave:
1.  **Innovación Pedagógica:** Las actividades deben ser **novedosas, poco convencionales y creativas**. Busca inspiración en temas de actualidad (tecnología, medio ambiente, cultura popular, etc.) para que sean significativas.
2.  **Enfoque Matemático:** El núcleo de cada actividad debe ser el concepto matemático. Los elementos contextuales o lúdicos deben servir para potenciar el aprendizaje matemático, no para opacarlo. La logística debe ser mínima.
3.  **Diferenciación Clara:** La actividad de "Fortalecer" debe ser fundamentalmente diferente en enfoque y ejecución a la de "Avanzar".
4.  **Tono de la redacción:** La redacción de las actividades debe ser impersonal, es decir no se nombra el sujeto (ej "El docente", "El estudiante") ni personas específicas.

### 1. Recomendación para FORTALECER
-   **Objetivo:** Tu punto de partida debe ser el análisis de las **opciones de respuesta incorrectas**. La recomendación debe enfocarse en corregir el error conceptual o procedimental específico que lleva a un estudiante a elegir uno de los distractores.
-   **Verbos Clave Sugeridos:** Utiliza verbos descritos en CRITERIO COGNITIVO PARA MATEMÁTICAS que están mas adelante.
-   **Párrafo Inicial:** Describe la estrategia didáctica, explicando cómo la actividad propuesta ataca directamente la raíz del error más común (identificado en los distractores).
-   **Actividad Propuesta:** Diseña una experiencia lúdica y significativa. Debe estar profundamente contextualizada en una situación cotidiana o de interés para los estudiantes.
-   **Preguntas Orientadoras:** Formula preguntas que guíen el aprendizaje desde lo más básico (concreto) hacia la comprensión del concepto.
-   **Edad de los evaluados:**  Ajustar el nivel cognitivo de las actividades de acuerdo a la edad de los estudiantes. (para gado 3 niños de 9 a 11 años, grado 6 de 11 a 13 años, grado noveno de 13 a 15 años)

### 2. Recomendación para AVANZAR
-   **Objetivo:** Desarrollar procesos cognitivos más complejos que permitan **ampliar, profundizar o transferir** el aprendizaje evaluado.
-   **Verbos Clave Sugeridos:** Emplea verbos de mayor nivel descritos en CRITERIO COGNITIVO PARA MATEMÁTICAS que están mas adelante.
-   **Párrafo Inicial:** Describe la estrategia para complejizar el aprendizaje. Incluye múltiples vías en las que se puede profundizar el conocimiento (ej., "se puede transferir a un problema de finanzas personales, a un desafío de diseño o a un análisis de datos simple...").
-   **Actividad Propuesta:** Crea una actividad totalmente diferente a la de fortalecer, orientado con el objetivo de la recomendación,con un desafío intelectual estimulante. Integra elementos de la actualidad de forma creativa.
-   **Preguntas Orientadoras:** Formula preguntas que progresen en dificultad, facilitando el paso de representaciones concretas a abstractas y fomentando el pensamiento crítico y la generalización.
-   **Edad de los evaluados:**  Ajustar el nivel cognitivo de las actividades de acuerdo a la edad de los estudiantes. (para gado 3 niños de 9 a 11 años, grado 6 de 11 a 13 años, grado noveno de 13 a 15 años)
-   **Estrategias para crear actividades de avanzar:** a. Progresar a partir del tipo de número utilizado en el objeto matemático; por ejemplo, si se trabaja con números naturales, avanzar hacia el uso de fracciones o decimales. b. Ampliar el objeto matemático relacionado; por ejemplo, si se interpreta información de una tabla a un diagrama de barras, avanzar hacia la interpretación de un diagrama de barras a uno circular, o de una lista a un pictograma y viceversa. c. Promover un avance en las operaciones intelectuales o procesos de pensamiento, pasando de identificar a diferenciar o corregir, siempre manteniendo la competencia.

📘 CRITERIO COGNITIVO PARA MATEMÁTICAS
Identifica la competencia principal del ítem y selecciona los verbos cognitivos adecuados de las siguientes listas. Para FORTALECER, elige un verbo que refleje un proceso fundamental o de entrada. Para AVANZAR, selecciona un verbo que implique una mayor elaboración o transferencia del conocimiento.

1. Interpretación y Comunicación (Comprender y representar información)
Verbos de menor complejidad (FORTALECER): identificar, leer (datos, gráficos), reconocer, nombrar, contar, localizar, señalar.
Verbos de mayor complejidad (AVANZAR): representar (en gráficos, tablas), describir, comparar, clasificar, organizar, traducir (de lenguaje verbal a matemático).

2. Formulación y Solución de Problemas (Aplicar procedimientos y estrategias)
Verbos de menor complejidad (FORTALECER): calcular, medir, aplicar (una fórmula), resolver (operaciones directas), completar (secuencias), usar (un algoritmo).
Verbos de mayor complejidad (AVANZAR): formular (un plan o ecuación), plantear, modelar, diseñar (una estrategia), optimizar, descomponer (un problema).

3. Argumentación (Justificar y validar procesos y resultados)
Verbos de menor complejidad (FORTALECER): verificar, explicar (los pasos), mostrar, relacionar, ejemplificar.
Verbos de mayor complejidad (AVANZAR): justificar (un método), validar (un resultado), probar, generalizar, demostrar, evaluar (la pertinencia de una solución).

✍️ FORMATO DE SALIDA DE LAS RECOMENDACIONES
**IMPORTANTE: Responde de forma directa y concreta. No incluyas frases de cierre o resúmenes. Cada recomendación debe seguir esta estructura exacta:**

RECOMENDACIÓN PARA [FORTALECER/AVANZAR] EL APRENDIZAJE EVALUADO EN EL ÍTEM
Para [Fortalecer/Avanzar] la habilidad de [verbo clave] en situaciones relacionadas con [frase del aprendizaje priorizado], se sugiere [descripción concreta de la sugerencia].
Una actividad que se puede hacer es: [Descripción detallada de la actividad].
Las preguntas orientadoras para esta actividad, entre otras, pueden ser:
- [Pregunta 1]
- [Pregunta 2]
- [Pregunta 3]
- [Pregunta 4]
- [Pregunta 5]
"""


# --- INTERFAZ PRINCIPAL DE STREAMLIT ---

st.title("🤖 Ensamblador de Fichas Técnicas con IA")
st.markdown("Una aplicación para enriquecer datos pedagógicos y generar fichas personalizadas.")

# Inicializar session_state para guardar los datos entre ejecuciones
if 'df_enriquecido' not in st.session_state:
    st.session_state.df_enriquecido = None
if 'zip_buffer' not in st.session_state:
    st.session_state.zip_buffer = None

# --- PASO 0: Clave API ---
st.sidebar.header("🔑 Configuración Obligatoria")
api_key = st.sidebar.text_input("Ingresa tu Clave API de Google AI", type="password")

# --- PASO 1: Carga de Archivos ---
st.header("Paso 1: Carga tus Archivos")
col1, col2 = st.columns(2)
with col1:
    archivo_excel = st.file_uploader("Sube tu Excel con los datos base", type=["xlsx"])
with col2:
    archivo_plantilla = st.file_uploader("Sube tu Plantilla de Word", type=["docx"])

# --- PASO 2: Enriquecimiento con IA ---
st.header("Paso 2: Enriquece tus Datos con IA")
if st.button("🤖 Iniciar Análisis y Generación", disabled=(not api_key or not archivo_excel)):
    if not api_key:
        st.error("Por favor, ingresa tu clave API en la barra lateral izquierda.")
    elif not archivo_excel:
        st.warning("Por favor, sube un archivo Excel para continuar.")
    else:
        model = setup_model(api_key)
        if model:
            with st.spinner("Procesando archivo Excel y preparando datos..."):
                df = pd.read_excel(archivo_excel)
                # Limpieza de HTML
                for col in df.columns:
                    if df[col].dtype == 'object':
                        df[col] = df[col].apply(limpiar_html)
                st.success("Datos limpios y listos para el análisis.")

            total_filas = len(df)
            
            # Proceso de Análisis
            with st.spinner("Generando Análisis de Ítems... Esto puede tardar varios minutos."):
                que_evalua_lista, just_correcta_lista, an_distractores_lista = [], [], []
                progress_bar_analisis = st.progress(0, text="Iniciando Análisis...")
                for i, fila in df.iterrows():
                    prompt = construir_prompt_analisis(fila)
                    try:
                        response = model.generate_content(prompt)
                        texto_completo = response.text.strip()
                        # Separación robusta
                        header_que_evalua = "Qué Evalúa:"
                        header_correcta = "Ruta Cognitiva Correcta:"
                        header_distractores = "Análisis de Opciones No Válidas:"
                        idx_correcta = texto_completo.find(header_correcta)
                        idx_distractores = texto_completo.find(header_distractores)
                        que_evalua = texto_completo[len(header_que_evalua):idx_correcta].strip() if idx_correcta != -1 else texto_completo
                        just_correcta = texto_completo[idx_correcta:idx_distractores].strip() if idx_correcta != -1 and idx_distractores != -1 else (texto_completo[idx_correcta:].strip() if idx_correcta != -1 else "ERROR")
                        an_distractores = texto_completo[idx_distractores:].strip() if idx_distractores != -1 else "ERROR"
                    except Exception as e:
                        st.warning(f"Error en fila {i+1} (Análisis): {e}")
                        que_evalua, just_correcta, an_distractores = "ERROR API", "ERROR API", "ERROR API"
                    
                    que_evalua_lista.append(que_evalua)
                    just_correcta_lista.append(just_correcta)
                    an_distractores_lista.append(an_distractores)
                    progress_bar_analisis.progress((i + 1) / total_filas, text=f"Analizando Ítem {i+1}/{total_filas}")
                    time.sleep(1) # Control de velocidad de la API
                
                df["Que_Evalua"] = que_evalua_lista
                df["Justificacion_Correcta"] = just_correcta_lista
                df["Analisis_Distractores"] = an_distractores_lista
                st.success("Análisis de Ítems completado.")

            # Proceso de Recomendaciones
            with st.spinner("Generando Recomendaciones Pedagógicas... Esto también puede tardar."):
                fortalecer_lista, avanzar_lista = [], []
                progress_bar_recom = st.progress(0, text="Iniciando Recomendaciones...")
                for i, fila in df.iterrows():
                    prompt = construir_prompt_recomendaciones(fila)
                    try:
                        response = model.generate_content(prompt)
                        texto_completo = response.text.strip()
                        # Separación robusta
                        titulo_avanzar = "RECOMENDACIÓN PARA AVANZAR"
                        idx_avanzar = texto_completo.upper().find(titulo_avanzar)
                        if idx_avanzar != -1:
                            fortalecer = texto_completo[:idx_avanzar].strip()
                            avanzar = texto_completo[idx_avanzar:].strip()
                        else:
                            fortalecer, avanzar = texto_completo, "ERROR: No se encontró 'AVANZAR'"
                    except Exception as e:
                        st.warning(f"Error en fila {i+1} (Recomendaciones): {e}")
                        fortalecer, avanzar = "ERROR API", "ERROR API"
                    
                    fortalecer_lista.append(fortalecer)
                    avanzar_lista.append(avanzar)
                    progress_bar_recom.progress((i + 1) / total_filas, text=f"Generando Recomendación {i+1}/{total_filas}")
                    time.sleep(1) # Control de velocidad

                df["Recomendacion_Fortalecer"] = fortalecer_lista
                df["Recomendacion_Avanzar"] = avanzar_lista
                st.success("Recomendaciones generadas con éxito.")
            
            # Guardar el resultado en el estado de la sesión
            st.session_state.df_enriquecido = df
            st.balloons()

# --- PASO 3: Vista Previa y Verificación ---
if st.session_state.df_enriquecido is not None:
    st.header("Paso 3: Verifica los Datos Enriquecidos")
    st.dataframe(st.session_state.df_enriquecido.head())
    
    # Opción para descargar el Excel enriquecido
    output_excel = BytesIO()
    with pd.ExcelWriter(output_excel, engine='openpyxl') as writer:
        st.session_state.df_enriquecido.to_excel(writer, index=False, sheet_name='Datos Enriquecidos')
    output_excel.seek(0)
    
    st.download_button(
        label="📥 Descargar Excel Enriquecido",
        data=output_excel,
        file_name="excel_enriquecido_con_ia.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# --- PASO 4: Ensamblaje de Fichas ---
if st.session_state.df_enriquecido is not None:
    st.header("Paso 4: Ensambla las Fichas Técnicas")
    if not archivo_plantilla:
        st.warning("Por favor, sube una plantilla de Word para continuar con el ensamblaje.")
    else:
        columna_nombre_archivo = st.text_input(
            "Escribe el nombre de la columna para nombrar los archivos (ej. ItemId)", 
            value="ItemId"
        )
        
        if st.button("📄 Ensamblar Fichas Técnicas", type="primary"):
            df_final = st.session_state.df_enriquecido
            if columna_nombre_archivo not in df_final.columns:
                st.error(f"La columna '{columna_nombre_archivo}' no existe en el Excel. Por favor, elige una de: {', '.join(df_final.columns)}")
            else:
                with st.spinner("Ensamblando todas las fichas en un archivo .zip..."):
                    plantilla_bytes = BytesIO(archivo_plantilla.getvalue())
                    zip_buffer = BytesIO()
                    
                    with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED, False) as zip_file:
                        total_docs = len(df_final)
                        progress_bar_zip = st.progress(0, text="Iniciando ensamblaje...")
                        for i, fila in df_final.iterrows():
                            doc = DocxTemplate(plantilla_bytes)
                            contexto = fila.to_dict()
                            doc.render(contexto)
                            
                            doc_buffer = BytesIO()
                            doc.save(doc_buffer)
                            doc_buffer.seek(0)
                            
                            nombre_base = str(fila[columna_nombre_archivo]).replace('/', '_').replace('\\', '_')
                            nombre_archivo_salida = f"{nombre_base}.docx"
                            
                            zip_file.writestr(nombre_archivo_salida, doc_buffer.getvalue())
                            progress_bar_zip.progress((i + 1) / total_docs, text=f"Añadiendo ficha {i+1}/{total_docs} al .zip")
                    
                    st.session_state.zip_buffer = zip_buffer
                    st.success("¡Ensamblaje completado!")

# --- PASO 5: Descarga Final ---
if st.session_state.zip_buffer:
    st.header("Paso 5: Descarga el Resultado Final")
    st.download_button(
        label="📥 Descargar TODAS las fichas (.zip)",
        data=st.session_state.zip_buffer,
        file_name="fichas_tecnicas_generadas.zip",
        mime="application/zip"
    )
