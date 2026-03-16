import os, json, ast
import streamlit as st
import google.generativeai as genai
from docx import Document
import openpyxl
import PyPDF2
from io import BytesIO
from datetime import datetime
import pytz

st.set_page_config(page_title="Oro Asistente", page_icon="🏆")

# ==========================================
# CONEXIÓN CON AUTO-DETECTOR DE MODELOS
# ==========================================
try:
    LLAVE_GEMINI = st.secrets["LLAVE_GEMINI"]
    genai.configure(api_key=LLAVE_GEMINI)
    
    # 1. Le preguntamos a Google qué modelos te permite usar
    modelos_disponibles = []
    for m in genai.list_models():
        if 'generateContent' in m.supported_generation_methods:
            # Quitamos la palabra 'models/' para que la librería no se confunda
            nombre_limpio = m.name.replace("models/", "")
            modelos_disponibles.append(nombre_limpio)
            
    if not modelos_disponibles:
        st.error("❌ Tu llave es correcta, pero Google no te habilitó modelos de texto. Crea una nueva en Google AI Studio.")
        st.stop()
        
    # 2. Elegimos automáticamente el mejor modelo que tengas disponible
    MODELO_ELEGIDO = modelos_disponibles[0] # Usamos el primero por defecto
    for m in modelos_disponibles:
        if 'flash' in m:
            MODELO_ELEGIDO = m
            break
            
except Exception as e:
    st.error(f"🔑 Error configurando la IA: {e}")
    st.stop()

st.markdown("""
    <style>
    .stButton>button { width: 100%; border-radius: 10px; height: 3.5em; background-color: #007bff; color: white; font-weight: bold; }
    h1 { text-align: center; color: #1e3a8a; }
    .footer { text-align: center; font-size: 12px; color: gray; margin-top: 50px; }
    </style>
    """, unsafe_allow_html=True)

st.title("🏆 Oro Asistente")

# Mostramos sutilmente qué modelo está usando para estar seguros
st.caption(f"🧠 Conectado exitosamente al modelo: {MODELO_ELEGIDO}")

# ==========================================
# FUNCIONES DE INTELIGENCIA ARTIFICIAL
# ==========================================

def solicitar_resumen_estructurado(texto):
    prompt = (
        "Analiza el documento.\n\n"
        "Responde UNICAMENTE con un objeto JSON válido. No uses markdown.\n"
        'Estructura EXACTA: {"tipo": "...", "datos": {"titulo": "...", "resumen_ejecutivo": "...", '
        '"detalles": {"puntos_clave": ["Punto 1", "Punto 2"], "metricas_principales": {"Dato": "Valor"}}}}\n\n'
        f"CONTENIDO:\n{texto[:10000]}"
    )
    try:
        model = genai.GenerativeModel(MODELO_ELEGIDO)
        respuesta = model.generate_content(prompt)
        res_raw = respuesta.text
        inicio, fin = res_raw.find("{"), res_raw.rfind("}") + 1
        if inicio != -1:
            return json.loads(res_raw[inicio:fin], strict=False)
    except Exception as e:
        st.error(f"Error procesando resumen: {e}")
    return None

def solicitar_informe_ia(texto, instruccion_extra=""):
    instruccion_base = (
        "Actúa como un analista experto. Escribe en texto plano basándote en los datos. "
        "Usa párrafos cortos y evita usar asteriscos o formato markdown."
    )
    if instruccion_extra:
        instruccion_base = f"INSTRUCCIÓN DEL USUARIO: {instruccion_extra}\n\n{instruccion_base}"
        
    prompt = f"{instruccion_base}\n\nDATOS:\n{texto[:10000]}"
    
    try:
        model = genai.GenerativeModel(MODELO_ELEGIDO)
        return model.generate_content(prompt).text
    except Exception as e:
        return f"Error al generar respuesta: {e}"

# ==========================================
# PROCESAMIENTO DE ARCHIVOS
# ==========================================

archivo = st.file_uploader("📂 Sube tu archivo (Word, Excel o PDF)", type=["docx", "xlsx", "pdf"])

if archivo:
    texto_extraido = ""
    try:
        if archivo.name.endswith(".docx"):
            doc = Document(archivo)
            texto_extraido = "\n".join([p.text for p in doc.paragraphs])
        elif archivo.name.endswith(".xlsx"):
            wb = openpyxl.load_workbook(archivo, data_only=True)
            for sheet in wb.worksheets:
                for row in sheet.iter_rows(values_only=True):
                    texto_extraido += " | ".join([str(c) for c in row if c]) + "\n"
        elif archivo.name.endswith(".pdf"):
            reader = PyPDF2.PdfReader(archivo)
            for page in reader.pages:
                texto_extraido += page.extract_text() + "\n"
                
        st.success("✅ Documento extraído correctamente")

        col1, col2 = st.columns(2)
        
        with col1:
            if st.button("📝 GENERAR RESUMEN"):
                with st.spinner("Analizando con IA..."):
                    data = solicitar_resumen_estructurado(texto_extraido)
                    if data:
                        info = data.get("datos", {})
                        tipo = data.get("tipo", "Documento")
                        
                        st.markdown(f"📄 **Análisis de {tipo.capitalize()}**")
                        st.markdown(f"🏆 **{info.get('titulo', 'Sin título')}**")
                        st.markdown(f"📝 **Resumen Ejecutivo:**\n{info.get('resumen_ejecutivo', 'No disponible')}")
                        
                        st.markdown("📊 **Métricas Principales:**")
                        metricas = info.get("detalles", {}).get("metricas_principales", {})
                        for clave, valor in metricas.items():
                            st.markdown(f"🔹 **{str(clave).replace('_', ' ').title()}:** {valor}")
                            
                        puntos = info.get("detalles", {}).get("puntos_clave", [])
                        if puntos:
                            st.markdown("📌 **Puntos Clave:**")
                            for p in puntos:
                                if isinstance(p, dict):
                                    v = list(p.values())[0] if p.values() else ""
                                    st.markdown(f"🔸 {v}")
                                else:
                                    st.markdown(f"🔸 {p}")
                    else:
                        st.error("Error al estructurar los datos.")
                        
        with col2:
            if st.button("📄 INFORME EJECUTIVO"):
                with st.spinner("Redactando informe..."):
                    informe = solicitar_informe_ia(texto_extraido)
                    texto_limpio_informe = informe.replace('*', '').replace('#', '')
                    
                    doc_out = Document()
                    doc_out.add_heading('Informe Ejecutivo', 0)
                    for parrafo in texto_limpio_informe.split('\n'):
                        if parrafo.strip(): doc_out.add_paragraph(parrafo.strip())
                        
                    buffer = BytesIO()
                    doc_out.save(buffer)
                    st.download_button("📥 DESCARGAR WORD", buffer.getvalue(), "Informe_Oro.docx")

    except Exception as e:
        st.error(f"Error leyendo el archivo: {e}")

# ==========================================
# SECCIÓN DE MODIFICACIONES (CHAT)
# ==========================================
st.divider()

st.subheader("✍️ Modificaciones específicas")
instruccion = st.text_input("¿Qué quieres que busque, corrija o resuma del archivo?")

if instruccion and archivo:
    with st.spinner("Procesando tu solicitud..."):
        respuesta = solicitar_informe_ia(texto_extraido, instruccion)
        st.info(respuesta)

zona_horaria = pytz.timezone('America/Caracas')
hora_actual = datetime.now(zona_horaria).strftime("%Y-%m-%d %I:%M:%S %p")
st.markdown(f"<p class='footer'>Última actualización de la App: {hora_actual}</p>", unsafe_allow_html=True)
