import os, time, json, ast
import streamlit as st
import google.generativeai as genai
from docx import Document
import openpyxl
import PyPDF2
from io import BytesIO
from datetime import datetime
import pytz

# Configuración de la página
st.set_page_config(page_title="Oro Asistente", page_icon="🏆")

# ==========================================
# CONEXIÓN SEGURA
# ==========================================
try:
    # Usamos la llave que guardaste en Secrets
    genai.configure(api_key=st.secrets["LLAVE_GEMINI"])
except Exception as e:
    st.error("🔑 Error de configuración: Revisa los Secrets de Streamlit.")
    st.stop()

st.title("🏆 Oro Asistente")

# ==========================================
# FUNCIONES DE IA (MODELO COMPATIBILIDAD TOTAL)
# ==========================================

def solicitar_ia_oro(prompt_texto):
    # Probamos con el nombre técnico más compatible de todos
    modelos_a_probar = ['gemini-1.0-pro', 'gemini-1.5-flash-latest']
    
    for nombre_modelo in modelos_a_probar:
        try:
            model = genai.GenerativeModel(nombre_modelo)
            respuesta = model.generate_content(prompt_texto)
            return respuesta.text
        except Exception:
            continue # Si falla uno, intenta el siguiente
    return None

# ==========================================
# LÓGICA DE LA APP
# ==========================================

archivo = st.file_uploader("📂 Sube tu archivo", type=["docx", "xlsx", "pdf"])

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
                    texto_extraido += " ".join([str(c) for c in row if c]) + "\n"
        elif archivo.name.endswith(".pdf"):
            reader = PyPDF2.PdfReader(archivo)
            for page in reader.pages:
                texto_extraido += page.extract_text() + "\n"
        
        st.success("✅ Documento cargado")

        if st.button("📝 GENERAR ANÁLISIS"):
            with st.spinner("La IA está trabajando..."):
                prompt = f"Analiza este texto y haz un resumen ejecutivo profesional:\n\n{texto_extraido[:8000]}"
                resultado = solicitar_ia_oro(prompt)
                
                if resultado:
                    st.markdown("### 📄 Resumen Ejecutivo")
                    st.write(resultado)
                else:
                    st.error("❌ Google no respondió. Prueba crear una llave nueva en una región diferente.")

    except Exception as e:
        st.error(f"Error: {e}")

st.divider()
zona_horaria = pytz.timezone('America/Caracas')
st.caption(f"Actualizado: {datetime.now(zona_horaria).strftime('%I:%M %p')}")
