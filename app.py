import os, time, json, ast
import streamlit as st
import google.generativeai as genai
from docx import Document
import openpyxl
import PyPDF2
from fpdf import FPDF
from io import BytesIO
from datetime import datetime
import pytz

# Configuración de la página
st.set_page_config(page_title="Oro Asistente", page_icon="🏆")

# ==========================================
# CONEXIÓN SEGURA CON LA IA
# ==========================================
try:
    # Busca la llave en la configuración de Streamlit (Secrets)
    LLAVE_GEMINI = st.secrets["LLAVE_GEMINI"]
    genai.configure(api_key=LLAVE_GEMINI)
except Exception:
    st.error("🔑 Error: No se encontró la llave en los Secretos de Streamlit.")
    st.stop()

# Estilos visuales
st.markdown("""
    <style>
    .stButton>button { width: 100%; border-radius: 10px; height: 3.5em; background-color: #007bff; color: white; font-weight: bold; }
    h1 { text-align: center; color: #1e3a8a; }
    .footer { text-align: center; font-size: 12px; color: gray; margin-top: 50px; }
    </style>
    """, unsafe_allow_html=True)

st.title("🏆 Oro Asistente")

# ==========================================
# FUNCIONES DE INTELIGENCIA ARTIFICIAL
# ==========================================

def solicitar_resumen_estructurado(texto, orden_especifica=None):
    instruccion = orden_especifica if orden_especifica else "Analiza el documento."
    prompt = (
        f"INSTRUCCIÓN: {instruccion}\n\n"
        "Responde UNICAMENTE con un objeto JSON válido. No uses markdown.\n"
        'Estructura EXACTA: {"tipo": "...", "datos": {"titulo": "...", "resumen_ejecutivo": "...", '
        '"detalles": {"puntos_clave": ["Punto 1", "Punto 2"], "metricas_principales": {"Dato": "Valor"}}}, "cambios": []}\n\n'
        f"CONTENIDO:\n{texto[:10000]}"
    )
    try:
        model = genai.GenerativeModel('gemini-1.5-flash')
        respuesta = model.generate_content(prompt)
        res_raw = respuesta.text
        inicio, fin = res_raw.find("{"), res_raw.rfind("}") + 1
        if inicio != -1:
            return json.loads(res_raw[inicio:fin], strict=False)
    except Exception as e:
        st.error(f"Error en IA: {e}")
    return None

def solicitar_informe_ia(texto):
    prompt = (
        "Escribe un informe ejecutivo profesional en texto plano basándote en los datos. "
        "Usa párrafos cortos y evita usar asteriscos o formato markdown.\n\n"
        f"DATOS:\n{texto[:10000]}"
    )
    try:
        model = genai.GenerativeModel('gemini-1.5-flash')
        return model.generate_content(prompt).text
    except Exception as e:
        return f"Error al redactar: {e}"

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
                
        st.success("✅ Documento listo para analizar")

        col1, col2 = st.columns(2)
        
        with col1:
            if st.button("📝 GENERAR RESUMEN"):
                with st.spinner("Analizando..."):
                    data = solicitar_resumen_estructurado(texto_extraido)
                    if data:
                        info = data.get("datos", {})
                        st.subheader(f"🏆 {info.get('titulo', 'Resumen')}")
                        st.write(info.get('resumen_ejecutivo', ''))
                        st.info("📊 Métricas y Puntos Clave generados.")
                    else:
                        st.error("No se pudo estructurar el resumen.")
                        
        with col2:
            if st.button("📄 DESCARGAR INFORME"):
                with st.spinner("Preparando Word..."):
                    informe = solicitar_informe_ia(texto_extraido)
                    doc_out = Document()
                    doc_out.add_heading('Informe Ejecutivo', 0)
                    doc_out.add_paragraph(informe)
                    buffer = BytesIO()
                    doc_out.save(buffer)
                    st.download_button("📥 GUARDAR WORD", buffer.getvalue(), "Informe_Oro.docx")

    except Exception as e:
        st.error(f"Error al leer archivo: {e}")

# Pie de página
st.divider()
zona_horaria = pytz.timezone('America/Caracas')
hora_actual = datetime.now(zona_horaria).strftime("%Y-%m-%d %I:%M:%S %p")
st.markdown(f"<p class='footer'>Actualizado: {hora_actual}</p>", unsafe_allow_html=True)
