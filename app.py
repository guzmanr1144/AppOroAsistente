import streamlit as st
import requests
from docx import Document
import openpyxl
import PyPDF2
from io import BytesIO

st.set_page_config(page_title="Oro Asistente 2026", page_icon="🏆")

# LLAVE DE EMERGENCIA - ASEGÚRATE QUE SEA ESTA
LLAVE_GEMINI = "AIzaSyADVQhbwbz6SZR-pT1rfpbf-tqJnFxRg-o"

st.markdown("""
    <style>
    .stButton>button { width: 100%; border-radius: 15px; height: 3.5em; background-color: #007bff; color: white; font-weight: bold; }
    </style>
    """, unsafe_allow_html=True)

st.title("🏆 Oro Asistente Atletas")

def consultar_ia(texto, orden):
    # Nueva URL de conexión más estable
    url = f"https://generativelanguage.googleapis.com/v1/models/gemini-1.5-flash:generateContent?key={LLAVE_GEMINI}"
    headers = {'Content-Type': 'application/json'}
    payload = {
        "contents": [{"parts": [{"text": f"{orden}\n\nDATOS DEL DOCUMENTO:\n{texto[:8000]}"}]}]
    }
    try:
        r = requests.post(url, json=payload, headers=headers, timeout=30)
        respuesta = r.json()
        return respuesta["candidates"][0]["content"]["parts"][0]["text"]
    except Exception as e:
        return f"⚠️ Error: No se pudo conectar con la IA. Detalle: {str(e)}"

archivo = st.file_uploader("📂 Sube tu archivo aquí", type=["docx", "xlsx", "pdf"])

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
                texto_extraido += page.extract_text()
        
        st.success("✅ Documento listo")

        col1, col2 = st.columns(2)
        with col1:
            if st.button("📝 RESUMEN"):
                with st.spinner("Analizando..."):
                    res = consultar_ia(texto_extraido, "Haz un resumen ejecutivo muy estético con emojis. No uses asteriscos.")
                    st.markdown(res)
        with col2:
            if st.button("📄 INFORME"):
                with st.spinner("Redactando..."):
                    informe = consultar_ia(texto_extraido, "Escribe un informe profesional y limpio. Sin asteriscos.")
                    doc_out = Document()
                    doc_out.add_heading("Informe Caracas 2026", 0)
                    doc_out.add_paragraph(informe)
                    buffer = BytesIO()
                    doc_out.save(buffer)
                    st.download_button("📥 DESCARGAR", buffer.getvalue(), "Informe.docx")
    except:
        st.error("Error al leer el archivo.")
