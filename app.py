import streamlit as st
import requests
from docx import Document
import openpyxl
import PyPDF2
from io import BytesIO

# Configuración para que se vea bien en móviles
st.set_page_config(page_title="Oro Asistente 2026", page_icon="🏆", layout="centered")

LLAVE_GEMINI = "AIzaSyADVQhbwbz6SZR-pT1rfpbf-tqJnFxRg-o"

# Diseño estético con CSS
st.markdown("""
    <style>
    .stButton>button { width: 100%; border-radius: 20px; height: 3.5em; background-color: #007bff; color: white; font-weight: bold; border: none; }
    .stFileInfo { background-color: #e8f0fe; border-radius: 15px; }
    h1 { color: #1e3a8a; text-align: center; font-size: 24px; }
    </style>
    """, unsafe_allow_html=True)

st.title("🏆 Oro Asistente Atletas")
st.write("---")

# Función para hablar con la IA
def consultar_ia(texto, orden):
    url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key={LLAVE_GEMINI}"
    payload = {"contents": [{"parts": [{"text": f"{orden}\n\nDATOS:\n{texto[:10000]}"}]}]}
    try:
        r = requests.post(url, json=payload)
        return r.json()["candidates"][0]["content"]["parts"][0]["text"]
    except:
        return "⚠️ Error de conexión con la IA."

# Subida de archivos
archivo = st.file_uploader("📂 Sube tu Word, Excel o PDF aquí", type=["docx", "xlsx", "pdf"])

if archivo:
    texto_extraido = ""
    # Lógica de extracción
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

    st.success("✅ Documento cargado correctamente")

    # Botones principales
    if st.button("📝 VER RESUMEN ESTÉTICO"):
        with st.spinner("Analizando..."):
            res = consultar_ia(texto_extraido, "Haz un resumen muy estético con emojis, limpio y sin asteriscos.")
            st.markdown(res)

    if st.button("📄 GENERAR INFORME (WORD)"):
        with st.spinner("Redactando..."):
            informe = consultar_ia(texto_extraido, "Escribe un informe ejecutivo profesional en texto limpio, sin asteriscos ni símbolos raros.")
            doc_out = Document()
            doc_out.add_heading("Informe Ejecutivo - Caracas 2026", 0)
            doc_out.add_paragraph(informe)
            buffer = BytesIO()
            doc_out.save(buffer)
            st.download_button("📥 DESCARGAR ARCHIVO", buffer.getvalue(), "Informe_Atletas.docx")

st.write("---")
st.caption("Versión 6.5 - Optimizada para dispositivos móviles")
