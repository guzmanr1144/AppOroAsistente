import streamlit as st
import requests
from docx import Document
import openpyxl
import PyPDF2
from io import BytesIO

st.set_page_config(page_title="Oro Asistente", page_icon="🏆")

LLAVE_GEMINI = "AIzaSyADVQhbwbz6SZR-pT1rfpbf-tqJnFxRg-o"

# Estilo visual
st.markdown("""
    <style>
    .stButton>button { width: 100%; border-radius: 15px; height: 3.5em; background-color: #007bff; color: white; font-weight: bold; }
    h1 { text-align: center; color: #1e3a8a; }
    </style>
    """, unsafe_allow_html=True)

# Título actualizado como pediste
st.title("🏆 Oro Asistente")

def consultar_ia(texto, orden):
    url = f"https://generativelanguage.googleapis.com/v1/models/gemini-1.5-flash:generateContent?key={LLAVE_GEMINI}"
    headers = {'Content-Type': 'application/json'}
    payload = {
        "contents": [{"parts": [{"text": f"{orden}\n\nDATOS DEL DOCUMENTO:\n{texto[:8000]}"}]}]
    }
    try:
        r = requests.post(url, json=payload, headers=headers, timeout=30)
        return r.json()["candidates"][0]["content"]["parts"][0]["text"]
    except:
        return "⚠️ Error de conexión. Intenta de nuevo."

archivo = st.file_uploader("📂 Sube tu archivo", type=["docx", "xlsx", "pdf"])

if archivo:
    texto_extraido = ""
    # (El código de extracción se mantiene igual)
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
    
    st.success("✅ Documento cargado")

    col1, col2 = st.columns(2)
    with col1:
        if st.button("📝 RESUMEN"):
            with st.spinner("Analizando..."):
                res = consultar_ia(texto_extraido, "Haz un resumen ejecutivo estético con emojis.")
                st.markdown(res)
    with col2:
        if st.button("📄 INFORME"):
            with st.spinner("Redactando..."):
                informe = consultar_ia(texto_extraido, "Escribe un informe profesional sin asteriscos.")
                doc_out = Document()
                doc_out.add_paragraph(informe)
                buffer = BytesIO()
                doc_out.save(buffer)
                st.download_button("📥 DESCARGAR", buffer.getvalue(), "Informe.docx")

    # --- SECCIÓN PARA MODIFICAR COSAS ---
    st.divider()
    st.subheader("✍️ Modificaciones específicas")
    instruccion = st.text_input("¿Qué quieres que haga con la información? (Ej: 'Busca el atleta X y dime su edad' o 'Cambia el formato a lista')")
    
    if instruccion:
        with st.spinner("Procesando..."):
            respuesta_personalizada = consultar_ia(texto_extraido, instruccion)
            st.info(f"Resultado:\n\n{respuesta_personalizada}")
