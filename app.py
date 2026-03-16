import os, time, requests, json, shutil, ast
import streamlit as st
from docx import Document
import openpyxl
import PyPDF2
from fpdf import FPDF
from io import BytesIO
from datetime import datetime
import pytz

st.set_page_config(page_title="Oro Asistente", page_icon="🏆")

# Clave de API
LLAVE_GEMINI = "AIzaSyADVQhbwbz6SZR-pT1rfpbf-tqJnFxRg-o"

st.markdown("""
    <style>
    .stButton>button { width: 100%; border-radius: 10px; height: 3.5em; background-color: #007bff; color: white; font-weight: bold; }
    h1 { text-align: center; color: #1e3a8a; }
    .footer { text-align: center; font-size: 12px; color: gray; margin-top: 50px; }
    </style>
    """, unsafe_allow_html=True)

st.title("🏆 Oro Asistente")

# ==========================================
# CEREBRO IA CORREGIDO (BÚSQUEDA DINÁMICA)
# ==========================================

def obtener_modelo_valido():
    """Busca el mejor modelo disponible en la cuenta."""
    try:
        # Probamos primero con v1beta
        url_list = f"https://generativelanguage.googleapis.com/v1beta/models?key={LLAVE_GEMINI}"
        r_list = requests.get(url_list, timeout=10)
        if r_list.status_code == 200:
            modelos = r_list.json().get('models', [])
            # Buscamos gemini-1.5-flash o el primero que soporte generación
            for m in modelos:
                if "gemini-1.5-flash" in m['name'] and "generateContent" in m['supportedGenerationMethods']:
                    return m['name'] # Ya viene como "models/gemini-1.5-flash"
            
            # Si no está flash, el primero disponible
            for m in modelos:
                if "generateContent" in m['supportedGenerationMethods']:
                    return m['name']
    except:
        pass
    return "models/gemini-1.5-flash" # Fallback manual

def solicitar_ia(payload, endpoint="generateContent"):
    modelo = obtener_modelo_valido()
    # Limpieza: si el modelo ya tiene "models/", no lo agregamos de nuevo
    nombre_final = modelo if modelo.startswith("models/") else f"models/{modelo}"
    
    # Intentamos con v1beta, si falla, v1
    for version in ["v1beta", "v1"]:
        url = f"https://generativelanguage.googleapis.com/{version}/{nombre_final}:{endpoint}?key={LLAVE_GEMINI}"
        try:
            r = requests.post(url, json=payload, timeout=30)
            if r.status_code == 200:
                return r.json()
        except:
            continue
    return None

def solicitar_resumen_estructurado(texto, orden_especifica=None):
    instruccion = orden_especifica if orden_especifica else "Analiza el documento."
    
    payload = {
        "contents": [{"parts": [{"text": (
            f"INSTRUCCIÓN: {instruccion}\n\n"
            "Responde UNICAMENTE con un objeto JSON válido. No uses markdown.\n"
            'Estructura: {"tipo": "...", "datos": {"titulo": "...", "resumen_ejecutivo": "...", '
            '"detalles": {"puntos_clave": ["Punto 1"], "metricas_principales": {"Total": "X"}}}, "cambios": []}\n\n'
            f"CONTENIDO:\n{texto[:10000]}"
        )}]}],
        "safetySettings": [{"category": c, "threshold": "BLOCK_NONE"} for c in [
            "HARM_CATEGORY_HARASSMENT", "HARM_CATEGORY_HATE_SPEECH", 
            "HARM_CATEGORY_SEXUALLY_EXPLICIT", "HARM_CATEGORY_DANGEROUS_CONTENT"
        ]]
    }

    res_data = solicitar_ia(payload)
    if res_data and "candidates" in res_data:
        try:
            res_raw = res_data["candidates"][0]["content"]["parts"][0]["text"]
            inicio, fin = res_raw.find("{"), res_raw.rfind("}") + 1
            return json.loads(res_raw[inicio:fin], strict=False)
        except:
            return None
    return None

def solicitar_informe_ia(texto):
    instruccion = (
        "Actúa como un analista experto. Escribe un informe ejecutivo en texto plano. "
        "Sin asteriscos ni markdown. Máximo 2 páginas."
    )
    payload = {
        "contents": [{"parts": [{"text": f"{instruccion}\n\nDATOS:\n{texto[:10000]}"}]}]
    }
    res_data = solicitar_ia(payload)
    if res_data and "candidates" in res_data:
        return res_data["candidates"][0]["content"]["parts"][0]["text"]
    return "No se pudo generar el informe."

# ==========================================
# INTERFAZ Y PROCESAMIENTO
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
                    fila = [str(c) for c in row if c is not None]
                    if fila: texto_extraido += " | ".join(fila) + "\n"
        elif archivo.name.endswith(".pdf"):
            reader = PyPDF2.PdfReader(archivo)
            for page in reader.pages:
                ext = page.extract_text()
                if ext: texto_extraido += ext + "\n"
                
        st.success("✅ Documento cargado")

        col1, col2 = st.columns(2)
        
        with col1:
            if st.button("📝 GENERAR RESUMEN"):
                with st.spinner("Buscando modelo y analizando..."):
                    data = solicitar_resumen_estructurado(texto_extraido)
                    if data:
                        info = data.get("datos", {})
                        st.markdown(f"🏆 **{info.get('titulo', 'Resumen')}**")
                        st.write(info.get('resumen_ejecutivo', ''))
                        st.json(info.get('detalles', {}))
                    else:
                        st.error("Error al obtener respuesta de la IA.")
                        
        with col2:
            if st.button("📄 INFORME EJECUTIVO"):
                with st.spinner("Redactando..."):
                    informe = solicitar_informe_ia(texto_extraido)
                    st.text_area("Informe Generado", informe, height=300)

    except Exception as e:
        st.error(f"Error: {e}")

st.divider()
instruccion_usuario = st.text_input("¿Qué quieres saber del archivo?")
if instruccion_usuario and archivo:
    res = solicitar_informe_ia(f"ORDEN: {instruccion_usuario}\n\nTEXTO: {texto_extraido}")
    st.info(res)

zona_horaria = pytz.timezone('America/Caracas')
st.markdown(f"<p class='footer'>Actualizado: {datetime.now(zona_horaria).strftime('%I:%M %p')}</p>", unsafe_allow_html=True)
