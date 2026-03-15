import os, time, requests, json, shutil, ast
import streamlit as st
from docx import Document
import openpyxl
import PyPDF2
from fpdf import FPDF
from io import BytesIO

st.set_page_config(page_title="Oro Asistente", page_icon="🏆")

LLAVE_GEMINI = "AIzaSyADVQhbwbz6SZR-pT1rfpbf-tqJnFxRg-o"

st.markdown("""
    <style>
    .stButton>button { width: 100%; border-radius: 10px; height: 3.5em; background-color: #007bff; color: white; font-weight: bold; }
    h1 { text-align: center; color: #1e3a8a; }
    </style>
    """, unsafe_allow_html=True)

st.title("🏆 Oro Asistente")

# ==========================================
# CEREBRO IA INTACTO (IGUAL QUE EN TELEGRAM)
# ==========================================

def solicitar_resumen_estructurado(texto, orden_especifica=None):
    url_list = f"https://generativelanguage.googleapis.com/v1beta/models?key={LLAVE_GEMINI}"
    try:
        r_list = requests.get(url_list, timeout=10)
        modelos_disponibles = [m['name'] for m in r_list.json().get('models', []) if 'generateContent' in m.get('supportedGenerationMethods', [])]
        modelo_usar = modelos_disponibles[0] if modelos_disponibles else "gemini-1.5-flash"
    except:
        modelo_usar = "gemini-1.5-flash"

    instruccion = orden_especifica if orden_especifica else "Analiza el documento."

    payload = {
        "contents": [{"parts": [{"text": (
            f"INSTRUCCIÓN: {instruccion}\n\n"
            "Responde UNICAMENTE con un objeto JSON válido. No uses markdown, no digas 'aquí tienes el json'.\n"
            'Estructura EXACTA: {"tipo": "...", "datos": {"titulo": "...", "resumen_ejecutivo": "...", '
            '"detalles": {"puntos_clave": ["Punto breve 1", "Punto breve 2"], "metricas_principales": {"Total": "X"}}}, "cambios": []}\n\n'
            "REGLA 1: En 'puntos_clave' escribe máximo 5 viñetas de texto simple. NUNCA copies toda la tabla ni uses diccionarios ahí.\n"
            "REGLA 2: En 'metricas_principales' usa solo valores simples (números o texto corto), NO anides listas ni diccionarios.\n"
            "REGLA 3: La lista 'cambios' DEBE estar vacía [] a menos que haya una orden explícita del usuario para reemplazar palabras.\n"
            f"CONTENIDO:\n{texto[:10000]}"
        )}]}],
        "safetySettings": [{"category": c, "threshold": "BLOCK_NONE"} for c in [
            "HARM_CATEGORY_HARASSMENT", "HARM_CATEGORY_HATE_SPEECH", 
            "HARM_CATEGORY_SEXUALLY_EXPLICIT", "HARM_CATEGORY_DANGEROUS_CONTENT"
        ]]
    }

    try:
        url = f"https://generativelanguage.googleapis.com/v1beta/{modelo_usar}:generateContent?key={LLAVE_GEMINI}"
        r = requests.post(url, json=payload, timeout=30)
        res_data = r.json()

        if "candidates" in res_data:
            res_raw = res_data["candidates"][0]["content"]["parts"][0]["text"]
            inicio = res_raw.find("{")
            fin = res_raw.rfind("}") + 1
            if inicio != -1 and fin != 0:
                res_clean = res_raw[inicio:fin]
                try:
                    return json.loads(res_clean, strict=False)
                except json.JSONDecodeError:
                    try:
                        return ast.literal_eval(res_clean)
                    except:
                        pass
    except Exception as e:
        st.error(f"Excepción IA: {str(e)}")
    return None

def solicitar_informe_ia(texto):
    url_list = f"https://generativelanguage.googleapis.com/v1beta/models?key={LLAVE_GEMINI}"
    try:
        r_list = requests.get(url_list, timeout=10)
        modelos = [m['name'] for m in r_list.json().get('models', []) if 'generateContent' in m.get('supportedGenerationMethods', [])]
        modelo = modelos[0] if modelos else "gemini-1.5-flash"
    except: 
        modelo = "gemini-1.5-flash"

    instruccion = (
        "Actúa como un analista experto y multidisciplinario. Escribe un informe ejecutivo en texto plano basado en los siguientes datos. "
        "Identifica automáticamente de qué trata el documento y adapta tu tono para que sea institucional y profesional. "
        "Organiza la información de manera lógica, usa párrafos cortos y resalta los puntos más importantes de forma fácil de entender. "
        "No lo hagas demasiado largo, máximo una o dos páginas. Escribe en texto COMPLETAMENTE PLANO sin usar asteriscos (*), almohadillas (#), ni ningún tipo de markdown."
    )
    payload = {
        "contents": [{"parts": [{"text": f"{instruccion}\n\nDATOS A ANALIZAR:\n{texto[:10000]}"}]}],
        "safetySettings": [{"category": c, "threshold": "BLOCK_NONE"} for c in [
            "HARM_CATEGORY_HARASSMENT", "HARM_CATEGORY_HATE_SPEECH", 
            "HARM_CATEGORY_SEXUALLY_EXPLICIT", "HARM_CATEGORY_DANGEROUS_CONTENT"
        ]]
    }
    try:
        r = requests.post(f"https://generativelanguage.googleapis.com/v1beta/{modelo}:generateContent?key={LLAVE_GEMINI}", json=payload, timeout=30)
        return r.json()["candidates"][0]["content"]["parts"][0]["text"]
    except Exception: 
        return "No se pudo generar el informe."

# ==========================================
# INTERFAZ Y MANEJO DE ARCHIVOS
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
                
        st.success("✅ Documento extraído correctamente")

        col1, col2 = st.columns(2)
        
        with col1:
            if st.button("📝 GENERAR RESUMEN"):
                with st.spinner("Conectando con IA..."):
                    data = solicitar_resumen_estructurado(texto_extraido)
                    if data:
                        info = data.get("datos", {})
                        tipo = data.get("tipo", "Documento")
                        
                        # Formato idéntico al de Telegram
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
                        st.error("La IA no devolvió el formato correcto.")
                        
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
        st.error(f"Error procesando el archivo: {e}")
