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

st.set_page_config(page_title="Oro Asistente", page_icon="🏆")

# Configuración OFICIAL de Google
LLAVE_GEMINI = "AIzaSyADVQhbwbz6SZR-pT1rfpbf-tqJnFxRg-o"
genai.configure(api_key=LLAVE_GEMINI)

st.markdown("""
    <style>
    .stButton>button { width: 100%; border-radius: 10px; height: 3.5em; background-color: #007bff; color: white; font-weight: bold; }
    h1 { text-align: center; color: #1e3a8a; }
    .footer { text-align: center; font-size: 12px; color: gray; margin-top: 50px; }
    </style>
    """, unsafe_allow_html=True)

st.title("🏆 Oro Asistente")

# ==========================================
# CEREBRO IA: LIBRERÍA OFICIAL DE GOOGLE
# ==========================================

def solicitar_resumen_estructurado(texto, orden_especifica=None):
    instruccion = orden_especifica if orden_especifica else "Analiza el documento."
    
    prompt = (
        f"INSTRUCCIÓN: {instruccion}\n\n"
        "Responde UNICAMENTE con un objeto JSON válido. No uses markdown, no digas 'aquí tienes el json'.\n"
        'Estructura EXACTA: {"tipo": "...", "datos": {"titulo": "...", "resumen_ejecutivo": "...", '
        '"detalles": {"puntos_clave": ["Punto breve 1", "Punto breve 2"], "metricas_principales": {"Total": "X"}}}, "cambios": []}\n\n'
        "REGLA 1: En 'puntos_clave' escribe máximo 5 viñetas de texto simple.\n"
        "REGLA 2: En 'metricas_principales' usa solo valores simples.\n"
        "REGLA 3: La lista 'cambios' DEBE estar vacía [] a menos que haya una orden explícita.\n\n"
        f"CONTENIDO:\n{texto[:10000]}"
    )

    try:
        # Usamos el modelo estable directamente a través de la librería
        model = genai.GenerativeModel('gemini-1.5-flash')
        respuesta = model.generate_content(prompt)
        res_raw = respuesta.text
        
        inicio = res_raw.find("{")
        fin = res_raw.rfind("}") + 1
        if inicio != -1 and fin != 0:
            res_clean = res_raw[inicio:fin]
            try:
                return json.loads(res_clean, strict=False)
            except json.JSONDecodeError:
                return ast.literal_eval(res_clean)
    except Exception as e:
        st.error(f"Error procesando la IA: {str(e)}")
    return None

def solicitar_informe_ia(texto):
    instruccion = (
        "Actúa como un analista experto y multidisciplinario. Escribe un informe ejecutivo en texto plano basado en los siguientes datos. "
        "Organiza la información de manera lógica, usa párrafos cortos y resalta los puntos más importantes. "
        "No lo hagas demasiado largo. Escribe en texto COMPLETAMENTE PLANO sin usar asteriscos, almohadillas, ni markdown."
    )
    prompt = f"{instruccion}\n\nDATOS A ANALIZAR:\n{texto[:10000]}"
    
    try:
        model = genai.GenerativeModel('gemini-1.5-flash')
        respuesta = model.generate_content(prompt)
        return respuesta.text
    except Exception as e:
        return f"No se pudo generar el informe. Detalle: {str(e)}"

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
                        st.error("Error al procesar los datos.")
                        
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
# SECCIÓN DE MODIFICACIONES Y FECHA
# ==========================================
st.divider()

st.subheader("✍️ Modificaciones específicas")
instruccion = st.text_input("¿Qué quieres que busque o resuma del archivo?")

if instruccion and archivo:
    with st.spinner("Procesando..."):
        respuesta = solicitar_informe_ia(f"INSTRUCCIÓN: {instruccion}\n\nTEXTO:\n{texto_extraido}")
        st.info(respuesta)

zona_horaria = pytz.timezone('America/Caracas')
hora_actual = datetime.now(zona_horaria).strftime("%Y-%m-%d %I:%M:%S %p")
st.markdown(f"<p class='footer'>Última actualización de la App: {hora_actual}</p>", unsafe_allow_html=True)
