import os, json, ast
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

# ==========================================
# CONEXIÓN CON AUTO-DETECTOR DE MODELOS
# ==========================================
try:
    LLAVE_GEMINI = st.secrets["LLAVE_GEMINI"]
    genai.configure(api_key=LLAVE_GEMINI)
    
    modelos_disponibles = []
    for m in genai.list_models():
        if 'generateContent' in m.supported_generation_methods:
            nombre_limpio = m.name.replace("models/", "")
            modelos_disponibles.append(nombre_limpio)
            
    if not modelos_disponibles:
        st.error("❌ Google no habilitó modelos de texto para esta llave.")
        st.stop()
        
    MODELO_ELEGIDO = modelos_disponibles[0]
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

def solicitar_informe_ia(texto):
    prompt = (
        "Actúa como un analista experto. Escribe un informe ejecutivo en texto plano basándote en los datos. "
        "Usa párrafos cortos y evita usar asteriscos o formato markdown.\n\n"
        f"DATOS:\n{texto[:10000]}"
    )
    try:
        model = genai.GenerativeModel(MODELO_ELEGIDO)
        return model.generate_content(prompt).text
    except Exception as e:
        return f"Error al generar respuesta: {e}"

def solicitar_modificacion(texto, instruccion):
    prompt = (
        "Eres un editor de documentos profesional.\n"
        f"INSTRUCCIÓN DEL USUARIO: {instruccion}\n\n"
        "REGLA OBLIGATORIA: Debes devolver EL DOCUMENTO COMPLETO, desde la primera palabra hasta la última. "
        "Aplica la instrucción del usuario, pero MANTÉN TODO EL RESTO DEL TEXTO INTACTO. "
        "NO hagas un resumen, NO recortes partes del texto y NO omitas nada. "
        "No agregues comentarios tuyos al principio ni al final.\n\n"
        f"TEXTO ORIGINAL COMPLETAMENTE:\n{texto[:10000]}"
    )
    try:
        model = genai.GenerativeModel(MODELO_ELEGIDO)
        respuesta = model.generate_content(prompt, generation_config={"max_output_tokens": 8192})
        return respuesta.text
    except Exception as e:
        return f"Error al modificar el texto: {e}"

# ==========================================
# PROCESAMIENTO DE ARCHIVOS
# ==========================================

archivo = st.file_uploader("📂 Sube tu archivo (Word, Excel o PDF)", type=["docx", "xlsx", "pdf"])

if archivo:
    texto_extraido = ""
    try:
        if archivo.name.endswith(".docx"):
            doc = Document(archivo)
            texto_docx = []
            
            # Extraer párrafos normales
            for p in doc.paragraphs:
                if p.text.strip():
                    texto_docx.append(p.text)
                    
            # Extraer contenido de las tablas (LA SOLUCIÓN AL ERROR)
            for table in doc.tables:
                for row in table.rows:
                    fila = [cell.text.strip() for cell in row.cells if cell.text.strip()]
                    if fila:
                        texto_docx.append(" | ".join(fila))
                        
            texto_extraido = "\n".join(texto_docx)
            
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
        
        with st.expander("👁️ Ver vista previa del documento"):
            st.text(texto_extraido[:1500] + "\n... (texto acortado para la vista previa)")

        col1, col2 = st.columns(2)
        
        with col1:
            if st.button("📝 GENERAR RESUMEN"):
                with st.spinner("Analizando con IA..."):
                    data = solicitar_resumen_estructurado(texto_extraido)
                    if data:
                        info = data.get("datos", {})
                        st.markdown(f"🏆 **{info.get('titulo', 'Sin título')}**")
                        st.markdown(f"📝 **Resumen Ejecutivo:**\n{info.get('resumen_ejecutivo', 'No disponible')}")
                        st.markdown("📊 **Métricas Principales:**")
                        for clave, valor in info.get("detalles", {}).get("metricas_principales", {}).items():
                            st.markdown(f"🔹 **{str(clave).replace('_', ' ').title()}:** {valor}")
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

st.subheader("✍️ Modificaciones y Correcciones")
instruccion = st.text_input("¿Qué quieres que corrija, modifique o cambie del archivo?")

if instruccion and archivo:
    with st.spinner("Aplicando los cambios al documento completo..."):
        texto_modificado = solicitar_modificacion(texto_extraido, instruccion)
        st.success("✅ Cambios aplicados correctamente")
        
        with st.expander("Ver texto modificado"):
            st.write(texto_modificado)
        
        st.markdown("### 📥 Descargar documento modificado")
        c_w, c_p, c_e = st.columns(3)
        
        # BOTÓN WORD
        doc_mod = Document()
        for parrafo in texto_modificado.split('\n'):
            if parrafo.strip():
                doc_mod.add_paragraph(parrafo.strip())
        buf_w = BytesIO()
        doc_mod.save(buf_w)
        c_w.download_button("📄 WORD", buf_w.getvalue(), "Documento_Modificado.docx", key="w_mod")
        
        # BOTÓN PDF
        try:
            pdf = FPDF()
            pdf.add_page()
            pdf.set_font("Arial", size=11)
            for linea in texto_modificado.split('\n'):
                texto_limpio = linea.encode('latin-1', 'replace').decode('latin-1')
                pdf.multi_cell(0, 8, txt=texto_limpio)
            
            salida_pdf = pdf.output(dest='S')
            if isinstance(salida_pdf, str):
                pdf_bytes = salida_pdf.encode('latin-1')
            else:
                pdf_bytes = bytes(salida_pdf)
                
            c_p.download_button("📕 PDF", pdf_bytes, "Documento_Modificado.pdf", key="p_mod")
        except Exception as e:
            c_p.error("Error al crear PDF")
            
        # BOTÓN EXCEL
        wb = openpyxl.Workbook()
        ws = wb.active
        for i, linea in enumerate(texto_modificado.split('\n')):
            ws.cell(row=i+1, column=1, value=linea)
        buf_e = BytesIO()
        wb.save(buf_e)
        c_e.download_button("📊 EXCEL", buf_e.getvalue(), "Documento_Modificado.xlsx", key="e_mod")

zona_horaria = pytz.timezone('America/Caracas')
hora_actual = datetime.now(zona_horaria).strftime("%Y-%m-%d %I:%M:%S %p")
st.markdown(f"<p class='footer'>Última actualización de la App: {hora_actual}</p>", unsafe_allow_html=True)
