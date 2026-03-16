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
# CONEXIÓN SEGURA
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
        st.error("❌ Google no habilitó modelos de texto.")
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
# LIMPIADOR DE ERRORES DE LA IA
# ==========================================
def limpiar_json_ia(texto_raw):
    # Quita las marcas de código que la IA a veces pone por error
    texto = texto_raw.strip()
    if texto.startswith("```json"):
        texto = texto[7:]
    elif texto.startswith("```"):
        texto = texto[3:]
    if texto.endswith("```"):
        texto = texto[:-3]
    return texto.strip()

# ==========================================
# FUNCIONES DE INTELIGENCIA ARTIFICIAL
# ==========================================

def solicitar_resumen_estructurado(texto):
    prompt = (
        "Analiza el documento.\n\n"
        "Responde UNICAMENTE con un objeto JSON válido. NO USES MARKDOWN NI COMILLAS INVERTIDAS.\n"
        'Estructura EXACTA: {"tipo": "...", "datos": {"titulo": "...", "resumen_ejecutivo": "...", '
        '"detalles": {"puntos_clave": ["Punto 1", "Punto 2"], "metricas_principales": {"Dato": "Valor"}}}}\n\n'
        f"CONTENIDO:\n{texto[:10000]}"
    )
    try:
        model = genai.GenerativeModel(MODELO_ELEGIDO)
        respuesta = model.generate_content(prompt)
        res_clean = limpiar_json_ia(respuesta.text)
        
        inicio = res_clean.find("{")
        fin = res_clean.rfind("}") + 1
        if inicio != -1 and fin != 0:
            return json.loads(res_clean[inicio:fin], strict=False)
    except Exception as e:
        st.error(f"Error procesando resumen: La IA no devolvió el formato correcto.")
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

def solicitar_lista_cambios(texto, instruccion):
    prompt = (
        f"INSTRUCCIÓN: {instruccion}\n\n"
        "Determina qué palabras exactas deben ser reemplazadas.\n"
        "REGLA DE ORO: La palabra a 'buscar' debe estar escrita EXACTAMENTE IGUAL a como aparece en el texto original (respeta mayúsculas, minúsculas y acentos).\n"
        "Responde ÚNICAMENTE con un arreglo JSON válido. NO USES MARKDOWN.\n"
        'Estructura EXACTA: [{"buscar": "palabra original", "reemplazar": "palabra nueva"}]\n'
        "Si la orden pide redactar algo nuevo o no hay cambios exactos, devuelve []\n\n"
        f"TEXTO ORIGINAL:\n{texto[:8000]}"
    )
    try:
        model = genai.GenerativeModel(MODELO_ELEGIDO)
        respuesta = model.generate_content(prompt)
        res_clean = limpiar_json_ia(respuesta.text)
        
        inicio = res_clean.find("[")
        fin = res_clean.rfind("]") + 1
        if inicio != -1 and fin != 0:
            return json.loads(res_clean[inicio:fin], strict=False)
    except Exception:
        pass
    return []

# ==========================================
# FUNCIONES DE REEMPLAZO CON FORMATO INTACTO
# ==========================================

def buscar_y_reemplazar_docx(archivo_original, cambios):
    doc = Document(archivo_original)
    for c in cambios:
        buscar = str(c.get("buscar", ""))
        reemplazar = str(c.get("reemplazar", ""))
        if not buscar: continue
        
        for p in doc.paragraphs:
            if buscar in p.text:
                p.text = p.text.replace(buscar, reemplazar)
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        if buscar in p.text:
                            p.text = p.text.replace(buscar, reemplazar)
    buf = BytesIO()
    doc.save(buf)
    return buf.getvalue()

def buscar_y_reemplazar_xlsx(archivo_original, cambios):
    wb = openpyxl.load_workbook(archivo_original)
    for c in cambios:
        buscar = str(c.get("buscar", ""))
        reemplazar = str(c.get("reemplazar", ""))
        if not buscar: continue
        
        for sheet in wb.worksheets:
            for row in sheet.iter_rows():
                for cell in row:
                    if cell.value is not None:
                        valor_celda = str(cell.value)
                        if buscar in valor_celda:
                            cell.value = valor_celda.replace(buscar, reemplazar)
    buf = BytesIO()
    wb.save(buf)
    return buf.getvalue()

def generar_texto_completo(texto, instruccion):
    prompt = (
        f"INSTRUCCIÓN: {instruccion}\n\n"
        "Crea un texto nuevo basado en el original. Escribe en texto plano.\n\n"
        f"TEXTO ORIGINAL:\n{texto[:8000]}"
    )
    try:
        model = genai.GenerativeModel(MODELO_ELEGIDO)
        return model.generate_content(prompt).text
    except Exception:
        return "Error al generar texto nuevo."

# ==========================================
# PROCESAMIENTO DE ARCHIVOS
# ==========================================

archivo = st.file_uploader("📂 Sube tu archivo (Word, Excel o PDF)", type=["docx", "xlsx", "pdf"])

if archivo:
    texto_extraido = ""
    try:
        if archivo.name.endswith(".docx"):
            doc = Document(archivo)
            texto_docx = [p.text for p in doc.paragraphs if p.text.strip()]
            for table in doc.tables:
                for row in table.rows:
                    fila = [cell.text.strip() for cell in row.cells if cell.text.strip()]
                    if fila: texto_docx.append(" | ".join(fila))
            texto_extraido = "\n".join(texto_docx)
            
        elif archivo.name.endswith(".xlsx"):
            wb = openpyxl.load_workbook(archivo, data_only=True)
            for sheet in wb.worksheets:
                for row in sheet.iter_rows(values_only=True):
                    fila = [str(c) for c in row if c is not None]
                    if fila: texto_extraido += " | ".join(fila) + "\n"
                    
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
                        st.markdown(f"🏆 **{info.get('titulo', 'Sin título')}**")
                        st.markdown(f"📝 **Resumen Ejecutivo:**\n{info.get('resumen_ejecutivo', '')}")
                        st.markdown("📊 **Métricas Principales:**")
                        for clave, valor in info.get("detalles", {}).get("metricas_principales", {}).items():
                            st.markdown(f"🔹 **{str(clave).replace('_', ' ').title()}:** {valor}")
                    else:
                        st.error("Error: La IA no envió el formato correcto.")
        with col2:
            if st.button("📄 INFORME EJECUTIVO"):
                with st.spinner("Redactando informe..."):
                    informe = solicitar_informe_ia(texto_extraido)
                    doc_out = Document()
                    doc_out.add_paragraph(informe)
                    buffer = BytesIO()
                    doc_out.save(buffer)
                    st.download_button("📥 DESCARGAR WORD", buffer.getvalue(), "Informe_Oro.docx")

    except Exception as e:
        st.error(f"Error leyendo el archivo: {e}")

# ==========================================
# SECCIÓN DE MODIFICACIONES INTELIGENTES
# ==========================================
st.divider()

st.subheader("✍️ Modificaciones con Formato Original")
instruccion = st.text_input("¿Qué quieres que busque y reemplace en el archivo original?")

if instruccion and archivo:
    with st.spinner("Buscando y aplicando los cambios..."):
        cambios = solicitar_lista_cambios(texto_extraido, instruccion)
        
        if cambios:
            st.success("✅ Cambios detectados y listos para descargar.")
            for c in cambios:
                st.write(f"🔄 Se cambiará: **{c.get('buscar')}** por **{c.get('reemplazar')}**")
            
            archivo.seek(0)
            
            if archivo.name.endswith(".docx"):
                doc_modificado = buscar_y_reemplazar_docx(archivo, cambios)
                st.download_button("📄 DESCARGAR WORD INTACTO", doc_modificado, f"Corregido_{archivo.name}")
                
            elif archivo.name.endswith(".xlsx"):
                xls_modificado = buscar_y_reemplazar_xlsx(archivo, cambios)
                st.download_button("📊 DESCARGAR EXCEL INTACTO", xls_modificado, f"Corregido_{archivo.name}")
                
            elif archivo.name.endswith(".pdf"):
                st.info("⚠️ Los PDF no mantienen formato. Aquí tienes el texto nuevo:")
                texto_nuevo = texto_extraido
                for c in cambios:
                    texto_nuevo = texto_nuevo.replace(c.get("buscar", ""), c.get("reemplazar", ""))
                doc_out = Document()
                doc_out.add_paragraph(texto_nuevo)
                buf = BytesIO()
                doc_out.save(buf)
                st.download_button("📄 DESCARGAR WORD", buf.getvalue(), "Corregido_PDF.docx")
                
        else:
            st.warning("⚠️ No se encontró la palabra exacta para reemplazar, pero aquí tienes una respuesta nueva:")
            texto_nuevo = generar_texto_completo(texto_extraido, instruccion)
            st.write(texto_nuevo)

zona_horaria = pytz.timezone('America/Caracas')
hora_actual = datetime.now(zona_horaria).strftime("%Y-%m-%d %I:%M:%S %p")
st.markdown(f"<p class='footer'>Última actualización de la App: {hora_actual}</p>", unsafe_allow_html=True)
