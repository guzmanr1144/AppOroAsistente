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
    except Exception:
        pass
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
        "Determina qué palabras o frases exactas deben ser reemplazadas en el texto original "
        "para cumplir con la orden del usuario.\n"
        "Responde ÚNICAMENTE con un arreglo JSON válido. No uses markdown.\n"
        'Estructura EXACTA: [{"buscar": "palabra original", "reemplazar": "palabra nueva"}]\n'
        "Si la orden pide redactar algo nuevo o no hay cambios exactos, devuelve []\n\n"
        f"TEXTO ORIGINAL:\n{texto[:8000]}"
    )
    try:
        model = genai.GenerativeModel(MODELO_ELEGIDO)
        respuesta = model.generate_content(prompt)
        res_raw = respuesta.text
        inicio, fin = res_raw.find("["), res_raw.rfind("]") + 1
        if inicio != -1:
            return json.loads(res_raw[inicio:fin], strict=False)
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
        
        # Cambiar en párrafos
        for p in doc.paragraphs:
            if buscar in p.text:
                p.text = p.text.replace(buscar, reemplazar)
        # Cambiar en tablas manteniendo celdas
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
        
        # Cambiar en celdas manteniendo estructura y colores
        for sheet in wb.worksheets:
            for row in sheet.iter_rows():
                for cell in row:
                    if cell.value and isinstance(cell.value, str):
                        if buscar in cell.value:
                            cell.value = cell.value.replace(buscar, reemplazar)
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
                    texto_extraido += " | ".join([str(c) for c in row if c]) + "\n"
                    
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
                    else:
                        st.error("Error al estructurar los datos.")
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
instruccion = st.text_input("¿Qué palabra o frase quieres cambiar del archivo original?")

if instruccion and archivo:
    with st.spinner("Buscando y aplicando los cambios..."):
        cambios = solicitar_lista_cambios(texto_extraido, instruccion)
        
        if cambios:
            st.success("✅ Cambios detectados y aplicados directamente en tu archivo original.")
            for c in cambios:
                st.write(f"🔄 Cambiado: **{c.get('buscar')}** ➡️ **{c.get('reemplazar')}**")
            
            # Reiniciamos el archivo original para editarlo
            archivo.seek(0)
            
            if archivo.name.endswith(".docx"):
                doc_modificado = buscar_y_reemplazar_docx(archivo, cambios)
                st.download_button("📄 DESCARGAR WORD CON FORMATO INTACTO", doc_modificado, f"Corregido_{archivo.name}")
                
            elif archivo.name.endswith(".xlsx"):
                xls_modificado = buscar_y_reemplazar_xlsx(archivo, cambios)
                st.download_button("📊 DESCARGAR EXCEL CON FORMATO INTACTO", xls_modificado, f"Corregido_{archivo.name}")
                
            elif archivo.name.endswith(".pdf"):
                st.info("⚠️ Los PDF no se pueden editar. Te entregamos un Word generado desde cero.")
                texto_nuevo = texto_extraido
                for c in cambios:
                    texto_nuevo = texto_nuevo.replace(c.get("buscar", ""), c.get("reemplazar", ""))
                doc_out = Document()
                doc_out.add_paragraph(texto_nuevo)
                buf = BytesIO()
                doc_out.save(buf)
                st.download_button("📄 DESCARGAR WORD", buf.getvalue(), "Corregido_PDF.docx")
                
        else:
            st.info("No se detectaron reemplazos exactos. Generando respuesta nueva...")
            texto_nuevo = generar_texto_completo(texto_extraido, instruccion)
            st.write(texto_nuevo)

zona_horaria = pytz.timezone('America/Caracas')
hora_actual = datetime.now(zona_horaria).strftime("%Y-%m-%d %I:%M:%S %p")
st.markdown(f"<p class='footer'>Última actualización de la App: {hora_actual}</p>", unsafe_allow_html=True)
