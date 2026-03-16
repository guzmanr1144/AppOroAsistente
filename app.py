import os, json, ast
import streamlit as st
import google.generativeai as genai
from docx import Document
import openpyxl
import PyPDF2
from io import BytesIO
from datetime import datetime
import pytz

st.set_page_config(page_title="Oro Asistente", page_icon="🏆")

# ==========================================
# CONEXIÓN Y BUSCADOR DE MODELOS
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
            
except Exception as e:
    st.error(f"🔑 Error configurando la IA: {e}")
    st.stop()

# MENÚ LATERAL PARA ELEGIR MODELO
st.sidebar.title("⚙️ Configuración")
st.sidebar.info("Si un modelo te da error de límite (Quota exceeded), elige otro de esta lista.")

indice_por_defecto = 0
for i, m in enumerate(modelos_disponibles):
    if '1.5-flash' in m:
        indice_por_defecto = i
        break

MODELO_ELEGIDO = st.sidebar.selectbox("🧠 Modelo de IA:", modelos_disponibles, index=indice_por_defecto)

st.markdown("""
    <style>
    .stButton>button { width: 100%; border-radius: 10px; height: 3.5em; background-color: #007bff; color: white; font-weight: bold; }
    h1 { text-align: center; color: #1e3a8a; }
    .footer { text-align: center; font-size: 12px; color: gray; margin-top: 50px; }
    </style>
    """, unsafe_allow_html=True)

st.title("🏆 Oro Asistente")

# ==========================================
# EXTRACTOR BLINDADO DE DATOS
# ==========================================
def extraer_json_seguro(texto_ia, es_lista=False):
    t = texto_ia.replace("```json", "").replace("```", "").strip()
    char_inicio = "[" if es_lista else "{"
    char_fin = "]" if es_lista else "}"
    
    inicio = t.find(char_inicio)
    fin = t.rfind(char_fin) + 1
    
    if inicio != -1 and fin != 0:
        json_str = t[inicio:fin]
        try:
            return json.loads(json_str, strict=False)
        except Exception:
            try:
                return ast.literal_eval(json_str)
            except Exception:
                pass
    return None

# ==========================================
# FUNCIONES DE INTELIGENCIA ARTIFICIAL
# ==========================================
def solicitar_resumen_estructurado(texto):
    prompt = (
        "Eres un sistema automatizado. Analiza el siguiente documento.\n"
        "Devuelve ÚNICAMENTE un JSON válido. NO escribas saludos ni formato markdown.\n"
        'Estructura EXACTA obligatoria:\n'
        '{"tipo": "Documento", "datos": {"titulo": "...", "resumen_ejecutivo": "...", '
        '"detalles": {"puntos_clave": ["Punto 1"], "metricas_principales": {"Dato": "Valor"}}}}\n\n'
        f"TEXTO A ANALIZAR:\n{texto[:10000]}"
    )
    try:
        model = genai.GenerativeModel(MODELO_ELEGIDO)
        respuesta = model.generate_content(prompt)
        return extraer_json_seguro(respuesta.text, es_lista=False)
    except Exception as e:
        st.error(f"Error de conexión con la IA: {e}")
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
        f"INSTRUCCIÓN DEL USUARIO: {instruccion}\n\n"
        "Busca en el texto original qué palabra o frase exacta debe ser reemplazada.\n"
        "Devuelve ÚNICAMENTE un arreglo JSON con este formato exacto: [{\"buscar\": \"palabra_vieja\", \"reemplazar\": \"palabra_nueva\"}]\n"
        "REGLA DE ORO: La palabra en 'buscar' debe existir EXACTAMENTE en el texto original (respeta mayúsculas y acentos).\n\n"
        f"TEXTO ORIGINAL:\n{texto[:8000]}"
    )
    try:
        model = genai.GenerativeModel(MODELO_ELEGIDO)
        respuesta = model.generate_content(prompt)
        return extraer_json_seguro(respuesta.text, es_lista=True)
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
        
        with st.expander("👁️ Ver vista previa (Ayuda para saber qué copiar)"):
            st.text(texto_extraido[:1500] + "\n... (texto acortado)")

        col1, col2 = st.columns(2)
        with col1:
            if st.button("📝 GENERAR RESUMEN"):
                with st.spinner("Analizando con IA..."):
                    data = solicitar_resumen_estructurado(texto_extraido)
                    if data and isinstance(data, dict):
                        info = data.get("datos", {})
                        st.markdown(f"🏆 **{info.get('titulo', 'Sin título')}**")
                        st.markdown(f"📝 **Resumen Ejecutivo:**\n{info.get('resumen_ejecutivo', '')}")
                        st.markdown("📊 **Métricas Principales:**")
                        for clave, valor in info.get("detalles", {}).get("metricas_principales", {}).items():
                            st.markdown(f"🔹 **{str(clave).replace('_', ' ').title()}:** {valor}")
                    else:
                        st.error("Error: La IA tuvo un problema leyendo la estructura. Intenta de nuevo.")
                        
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

st.subheader("✍️ Edición Quirúrgica (Mantiene tu diseño original)")
instruccion = st.text_input("Escribe exactamente qué palabra quieres cambiar y por cuál (Ej: Cambia 'AJEDREZ' por 'Softbol')")

if instruccion and archivo:
    with st.spinner("Buscando la palabra en tu documento original..."):
        cambios = solicitar_lista_cambios(texto_extraido, instruccion)
        
        if cambios and len(cambios) > 0:
            st.success("✅ ¡Palabra encontrada! Listo para descargar sin perder tu formato.")
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
                st.info("⚠️ Los PDF no mantienen formato. Te damos un archivo de Word con el texto nuevo.")
                texto_nuevo = texto_extraido
                for c in cambios:
                    texto_nuevo = texto_nuevo.replace(c.get("buscar", ""), c.get("reemplazar", ""))
                doc_out = Document()
                doc_out.add_paragraph(texto_nuevo)
                buf = BytesIO()
                doc_out.save(buf)
                st.download_button("📄 DESCARGAR WORD", buf.getvalue(), "Corregido_PDF.docx")
                
        else:
            st.warning("⚠️ No encontré esa palabra exacta. Asegúrate de copiarla idéntica a como aparece en el archivo original (revisa la Vista Previa para ayudarte).")

zona_horaria = pytz.timezone('America/Caracas')
hora_actual = datetime.now(zona_horaria).strftime("%Y-%m-%d %I:%M:%S %p")
st.markdown(f"<p class='footer'>Última actualización de la App: {hora_actual}</p>", unsafe_allow_html=True)
