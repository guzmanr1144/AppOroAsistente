import os, json, ast, re
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
# CONEXIÓN Y SELECCIÓN DE MODELO
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

st.sidebar.title("⚙️ Configuración")
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
# UTILIDADES DE DATOS
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
        except:
            try: return ast.literal_eval(json_str)
            except: pass
    return None

# ==========================================
# FUNCIONES DE INTELIGENCIA ARTIFICIAL
# ==========================================
def solicitar_resumen_estructurado(texto):
    prompt = (
        "Analiza este listado deportivo. Extrae métricas clave.\n"
        "REGLA CRÍTICA: En 'metricas_principales', NO uses listas ni diccionarios. "
        "Escribe los datos como texto simple. Ejemplo: 'Softbol: 2, Atletismo: 5'.\n"
        "Responde solo con JSON.\n"
        '{"tipo": "Listado", "datos": {"titulo": "...", "resumen_ejecutivo": "...", '
        '"detalles": {"puntos_clave": ["..."], "metricas_principales": {"Total": "X", "Disciplinas": "...", "Municipios": "..."}}}}\n\n'
        f"CONTENIDO:\n{texto[:10000]}"
    )
    try:
        model = genai.GenerativeModel(MODELO_ELEGIDO)
        respuesta = model.generate_content(prompt)
        return extraer_json_seguro(respuesta.text, es_lista=False)
    except: return None

def solicitar_informe_ia(texto):
    prompt = f"Escribe un informe ejecutivo en texto plano, párrafos cortos, sin asteriscos.\n\nDATOS:\n{texto[:10000]}"
    try:
        model = genai.GenerativeModel(MODELO_ELEGIDO)
        return model.generate_content(prompt).text
    except: return "Error generando informe."

def solicitar_lista_cambios_aislada(instruccion):
    prompt = (
        f"Extrae la palabra vieja y la nueva de: '{instruccion}'\n"
        "Responde solo JSON: [{\"buscar\": \"texto_viejo\", \"reemplazar\": \"texto_nuevo\"}]"
    )
    try:
        model = genai.GenerativeModel(MODELO_ELEGIDO)
        respuesta = model.generate_content(prompt)
        return extraer_json_seguro(respuesta.text, es_lista=True)
    except: return []

# ==========================================
# PROCESAMIENTO Y REEMPLAZO
# ==========================================
def realizar_reemplazo_docx(archivo_original, cambios):
    doc = Document(archivo_original)
    for c in cambios:
        b, r = str(c["buscar"]), str(c["reemplazar"])
        if not b or not r or b == r: continue
        for p in doc.paragraphs: p.text = p.text.replace(b, r)
        for t in doc.tables:
            for row in t.rows:
                for cell in row.cells:
                    for p in cell.paragraphs: p.text = p.text.replace(b, r)
    buf = BytesIO()
    doc.save(buf)
    return buf.getvalue()

def realizar_reemplazo_xlsx(archivo_original, cambios):
    wb = openpyxl.load_workbook(archivo_original)
    for c in cambios:
        b, r = str(c["buscar"]), str(c["reemplazar"])
        if not b or not r or b == r: continue
        for sheet in wb.worksheets:
            for row in sheet.iter_rows():
                for cell in row:
                    if cell.value and b in str(cell.value):
                        cell.value = str(cell.value).replace(b, r)
    buf = BytesIO()
    wb.save(buf)
    return buf.getvalue()

def generar_pdf_basico(texto):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", size=11)
    for linea in texto.split('\n'):
        pdf.multi_cell(0, 10, txt=linea.encode('latin-1', 'replace').decode('latin-1'))
    return pdf.output(dest='S').encode('latin-1')

# ==========================================
# INTERFAZ PRINCIPAL
# ==========================================
archivo = st.file_uploader("📂 Sube tu archivo (Word, Excel o PDF)", type=["docx", "xlsx", "pdf"])

if archivo:
    texto_extraido = ""
    try:
        if archivo.name.endswith(".docx"):
            doc = Document(archivo)
            texto_extraido = "\n".join([p.text for p in doc.paragraphs])
            for t in doc.tables:
                for row in t.rows: texto_extraido += " | ".join([c.text.strip() for c in row.cells]) + "\n"
        elif archivo.name.endswith(".xlsx"):
            wb = openpyxl.load_workbook(archivo, data_only=True)
            for s in wb.worksheets:
                for r in s.iter_rows(values_only=True): texto_extraido += " | ".join([str(c) for c in r if c]) + "\n"
        elif archivo.name.endswith(".pdf"):
            reader = PyPDF2.PdfReader(archivo)
            for p in reader.pages: texto_extraido += p.extract_text() + "\n"
        
        st.success("✅ Archivo cargado.")
        
        col_res, col_inf = st.columns(2)
        with col_res:
            if st.button("📝 GENERAR RESUMEN"):
                data = solicitar_resumen_estructurado(texto_extraido)
                if data:
                    info = data.get("datos", {})
                    st.markdown(f"🏆 **{info.get('titulo', 'Sin título')}**")
                    st.write(info.get("resumen_ejecutivo", ""))
                    for k, v in info.get("detalles", {}).get("metricas_principales", {}).items():
                        st.markdown(f"🔹 **{k}:** {v}")
        with col_inf:
            if st.button("📄 INFORME WORD"):
                informe = solicitar_informe_ia(texto_extraido)
                doc_inf = Document(); doc_inf.add_paragraph(informe)
                buf = BytesIO(); doc_inf.save(buf)
                st.download_button("📥 Descargar Informe", buf.getvalue(), "Informe.docx")

        st.divider()
        st.subheader("✍️ Edición Quirúrgica")
        instruccion = st.text_input("Cambio a realizar (Ej: Cambia 'REMO' por 'CANOTAJE')")
        
        if instruccion:
            cambios = solicitar_lista_cambios_aislada(instruccion)
            if cambios:
                st.info(f"🔄 Cambio: {cambios[0]['buscar']} ➡️ {cambios[0]['reemplazar']}")
                archivo.seek(0)
                
                # Preparamos los archivos corregidos
                if archivo.name.endswith(".docx"):
                    final_file = realizar_reemplazo_docx(archivo, cambios)
                elif archivo.name.endswith(".xlsx"):
                    final_file = realizar_reemplazo_xlsx(archivo, cambios)
                else:
                    final_file = texto_extraido.replace(cambios[0]['buscar'], cambios[0]['reemplazar']).encode()

                st.markdown("### 📥 Opciones de descarga")
                d1, d2, d3 = st.columns(3)
                d1.download_button("📄 WORD", final_file if archivo.name.endswith(".docx") else b"", "Corregido.docx")
                d2.download_button("📊 EXCEL", final_file if archivo.name.endswith(".xlsx") else b"", "Corregido.xlsx")
                d3.download_button("📕 PDF", generar_pdf_basico(texto_extraido.replace(cambios[0]['buscar'], cambios[0]['reemplazar'])), "Corregido.pdf")
            else:
                st.warning("⚠️ No se identificó el cambio. Intenta: Cambia 'X' por 'Y'.")
    except Exception as e:
        st.error(f"Error: {e}")

zona_horaria = pytz.timezone('America/Caracas')
st.markdown(f"<p class='footer'>Actualizado: {datetime.now(zona_horaria).strftime('%I:%M %p')}</p>", unsafe_allow_html=True)
