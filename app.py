import os, json, ast, re
import streamlit as st
import google.generativeai as genai
from docx import Document
from docx.shared import Pt, RGBColor
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import PyPDF2
from fpdf import FPDF
from io import BytesIO
from datetime import datetime
import pytz

# ==========================================
# CONFIGURACIÓN DE PÁGINA
# ==========================================
st.set_page_config(
    page_title="Oro Asistente",
    page_icon="🏆",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ==========================================
# ESTILOS CSS
# ==========================================
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap');

html, body, [class*="css"] { font-family: 'Inter', sans-serif; }

.main { background: #0f1117; }

/* Tarjeta de métricas */
.metric-card {
    background: linear-gradient(135deg, #1e2530, #252d3a);
    border: 1px solid #2e3a4e;
    border-radius: 12px;
    padding: 1rem 1.2rem;
    margin-bottom: 0.7rem;
}
.metric-label { color: #8b9ab5; font-size: 0.78rem; font-weight: 500; text-transform: uppercase; letter-spacing: 0.05em; }
.metric-value { color: #e8edf5; font-size: 1.4rem; font-weight: 700; margin-top: 0.2rem; }

/* Resumen */
.summary-box {
    background: linear-gradient(135deg, #1a2332, #1e2d40);
    border: 1px solid #2a4a6b;
    border-left: 4px solid #3b82f6;
    border-radius: 12px;
    padding: 1.5rem;
    margin: 1rem 0;
    color: #c8d8ec;
    line-height: 1.7;
    font-size: 0.95rem;
}
.summary-title {
    color: #60a5fa;
    font-size: 1.1rem;
    font-weight: 700;
    margin-bottom: 0.8rem;
    display: flex;
    align-items: center;
    gap: 0.5rem;
}

/* Tag de puntos clave */
.tag {
    display: inline-block;
    background: #1e3a5f;
    color: #60a5fa;
    border: 1px solid #2a5080;
    border-radius: 20px;
    padding: 0.3rem 0.8rem;
    font-size: 0.8rem;
    margin: 0.2rem;
    font-weight: 500;
}

/* Header */
.app-header {
    text-align: center;
    padding: 2rem 0 1rem 0;
}
.app-title {
    font-size: 2.5rem;
    font-weight: 800;
    background: linear-gradient(135deg, #f59e0b, #fbbf24, #f59e0b);
    -webkit-background-clip: text;
    -webkit-text-fill-color: transparent;
    background-clip: text;
}
.app-subtitle {
    color: #6b7280;
    font-size: 0.95rem;
    margin-top: 0.3rem;
}

/* Botones */
.stButton > button {
    background: linear-gradient(135deg, #1d4ed8, #2563eb) !important;
    color: white !important;
    border: none !important;
    border-radius: 10px !important;
    font-weight: 600 !important;
    height: 3rem !important;
    transition: all 0.2s !important;
}
.stButton > button:hover {
    background: linear-gradient(135deg, #1e40af, #1d4ed8) !important;
    transform: translateY(-1px) !important;
    box-shadow: 0 4px 15px rgba(37, 99, 235, 0.4) !important;
}

/* Uploader */
[data-testid="stFileUploader"] {
    background: #1e2530;
    border: 2px dashed #2e3a4e;
    border-radius: 14px;
    padding: 1rem;
}

/* Divider */
.custom-divider {
    height: 1px;
    background: linear-gradient(90deg, transparent, #2e3a4e, transparent);
    margin: 1.5rem 0;
}

/* Success/warning boxes */
.info-box {
    background: #0f2a1e;
    border: 1px solid #166534;
    border-radius: 10px;
    padding: 0.8rem 1rem;
    color: #4ade80;
    font-size: 0.9rem;
    margin: 0.5rem 0;
}
.warn-box {
    background: #2a1a0a;
    border: 1px solid #92400e;
    border-radius: 10px;
    padding: 0.8rem 1rem;
    color: #fbbf24;
    font-size: 0.9rem;
    margin: 0.5rem 0;
}

/* Sidebar */
[data-testid="stSidebar"] {
    background: #111827 !important;
}

.footer {
    text-align: center;
    font-size: 0.75rem;
    color: #374151;
    padding: 2rem 0 0.5rem 0;
    border-top: 1px solid #1f2937;
    margin-top: 3rem;
}
</style>
""", unsafe_allow_html=True)

# ==========================================
# CONEXIÓN GEMINI
# ==========================================
try:
    LLAVE_GEMINI = st.secrets["LLAVE_GEMINI"]
    genai.configure(api_key=LLAVE_GEMINI)
    modelos_disponibles = []
    for m in genai.list_models():
        if 'generateContent' in m.supported_generation_methods:
            nombre = m.name.replace("models/", "")
            modelos_disponibles.append(nombre)
    if not modelos_disponibles:
        st.error("❌ No se encontraron modelos disponibles.")
        st.stop()
except Exception as e:
    st.error(f"🔑 Error configurando la IA: {e}")
    st.stop()

# ==========================================
# SIDEBAR
# ==========================================
with st.sidebar:
    st.markdown("### ⚙️ Configuración")
    idx_default = next((i for i, m in enumerate(modelos_disponibles) if '1.5-flash' in m), 0)
    MODELO_ELEGIDO = st.selectbox("🧠 Modelo IA:", modelos_disponibles, index=idx_default)
    
    st.markdown("---")
    st.markdown("### 📌 Guía rápida")
    st.markdown("""
    1. **Sube** un archivo Word, Excel o PDF
    2. **Resumen** → análisis instantáneo
    3. **Edición** → corrige palabras en el doc
    4. **Exporta** en Word, Excel o PDF
    """)
    
    st.markdown("---")
    st.markdown("### 🆕 Funciones disponibles")
    st.markdown("""
    - 📝 Resumen estructurado
    - ✍️ Corrección quirúrgica
    - 📊 Exportar a Excel formateado
    - 📄 Exportar a Word con estilo
    - 📕 Exportar a PDF
    - 💬 Pregunta libre sobre el doc
    """)

# ==========================================
# SESSION STATE
# ==========================================
for key, val in {
    "texto_extraido": "",
    "nombre_archivo": "",
    "archivo_bytes": None,
    "resumen_data": None,
    "historial_chat": [],
    "cambios_aplicados": None,
    "archivo_tipo": "",
}.items():
    if key not in st.session_state:
        st.session_state[key] = val

# ==========================================
# HEADER
# ==========================================
st.markdown("""
<div class="app-header">
    <div class="app-title">🏆 Oro Asistente</div>
    <div class="app-subtitle">Analiza, edita y exporta documentos deportivos con IA</div>
</div>
""", unsafe_allow_html=True)

# ==========================================
# UTILIDADES JSON
# ==========================================
def extraer_json_seguro(texto, es_lista=False):
    t = texto.replace("```json", "").replace("```", "").strip()
    c1, c2 = ("[", "]") if es_lista else ("{", "}")
    inicio = t.find(c1)
    fin = t.rfind(c2) + 1
    if inicio != -1 and fin > 0:
        try:
            return json.loads(t[inicio:fin], strict=False)
        except:
            try:
                return ast.literal_eval(t[inicio:fin])
            except:
                pass
    return None

# ==========================================
# FUNCIONES IA
# ==========================================
def solicitar_resumen_estructurado(texto):
    prompt = (
        "Eres un analista deportivo profesional. Analiza este documento y devuelve SOLO un JSON.\n"
        "El resumen_ejecutivo debe ser amigable, directo y máximo 3 oraciones. "
        "metricas_principales deben ser strings simples (no objetos).\n"
        "Formato exacto:\n"
        '{"titulo": "...", "emoji_categoria": "⚽", "resumen_ejecutivo": "...", '
        '"metricas": {"Clave1": "Valor1", "Clave2": "Valor2"}, '
        '"puntos_clave": ["punto 1", "punto 2", "punto 3"], '
        '"hallazgo_destacado": "Una observación importante o curiosa del documento"}\n\n'
        f"DOCUMENTO:\n{texto[:12000]}"
    )
    try:
        model = genai.GenerativeModel(MODELO_ELEGIDO)
        resp = model.generate_content(prompt)
        return extraer_json_seguro(resp.text)
    except Exception as e:
        return None

def solicitar_informe_word(texto):
    prompt = (
        "Escribe un informe ejecutivo deportivo profesional. "
        "Usa párrafos cortos y claros, sin asteriscos ni markdown. "
        "Incluye: introducción, hallazgos principales, análisis y conclusión.\n\n"
        f"DATOS:\n{texto[:12000]}"
    )
    try:
        model = genai.GenerativeModel(MODELO_ELEGIDO)
        return model.generate_content(prompt).text
    except:
        return "Error generando el informe."

def extraer_cambio_con_regex(instruccion):
    """
    Fallback: detecta patrones comunes sin necesidad de IA.
    Soporta variantes como:
      - cambia 'X' por 'Y'
      - reemplaza X por Y
      - X → Y  /  X -> Y
      - sustituye X con Y
    """
    patrones = [
        r"(?:cambia|reemplaza|sustituye|cambie|reemplaz[ao])\s+['\"]?(.+?)['\"]?\s+(?:por|con|a)\s+['\"]?(.+?)['\"]?\s*$",
        r"['\"](.+?)['\"]\s*(?:→|->|=>|por|con)\s*['\"]?(.+?)['\"]?\s*$",
        r"(.+?)\s*(?:→|->|=>)\s*(.+)",
    ]
    texto = instruccion.strip()
    for pat in patrones:
        m = re.search(pat, texto, re.IGNORECASE)
        if m:
            buscar = m.group(1).strip().strip("'\"")
            reemplazar = m.group(2).strip().strip("'\"")
            if buscar and reemplazar:
                return [{"buscar": buscar, "reemplazar": reemplazar}]
    return []

def solicitar_cambios(instruccion):
    prompt = (
        "Eres un asistente que extrae instrucciones de edición de texto.\n"
        "El usuario quiere cambiar una palabra o frase por otra en un documento.\n\n"
        f"INSTRUCCIÓN DEL USUARIO: \"{instruccion}\"\n\n"
        "Extrae el texto a buscar y el texto de reemplazo.\n"
        "REGLAS IMPORTANTES:\n"
        "- Si el usuario dice 'cambia X por Y', buscar=X y reemplazar=Y\n"
        "- Si dice 'reemplaza X con Y', buscar=X y reemplazar=Y\n"
        "- Si dice 'X → Y' o 'X -> Y', buscar=X y reemplazar=Y\n"
        "- Devuelve los valores EXACTOS sin modificar mayúsculas/minúsculas\n"
        "- Si hay múltiples cambios, incluye todos en el array\n\n"
        "Responde ÚNICAMENTE con este JSON (sin texto adicional, sin markdown):\n"
        '[{"buscar": "texto_exacto_a_buscar", "reemplazar": "texto_nuevo"}]'
    )
    try:
        model = genai.GenerativeModel(MODELO_ELEGIDO)
        resp = model.generate_content(prompt)
        resultado = extraer_json_seguro(resp.text, es_lista=True)
        # Validar que el resultado tiene la estructura correcta
        if resultado and isinstance(resultado, list):
            validos = [
                c for c in resultado
                if isinstance(c, dict)
                and "buscar" in c and "reemplazar" in c
                and str(c["buscar"]).strip()
                and str(c["reemplazar"]).strip()
            ]
            if validos:
                return validos
    except:
        pass
    # Fallback: intentar con regex
    return extraer_cambio_con_regex(instruccion)

def preguntar_al_documento(pregunta, texto):
    historial = st.session_state.historial_chat
    contexto = "\n".join([f"{m['rol']}: {m['texto']}" for m in historial[-6:]])
    prompt = (
        f"Eres un asistente experto analizando este documento deportivo.\n"
        f"DOCUMENTO:\n{texto[:10000]}\n\n"
        f"CONVERSACIÓN PREVIA:\n{contexto}\n\n"
        f"PREGUNTA: {pregunta}\n"
        "Responde de forma concisa y directa en español."
    )
    try:
        model = genai.GenerativeModel(MODELO_ELEGIDO)
        return model.generate_content(prompt).text
    except:
        return "No pude procesar tu pregunta."

def detectar_anomalias(texto):
    prompt = (
        "Analiza este documento deportivo y detecta posibles inconsistencias, "
        "datos duplicados, errores o anomalías. Sé breve y directo.\n"
        "Devuelve SOLO JSON:\n"
        '{"anomalias": ["anomalía 1", "anomalía 2"], "nivel_calidad": "Alto/Medio/Bajo", '
        '"recomendacion": "texto breve"}\n\n'
        f"DOCUMENTO:\n{texto[:10000]}"
    )
    try:
        model = genai.GenerativeModel(MODELO_ELEGIDO)
        resp = model.generate_content(prompt)
        return extraer_json_seguro(resp.text)
    except:
        return None

# ==========================================
# EXPORTADORES
# ==========================================
def exportar_word(texto, resumen_data=None):
    """Genera Word con formato profesional."""
    doc = Document()
    
    # Estilos del documento
    style = doc.styles['Normal']
    style.font.name = 'Calibri'
    style.font.size = Pt(11)
    
    # Título principal
    titulo = doc.add_heading('', 0)
    run = titulo.add_run('🏆 INFORME EJECUTIVO')
    run.font.color.rgb = RGBColor(0x1E, 0x40, 0xAF)
    run.font.size = Pt(22)
    
    # Fecha
    zona = pytz.timezone('America/Caracas')
    fecha = datetime.now(zona).strftime('%d de %B de %Y, %I:%M %p')
    p_fecha = doc.add_paragraph()
    run_fecha = p_fecha.add_run(f'Generado: {fecha}')
    run_fecha.font.size = Pt(9)
    run_fecha.font.color.rgb = RGBColor(0x6B, 0x72, 0x80)
    
    doc.add_paragraph()
    
    # Si hay resumen estructurado, incluirlo
    if resumen_data:
        h2 = doc.add_heading('Resumen Ejecutivo', level=1)
        h2.runs[0].font.color.rgb = RGBColor(0x1E, 0x40, 0xAF)
        doc.add_paragraph(resumen_data.get("resumen_ejecutivo", ""))
        
        if resumen_data.get("metricas"):
            doc.add_heading('Métricas Clave', level=2)
            tabla = doc.add_table(rows=1, cols=2)
            tabla.style = 'Table Grid'
            hdr = tabla.rows[0].cells
            hdr[0].text = 'Indicador'
            hdr[1].text = 'Valor'
            for cell in hdr:
                for run in cell.paragraphs[0].runs:
                    run.font.bold = True
            for k, v in resumen_data["metricas"].items():
                row = tabla.add_row().cells
                row[0].text = str(k)
                row[1].text = str(v)
            doc.add_paragraph()
        
        if resumen_data.get("puntos_clave"):
            doc.add_heading('Puntos Clave', level=2)
            for punto in resumen_data["puntos_clave"]:
                p = doc.add_paragraph(style='List Bullet')
                p.add_run(punto)
        
        if resumen_data.get("hallazgo_destacado"):
            doc.add_paragraph()
            doc.add_heading('💡 Hallazgo Destacado', level=2)
            p_hall = doc.add_paragraph()
            run_hall = p_hall.add_run(resumen_data["hallazgo_destacado"])
            run_hall.font.italic = True
            run_hall.font.color.rgb = RGBColor(0x1D, 0x4E, 0xD8)
        
        doc.add_page_break()
    
    # Contenido completo
    doc.add_heading('Contenido del Documento', level=1)
    for linea in texto.split('\n'):
        if linea.strip():
            doc.add_paragraph(linea.strip())
    
    buf = BytesIO()
    doc.save(buf)
    return buf.getvalue()


def exportar_excel(texto, resumen_data=None):
    """Genera Excel formateado profesionalmente."""
    wb = openpyxl.Workbook()
    
    # ---- HOJA 1: RESUMEN ----
    ws_res = wb.active
    ws_res.title = "📊 Resumen"
    
    # Colores
    azul_oscuro = "1E3A5F"
    azul_medio  = "2563EB"
    azul_claro  = "DBEAFE"
    blanco      = "FFFFFF"
    gris_claro  = "F8FAFC"
    
    def header_cell(ws, row, col, texto_cell, bg=azul_oscuro, fg=blanco, size=12, bold=True):
        cell = ws.cell(row=row, column=col, value=texto_cell)
        cell.fill = PatternFill("solid", fgColor=bg)
        cell.font = Font(color=fg, bold=bold, size=size)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        return cell
    
    def data_cell(ws, row, col, texto_cell, bg=blanco, bold=False, align="left"):
        cell = ws.cell(row=row, column=col, value=texto_cell)
        cell.fill = PatternFill("solid", fgColor=bg)
        cell.font = Font(bold=bold, size=11)
        cell.alignment = Alignment(horizontal=align, vertical="center", wrap_text=True)
        return cell
    
    thin = Border(
        left=Side(style='thin', color='CBD5E1'),
        right=Side(style='thin', color='CBD5E1'),
        top=Side(style='thin', color='CBD5E1'),
        bottom=Side(style='thin', color='CBD5E1')
    )
    
    # Título
    ws_res.merge_cells("A1:D1")
    header_cell(ws_res, 1, 1, "🏆 ORO ASISTENTE — REPORTE EJECUTIVO", bg=azul_oscuro, size=14)
    ws_res.row_dimensions[1].height = 40
    
    zona = pytz.timezone('America/Caracas')
    fecha = datetime.now(zona).strftime('%d/%m/%Y %I:%M %p')
    ws_res.merge_cells("A2:D2")
    data_cell(ws_res, 2, 1, f"Generado: {fecha}", bg=azul_claro, align="center")
    
    fila = 4
    if resumen_data:
        titulo_doc = resumen_data.get("titulo", "Sin título")
        ws_res.merge_cells(f"A{fila}:D{fila}")
        header_cell(ws_res, fila, 1, titulo_doc, bg=azul_medio, size=12)
        ws_res.row_dimensions[fila].height = 30
        fila += 1
        
        resumen_ej = resumen_data.get("resumen_ejecutivo", "")
        if resumen_ej:
            ws_res.merge_cells(f"A{fila}:D{fila+2}")
            cell = ws_res.cell(row=fila, column=1, value=resumen_ej)
            cell.fill = PatternFill("solid", fgColor="EFF6FF")
            cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
            cell.font = Font(italic=True, size=11)
            ws_res.row_dimensions[fila].height = 60
            fila += 3
        
        # Métricas
        if resumen_data.get("metricas"):
            fila += 1
            ws_res.merge_cells(f"A{fila}:D{fila}")
            header_cell(ws_res, fila, 1, "📈 MÉTRICAS CLAVE", bg="1E40AF", size=11)
            fila += 1
            header_cell(ws_res, fila, 1, "Indicador", bg="DBEAFE", fg="1E3A5F", size=10)
            header_cell(ws_res, fila, 2, "Valor",     bg="DBEAFE", fg="1E3A5F", size=10)
            ws_res.merge_cells(f"C{fila}:D{fila}")
            fila += 1
            for idx, (k, v) in enumerate(resumen_data["metricas"].items()):
                bg = gris_claro if idx % 2 == 0 else blanco
                data_cell(ws_res, fila, 1, k, bg=bg, bold=True)
                ws_res.merge_cells(f"B{fila}:C{fila}")
                data_cell(ws_res, fila, 2, str(v), bg=bg, align="center")
                for c in range(1, 4):
                    ws_res.cell(row=fila, column=c).border = thin
                fila += 1
        
        # Puntos clave
        if resumen_data.get("puntos_clave"):
            fila += 1
            ws_res.merge_cells(f"A{fila}:D{fila}")
            header_cell(ws_res, fila, 1, "✅ PUNTOS CLAVE", bg="1E40AF", size=11)
            fila += 1
            for i, punto in enumerate(resumen_data["puntos_clave"], 1):
                ws_res.merge_cells(f"A{fila}:D{fila}")
                cell = ws_res.cell(row=fila, column=1, value=f"{i}. {punto}")
                cell.fill = PatternFill("solid", fgColor=gris_claro if i%2==0 else blanco)
                cell.font = Font(size=11)
                cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
                cell.border = thin
                ws_res.row_dimensions[fila].height = 22
                fila += 1
        
        # Hallazgo
        if resumen_data.get("hallazgo_destacado"):
            fila += 1
            ws_res.merge_cells(f"A{fila}:D{fila}")
            header_cell(ws_res, fila, 1, "💡 HALLAZGO DESTACADO", bg="F59E0B", fg=blanco, size=11)
            fila += 1
            ws_res.merge_cells(f"A{fila}:D{fila+1}")
            cell = ws_res.cell(row=fila, column=1, value=resumen_data["hallazgo_destacado"])
            cell.fill = PatternFill("solid", fgColor="FFFBEB")
            cell.font = Font(italic=True, size=11, color="92400E")
            cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
            ws_res.row_dimensions[fila].height = 45
    
    # Anchos de columna
    ws_res.column_dimensions['A'].width = 30
    ws_res.column_dimensions['B'].width = 25
    ws_res.column_dimensions['C'].width = 20
    ws_res.column_dimensions['D'].width = 20
    
    # ---- HOJA 2: DATOS RAW ----
    ws_data = wb.create_sheet("📄 Datos")
    header_cell(ws_data, 1, 1, "Contenido Extraído del Documento", bg=azul_oscuro, size=12)
    ws_data.column_dimensions['A'].width = 120
    ws_data.merge_cells("A1:B1")
    
    for i, linea in enumerate(texto.split('\n'), start=2):
        if linea.strip():
            cell = ws_data.cell(row=i, column=1, value=linea.strip())
            cell.alignment = Alignment(wrap_text=True, vertical="center")
            cell.fill = PatternFill("solid", fgColor=gris_claro if i%2==0 else blanco)
            ws_data.row_dimensions[i].height = 18
    
    buf = BytesIO()
    wb.save(buf)
    return buf.getvalue()


def exportar_pdf(texto, resumen_data=None):
    """Genera PDF profesional."""
    pdf = FPDF()
    pdf.add_page()
    pdf.set_auto_page_break(auto=True, margin=15)
    
    # Título
    pdf.set_fill_color(30, 58, 95)
    pdf.rect(0, 0, 210, 35, 'F')
    pdf.set_text_color(255, 255, 255)
    pdf.set_font("Helvetica", 'B', 18)
    pdf.set_xy(10, 10)
    pdf.cell(190, 12, "INFORME EJECUTIVO - ORO ASISTENTE", align='C')
    
    zona = pytz.timezone('America/Caracas')
    fecha = datetime.now(zona).strftime('%d/%m/%Y %I:%M %p')
    pdf.set_font("Helvetica", '', 9)
    pdf.set_xy(10, 24)
    pdf.cell(190, 8, f"Generado: {fecha}", align='C')
    pdf.ln(20)
    
    pdf.set_text_color(30, 30, 30)
    
    if resumen_data:
        # Título del documento
        titulo_doc = resumen_data.get("titulo", "")
        if titulo_doc:
            pdf.set_fill_color(37, 99, 235)
            pdf.set_text_color(255, 255, 255)
            pdf.set_font("Helvetica", 'B', 13)
            pdf.cell(0, 10, titulo_doc[:80], fill=True, ln=True, align='C')
            pdf.ln(4)
        
        # Resumen ejecutivo
        res_ej = resumen_data.get("resumen_ejecutivo", "")
        if res_ej:
            pdf.set_fill_color(219, 234, 254)
            pdf.set_text_color(30, 58, 95)
            pdf.set_font("Helvetica", 'I', 10)
            pdf.multi_cell(0, 7, res_ej, fill=True)
            pdf.ln(5)
        
        # Métricas
        if resumen_data.get("metricas"):
            pdf.set_fill_color(30, 58, 95)
            pdf.set_text_color(255, 255, 255)
            pdf.set_font("Helvetica", 'B', 11)
            pdf.cell(0, 8, "  METRICAS CLAVE", fill=True, ln=True)
            pdf.ln(2)
            toggle = False
            for k, v in resumen_data["metricas"].items():
                pdf.set_fill_color(248, 250, 252 if toggle else 255)
                pdf.set_text_color(30, 30, 30)
                pdf.set_font("Helvetica", 'B', 10)
                pdf.cell(80, 8, f"  {k}", fill=True)
                pdf.set_font("Helvetica", '', 10)
                pdf.cell(110, 8, str(v), fill=True, ln=True)
                toggle = not toggle
            pdf.ln(5)
        
        # Puntos clave
        if resumen_data.get("puntos_clave"):
            pdf.set_fill_color(30, 64, 175)
            pdf.set_text_color(255, 255, 255)
            pdf.set_font("Helvetica", 'B', 11)
            pdf.cell(0, 8, "  PUNTOS CLAVE", fill=True, ln=True)
            pdf.ln(2)
            pdf.set_text_color(30, 30, 30)
            for i, punto in enumerate(resumen_data["puntos_clave"], 1):
                pdf.set_font("Helvetica", '', 10)
                pdf.set_fill_color(239, 246, 255 if i%2==0 else 255)
                pdf.multi_cell(0, 7, f"  {i}. {punto}", fill=True)
            pdf.ln(5)
        
        # Hallazgo
        hallazgo = resumen_data.get("hallazgo_destacado", "")
        if hallazgo:
            pdf.set_fill_color(245, 158, 11)
            pdf.set_text_color(255, 255, 255)
            pdf.set_font("Helvetica", 'B', 11)
            pdf.cell(0, 8, "  HALLAZGO DESTACADO", fill=True, ln=True)
            pdf.set_fill_color(255, 251, 235)
            pdf.set_text_color(146, 64, 14)
            pdf.set_font("Helvetica", 'I', 10)
            pdf.multi_cell(0, 7, f"  {hallazgo}", fill=True)
            pdf.ln(5)
        
        pdf.add_page()
    
    # Contenido
    pdf.set_fill_color(30, 58, 95)
    pdf.set_text_color(255, 255, 255)
    pdf.set_font("Helvetica", 'B', 11)
    pdf.cell(0, 8, "  CONTENIDO DEL DOCUMENTO", fill=True, ln=True)
    pdf.ln(3)
    pdf.set_text_color(30, 30, 30)
    pdf.set_font("Helvetica", '', 9)
    
    for linea in texto.split('\n'):
        if linea.strip():
            safe = linea.strip().encode('latin-1', 'replace').decode('latin-1')
            pdf.multi_cell(0, 6, safe)
    
    return pdf.output()


# ==========================================
# REEMPLAZOS PRESERVANDO FORMATO
# ==========================================
def reemplazar_docx_preservando_formato(archivo_bytes, cambios):
    """Reemplaza texto en DOCX iterando runs para preservar formato."""
    doc = Document(BytesIO(archivo_bytes))
    conteo = 0
    
    for c in cambios:
        buscar = str(c["buscar"])
        reemplazar = str(c["reemplazar"])
        if not buscar or buscar.lower() == reemplazar.lower():
            continue
        regex = re.compile(re.escape(buscar), re.IGNORECASE)
        
        def reemplazar_en_parrafo(parrafo):
            nonlocal conteo
            texto_completo = parrafo.text
            if not regex.search(texto_completo):
                return
            nuevo_texto, n = regex.subn(reemplazar, texto_completo)
            conteo += n
            # Preservar formato del primer run, limpiar el resto
            if parrafo.runs:
                parrafo.runs[0].text = nuevo_texto
                for run in parrafo.runs[1:]:
                    run.text = ""
        
        for parrafo in doc.paragraphs:
            reemplazar_en_parrafo(parrafo)
        
        for tabla in doc.tables:
            for fila in tabla.rows:
                for celda in fila.cells:
                    for parrafo in celda.paragraphs:
                        reemplazar_en_parrafo(parrafo)
    
    buf = BytesIO()
    doc.save(buf)
    return buf.getvalue(), conteo


def reemplazar_xlsx_preservando_formato(archivo_bytes, cambios):
    """Reemplaza texto en XLSX preservando estilos."""
    wb = openpyxl.load_workbook(BytesIO(archivo_bytes))
    conteo = 0
    
    for c in cambios:
        buscar = str(c["buscar"])
        reemplazar_val = str(c["reemplazar"])
        if not buscar or buscar.lower() == reemplazar_val.lower():
            continue
        regex = re.compile(re.escape(buscar), re.IGNORECASE)
        
        for sheet in wb.worksheets:
            for row in sheet.iter_rows():
                for cell in row:
                    if cell.value and isinstance(cell.value, str):
                        if regex.search(cell.value):
                            nuevo, n = regex.subn(reemplazar_val, cell.value)
                            cell.value = nuevo
                            conteo += n
    
    buf = BytesIO()
    wb.save(buf)
    return buf.getvalue(), conteo


# ==========================================
# SUBIDA DE ARCHIVO
# ==========================================
st.markdown('<div class="custom-divider"></div>', unsafe_allow_html=True)

col_up, col_info = st.columns([2, 1])
with col_up:
    archivo = st.file_uploader(
        "📂 Sube tu archivo",
        type=["docx", "xlsx", "pdf"],
        help="Soporta Word (.docx), Excel (.xlsx) y PDF"
    )

if archivo and archivo.name != st.session_state.nombre_archivo:
    with st.spinner("📖 Leyendo archivo..."):
        contenido = archivo.read()
        st.session_state.archivo_bytes = contenido
        st.session_state.nombre_archivo = archivo.name
        st.session_state.archivo_tipo = archivo.name.split('.')[-1].lower()
        st.session_state.resumen_data = None
        st.session_state.historial_chat = []
        texto = ""
        try:
            if archivo.name.endswith(".docx"):
                doc = Document(BytesIO(contenido))
                texto = "\n".join([p.text for p in doc.paragraphs if p.text.strip()])
                for t in doc.tables:
                    for row in t.rows:
                        texto += " | ".join([c.text.strip() for c in row.cells]) + "\n"
            elif archivo.name.endswith(".xlsx"):
                wb = openpyxl.load_workbook(BytesIO(contenido), data_only=True)
                for s in wb.worksheets:
                    for r in s.iter_rows(values_only=True):
                        linea = " | ".join([str(c) for c in r if c is not None])
                        if linea.strip():
                            texto += linea + "\n"
            elif archivo.name.endswith(".pdf"):
                reader = PyPDF2.PdfReader(BytesIO(contenido))
                for p in reader.pages:
                    t = p.extract_text()
                    if t:
                        texto += t + "\n"
            st.session_state.texto_extraido = texto
        except Exception as e:
            st.error(f"Error leyendo el archivo: {e}")

# ==========================================
# PANEL PRINCIPAL
# ==========================================
if st.session_state.texto_extraido:
    texto = st.session_state.texto_extraido
    tipo = st.session_state.archivo_tipo
    
    with col_info:
        palabras = len(texto.split())
        lineas = len([l for l in texto.split('\n') if l.strip()])
        st.markdown(f"""
        <div class="metric-card">
            <div class="metric-label">📄 Archivo cargado</div>
            <div class="metric-value" style="font-size:1rem">{st.session_state.nombre_archivo}</div>
        </div>
        <div class="metric-card">
            <div class="metric-label">📊 Estadísticas</div>
            <div class="metric-value" style="font-size:0.95rem">{palabras:,} palabras · {lineas:,} líneas</div>
        </div>
        """, unsafe_allow_html=True)
    
    st.markdown('<div class="custom-divider"></div>', unsafe_allow_html=True)
    
    # ---- TABS ----
    tab1, tab2, tab3, tab4 = st.tabs(["📊 Resumen", "✍️ Edición", "💬 Preguntar", "🔍 Calidad"])
    
    # ========== TAB 1: RESUMEN ==========
    with tab1:
        if st.button("⚡ Generar Resumen Inteligente", use_container_width=True):
            with st.spinner("🧠 Analizando documento..."):
                data = solicitar_resumen_estructurado(texto)
                st.session_state.resumen_data = data
        
        data = st.session_state.resumen_data
        if data:
            emoji = data.get("emoji_categoria", "📋")
            titulo_doc = data.get("titulo", "Documento analizado")
            
            # Encabezado del resumen
            st.markdown(f"""
            <div class="summary-box">
                <div class="summary-title">{emoji} {titulo_doc}</div>
                {data.get("resumen_ejecutivo", "")}
            </div>
            """, unsafe_allow_html=True)
            
            # Métricas en columnas
            metricas = data.get("metricas", {})
            if metricas:
                items = list(metricas.items())
                cols = st.columns(min(len(items), 4))
                for i, (k, v) in enumerate(items):
                    with cols[i % len(cols)]:
                        st.markdown(f"""
                        <div class="metric-card">
                            <div class="metric-label">{k}</div>
                            <div class="metric-value">{v}</div>
                        </div>
                        """, unsafe_allow_html=True)
            
            # Puntos clave como tags
            puntos = data.get("puntos_clave", [])
            if puntos:
                st.markdown("**📌 Puntos clave:**")
                tags_html = "".join([f'<span class="tag">✓ {p}</span>' for p in puntos])
                st.markdown(tags_html, unsafe_allow_html=True)
            
            # Hallazgo destacado
            hallazgo = data.get("hallazgo_destacado", "")
            if hallazgo:
                st.markdown(f"""
                <div style="background:#1a2a1a;border:1px solid #166534;border-left:4px solid #22c55e;
                border-radius:10px;padding:1rem;margin-top:1rem;color:#86efac;">
                    💡 <strong>Hallazgo:</strong> {hallazgo}
                </div>
                """, unsafe_allow_html=True)
            
            # Exportar
            st.markdown('<div class="custom-divider"></div>', unsafe_allow_html=True)
            st.markdown("### 📥 Exportar Informe")
            c1, c2, c3 = st.columns(3)
            
            with c1:
                word_bytes = exportar_word(texto, data)
                st.download_button("📄 Descargar Word", word_bytes, "Informe_Oro.docx",
                                   mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                                   use_container_width=True)
            with c2:
                excel_bytes = exportar_excel(texto, data)
                st.download_button("📊 Descargar Excel", excel_bytes, "Informe_Oro.xlsx",
                                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                   use_container_width=True)
            with c3:
                pdf_bytes = exportar_pdf(texto, data)
                st.download_button("📕 Descargar PDF", bytes(pdf_bytes), "Informe_Oro.pdf",
                                   mime="application/pdf",
                                   use_container_width=True)
        else:
            st.info("👆 Haz clic en **Generar Resumen Inteligente** para analizar el documento.")
    
    # ========== TAB 2: EDICIÓN ==========
    with tab2:
        st.markdown("#### ✍️ Corrección Quirúrgica")
        st.caption("Escribe una instrucción natural. Ej: *Cambia 'atletismo' por 'BEISBOL'* o *Reemplaza 'municipio' por 'ciudad'*")
        
        instruccion = st.text_input("✏️ Instrucción de cambio:", placeholder="Cambia 'X' por 'Y'")
        
        if instruccion:
            with st.spinner("🔍 Procesando cambio..."):
                cambios = solicitar_cambios(instruccion)
            
            if cambios:
                st.markdown(f"**Cambios detectados:** `{cambios[0]['buscar']}` → `{cambios[0]['reemplazar']}`")
                
                archivo_bytes = st.session_state.archivo_bytes
                
                if tipo == "docx":
                    final_bytes, n = reemplazar_docx_preservando_formato(archivo_bytes, cambios)
                elif tipo == "xlsx":
                    final_bytes, n = reemplazar_xlsx_preservando_formato(archivo_bytes, cambios)
                else:
                    txt_mod = texto
                    n = 0
                    for c in cambios:
                        txt_mod, count = re.compile(re.escape(c['buscar']), re.IGNORECASE).subn(c['reemplazar'], txt_mod)
                        n += count
                    final_bytes = txt_mod.encode()
                
                if n > 0:
                    st.markdown(f'<div class="info-box">✅ {n} cambio(s) realizados correctamente.</div>', unsafe_allow_html=True)
                    st.session_state.cambios_aplicados = final_bytes
                    
                    c1, c2, c3 = st.columns(3)
                    with c1:
                        if tipo == "docx":
                            st.download_button("📄 Word corregido", final_bytes, "Corregido.docx", use_container_width=True)
                        else:
                            # Generar Word desde texto corregido
                            txt_corr = texto
                            for c in cambios:
                                txt_corr = re.compile(re.escape(c['buscar']), re.IGNORECASE).sub(c['reemplazar'], txt_corr)
                            st.download_button("📄 Word", exportar_word(txt_corr), "Corregido.docx", use_container_width=True)
                    with c2:
                        if tipo == "xlsx":
                            st.download_button("📊 Excel corregido", final_bytes, "Corregido.xlsx", use_container_width=True)
                        else:
                            txt_corr = texto
                            for c in cambios:
                                txt_corr = re.compile(re.escape(c['buscar']), re.IGNORECASE).sub(c['reemplazar'], txt_corr)
                            st.download_button("📊 Excel", exportar_excel(txt_corr), "Corregido.xlsx", use_container_width=True)
                    with c3:
                        txt_corr = texto
                        for c in cambios:
                            txt_corr = re.compile(re.escape(c['buscar']), re.IGNORECASE).sub(c['reemplazar'], txt_corr)
                        pdf_c = exportar_pdf(txt_corr)
                        st.download_button("📕 PDF", bytes(pdf_c), "Corregido.pdf", use_container_width=True)
                else:
                    st.markdown(f'<div class="warn-box">⚠️ No encontré "{cambios[0]["buscar"]}" en el documento.</div>', unsafe_allow_html=True)
            else:
                st.markdown('<div class="warn-box">⚠️ No entendí la instrucción. Prueba: Cambia \'X\' por \'Y\'</div>', unsafe_allow_html=True)
    
    # ========== TAB 3: CHAT ==========
    with tab3:
        st.markdown("#### 💬 Pregunta sobre el documento")
        st.caption("Haz cualquier pregunta sobre el contenido del archivo cargado.")
        
        # Mostrar historial
        for msg in st.session_state.historial_chat:
            with st.chat_message("user" if msg["rol"] == "Usuario" else "assistant"):
                st.write(msg["texto"])
        
        pregunta = st.chat_input("Escribe tu pregunta aquí...")
        if pregunta:
            st.session_state.historial_chat.append({"rol": "Usuario", "texto": pregunta})
            with st.spinner("🤔 Pensando..."):
                respuesta = preguntar_al_documento(pregunta, texto)
            st.session_state.historial_chat.append({"rol": "Asistente", "texto": respuesta})
            st.rerun()
    
    # ========== TAB 4: CALIDAD ==========
    with tab4:
        st.markdown("#### 🔍 Análisis de Calidad")
        st.caption("Detecta inconsistencias, datos duplicados y posibles errores en el documento.")
        
        if st.button("🔎 Analizar Calidad del Documento", use_container_width=True):
            with st.spinner("Revisando inconsistencias..."):
                resultado = detectar_anomalias(texto)
            
            if resultado:
                nivel = resultado.get("nivel_calidad", "?")
                color = {"Alto": "#22c55e", "Medio": "#f59e0b", "Bajo": "#ef4444"}.get(nivel, "#6b7280")
                
                st.markdown(f"""
                <div style="text-align:center;margin:1rem 0">
                    <span style="background:{color}22;color:{color};border:1px solid {color};
                    border-radius:20px;padding:0.5rem 1.5rem;font-weight:700;font-size:1.1rem;">
                        Calidad del documento: {nivel}
                    </span>
                </div>
                """, unsafe_allow_html=True)
                
                anomalias = resultado.get("anomalias", [])
                if anomalias:
                    st.markdown("**⚠️ Posibles anomalías:**")
                    for a in anomalias:
                        st.markdown(f"- {a}")
                else:
                    st.success("✅ No se detectaron anomalías significativas.")
                
                rec = resultado.get("recomendacion", "")
                if rec:
                    st.info(f"💡 **Recomendación:** {rec}")
        else:
            st.info("👆 Haz clic para analizar la calidad del documento.")

else:
    # Estado vacío
    st.markdown("""
    <div style="text-align:center;padding:3rem;color:#4b5563;">
        <div style="font-size:4rem;margin-bottom:1rem">📂</div>
        <div style="font-size:1.1rem;font-weight:600;color:#6b7280">Sube un archivo para comenzar</div>
        <div style="font-size:0.85rem;margin-top:0.5rem;color:#374151">Formatos soportados: .docx · .xlsx · .pdf</div>
    </div>
    """, unsafe_allow_html=True)

# ==========================================
# FOOTER
# ==========================================
zona_horaria = pytz.timezone('America/Caracas')
hora = datetime.now(zona_horaria).strftime('%I:%M %p')
st.markdown(f"<p class='footer'>🏆 Oro Asistente · {hora} VET · Powered by Gemini</p>", unsafe_allow_html=True)
