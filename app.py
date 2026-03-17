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
    layout="centered",
    initial_sidebar_state="collapsed"
)

# ==========================================
# ESTILOS CSS
# ==========================================
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Outfit:wght@300;400;500;600;700;800&family=JetBrains+Mono:wght@400;600&display=swap');

/* ── RESET & BASE ── */
html, body, [class*="css"] {
    font-family: 'Outfit', sans-serif !important;
    -webkit-tap-highlight-color: transparent;
}

/* ── FONDO PRINCIPAL ── */
.stApp {
    background: linear-gradient(160deg, #0a0e1a 0%, #0d1525 50%, #0a1020 100%);
    min-height: 100vh;
}
.main .block-container {
    padding: 1rem 1rem 4rem 1rem !important;
    max-width: 480px !important;
    margin: 0 auto !important;
}

/* ── OCULTAR elementos de escritorio innecesarios ── */
#MainMenu, footer, header { visibility: hidden; }
[data-testid="stToolbar"] { display: none; }

/* ── HEADER ── */
.oro-header {
    text-align: center;
    padding: 1.5rem 0 0.5rem 0;
    position: relative;
}
.oro-logo {
    font-size: 3rem;
    line-height: 1;
    filter: drop-shadow(0 0 20px rgba(251,191,36,0.5));
    animation: pulse-glow 3s ease-in-out infinite;
}
@keyframes pulse-glow {
    0%,100% { filter: drop-shadow(0 0 15px rgba(251,191,36,0.4)); }
    50%      { filter: drop-shadow(0 0 30px rgba(251,191,36,0.8)); }
}
.oro-title {
    font-size: 1.9rem;
    font-weight: 800;
    background: linear-gradient(135deg, #fbbf24 0%, #f59e0b 40%, #fde68a 70%, #f59e0b 100%);
    -webkit-background-clip: text;
    -webkit-text-fill-color: transparent;
    background-clip: text;
    letter-spacing: -0.02em;
    margin: 0.2rem 0 0.1rem 0;
}
.oro-subtitle {
    color: #4b6080;
    font-size: 0.82rem;
    font-weight: 400;
    letter-spacing: 0.03em;
}

/* ── UPLOAD ZONE ── */
.upload-zone {
    background: linear-gradient(135deg, #111827, #162032);
    border: 2px dashed #2a4a6b;
    border-radius: 20px;
    padding: 2rem 1rem;
    text-align: center;
    margin: 1rem 0;
    transition: border-color 0.3s;
}
.upload-icon { font-size: 2.5rem; margin-bottom: 0.5rem; }
.upload-text { color: #60a5fa; font-weight: 600; font-size: 1rem; }
.upload-hint { color: #374151; font-size: 0.78rem; margin-top: 0.3rem; }

[data-testid="stFileUploader"] {
    background: transparent !important;
    border: none !important;
}
[data-testid="stFileUploader"] > div {
    background: linear-gradient(135deg, #0f1927, #162032) !important;
    border: 2px dashed #1e3a5f !important;
    border-radius: 20px !important;
    padding: 1.5rem !important;
}
[data-testid="stFileUploader"] label {
    color: #60a5fa !important;
    font-weight: 600 !important;
    font-size: 1rem !important;
}

/* ── FILE BADGE ── */
.file-badge {
    display: flex;
    align-items: center;
    gap: 0.8rem;
    background: linear-gradient(135deg, #0f2037, #132840);
    border: 1px solid #1e4a7a;
    border-radius: 16px;
    padding: 1rem 1.2rem;
    margin: 0.8rem 0;
}
.file-icon { font-size: 2rem; }
.file-info-name { color: #93c5fd; font-weight: 600; font-size: 0.9rem; word-break: break-all; }
.file-info-stats { color: #4b6080; font-size: 0.75rem; margin-top: 0.2rem; }

/* ── NAVEGACIÓN TIPO APP MÓVIL ── */
.nav-bar {
    display: flex;
    gap: 0.4rem;
    background: #0d1525;
    border: 1px solid #1e3a5f;
    border-radius: 16px;
    padding: 0.4rem;
    margin: 1rem 0;
}
.nav-btn {
    flex: 1;
    text-align: center;
    padding: 0.6rem 0.2rem;
    border-radius: 12px;
    cursor: pointer;
    transition: all 0.2s;
    font-size: 0.65rem;
    color: #4b6080;
    font-weight: 500;
}
.nav-btn .nav-icon { font-size: 1.3rem; display: block; margin-bottom: 0.2rem; }
.nav-btn.active {
    background: linear-gradient(135deg, #1e40af, #2563eb);
    color: white;
    box-shadow: 0 4px 15px rgba(37,99,235,0.35);
}

/* ── SECCIÓN CONTENT ── */
.section-title {
    color: #e2e8f0;
    font-size: 1.05rem;
    font-weight: 700;
    margin: 1.2rem 0 0.5rem 0;
    display: flex;
    align-items: center;
    gap: 0.5rem;
}
.section-hint {
    color: #374151;
    font-size: 0.78rem;
    margin-bottom: 1rem;
    line-height: 1.5;
}

/* ── CARD RESUMEN ── */
.summary-card {
    background: linear-gradient(135deg, #0f1e33, #132840);
    border: 1px solid #1e4a7a;
    border-left: 4px solid #3b82f6;
    border-radius: 18px;
    padding: 1.2rem;
    margin: 0.8rem 0;
    color: #bfdbfe;
    line-height: 1.7;
    font-size: 0.9rem;
}
.summary-card-title {
    color: #60a5fa;
    font-size: 1rem;
    font-weight: 700;
    margin-bottom: 0.6rem;
}

/* ── METRIC PILL ── */
.metrics-grid {
    display: grid;
    grid-template-columns: 1fr 1fr;
    gap: 0.6rem;
    margin: 0.8rem 0;
}
.metric-pill {
    background: linear-gradient(135deg, #111827, #162032);
    border: 1px solid #1e3a5f;
    border-radius: 14px;
    padding: 0.8rem 1rem;
    text-align: center;
}
.metric-pill-label {
    color: #4b6080;
    font-size: 0.68rem;
    font-weight: 600;
    text-transform: uppercase;
    letter-spacing: 0.05em;
}
.metric-pill-value {
    color: #e2e8f0;
    font-size: 1.2rem;
    font-weight: 800;
    margin-top: 0.2rem;
    font-family: 'JetBrains Mono', monospace;
}

/* ── TAGS ── */
.tags-wrap { display: flex; flex-wrap: wrap; gap: 0.4rem; margin: 0.6rem 0; }
.tag {
    background: #0f2037;
    color: #60a5fa;
    border: 1px solid #1e4a7a;
    border-radius: 20px;
    padding: 0.35rem 0.8rem;
    font-size: 0.75rem;
    font-weight: 500;
}

/* ── HALLAZGO ── */
.hallazgo-card {
    background: #0a1f0e;
    border: 1px solid #14532d;
    border-left: 4px solid #22c55e;
    border-radius: 14px;
    padding: 1rem 1.1rem;
    color: #86efac;
    font-size: 0.85rem;
    margin: 0.8rem 0;
    line-height: 1.6;
}

/* ── BOTONES GRANDES TÁCTILES ── */
.stButton > button {
    background: linear-gradient(135deg, #1d4ed8, #2563eb) !important;
    color: white !important;
    border: none !important;
    border-radius: 14px !important;
    font-weight: 700 !important;
    font-size: 0.95rem !important;
    height: 3.2rem !important;
    width: 100% !important;
    transition: all 0.15s !important;
    letter-spacing: 0.01em !important;
    font-family: 'Outfit', sans-serif !important;
}
.stButton > button:active {
    transform: scale(0.97) !important;
    box-shadow: 0 2px 10px rgba(37,99,235,0.3) !important;
}

/* ── BOTONES DESCARGA ── */
[data-testid="stDownloadButton"] > button {
    background: linear-gradient(135deg, #065f46, #059669) !important;
    color: white !important;
    border: none !important;
    border-radius: 14px !important;
    font-weight: 700 !important;
    font-size: 0.88rem !important;
    height: 3rem !important;
    width: 100% !important;
    font-family: 'Outfit', sans-serif !important;
}
[data-testid="stDownloadButton"] > button:active {
    transform: scale(0.97) !important;
}

/* ── DESCARGA GRID ── */
.download-grid { display: flex; flex-direction: column; gap: 0.6rem; margin-top: 0.8rem; }

/* ── INPUT TEXTO ── */
.stTextInput > div > div > input {
    background: #0f1927 !important;
    border: 2px solid #1e3a5f !important;
    border-radius: 14px !important;
    color: #e2e8f0 !important;
    font-size: 1rem !important;
    padding: 0.8rem 1rem !important;
    font-family: 'Outfit', sans-serif !important;
    height: 3.2rem !important;
}
.stTextInput > div > div > input:focus {
    border-color: #3b82f6 !important;
    box-shadow: 0 0 0 3px rgba(59,130,246,0.15) !important;
}

/* ── CHAT ── */
[data-testid="stChatInput"] textarea {
    background: #0f1927 !important;
    border: 2px solid #1e3a5f !important;
    border-radius: 14px !important;
    color: #e2e8f0 !important;
    font-family: 'Outfit', sans-serif !important;
    font-size: 0.95rem !important;
}
[data-testid="stChatMessageContent"] {
    font-size: 0.9rem !important;
}

/* ── INFO / WARN BOXES ── */
.info-box {
    background: #052e16;
    border: 1px solid #15803d;
    border-radius: 12px;
    padding: 0.9rem 1rem;
    color: #4ade80;
    font-size: 0.88rem;
    margin: 0.6rem 0;
    display: flex;
    align-items: center;
    gap: 0.5rem;
}
.warn-box {
    background: #1c1003;
    border: 1px solid #b45309;
    border-radius: 12px;
    padding: 0.9rem 1rem;
    color: #fbbf24;
    font-size: 0.88rem;
    margin: 0.6rem 0;
    display: flex;
    align-items: center;
    gap: 0.5rem;
}

/* ── CAMBIOS LISTA ── */
.cambio-item {
    background: #0d1525;
    border: 1px solid #1e3a5f;
    border-radius: 10px;
    padding: 0.6rem 0.9rem;
    margin: 0.3rem 0;
    font-size: 0.82rem;
    color: #93c5fd;
    font-family: 'JetBrains Mono', monospace;
    display: flex;
    align-items: center;
    gap: 0.5rem;
}
.cambio-num { color: #4b6080; font-size: 0.7rem; min-width: 1.2rem; }
.cambio-arrow { color: #f59e0b; }

/* ── CALIDAD BADGE ── */
.calidad-badge {
    text-align: center;
    padding: 1.2rem;
    margin: 0.8rem 0;
}
.calidad-pill {
    display: inline-block;
    padding: 0.6rem 2rem;
    border-radius: 30px;
    font-weight: 800;
    font-size: 1rem;
    letter-spacing: 0.03em;
}
.anomalia-item {
    background: #1c0f03;
    border: 1px solid #92400e;
    border-radius: 10px;
    padding: 0.7rem 1rem;
    color: #fcd34d;
    font-size: 0.84rem;
    margin: 0.4rem 0;
    display: flex;
    gap: 0.5rem;
    line-height: 1.5;
}

/* ── TABS OCULTAR estilo nativo, usar nuestro nav ── */
[data-testid="stTabs"] [role="tablist"] { display: none !important; }
[data-testid="stTabPanel"] { padding-top: 0 !important; }

/* ── EXPANDER ── */
[data-testid="stExpander"] {
    background: #0d1525 !important;
    border: 1px solid #1e3a5f !important;
    border-radius: 14px !important;
}

/* ── SPINNER ── */
[data-testid="stSpinner"] { color: #60a5fa !important; }

/* ── DIVIDER ── */
.oro-divider {
    height: 1px;
    background: linear-gradient(90deg, transparent, #1e3a5f, transparent);
    margin: 1.2rem 0;
}

/* ── ESTADO VACÍO ── */
.empty-state {
    text-align: center;
    padding: 3rem 1rem;
}
.empty-icon { font-size: 4rem; margin-bottom: 1rem; opacity: 0.5; }
.empty-title { color: #374151; font-size: 1rem; font-weight: 600; }
.empty-hint { color: #1f2937; font-size: 0.8rem; margin-top: 0.4rem; }
.format-badges { display: flex; justify-content: center; gap: 0.5rem; margin-top: 1rem; }
.format-badge {
    background: #111827;
    border: 1px solid #1f2937;
    border-radius: 8px;
    padding: 0.3rem 0.7rem;
    color: #374151;
    font-size: 0.75rem;
    font-family: 'JetBrains Mono', monospace;
}

/* ── FOOTER ── */
.oro-footer {
    text-align: center;
    font-size: 0.72rem;
    color: #1f2937;
    padding: 1.5rem 0 0.5rem 0;
    border-top: 1px solid #111827;
    margin-top: 2rem;
}

/* ── SELECTBOX SIDEBAR ── */
[data-testid="stSidebar"] {
    background: #080e1a !important;
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
    st.caption("Oro Asistente v2 · Mobile First")
    st.caption("Sube un archivo · Analiza · Edita · Exporta")

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
    "lista_cambios": [],
    "texto_modificado": "",
}.items():
    if key not in st.session_state:
        st.session_state[key] = val

# ==========================================
# HEADER
# ==========================================
st.markdown("""
<div class="oro-header">
    <div class="oro-logo">🏆</div>
    <div class="oro-title">Oro Asistente</div>
    <div class="oro-subtitle">ANALIZA · EDITA · EXPORTA</div>
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
def exportar_word(texto, resumen_data=None, archivo_bytes=None, archivo_tipo=None, cambios=None):
    """
    Genera Word profesional con conversion cruzada.
    - DOCX original: aplica cambios preservando formato.
    - XLSX original: convierte cada hoja en tabla Word real.
    - PDF/texto: genera informe con resumen estructurado.
    """
    zona = pytz.timezone('America/Caracas')
    fecha = datetime.now(zona).strftime('%d de %B de %Y, %I:%M %p')
    cambios = cambios or []

    if archivo_tipo == "docx" and archivo_bytes and cambios:
        resultado, _ = reemplazar_docx_preservando_formato(archivo_bytes, cambios)
        return resultado

    if archivo_tipo == "xlsx" and archivo_bytes:
        doc = Document()
        doc.styles['Normal'].font.name = 'Calibri'
        titulo_h = doc.add_heading('', 0)
        r = titulo_h.add_run('Reporte Exportado desde Excel')
        r.font.color.rgb = RGBColor(0x1E, 0x40, 0xAF)
        r.font.size = Pt(20)
        doc.add_paragraph().add_run(f'Generado: {fecha}').font.size = Pt(9)
        doc.add_paragraph()
        bytes_usar = archivo_bytes
        if cambios:
            bytes_usar, _ = reemplazar_xlsx_preservando_formato(archivo_bytes, cambios)
        wb = openpyxl.load_workbook(BytesIO(bytes_usar), data_only=True)
        for sheet in wb.worksheets:
            doc.add_heading(f'Hoja: {sheet.title}', level=1)
            filas = [f for f in sheet.iter_rows(values_only=True) if any(c is not None for c in f)]
            if not filas:
                doc.add_paragraph('(Hoja vacia)')
                continue
            n_cols = max(len(f) for f in filas)
            tabla = doc.add_table(rows=len(filas), cols=n_cols)
            tabla.style = 'Table Grid'
            for i, fila in enumerate(filas):
                for j in range(n_cols):
                    val = fila[j] if j < len(fila) else ""
                    cell = tabla.cell(i, j)
                    cell.text = str(val) if val is not None else ""
                    if i == 0:
                        for run in cell.paragraphs[0].runs:
                            run.font.bold = True
            doc.add_paragraph()
        buf = BytesIO()
        doc.save(buf)
        return buf.getvalue()

    doc = Document()
    doc.styles['Normal'].font.name = 'Calibri'
    doc.styles['Normal'].font.size = Pt(11)
    titulo_h = doc.add_heading('', 0)
    run_t = titulo_h.add_run('INFORME EJECUTIVO')
    run_t.font.color.rgb = RGBColor(0x1E, 0x40, 0xAF)
    run_t.font.size = Pt(22)
    doc.add_paragraph().add_run(f'Generado: {fecha}').font.size = Pt(9)
    doc.add_paragraph()

    if resumen_data:
        h2 = doc.add_heading('Resumen Ejecutivo', level=1)
        h2.runs[0].font.color.rgb = RGBColor(0x1E, 0x40, 0xAF)
        doc.add_paragraph(resumen_data.get("resumen_ejecutivo", ""))
        if resumen_data.get("metricas"):
            doc.add_heading('Metricas Clave', level=2)
            tabla = doc.add_table(rows=1, cols=2)
            tabla.style = 'Table Grid'
            hdr = tabla.rows[0].cells
            hdr[0].text, hdr[1].text = 'Indicador', 'Valor'
            for cell in hdr:
                for run in cell.paragraphs[0].runs:
                    run.font.bold = True
            for k, v in resumen_data["metricas"].items():
                row = tabla.add_row().cells
                row[0].text, row[1].text = str(k), str(v)
            doc.add_paragraph()
        if resumen_data.get("puntos_clave"):
            doc.add_heading('Puntos Clave', level=2)
            for punto in resumen_data["puntos_clave"]:
                doc.add_paragraph(style='List Bullet').add_run(punto)
        if resumen_data.get("hallazgo_destacado"):
            doc.add_paragraph()
            doc.add_heading('Hallazgo Destacado', level=2)
            run_h = doc.add_paragraph().add_run(resumen_data["hallazgo_destacado"])
            run_h.font.italic = True
            run_h.font.color.rgb = RGBColor(0x1D, 0x4E, 0xD8)
        doc.add_page_break()

    doc.add_heading('Contenido del Documento', level=1)
    for linea in texto.split('\n'):
        linea_limpia = linea.strip().replace('*', '').replace('#', '')
        if linea_limpia:
            doc.add_paragraph(linea_limpia)
    buf = BytesIO()
    doc.save(buf)
    return buf.getvalue()

def exportar_excel(texto, resumen_data=None, archivo_bytes=None, archivo_tipo=None, cambios=None):
    """
    Genera Excel profesional con conversion cruzada.
    - XLSX original: aplica cambios y devuelve el archivo original formateado.
    - DOCX original: extrae cada tabla como hoja separada.
    - PDF/texto: vuelca lineas en una hoja con resumen.
    """
    cambios = cambios or []

    if archivo_tipo == "xlsx" and archivo_bytes and cambios:
        resultado, _ = reemplazar_xlsx_preservando_formato(archivo_bytes, cambios)
        return resultado

    if archivo_tipo == "docx" and archivo_bytes:
        bytes_usar = archivo_bytes
        if cambios:
            bytes_usar, _ = reemplazar_docx_preservando_formato(archivo_bytes, cambios)
        wb = openpyxl.Workbook()
        wb.remove(wb.active)
        doc_src = Document(BytesIO(bytes_usar))
        azul = "1E3A5F"
        blanco = "FFFFFF"
        gris = "F8FAFC"
        for i, tabla in enumerate(doc_src.tables):
            ws = wb.create_sheet(title=f"Tabla_{i+1}")
            filas_limpias = []
            for row in tabla.rows:
                vistas = set()
                fila = []
                for cell in row.cells:
                    if cell._tc not in vistas:
                        vistas.add(cell._tc)
                        fila.append(cell.text.strip())
                if any(fila):
                    filas_limpias.append(fila)
            for r_idx, fila in enumerate(filas_limpias, 1):
                for c_idx, val in enumerate(fila, 1):
                    cell = ws.cell(row=r_idx, column=c_idx, value=val)
                    if r_idx == 1:
                        cell.fill = PatternFill("solid", fgColor=azul)
                        cell.font = Font(color=blanco, bold=True, size=10)
                    else:
                        cell.fill = PatternFill("solid", fgColor=gris if r_idx%2==0 else blanco)
                        cell.font = Font(size=10)
                    cell.alignment = Alignment(wrap_text=True, vertical="center")
            for col in ws.columns:
                ws.column_dimensions[col[0].column_letter].width = 22
        if not wb.sheetnames:
            ws = wb.create_sheet("Datos")
            ws.cell(1, 1, "No se encontraron tablas estructuradas en el documento.")
        buf = BytesIO()
        wb.save(buf)
        return buf.getvalue()

    wb = openpyxl.Workbook()
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
        left=Side(style='thin', color='CBD5E1'), right=Side(style='thin', color='CBD5E1'),
        top=Side(style='thin', color='CBD5E1'), bottom=Side(style='thin', color='CBD5E1')
    )

    ws_res = wb.active
    ws_res.title = "Resumen"
    zona = pytz.timezone('America/Caracas')
    fecha = datetime.now(zona).strftime('%d/%m/%Y %I:%M %p')
    ws_res.merge_cells("A1:D1")
    header_cell(ws_res, 1, 1, "ORO ASISTENTE - REPORTE EJECUTIVO", bg=azul_oscuro, size=14)
    ws_res.row_dimensions[1].height = 40
    ws_res.merge_cells("A2:D2")
    data_cell(ws_res, 2, 1, f"Generado: {fecha}", bg=azul_claro, align="center")

    fila = 4
    if resumen_data:
        titulo_doc = resumen_data.get("titulo", "Sin titulo")
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
        if resumen_data.get("metricas"):
            fila += 1
            ws_res.merge_cells(f"A{fila}:D{fila}")
            header_cell(ws_res, fila, 1, "METRICAS CLAVE", bg="1E40AF", size=11)
            fila += 1
            header_cell(ws_res, fila, 1, "Indicador", bg="DBEAFE", fg="1E3A5F", size=10)
            header_cell(ws_res, fila, 2, "Valor", bg="DBEAFE", fg="1E3A5F", size=10)
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
        if resumen_data.get("puntos_clave"):
            fila += 1
            ws_res.merge_cells(f"A{fila}:D{fila}")
            header_cell(ws_res, fila, 1, "PUNTOS CLAVE", bg="1E40AF", size=11)
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
        if resumen_data.get("hallazgo_destacado"):
            fila += 1
            ws_res.merge_cells(f"A{fila}:D{fila}")
            header_cell(ws_res, fila, 1, "HALLAZGO DESTACADO", bg="F59E0B", fg=blanco, size=11)
            fila += 1
            ws_res.merge_cells(f"A{fila}:D{fila+1}")
            cell = ws_res.cell(row=fila, column=1, value=resumen_data["hallazgo_destacado"])
            cell.fill = PatternFill("solid", fgColor="FFFBEB")
            cell.font = Font(italic=True, size=11, color="92400E")
            cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
            ws_res.row_dimensions[fila].height = 45

    for col in ['A','B','C','D']:
        ws_res.column_dimensions[col].width = 28

    ws_data = wb.create_sheet("Datos")
    header_cell(ws_data, 1, 1, "Contenido Extraido del Documento", bg=azul_oscuro, size=12)
    ws_data.merge_cells("A1:B1")
    ws_data.column_dimensions['A'].width = 120
    for i, linea in enumerate(texto.split('\n'), start=2):
        if linea.strip():
            cell = ws_data.cell(row=i, column=1, value=linea.strip())
            cell.alignment = Alignment(wrap_text=True, vertical="center")
            cell.fill = PatternFill("solid", fgColor=gris_claro if i%2==0 else blanco)
            ws_data.row_dimensions[i].height = 18

    buf = BytesIO()
    wb.save(buf)
    return buf.getvalue()

def safe_text(t):
    """Convierte texto a latin-1 seguro para fpdf versión antigua."""
    return str(t).encode('latin-1', 'replace').decode('latin-1')

def pdf_seccion_header(pdf, titulo, r, g, b):
    """Dibuja un encabezado de sección con fondo de color."""
    pdf.set_fill_color(r, g, b)
    pdf.set_text_color(255, 255, 255)
    pdf.set_font("Arial", 'B', 11)
    pdf.cell(190, 8, safe_text(titulo), border=0, ln=1, fill=True)
    pdf.ln(2)
    pdf.set_text_color(30, 30, 30)

def exportar_pdf(texto, resumen_data=None):
    """Genera PDF compatible con fpdf (versión clásica)."""
    pdf = FPDF()
    pdf.set_margins(10, 10, 10)
    pdf.add_page()
    pdf.set_auto_page_break(auto=True, margin=15)

    # ---- Encabezado ----
    pdf.set_fill_color(30, 58, 95)
    pdf.rect(0, 0, 210, 32, 'F')
    pdf.set_text_color(255, 255, 255)
    pdf.set_font("Arial", 'B', 16)
    pdf.set_xy(10, 8)
    pdf.cell(190, 10, "INFORME EJECUTIVO - ORO ASISTENTE", ln=1, align='C')
    zona = pytz.timezone('America/Caracas')
    fecha = datetime.now(zona).strftime('%d/%m/%Y %I:%M %p')
    pdf.set_font("Arial", '', 9)
    pdf.set_xy(10, 20)
    pdf.cell(190, 8, safe_text(f"Generado: {fecha}"), ln=1, align='C')
    pdf.set_xy(10, 35)
    pdf.set_text_color(30, 30, 30)

    if resumen_data:
        # Título del documento
        titulo_doc = resumen_data.get("titulo", "")
        if titulo_doc:
            pdf.set_fill_color(37, 99, 235)
            pdf.set_text_color(255, 255, 255)
            pdf.set_font("Arial", 'B', 12)
            pdf.cell(190, 10, safe_text(titulo_doc[:90]), border=0, ln=1, align='C', fill=True)
            pdf.ln(3)
            pdf.set_text_color(30, 30, 30)

        # Resumen ejecutivo
        res_ej = resumen_data.get("resumen_ejecutivo", "")
        if res_ej:
            pdf.set_font("Arial", 'I', 10)
            pdf.multi_cell(190, 6, safe_text(res_ej))
            pdf.ln(4)

        # Métricas
        metricas = resumen_data.get("metricas", {})
        if metricas:
            pdf_seccion_header(pdf, "  METRICAS CLAVE", 30, 58, 95)
            toggle = False
            for k, v in metricas.items():
                r_bg, g_bg, b_bg = (245, 247, 250) if toggle else (255, 255, 255)
                pdf.set_fill_color(r_bg, g_bg, b_bg)
                pdf.set_font("Arial", 'B', 10)
                pdf.cell(85, 8, safe_text(f"  {k}"), border=0, fill=True)
                pdf.set_font("Arial", '', 10)
                pdf.cell(105, 8, safe_text(str(v)), border=0, ln=1, fill=True)
                toggle = not toggle
            pdf.ln(4)

        # Puntos clave
        puntos = resumen_data.get("puntos_clave", [])
        if puntos:
            pdf_seccion_header(pdf, "  PUNTOS CLAVE", 30, 64, 175)
            pdf.set_font("Arial", '', 10)
            for i, punto in enumerate(puntos, 1):
                pdf.multi_cell(190, 7, safe_text(f"  {i}. {punto}"))
            pdf.ln(4)

        # Hallazgo
        hallazgo = resumen_data.get("hallazgo_destacado", "")
        if hallazgo:
            pdf_seccion_header(pdf, "  HALLAZGO DESTACADO", 180, 120, 10)
            pdf.set_font("Arial", 'I', 10)
            pdf.multi_cell(190, 7, safe_text(f"  {hallazgo}"))
            pdf.ln(4)

        pdf.add_page()

    # ---- Contenido del documento ----
    pdf_seccion_header(pdf, "  CONTENIDO DEL DOCUMENTO", 30, 58, 95)
    pdf.set_font("Arial", '', 9)
    for linea in texto.split('\n'):
        linea = linea.strip()
        if linea:
            pdf.multi_cell(190, 5, safe_text(linea))

    return pdf.output(dest='S').encode('latin-1')


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
st.markdown('<div class="oro-divider"></div>', unsafe_allow_html=True)

archivo = st.file_uploader(
    "📎 Toca aquí para subir tu archivo",
    type=["docx", "xlsx", "pdf"],
    help="Word, Excel o PDF — máx 200MB",
    label_visibility="visible"
)

if archivo and archivo.name != st.session_state.nombre_archivo:
    with st.spinner("📖 Leyendo archivo..."):
        contenido = archivo.read()
        st.session_state.archivo_bytes = contenido
        st.session_state.nombre_archivo = archivo.name
        st.session_state.archivo_tipo = archivo.name.split('.')[-1].lower()
        st.session_state.resumen_data = None
        st.session_state.historial_chat = []
        st.session_state.lista_cambios = []
        st.session_state.cambios_aplicados = None
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

    # ── File Badge ──
    palabras = len(texto.split())
    lineas = len([l for l in texto.split('\n') if l.strip()])
    ext_icons = {"docx": "📄", "xlsx": "📊", "pdf": "📕"}
    ext_icon = ext_icons.get(tipo, "📎")
    st.markdown(f"""
    <div class="file-badge">
        <div class="file-icon">{ext_icon}</div>
        <div>
            <div class="file-info-name">{st.session_state.nombre_archivo}</div>
            <div class="file-info-stats">📝 {palabras:,} palabras &nbsp;·&nbsp; 📋 {lineas:,} líneas</div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    # ── Navegación tipo app con radio buttons estilizados ──
    st.markdown('<div class="oro-divider"></div>', unsafe_allow_html=True)

    # Usamos radio horizontal como nav
    nav = st.radio(
        "nav",
        ["📊 Resumen", "✍️ Editar", "💬 Chat", "🔍 Calidad"],
        horizontal=True,
        label_visibility="collapsed",
        key="nav_principal"
    )

    st.markdown('<div class="oro-divider"></div>', unsafe_allow_html=True)

    # ═══════════════════════════════════════
    # PANTALLA 1 — RESUMEN
    # ═══════════════════════════════════════
    if nav == "📊 Resumen":
        st.markdown('<div class="section-title">📊 Análisis del documento</div>', unsafe_allow_html=True)
        st.markdown('<div class="section-hint">La IA analiza el contenido y te da un resumen claro y profesional.</div>', unsafe_allow_html=True)

        if st.button("⚡ Generar Resumen", use_container_width=True):
            with st.spinner("🧠 Analizando con IA..."):
                data = solicitar_resumen_estructurado(texto)
                st.session_state.resumen_data = data

        data = st.session_state.resumen_data
        if data:
            emoji = data.get("emoji_categoria", "📋")
            titulo_doc = data.get("titulo", "Documento analizado")

            st.markdown(f"""
            <div class="summary-card">
                <div class="summary-card-title">{emoji} {titulo_doc}</div>
                {data.get("resumen_ejecutivo", "")}
            </div>
            """, unsafe_allow_html=True)

            # Métricas en grid 2 columnas
            metricas = data.get("metricas", {})
            if metricas:
                items = list(metricas.items())
                pills_html = '<div class="metrics-grid">'
                for k, v in items[:4]:
                    pills_html += f'<div class="metric-pill"><div class="metric-pill-label">{k}</div><div class="metric-pill-value">{v}</div></div>'
                pills_html += '</div>'
                st.markdown(pills_html, unsafe_allow_html=True)

            # Puntos clave
            puntos = data.get("puntos_clave", [])
            if puntos:
                tags_html = '<div class="tags-wrap">' + "".join([f'<span class="tag">✓ {p}</span>' for p in puntos]) + '</div>'
                st.markdown(tags_html, unsafe_allow_html=True)

            # Hallazgo
            hallazgo = data.get("hallazgo_destacado", "")
            if hallazgo:
                st.markdown(f'<div class="hallazgo-card">💡 <strong>Hallazgo:</strong> {hallazgo}</div>', unsafe_allow_html=True)

            # Exportar — UNA columna en móvil
            st.markdown('<div class="oro-divider"></div>', unsafe_allow_html=True)
            st.markdown('<div class="section-title">📥 Exportar informe</div>', unsafe_allow_html=True)

            word_bytes = exportar_word(texto, data, archivo_bytes=st.session_state.archivo_bytes, archivo_tipo=tipo, cambios=st.session_state.lista_cambios)
            st.download_button("📄 Descargar Word", word_bytes, "Informe_Oro.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True)

            excel_bytes = exportar_excel(texto, data, archivo_bytes=st.session_state.archivo_bytes, archivo_tipo=tipo, cambios=st.session_state.lista_cambios)
            st.download_button("📊 Descargar Excel", excel_bytes, "Informe_Oro.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True)

            pdf_bytes = exportar_pdf(texto, data)
            st.download_button("📕 Descargar PDF", pdf_bytes, "Informe_Oro.pdf",
                mime="application/pdf", use_container_width=True)
        else:
            st.markdown("""
            <div style="text-align:center;padding:2.5rem 0;color:#1e3a5f;">
                <div style="font-size:3rem">🧠</div>
                <div style="color:#374151;font-size:0.9rem;margin-top:0.5rem">
                    Toca <strong style="color:#3b82f6">Generar Resumen</strong> para analizar tu archivo
                </div>
            </div>
            """, unsafe_allow_html=True)

    # ═══════════════════════════════════════
    # PANTALLA 2 — EDICIÓN
    # ═══════════════════════════════════════
    elif nav == "✍️ Editar":
        st.markdown('<div class="section-title">✍️ Corregir palabras</div>', unsafe_allow_html=True)
        st.markdown('<div class="section-hint">Escribe qué quieres cambiar en lenguaje natural. Puedes hacer varios cambios seguidos.</div>', unsafe_allow_html=True)

        instruccion = st.text_input(
            "Instrucción",
            placeholder="Ej: cambia atletismo por BEISBOL",
            label_visibility="collapsed",
            key="input_edicion"
        )

        if instruccion:
            with st.spinner("🔍 Procesando..."):
                nuevos_cambios = solicitar_cambios(instruccion)

            if nuevos_cambios:
                st.session_state.lista_cambios.extend(nuevos_cambios)
                archivo_bytes_orig = st.session_state.archivo_bytes
                todos_cambios = st.session_state.lista_cambios

                if tipo == "docx":
                    final_bytes, n = reemplazar_docx_preservando_formato(archivo_bytes_orig, todos_cambios)
                elif tipo == "xlsx":
                    final_bytes, n = reemplazar_xlsx_preservando_formato(archivo_bytes_orig, todos_cambios)
                else:
                    txt_mod = texto
                    n = 0
                    for c in todos_cambios:
                        txt_mod, count = re.compile(re.escape(c["buscar"]), re.IGNORECASE).subn(c["reemplazar"], txt_mod)
                        n += count
                    final_bytes = txt_mod.encode()

                st.session_state.cambios_aplicados = final_bytes

                if n > 0:
                    st.markdown(f'<div class="info-box">✅ {n} cambio(s) aplicados correctamente</div>', unsafe_allow_html=True)
                else:
                    st.markdown(f'<div class="warn-box">⚠️ No encontré "{nuevos_cambios[0]["buscar"]}" en el documento</div>', unsafe_allow_html=True)
            else:
                st.markdown('<div class="warn-box">⚠️ No entendí la instrucción. Prueba: cambia X por Y</div>', unsafe_allow_html=True)

        # Historial de cambios
        if st.session_state.lista_cambios:
            st.markdown(f'<div class="section-title">📋 Cambios registrados ({len(st.session_state.lista_cambios)})</div>', unsafe_allow_html=True)
            for i, c in enumerate(st.session_state.lista_cambios, 1):
                st.markdown(f"""
                <div class="cambio-item">
                    <span class="cambio-num">{i}.</span>
                    <span style="color:#e2e8f0">{c['buscar']}</span>
                    <span class="cambio-arrow">→</span>
                    <span style="color:#fbbf24">{c['reemplazar']}</span>
                </div>
                """, unsafe_allow_html=True)

            if st.button("🗑️ Limpiar todos los cambios", use_container_width=True):
                st.session_state.lista_cambios = []
                st.session_state.cambios_aplicados = None
                st.rerun()

        # Descarga
        if st.session_state.cambios_aplicados:
            st.markdown('<div class="oro-divider"></div>', unsafe_allow_html=True)
            st.markdown('<div class="section-title">📥 Descargar corregido</div>', unsafe_allow_html=True)
            todos_cambios = st.session_state.lista_cambios

            word_out = exportar_word(texto, None, archivo_bytes=st.session_state.archivo_bytes, archivo_tipo=tipo, cambios=todos_cambios)
            st.download_button("📄 Word corregido", word_out, "Corregido.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True)

            excel_out = exportar_excel(texto, None, archivo_bytes=st.session_state.archivo_bytes, archivo_tipo=tipo, cambios=todos_cambios)
            st.download_button("📊 Excel corregido", excel_out, "Corregido.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True)

            txt_corr = texto
            for c in todos_cambios:
                txt_corr = re.compile(re.escape(c["buscar"]), re.IGNORECASE).sub(c["reemplazar"], txt_corr)
            pdf_c = exportar_pdf(txt_corr)
            st.download_button("📕 PDF corregido", pdf_c, "Corregido.pdf",
                mime="application/pdf", use_container_width=True)

    # ═══════════════════════════════════════
    # PANTALLA 3 — CHAT
    # ═══════════════════════════════════════
    elif nav == "💬 Chat":
        st.markdown('<div class="section-title">💬 Pregunta sobre el documento</div>', unsafe_allow_html=True)
        st.markdown('<div class="section-hint">Hazle cualquier pregunta al asistente sobre el contenido del archivo.</div>', unsafe_allow_html=True)

        for msg in st.session_state.historial_chat:
            with st.chat_message("user" if msg["rol"] == "Usuario" else "assistant"):
                st.write(msg["texto"])

        pregunta = st.chat_input("Escribe tu pregunta...")
        if pregunta:
            st.session_state.historial_chat.append({"rol": "Usuario", "texto": pregunta})
            with st.spinner("🤔 Pensando..."):
                respuesta = preguntar_al_documento(pregunta, texto)
            st.session_state.historial_chat.append({"rol": "Asistente", "texto": respuesta})
            st.rerun()

        if not st.session_state.historial_chat:
            st.markdown("""
            <div style="text-align:center;padding:2rem 0;color:#1e3a5f;">
                <div style="font-size:2.5rem">💬</div>
                <div style="color:#374151;font-size:0.85rem;margin-top:0.5rem">
                    Puedes preguntar cosas como:<br>
                    <em style="color:#1e3a5f">"¿Cuántos atletas hay en total?"</em><br>
                    <em style="color:#1e3a5f">"¿Qué municipios aparecen?"</em>
                </div>
            </div>
            """, unsafe_allow_html=True)

    # ═══════════════════════════════════════
    # PANTALLA 4 — CALIDAD
    # ═══════════════════════════════════════
    elif nav == "🔍 Calidad":
        st.markdown('<div class="section-title">🔍 Análisis de calidad</div>', unsafe_allow_html=True)
        st.markdown('<div class="section-hint">Detecta errores, inconsistencias o datos duplicados en el documento.</div>', unsafe_allow_html=True)

        if st.button("🔎 Analizar Calidad", use_container_width=True):
            with st.spinner("🔍 Revisando el documento..."):
                resultado = detectar_anomalias(texto)

            if resultado:
                nivel = resultado.get("nivel_calidad", "?")
                color_map = {"Alto": ("#22c55e", "#052e16"), "Medio": ("#f59e0b", "#1c1003"), "Bajo": ("#ef4444", "#1f0707")}
                color_fg, color_bg = color_map.get(nivel, ("#6b7280", "#111827"))
                st.markdown(f"""
                <div class="calidad-badge">
                    <span class="calidad-pill" style="background:{color_bg};color:{color_fg};border:2px solid {color_fg};">
                        Calidad {nivel}
                    </span>
                </div>
                """, unsafe_allow_html=True)

                anomalias = resultado.get("anomalias", [])
                if anomalias:
                    st.markdown('<div class="section-title">⚠️ Posibles problemas</div>', unsafe_allow_html=True)
                    for a in anomalias:
                        st.markdown(f'<div class="anomalia-item"><span>⚠️</span><span>{a}</span></div>', unsafe_allow_html=True)
                else:
                    st.markdown('<div class="info-box">✅ No se detectaron anomalías significativas</div>', unsafe_allow_html=True)

                rec = resultado.get("recomendacion", "")
                if rec:
                    st.markdown(f'<div class="hallazgo-card">💡 <strong>Recomendación:</strong> {rec}</div>', unsafe_allow_html=True)
            else:
                st.markdown('<div class="warn-box">⚠️ No se pudo analizar el documento</div>', unsafe_allow_html=True)

        else:
            st.markdown("""
            <div style="text-align:center;padding:2rem 0;color:#1e3a5f;">
                <div style="font-size:2.5rem">🔍</div>
                <div style="color:#374151;font-size:0.85rem;margin-top:0.5rem">
                    Toca <strong style="color:#3b82f6">Analizar Calidad</strong><br>para revisar el documento
                </div>
            </div>
            """, unsafe_allow_html=True)

else:
    # ── Estado vacío ──
    st.markdown("""
    <div class="empty-state">
        <div class="empty-icon">📂</div>
        <div class="empty-title">Sube un archivo para empezar</div>
        <div class="empty-hint">Toca el área de arriba para seleccionar tu documento</div>
        <div class="format-badges">
            <span class="format-badge">.docx</span>
            <span class="format-badge">.xlsx</span>
            <span class="format-badge">.pdf</span>
        </div>
    </div>
    """, unsafe_allow_html=True)

zona_horaria = pytz.timezone('America/Caracas')
hora = datetime.now(zona_horaria).strftime('%I:%M %p')
st.markdown(f"<p class='oro-footer'>🏆 Oro Asistente · {hora} VET · Powered by Gemini</p>", unsafe_allow_html=True)
