import os, json, ast, re
import streamlit as st
import google.generativeai as genai
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import PyPDF2
from fpdf import FPDF
try:
    import fitz  # pymupdf
    PYMUPDF_OK = True
except ImportError:
    PYMUPDF_OK = False
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
    background: linear-gradient(135deg, #0d1525, #111c2e) !important;
    color: #4b6080 !important;
    border: 1px solid #1e3a5f !important;
    border-radius: 14px !important;
    font-weight: 600 !important;
    font-size: 0.82rem !important;
    min-height: 3.8rem !important;
    width: 100% !important;
    transition: all 0.15s !important;
    letter-spacing: 0.01em !important;
    font-family: 'Outfit', sans-serif !important;
    line-height: 1.3 !important;
    white-space: pre-line !important;
}
.stButton > button:hover {
    background: linear-gradient(135deg, #132840, #1a3550) !important;
    color: #93c5fd !important;
    border-color: #2563eb !important;
}
.stButton > button:active {
    transform: scale(0.96) !important;
}
/* Botón acción principal (generar, confirmar, etc.) */
div[data-testid="stVerticalBlock"] > div:has(> div[data-testid="stButton"]) button[kind="primary"],
.btn-primary > button {
    background: linear-gradient(135deg, #1d4ed8, #2563eb) !important;
    color: white !important;
    border: none !important;
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
    padding: 0.5rem 0;
}

/* ── BOTONES DE NAVEGACIÓN ── */
.nav-tab-activo > button {
    background: linear-gradient(135deg, #065f46, #10b981) !important;
    color: white !important;
    border: none !important;
    font-weight: 700 !important;
    box-shadow: 0 4px 15px rgba(16,185,129,0.3) !important;
}
.nav-tab-inactivo > button {
    background: #021008 !important;
    color: #2d6a4f !important;
    border: 1.5px solid #0a3d1a !important;
    font-weight: 500 !important;
}
.nav-tab-inactivo > button:hover {
    background: #031510 !important;
    color: #34d399 !important;
    border-color: #10b981 !important;
}
/* Todos los botones de nav con misma altura */
.nav-tab-activo > button,
.nav-tab-inactivo > button {
    height: 3rem !important;
    border-radius: 12px !important;
    font-size: 0.85rem !important;
    font-family: 'Outfit', sans-serif !important;
    width: 100% !important;
    transition: all 0.15s !important;
}

/* ── GUÍA VISUAL ── */
.guia-card {
    background: linear-gradient(135deg, #021008, #031510);
    border: 1px solid #0a3d1a;
    border-left: 3px solid #10b981;
    border-radius: 14px;
    padding: 1rem 1.1rem;
    margin: 0.8rem 0 1rem 0;
    display: flex;
    gap: 0.9rem;
    align-items: flex-start;
}
.guia-icon { font-size: 1.6rem; flex-shrink: 0; margin-top: 0.1rem; }
.guia-titulo { color: #34d399; font-weight: 700; font-size: 0.95rem; margin-bottom: 0.25rem; }
.guia-texto  { color: #2d6a4f; font-size: 0.8rem; line-height: 1.55; }
.guia-texto em { color: #10b981; font-style: normal; font-weight: 500; }

.main .block-container { padding-bottom: 2rem !important; }

/* ── SELECTBOX SIDEBAR ── */
[data-testid="stSidebar"] {
    background: #080e1a !important;
}
</style>
""", unsafe_allow_html=True)


# ── CSS dinámico según tema ──
_TEMAS = {
    "oscuro": {
        "bg1": "#0a0e1a", "bg2": "#0d1525", "bg3": "#111827",
        "card": "#111827", "card2": "#162032",
        "borde": "#1e3a5f", "borde2": "#2a4a6b",
        "acento1": "#3b82f6", "acento2": "#60a5fa",
        "acento_grad": "linear-gradient(135deg,#1d4ed8,#2563eb)",
        "titulo_grad": "linear-gradient(135deg,#fbbf24,#f59e0b,#fde68a,#f59e0b)",
        "texto": "#e2e8f0", "texto2": "#93c5fd", "texto3": "#4b6080",
    },
    "azul": {
        "bg1": "#020818", "bg2": "#030d24", "bg3": "#041230",
        "card": "#041230", "card2": "#061840",
        "borde": "#0c3a7a", "borde2": "#1050aa",
        "acento1": "#38bdf8", "acento2": "#7dd3fc",
        "acento_grad": "linear-gradient(135deg,#0369a1,#0ea5e9)",
        "titulo_grad": "linear-gradient(135deg,#38bdf8,#7dd3fc,#bae6fd)",
        "texto": "#e0f2fe", "texto2": "#7dd3fc", "texto3": "#1e5a8a",
    },
    "verde": {
        "bg1": "#010c06", "bg2": "#021008", "bg3": "#03160a",
        "card": "#041208", "card2": "#051a0c",
        "borde": "#0a3d1a", "borde2": "#0f5225",
        "acento1": "#10b981", "acento2": "#34d399",
        "acento_grad": "linear-gradient(135deg,#065f46,#10b981)",
        "titulo_grad": "linear-gradient(135deg,#10b981,#34d399,#6ee7b7,#10b981)",
        "texto": "#d1fae5", "texto2": "#34d399", "texto3": "#065f46",
    },
    "rosa": {
        "bg1": "#120008", "bg2": "#1a000f", "bg3": "#220015",
        "card": "#1a000f", "card2": "#280018",
        "borde": "#7c0040", "borde2": "#9d0050",
        "acento1": "#f472b6", "acento2": "#f9a8d4",
        "acento_grad": "linear-gradient(135deg,#be185d,#ec4899)",
        "titulo_grad": "linear-gradient(135deg,#f472b6,#f9a8d4,#fce7f3)",
        "texto": "#fce7f3", "texto2": "#f9a8d4", "texto3": "#7c0040",
    },
    "ambar": {
        "bg1": "#0f0800", "bg2": "#180d00", "bg3": "#1f1100",
        "card": "#1a0e00", "card2": "#251500",
        "borde": "#78350f", "borde2": "#92400e",
        "acento1": "#f59e0b", "acento2": "#fbbf24",
        "acento_grad": "linear-gradient(135deg,#b45309,#f59e0b)",
        "titulo_grad": "linear-gradient(135deg,#fbbf24,#fde68a,#f59e0b)",
        "texto": "#fef3c7", "texto2": "#fbbf24", "texto3": "#78350f",
    },
}

_t = _TEMAS.get(st.session_state.get("tema", "oscuro"), _TEMAS["oscuro"])
st.markdown(f"""
<style>
.stApp {{
    background: linear-gradient(160deg, {_t['bg1']} 0%, {_t['bg2']} 50%, {_t['bg1']} 100%) !important;
}}
.main .block-container {{ background: transparent !important; }}
.oro-title {{
    background: {_t['titulo_grad']} !important;
    -webkit-background-clip: text !important;
    -webkit-text-fill-color: transparent !important;
    background-clip: text !important;
}}
.file-badge, .metric-pill, .cambio-item {{
    background: {_t['card']} !important;
    border-color: {_t['borde']} !important;
}}
.summary-card {{
    background: {_t['card2']} !important;
    border-color: {_t['borde2']} !important;
    border-left-color: {_t['acento1']} !important;
}}
.tag {{
    background: {_t['card']} !important;
    color: {_t['acento2']} !important;
    border-color: {_t['borde']} !important;
}}
.hallazgo-card {{
    background: {_t['card']} !important;
    border-left-color: {_t['acento1']} !important;
    color: {_t['texto2']} !important;
}}
.metric-pill-value {{ color: {_t['texto']} !important; }}
.metric-pill-label, .file-info-stats, .section-hint {{ color: {_t['texto3']} !important; }}
.section-title, .file-info-name {{ color: {_t['texto']} !important; }}
.summary-card-title {{ color: {_t['acento2']} !important; }}
.summary-card {{ color: {_t['texto2']} !important; }}
.oro-divider {{ background: linear-gradient(90deg, transparent, {_t['borde']}, transparent) !important; }}
[data-testid="stFileUploader"] > div {{
    border-color: {_t['borde']} !important;
    background: {_t['card']} !important;
}}
[data-testid="stSidebar"] {{ background: {_t['bg1']} !important; }}
.stButton > button {{
    background: {_t['card']} !important;
    border-color: {_t['borde']} !important;
    color: {_t['texto3']} !important;
}}
.stButton > button:hover {{
    border-color: {_t['acento1']} !important;
    color: {_t['acento2']} !important;
}}
.stTextInput > div > div > input {{
    background: {_t['card']} !important;
    border-color: {_t['borde']} !important;
    color: {_t['texto']} !important;
}}
.stTextInput > div > div > input:focus {{
    border-color: {_t['acento1']} !important;
}}
[data-testid="stDownloadButton"] > button {{
    background: linear-gradient(135deg,#065f46,#059669) !important;
    color: white !important;
    border: none !important;
}}
</style>
""", unsafe_allow_html=True)

# ==========================================
# CONEXIÓN GEMINI
# ==========================================
try:
    LLAVE_GEMINI = st.secrets["LLAVE_GEMINI"]
    genai.configure(api_key=LLAVE_GEMINI)
except Exception as e:
    st.error(f"🔑 Error configurando la IA: {e}")
    st.stop()

# Modelos en orden de preferencia — el sistema prueba cada uno automáticamente
MODELOS_FALLBACK = [
    "gemini-3.1-flash-lite-preview",
    "gemini-3.1-flash-preview",
    "gemini-3.1-pro-preview",
]

def llamar_ia(prompt, es_json=False):
    """
    Llama a la IA probando los modelos en orden.
    Si uno falla, pasa automáticamente al siguiente.
    El usuario nunca ve qué modelo se está usando.
    """
    for modelo in MODELOS_FALLBACK:
        try:
            model = genai.GenerativeModel(modelo)
            resp = model.generate_content(prompt)
            texto = resp.text
            if es_json:
                return extraer_json_seguro(texto, es_lista=texto.strip().startswith("["))
            return texto
        except Exception:
            continue
    return None

# Sidebar mínimo — solo para info interna si hace falta
with st.sidebar:
    st.caption("Oro Asistente v2")

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
    "generando_resumen": False,
    "resumen_error": False,
    "tab_activa": "resumen",
    "tema": "verde",
    "preview_cambio": None,
    "edicion_counter": 0,
    "texto_corregido": "",
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
        "Eres un analista profesional experto en documentos de cualquier tipo. Analiza este documento y devuelve SOLO un JSON.\n"
        "Identifica automáticamente el tipo de documento y adapta el análisis. El resumen_ejecutivo debe ser amigable, directo y máximo 3 oraciones. "
        "metricas_principales deben ser strings simples (no objetos).\n"
        "Formato exacto:\n"
        '{"titulo": "...", "emoji_categoria": "📋", "resumen_ejecutivo": "...", '
        '"metricas": {"Clave1": "Valor1", "Clave2": "Valor2"}, '
        '"puntos_clave": ["punto 1", "punto 2", "punto 3"], '
        '"hallazgo_destacado": "Una observación importante o curiosa del documento"}\n\n'
        f"DOCUMENTO:\n{texto[:12000]}"
    )
    resultado = llamar_ia(prompt, es_json=False)
    if resultado:
        return extraer_json_seguro(resultado)
    return None

def solicitar_informe_word(texto):
    prompt = (
        "Eres un analista experto en documentos de cualquier índole. Escribe un informe ejecutivo profesional. "
        "Usa párrafos cortos y claros, sin asteriscos ni markdown. "
        "Incluye: introducción, hallazgos principales, análisis y conclusión.\n\n"
        f"DATOS:\n{texto[:12000]}"
    )
    return llamar_ia(prompt) or "No se pudo generar el informe."

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

def solicitar_cambios(instruccion, texto_doc=""):
    """
    Interpreta instrucciones de edición en lenguaje natural.
    Soporta: reemplazar, agregar datos, completar campos vacíos.
    Devuelve lista de {buscar, reemplazar} para aplicar en el documento.
    """
    contexto_doc = f"\n\nFRAGMENTO DEL DOCUMENTO (para contexto):\n{texto_doc[:3000]}" if texto_doc else ""
    prompt = (
        "Eres un asistente experto en edición de documentos.\n"
        "El usuario quiere modificar su documento con esta instrucción.\n\n"
        f"INSTRUCCIÓN: \"{instruccion}\"\n"
        f"{contexto_doc}\n\n"
        "REGLAS:\n"
        "- Si dice 'cambia X por Y' o 'reemplaza X con Y': buscar=X, reemplazar=Y\n"
        "- Si dice 'agrega el número XXXX a Juan Pérez': buscar el texto exacto de Juan Pérez "
        "en el documento y reemplazarlo por 'Juan Pérez XXXX' (o donde corresponda en la fila/línea)\n"
        "- Si dice 'agrega X a la fila/celda de Y': buscar el texto de Y y agregar X al final\n"
        "- Si dice 'completa el campo de Y con X': igual, buscar Y y añadir X\n"
        "- SIEMPRE usa texto que realmente exista en el documento como 'buscar'\n"
        "- Si hay múltiples personas/cambios, incluye TODOS en el array\n"
        "- NO inventes texto que no esté en el documento\n\n"
        "Responde ÚNICAMENTE con JSON array (sin texto adicional):\n"
        '[{"buscar": "texto_exacto_del_doc", "reemplazar": "texto_nuevo_completo"}]'
    )
    resp = llamar_ia(prompt)
    if resp:
        resultado = extraer_json_seguro(resp, es_lista=True)
        if resultado and isinstance(resultado, list):
            validos = [
                c for c in resultado
                if isinstance(c, dict)
                and "buscar" in c and "reemplazar" in c
                and str(c["buscar"]).strip()
                and str(c["reemplazar"]).strip()
                and c["buscar"] != c["reemplazar"]
            ]
            if validos:
                return validos
    return extraer_cambio_con_regex(instruccion)

def preguntar_al_documento(pregunta, texto):
    historial = st.session_state.historial_chat
    contexto = "\n".join([f"{m['rol']}: {m['texto']}" for m in historial[-6:]])
    prompt = (
        f"Eres un asistente experto en análisis de documentos de cualquier tipo.\n"
        f"DOCUMENTO:\n{texto[:10000]}\n\n"
        f"CONVERSACIÓN PREVIA:\n{contexto}\n\n"
        f"PREGUNTA: {pregunta}\n"
        "Responde de forma concisa y directa en español."
    )
    return llamar_ia(prompt) or "No pude procesar tu pregunta."

def detectar_anomalias(texto):
    prompt = (
        "Analiza este documento en detalle. Detecta todos los problemas que encuentres.\n"
        "Clasifica cada problema por nivel de gravedad:\n"
        "- CRITICO: errores graves que cambian el significado o hacen el documento inválido\n"
        "- ALTO: errores importantes como datos incorrectos, incoherencias lógicas\n"
        "- MEDIO: errores ortográficos, de formato, datos duplicados\n"
        "- LEVE: sugerencias de mejora, redacción mejorable\n\n"
        "Devuelve SOLO este JSON (sin texto extra):\n"
        '{"nivel_general": "Excelente/Bueno/Regular/Deficiente",'
        '"puntaje": 85,'
        '"criticos": ["descripción del problema"],'
        '"altos": ["descripción del problema"],'
        '"medios": ["descripción del problema"],'
        '"leves": ["descripción del problema"],'
        '"recomendacion": "recomendación principal en una oración"}\n\n'
        f"DOCUMENTO:\n{texto[:12000]}"
    )
    resp = llamar_ia(prompt)
    if resp:
        return extraer_json_seguro(resp)
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
    # Márgenes más ajustados
    for section in doc.sections:
        section.top_margin    = Inches(0.8)
        section.bottom_margin = Inches(0.8)
        section.left_margin   = Inches(1.0)
        section.right_margin  = Inches(1.0)

    sty = doc.styles['Normal']
    sty.font.name = 'Calibri'
    sty.font.size = Pt(11)

    # ── Encabezado con banda azul ──
    tabla_hdr = doc.add_table(rows=1, cols=1)
    tabla_hdr.style = 'Table Grid'
    cell_hdr = tabla_hdr.cell(0, 0)
    cell_hdr.paragraphs[0].clear()
    run_hdr = cell_hdr.paragraphs[0].add_run(
        resumen_data.get("titulo", "INFORME EJECUTIVO") if resumen_data else "INFORME EJECUTIVO"
    )
    run_hdr.font.bold = True
    run_hdr.font.size = Pt(16)
    run_hdr.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    cell_hdr.paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    # Fondo azul
    tc = cell_hdr._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), '1E3A5F')
    tcPr.append(shd)
    doc.add_paragraph()

    # Fecha pequeña
    p_fecha = doc.add_paragraph()
    r_fecha = p_fecha.add_run(f'Generado: {fecha}')
    r_fecha.font.size = Pt(9)
    r_fecha.font.color.rgb = RGBColor(0x6B, 0x72, 0x80)
    p_fecha.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    doc.add_paragraph()

    if resumen_data:
        # Resumen ejecutivo en caja gris clara
        if resumen_data.get("resumen_ejecutivo"):
            t_res = doc.add_table(rows=1, cols=1)
            t_res.style = 'Table Grid'
            c_res = t_res.cell(0, 0)
            c_res.paragraphs[0].clear()
            r_res = c_res.paragraphs[0].add_run(resumen_data["resumen_ejecutivo"])
            r_res.font.size = Pt(10)
            r_res.font.italic = True
            tcPr2 = c_res._tc.get_or_add_tcPr()
            shd2 = OxmlElement('w:shd')
            shd2.set(qn('w:val'), 'clear')
            shd2.set(qn('w:color'), 'auto')
            shd2.set(qn('w:fill'), 'EFF6FF')
            tcPr2.append(shd2)
            doc.add_paragraph()

        if resumen_data.get("metricas"):
            h2 = doc.add_heading('Métricas Clave', level=1)
            h2.runs[0].font.color.rgb = RGBColor(0x1E, 0x40, 0xAF)
            tabla_m = doc.add_table(rows=1, cols=2)
            tabla_m.style = 'Table Grid'
            hdr = tabla_m.rows[0].cells
            for ci, txt in enumerate(['Indicador', 'Valor']):
                hdr[ci].paragraphs[0].clear()
                r = hdr[ci].paragraphs[0].add_run(txt)
                r.font.bold = True
                r.font.color.rgb = RGBColor(0xFF,0xFF,0xFF)
                hdr[ci].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
                tcPr_h = hdr[ci]._tc.get_or_add_tcPr()
                shd_h = OxmlElement('w:shd')
                shd_h.set(qn('w:val'), 'clear'); shd_h.set(qn('w:color'), 'auto')
                shd_h.set(qn('w:fill'), '1E40AF')
                tcPr_h.append(shd_h)
            for idx, (k, v) in enumerate(resumen_data["metricas"].items()):
                row_m = tabla_m.add_row().cells
                row_m[0].text = str(k)
                row_m[1].text = str(v)
                fill = 'F8FAFC' if idx % 2 == 0 else 'FFFFFF'
                for ci2 in range(2):
                    tcPr_d = row_m[ci2]._tc.get_or_add_tcPr()
                    shd_d = OxmlElement('w:shd')
                    shd_d.set(qn('w:val'),'clear'); shd_d.set(qn('w:color'),'auto')
                    shd_d.set(qn('w:fill'), fill)
                    tcPr_d.append(shd_d)
            doc.add_paragraph()

        if resumen_data.get("puntos_clave"):
            h3 = doc.add_heading('Puntos Clave', level=1)
            h3.runs[0].font.color.rgb = RGBColor(0x1E, 0x40, 0xAF)
            for punto in resumen_data["puntos_clave"]:
                p_b = doc.add_paragraph(style='List Bullet')
                p_b.add_run(punto).font.size = Pt(11)

        if resumen_data.get("hallazgo_destacado"):
            doc.add_paragraph()
            h4 = doc.add_heading('💡 Hallazgo Destacado', level=1)
            h4.runs[0].font.color.rgb = RGBColor(0x1E, 0x40, 0xAF)
            t_hall = doc.add_table(rows=1, cols=1)
            t_hall.style = 'Table Grid'
            c_hall = t_hall.cell(0,0)
            c_hall.paragraphs[0].clear()
            r_hall = c_hall.paragraphs[0].add_run(resumen_data["hallazgo_destacado"])
            r_hall.font.italic = True
            r_hall.font.size = Pt(10)
            tcPr_hall = c_hall._tc.get_or_add_tcPr()
            shd_hall = OxmlElement('w:shd')
            shd_hall.set(qn('w:val'),'clear'); shd_hall.set(qn('w:color'),'auto')
            shd_hall.set(qn('w:fill'), 'F0FDF4')
            tcPr_hall.append(shd_hall)
        doc.add_page_break()

    h_cont = doc.add_heading('Contenido del Documento', level=1)
    h_cont.runs[0].font.color.rgb = RGBColor(0x1E, 0x40, 0xAF)
    for linea in texto.split('\n'):
        linea_limpia = linea.strip().replace('*', '').replace('#', '')
        if linea_limpia:
            p = doc.add_paragraph(linea_limpia)
            p.paragraph_format.space_after = Pt(2)
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

    raw = pdf.output(dest='S')
    if isinstance(raw, (bytes, bytearray)):
        return bytes(raw)
    return raw.encode('latin-1')


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




def reemplazar_pdf_original(archivo_bytes, cambios):
    """
    Edita el PDF original preservando formato exacto (fuente, tamaño, bold, color).
    Usa redact de pymupdf para borrar y reinsertar con los mismos atributos.
    """
    if not PYMUPDF_OK:
        return archivo_bytes, 0

    doc = fitz.open(stream=archivo_bytes, filetype="pdf")
    conteo = 0

    for c in cambios:
        buscar   = str(c["buscar"]).strip()
        reemplazar = str(c["reemplazar"]).strip()
        if not buscar or buscar.lower() == reemplazar.lower():
            continue

        for pagina in doc:
            instancias = pagina.search_for(buscar, quads=False)
            if not instancias:
                continue

            # Extraer dict completo de la página para leer atributos de span
            bloques_dict = pagina.get_text("dict")["blocks"]

            for rect in instancias:
                # ── Buscar el span que contiene el texto ──
                font_size  = 11.0
                font_name  = "helv"
                color      = (0.0, 0.0, 0.0)
                bold       = False
                italic     = False
                # Color de fondo del área (para el redact)
                bg_color   = None

                for bloque in bloques_dict:
                    for linea in bloque.get("lines", []):
                        for span in linea.get("spans", []):
                            if buscar.lower() in span["text"].lower():
                                font_size = span.get("size", 11.0)
                                font_name = span.get("font", "helv")
                                ci = span.get("color", 0)
                                color = (
                                    ((ci >> 16) & 0xFF) / 255,
                                    ((ci >> 8)  & 0xFF) / 255,
                                    (ci & 0xFF) / 255,
                                )
                                flags = span.get("flags", 0)
                                bold   = bool(flags & 2**4)
                                italic = bool(flags & 2**1)
                                # Color de fondo del span si lo tiene
                                origin = span.get("origin", rect.tl)
                                break

                # ── Determinar fuente compatible ──
                fn_lower = font_name.lower()
                if "bold" in fn_lower and "italic" in fn_lower:
                    use_font = "Times-BoldItalic"
                elif "bold" in fn_lower or bold:
                    use_font = "Helvetica-Bold"
                elif "italic" in fn_lower or italic:
                    use_font = "Helvetica-Oblique"
                elif "times" in fn_lower or "serif" in fn_lower:
                    use_font = "Times-Roman"
                elif "courier" in fn_lower or "mono" in fn_lower:
                    use_font = "Courier"
                else:
                    use_font = "Helvetica"

                # ── 1. Redact: borra el texto original con el fondo correcto ──
                # Detectar color de fondo real del área
                try:
                    pix = pagina.get_pixmap(clip=rect, dpi=72)
                    # Pixel central del área
                    cx = pix.width  // 2
                    cy = pix.height // 2
                    sample = pix.pixel(cx, cy)
                    bg = (sample[0]/255, sample[1]/255, sample[2]/255)
                except Exception:
                    bg = (1.0, 1.0, 1.0)

                pagina.add_redact_annot(rect, fill=bg)
                pagina.apply_redactions()

                # ── 2. Reinsertar texto nuevo con los mismos atributos ──
                # Calcular posición Y: baseline es bottom del rect menos pequeño margen
                baseline_y = rect.y1 - 1.5
                pagina.insert_text(
                    fitz.Point(rect.x0, baseline_y),
                    reemplazar,
                    fontname=use_font,
                    fontsize=font_size,
                    color=color,
                )
                conteo += 1

    buf = BytesIO()
    doc.save(buf)
    doc.close()
    return buf.getvalue(), conteo


# ==========================================
# SUBIDA DE ARCHIVO
# ==========================================


st.markdown('<div class="oro-divider"></div>', unsafe_allow_html=True)

# CSS para traducir el uploader al español
st.markdown("""
<style>
[data-testid="stFileUploaderDropzoneInstructions"] > div > span::after {
    content: "Arrastra tu archivo aquí";
}
[data-testid="stFileUploaderDropzoneInstructions"] > div > span {
    font-size: 0 !important;
}
[data-testid="stFileUploaderDropzoneInstructions"] > div > small::after {
    content: "Límite 200MB por archivo • DOCX, XLSX, PDF";
}
[data-testid="stFileUploaderDropzoneInstructions"] > div > small {
    font-size: 0 !important;
}
[data-testid="stFileUploadDropzone"] > div > button {
    visibility: hidden;
    position: relative;
}
[data-testid="stFileUploadDropzone"] > div > button::after {
    content: "Seleccionar archivo";
    visibility: visible;
    position: absolute;
    left: 0; right: 0;
    text-align: center;
}
</style>
""", unsafe_allow_html=True)

archivo = st.file_uploader(
    "📎 Toca aquí para subir tu archivo",
    type=["docx", "xlsx", "pdf"],
    help="Word, Excel o PDF — máx 200MB",
    label_visibility="visible"
)

if archivo and archivo.name != st.session_state.nombre_archivo:
    with st.spinner("📖 Cargando archivo..."):
        contenido = archivo.read()
        st.session_state.archivo_bytes     = contenido
        st.session_state.nombre_archivo    = archivo.name
        st.session_state.archivo_tipo      = archivo.name.split(".")[-1].lower()
        st.session_state.resumen_data      = None
        st.session_state.historial_chat    = []
        st.session_state.lista_cambios     = []
        st.session_state.cambios_aplicados = None
        st.session_state.texto_corregido   = ""
        st.session_state.preview_cambio    = None
        st.session_state.resumen_error     = False
        st.session_state.generando_resumen = False
        texto = ""
        try:
            if archivo.name.endswith(".docx"):
                doc = Document(BytesIO(contenido))
                partes = [p.text for p in doc.paragraphs if p.text.strip()]
                for t in doc.tables:
                    for row in t.rows:
                        celdas = list(dict.fromkeys([c.text.strip() for c in row.cells]))
                        if any(celdas):
                            partes.append(" | ".join(celdas))
                texto = "\n".join(partes)
            elif archivo.name.endswith(".xlsx"):
                wb = openpyxl.load_workbook(BytesIO(contenido), data_only=True, read_only=True)
                for s in wb.worksheets:
                    for r in s.iter_rows(values_only=True):
                        linea = " | ".join([str(c) for c in r if c is not None and str(c).strip()])
                        if linea.strip():
                            texto += linea + "\n"
                wb.close()
            elif archivo.name.endswith(".pdf"):
                reader = PyPDF2.PdfReader(BytesIO(contenido))
                for p in reader.pages:
                    t = p.extract_text()
                    if t:
                        texto += t + "\n"
            st.session_state.texto_extraido    = texto
            st.session_state.generando_resumen = False
            # NO auto-generar resumen — el usuario lo pide cuando quiera
        except Exception as e:
            st.error(f"Error leyendo el archivo: {e}")



# ==========================================
# PANEL PRINCIPAL
# ==========================================
# Seguridad anti-loop: si no hay texto, nunca generar resumen
if not st.session_state.get("texto_extraido") and st.session_state.get("generando_resumen"):
    st.session_state.generando_resumen = False

if st.session_state.texto_extraido:
    texto = st.session_state.texto_extraido
    tipo  = st.session_state.archivo_tipo
    texto_activo = st.session_state.texto_corregido if st.session_state.texto_corregido else texto

    # ── File badge ──
    palabras  = len(texto.split())
    ext_icons = {"docx":"📄","xlsx":"📊","pdf":"📕"}
    ext_icon  = ext_icons.get(tipo,"📎")
    cambios_n = len(st.session_state.lista_cambios)
    badge_extra = f' &nbsp;·&nbsp; ✏️ <strong style="color:#10b981">{cambios_n} cambio(s)</strong>' if cambios_n else ""
    st.markdown(f"""
    <div class="file-badge">
        <div class="file-icon">{ext_icon}</div>
        <div>
            <div class="file-info-name">{st.session_state.nombre_archivo}</div>
            <div class="file-info-stats">📝 {palabras:,} palabras{badge_extra}</div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    # ══════════════════════════════════════
    # RESUMEN
    # ══════════════════════════════════════
    # Si no hay resumen aún, mostrar botón prominente
    if not st.session_state.resumen_data and not st.session_state.generando_resumen:
        st.markdown("""
        <div style="text-align:center;padding:1.5rem 0 0.5rem">
            <div style="font-size:2.5rem">🧠</div>
            <div style="color:#34d399;font-weight:700;font-size:1rem;margin-top:0.5rem">
                Listo para analizar
            </div>
            <div style="color:#065f46;font-size:0.78rem;margin-top:0.2rem">
                Toca el botón para generar el resumen inteligente
            </div>
        </div>
        """, unsafe_allow_html=True)
        if st.button("⚡ Analizar documento", use_container_width=True):
            st.session_state.generando_resumen = True
            st.rerun()

    # Procesando — llamar a IA solo cuando el usuario lo pidió
    if st.session_state.generando_resumen:
        st.markdown("""
        <div style="text-align:center;padding:2rem 0;">
            <div style="font-size:2.8rem;animation:pulse-glow 1.2s ease-in-out infinite">🧠</div>
            <div style="color:#34d399;font-weight:700;font-size:1rem;margin-top:0.7rem">
                Analizando tu documento...
            </div>
            <div style="color:#065f46;font-size:0.78rem;margin-top:0.3rem">Esto puede tomar unos segundos ⏳</div>
        </div>
        """, unsafe_allow_html=True)
        texto_para_resumen = st.session_state.texto_corregido if st.session_state.texto_corregido else texto
        data_nueva = solicitar_resumen_estructurado(texto_para_resumen)
        st.session_state.generando_resumen = False
        if data_nueva:
            st.session_state.resumen_data  = data_nueva
            st.session_state.resumen_error = False
        else:
            st.session_state.resumen_error = True
        st.rerun()

    if st.session_state.get("resumen_error"):
        st.markdown("""
        <div style="text-align:center;padding:1.2rem;background:#1c0003;border:1px solid #7f1d1d;
        border-radius:14px;margin:0.8rem 0">
            <div style="font-size:1.8rem">⚠️</div>
            <div style="color:#fca5a5;font-weight:600;margin-top:0.3rem">No se pudo generar el resumen</div>
            <div style="color:#6b7280;font-size:0.78rem;margin-top:0.2rem">Problema temporal con la IA</div>
        </div>
        """, unsafe_allow_html=True)
        if st.button("🔄 Reintentar", use_container_width=True):
            st.session_state.resumen_error     = False
            st.session_state.generando_resumen = True
            st.rerun()

    data = st.session_state.resumen_data
    if data:
        emoji = data.get("emoji_categoria","📋")
        titulo_doc = data.get("titulo","Documento analizado")

        st.markdown(f"""
        <div class="summary-card">
            <div class="summary-card-title">{emoji} {titulo_doc}</div>
            {data.get("resumen_ejecutivo","")}
        </div>
        """, unsafe_allow_html=True)

        metricas = data.get("metricas",{})
        if metricas:
            items = list(metricas.items())
            pills = '<div class="metrics-grid">'
            for k,v in items[:4]:
                pills += f'<div class="metric-pill"><div class="metric-pill-label">{k}</div><div class="metric-pill-value">{v}</div></div>'
            pills += '</div>'
            st.markdown(pills, unsafe_allow_html=True)

        puntos = data.get("puntos_clave",[])
        if puntos:
            tags = '<div class="tags-wrap">'+"".join([f'<span class="tag">✓ {p}</span>' for p in puntos])+'</div>'
            st.markdown(tags, unsafe_allow_html=True)

        hallazgo = data.get("hallazgo_destacado","")
        if hallazgo:
            st.markdown(f'<div class="hallazgo-card">💡 <strong>Hallazgo:</strong> {hallazgo}</div>', unsafe_allow_html=True)

        # ── Exportar ──
        st.markdown('<div class="oro-divider"></div>', unsafe_allow_html=True)
        st.markdown('<div class="section-title">📥 Exportar informe</div>', unsafe_allow_html=True)
        ab = st.session_state.archivo_bytes
        ca = st.session_state.lista_cambios
        c1,c2,c3 = st.columns(3)
        with c1:
            wb2 = exportar_word(texto_activo, data, archivo_bytes=ab, archivo_tipo=tipo, cambios=ca)
            st.download_button("📄 Word", wb2, "Informe.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True)
        with c2:
            eb2 = exportar_excel(texto_activo, data, archivo_bytes=ab, archivo_tipo=tipo, cambios=ca)
            st.download_button("📊 Excel", eb2, "Informe.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True)
        with c3:
            pb2 = exportar_pdf(texto_activo, data)
            st.download_button("📕 PDF", pb2, "Informe.pdf",
                mime="application/pdf", use_container_width=True)

        # Regenerar
        if st.button("🔄 Regenerar resumen", use_container_width=True):
            st.session_state.generando_resumen = True
            st.session_state.resumen_data = None
            st.rerun()

    st.markdown('<div class="oro-divider"></div>', unsafe_allow_html=True)

    # ══════════════════════════════════════
    # BOTÓN EVALUAR
    # ══════════════════════════════════════
    if st.button("🔍 Evaluar documento", use_container_width=True):
        with st.spinner("Analizando errores e inconsistencias..."):
            resultado = detectar_anomalias(texto_activo)

        if resultado:
            nivel   = resultado.get("nivel_general","Regular")
            puntaje = resultado.get("puntaje",0)
            nivel_cfg = {
                "Excelente":("#10b981","#021008","🟢"),
                "Bueno":    ("#34d399","#021008","🟢"),
                "Regular":  ("#f59e0b","#1c1003","🟡"),
                "Deficiente":("#ef4444","#1f0707","🔴"),
            }
            cfg = nivel_cfg.get(nivel, nivel_cfg["Regular"])
            st.markdown(f"""
            <div style="text-align:center;padding:1.2rem 0 0.5rem">
                <div style="font-size:3rem">{cfg[2]}</div>
                <div style="color:{cfg[0]};font-size:1.3rem;font-weight:800">{nivel}</div>
                <div style="color:#2d6a4f;font-size:0.8rem;margin-top:0.2rem">
                    Puntaje: <strong style="color:{cfg[0]}">{puntaje}/100</strong>
                </div>
            </div>
            """, unsafe_allow_html=True)

            niveles_eval = [
                ("criticos","🔴 Crítico","#ef4444","#1f0707","#450a0a"),
                ("altos",   "🟠 Alto",   "#f97316","#1c0a03","#431407"),
                ("medios",  "🟡 Medio",  "#f59e0b","#1c1003","#451a03"),
                ("leves",   "🟢 Leve",   "#22c55e","#052e16","#14532d"),
            ]
            hay = False
            for key,label,cfg,cbg,cbrd in niveles_eval:
                items_e = resultado.get(key,[])
                if items_e:
                    hay = True
                    rows = "".join([f'<div style="color:#d1fae5;font-size:0.8rem;padding:0.25rem 0;border-bottom:1px solid {cbrd}">• {it}</div>' for it in items_e])
                    st.markdown(f"""
                    <div style="background:{cbg};border:1px solid {cbrd};border-left:4px solid {cfg};
                    border-radius:12px;padding:0.8rem 1rem;margin:0.5rem 0">
                        <div style="color:{cfg};font-weight:700;font-size:0.85rem;margin-bottom:0.4rem">{label}</div>
                        {rows}
                    </div>
                    """, unsafe_allow_html=True)
            if not hay:
                st.markdown('<div class="info-box">✅ ¡Sin problemas detectados! El documento se ve en buen estado 🎉</div>', unsafe_allow_html=True)
            rec = resultado.get("recomendacion","")
            if rec:
                st.markdown(f'<div class="hallazgo-card" style="margin-top:0.8rem">💡 <strong>Recomendación:</strong> {rec}</div>', unsafe_allow_html=True)
        else:
            st.markdown('<div class="warn-box">⚠️ No se pudo evaluar. Intenta de nuevo.</div>', unsafe_allow_html=True)

    st.markdown('<div class="oro-divider"></div>', unsafe_allow_html=True)

    # ══════════════════════════════════════
    # PREVIEW de cambio pendiente
    # ══════════════════════════════════════
    if st.session_state.preview_cambio:
        preview = st.session_state.preview_cambio
        st.markdown("""
        <div style="background:#021008;border:1px solid #10b981;border-radius:14px;
        padding:0.9rem 1rem;margin:0.5rem 0">
            <div style="color:#34d399;font-weight:700;font-size:0.88rem;margin-bottom:0.6rem">
                👁 Vista previa del cambio
            </div>
        """, unsafe_allow_html=True)
        for c in preview:
            bq = c["buscar"][:50]+("..." if len(c["buscar"])>50 else "")
            rq = c["reemplazar"][:50]+("..." if len(c["reemplazar"])>50 else "")
            idx = texto_activo.lower().find(c["buscar"].lower())
            if idx != -1:
                ini=max(0,idx-30); fin=min(len(texto_activo),idx+len(c["buscar"])+30)
                ca2=texto_activo[ini:idx]; cd=texto_activo[idx+len(c["buscar"]):fin]
                st.markdown(
                    f'<div style="font-size:0.78rem;margin-bottom:0.3rem">'
                    f'<span style="color:#6b7280;font-size:0.65rem;text-transform:uppercase">Antes: </span>'
                    f'<span style="color:#fca5a5;font-family:monospace">...{ca2}<mark style="background:#7f1d1d;color:#fca5a5;border-radius:3px;padding:0 3px">{bq}</mark>{cd}...</span>'
                    f'</div>'
                    f'<div style="font-size:0.78rem">'
                    f'<span style="color:#6b7280;font-size:0.65rem;text-transform:uppercase">Después: </span>'
                    f'<span style="color:#86efac;font-family:monospace">...{ca2}<mark style="background:#14532d;color:#86efac;border-radius:3px;padding:0 3px">{rq}</mark>{cd}...</span>'
                    f'</div>',
                    unsafe_allow_html=True)
            else:
                st.markdown(f'<div style="color:#fbbf24;font-size:0.8rem">⚠️ "{bq}" no encontrado en el documento</div>', unsafe_allow_html=True)
        st.markdown('</div>', unsafe_allow_html=True)

        cs, cn = st.columns(2)
        with cs:
            if st.button("✅ Confirmar", use_container_width=True):
                st.session_state.lista_cambios.extend(preview)
                st.session_state.preview_cambio = None
                todos_c = st.session_state.lista_cambios
                ab_orig = st.session_state.archivo_bytes
                if tipo == "docx":
                    final_bytes, n = reemplazar_docx_preservando_formato(ab_orig, todos_c)
                elif tipo == "xlsx":
                    final_bytes, n = reemplazar_xlsx_preservando_formato(ab_orig, todos_c)
                elif tipo == "pdf" and PYMUPDF_OK:
                    final_bytes, n = reemplazar_pdf_original(ab_orig, todos_c)
                else:
                    txt_m=texto_activo; n=0
                    for c2 in todos_c:
                        txt_m,cnt = re.compile(re.escape(c2["buscar"]),re.IGNORECASE).subn(c2["reemplazar"],txt_m)
                        n+=cnt
                    final_bytes = txt_m.encode()
                txt_c=texto_activo
                for c2 in todos_c:
                    txt_c = re.compile(re.escape(c2["buscar"]),re.IGNORECASE).sub(c2["reemplazar"],txt_c)
                st.session_state.texto_corregido = txt_c
                st.session_state.cambios_aplicados = final_bytes
                st.session_state.resumen_data = None
                st.session_state.generando_resumen = True
                st.session_state.edicion_counter += 1
                st.session_state.historial_chat.append({
                    "rol":"Asistente",
                    "texto":f"✅ ¡Listo! Cambié **{preview[0]['buscar']}** → **{preview[0]['reemplazar']}** en el documento. ¿Algo más?"
                })
                st.rerun()
        with cn:
            if st.button("❌ Cancelar", use_container_width=True):
                st.session_state.preview_cambio = None
                st.session_state.edicion_counter += 1
                st.rerun()

    # ── Descarga si hay cambios ──
    if st.session_state.cambios_aplicados:
        with st.expander(f"📥 Descargar documento corregido ({len(st.session_state.lista_cambios)} cambio(s))", expanded=False):
            fb = st.session_state.cambios_aplicados
            todos_c = st.session_state.lista_cambios
            if tipo == "docx":
                st.download_button("📄 Word corregido", fb, "Corregido.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)
            elif tipo == "xlsx":
                st.download_button("📊 Excel corregido", fb, "Corregido.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", use_container_width=True)
            elif tipo == "pdf" and PYMUPDF_OK:
                st.download_button("📕 PDF corregido", fb, "Corregido.pdf",
                    mime="application/pdf", use_container_width=True)
            wc = exportar_word(st.session_state.texto_corregido or texto, None)
            st.download_button("📄 Exportar como Word", wc, "Exportado.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)
            if st.button("🗑️ Limpiar todos los cambios", use_container_width=True):
                st.session_state.lista_cambios=[]
                st.session_state.cambios_aplicados=None
                st.session_state.texto_corregido=""
                st.session_state.preview_cambio=None
                st.rerun()

    # ══════════════════════════════════════
    # CHAT — siempre visible al fondo
    # ══════════════════════════════════════
    # Historial
    for msg in st.session_state.historial_chat:
        with st.chat_message("user" if msg["rol"]=="Usuario" else "assistant"):
            st.write(msg["texto"])

    if not st.session_state.historial_chat:
        st.markdown("""
        <div style="text-align:center;padding:1.5rem 0 0.5rem">
            <div style="font-size:2rem">💬</div>
            <div style="color:#065f46;font-size:0.85rem;margin-top:0.4rem;line-height:1.7">
                <strong style="color:#10b981">Conversa sobre el documento</strong><br>
                <span style="color:#2d6a4f">Edita, pregunta o pide cambios en lenguaje natural</span>
            </div>
            <div style="display:flex;justify-content:center;gap:0.5rem;flex-wrap:wrap;margin-top:0.8rem">
                <span style="background:#021008;border:1px solid #0a3d1a;border-radius:20px;padding:0.25rem 0.7rem;font-size:0.72rem;color:#10b981">cambia X por Y</span>
                <span style="background:#021008;border:1px solid #0a3d1a;border-radius:20px;padding:0.25rem 0.7rem;font-size:0.72rem;color:#10b981">¿cuántas personas hay?</span>
                <span style="background:#021008;border:1px solid #0a3d1a;border-radius:20px;padding:0.25rem 0.7rem;font-size:0.72rem;color:#10b981">resume en 3 puntos</span>
            </div>
        </div>
        """, unsafe_allow_html=True)

    entrada = st.chat_input("✍️ Escribe un cambio o una pregunta...")
    if entrada:
        st.session_state.historial_chat.append({"rol":"Usuario","texto":entrada})
        palabras_cambio = ["cambia","reemplaza","sustituye","corrige","agrega","añade","borra","elimina","pon","escribe","modifica","quita","actualiza"]
        es_cambio = any(p in entrada.lower() for p in palabras_cambio)
        if es_cambio:
            with st.spinner("🔍 Procesando cambio..."):
                nuevos = solicitar_cambios(entrada, texto_activo)
            if nuevos:
                st.session_state.preview_cambio = nuevos
                st.session_state.historial_chat.append({
                    "rol":"Asistente",
                    "texto":"Encontré el cambio 👆 Revisa la vista previa y confirma si es correcto."
                })
            else:
                st.session_state.historial_chat.append({
                    "rol":"Asistente",
                    "texto":"No encontré qué cambiar exactamente. Intenta: *cambia 'palabra' por 'nueva palabra'*"
                })
        else:
            with st.spinner("🤔 Pensando..."):
                resp = preguntar_al_documento(entrada, texto_activo)
            st.session_state.historial_chat.append({"rol":"Asistente","texto":resp})
        st.rerun()

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

st.markdown(f"<p class='oro-footer'>🏆 Oro Asistente · {hora} VET</p>", unsafe_allow_html=True)
