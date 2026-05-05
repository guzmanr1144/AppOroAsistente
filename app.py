import os, json, ast, re, warnings, copy
warnings.filterwarnings("ignore", category=DeprecationWarning)
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
    import fitz
    PYMUPDF_OK = True
except ImportError:
    PYMUPDF_OK = False
from io import BytesIO
from datetime import datetime
import pytz

st.set_page_config(page_title="Oro Asistente", page_icon="🏆", layout="centered", initial_sidebar_state="collapsed")

# ══════════════════════════════════════════════════════════════
# CSS — cacheado por tema
# ══════════════════════════════════════════════════════════════
@st.cache_data(show_spinner=False)
def _get_all_css(tema_key="noche"):
    _T = {
        "noche": {
            "bg1":"#0a0e1a","bg2":"#0d1222","bg3":"#10162a",
            "card":"#141828","card2":"#1a2035",
            "borde":"#2a3560","borde2":"#364480",
            "acento1":"#6b83f8","acento2":"#8fa0ff","acento3":"#4f6ef7",
            "titulo_grad":"linear-gradient(135deg,#6b83f8,#a78bfa,#6b83f8)",
            "texto":"#e8ecff","texto2":"#a8b8f0","texto3":"#5060a0",
            "sombra":"rgba(107,131,248,0.2)","sombra2":"rgba(107,131,248,0.08)",
        },
        "carbon": {
            "bg1":"#111418","bg2":"#161b22","bg3":"#1c2330",
            "card":"#1e2530","card2":"#242e3d",
            "borde":"#2d3f55","borde2":"#3a5070",
            "acento1":"#10b981","acento2":"#34d399","acento3":"#6ee7b7",
            "titulo_grad":"linear-gradient(135deg,#10b981,#06b6d4,#10b981)",
            "texto":"#d8f0e8","texto2":"#90c8b0","texto3":"#3a6050",
            "sombra":"rgba(16,185,129,0.2)","sombra2":"rgba(16,185,129,0.08)",
        },
        "cosmos": {
            "bg1":"#0d0818","bg2":"#120d22","bg3":"#18102c",
            "card":"#1a1030","card2":"#201540",
            "borde":"#3a2060","borde2":"#502888",
            "acento1":"#a78bfa","acento2":"#c4b0ff","acento3":"#7c3aed",
            "titulo_grad":"linear-gradient(135deg,#a78bfa,#f472b6,#a78bfa)",
            "texto":"#f0e8ff","texto2":"#c0a8f0","texto3":"#705898",
            "sombra":"rgba(167,139,250,0.2)","sombra2":"rgba(167,139,250,0.08)",
        },
    }
    t = _T.get(tema_key, _T["noche"])
    return f"""<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800;900&family=JetBrains+Mono:wght@400;500;600&display=swap');
*{{box-sizing:border-box}}
html,body,[class*="css"]{{font-family:'Inter',sans-serif!important;-webkit-tap-highlight-color:transparent}}
.stApp{{background:linear-gradient(145deg,{t['bg2']} 0%,{t['bg1']} 40%,{t['bg3']} 100%)!important;min-height:100vh}}
.main .block-container{{padding:.8rem .9rem 5rem .9rem!important;max-width:460px!important;margin:0 auto!important;background:transparent!important}}
#MainMenu,footer,header{{visibility:hidden}}[data-testid="stToolbar"]{{display:none}}

/* ── HEADER ── */
.oro-header{{text-align:center;padding:1.6rem 0 .5rem}}
.oro-logo-wrap{{position:relative;display:inline-block;margin-bottom:.3rem}}
.oro-logo{{font-size:2.8rem;display:block;filter:drop-shadow(0 4px 12px {t['sombra']});animation:float 4s ease-in-out infinite}}
.oro-logo-ring{{position:absolute;inset:-8px;border:2px solid {t['acento1']};border-radius:50%;opacity:.2;animation:spin 8s linear infinite}}
@keyframes float{{0%,100%{{transform:translateY(0)}}50%{{transform:translateY(-5px)}}}}
@keyframes spin{{to{{transform:rotate(360deg)}}}}
@keyframes fadeUp{{from{{opacity:0;transform:translateY(8px)}}to{{opacity:1;transform:translateY(0)}}}}
@keyframes shimmer{{0%,100%{{background-position:0% 50%}}50%{{background-position:100% 50%}}}}
.oro-title{{font-size:1.9rem;font-weight:900;background:{t['titulo_grad']}!important;background-size:200%!important;-webkit-background-clip:text!important;-webkit-text-fill-color:transparent!important;background-clip:text!important;letter-spacing:-.03em;animation:shimmer 5s ease infinite}}
.oro-badge{{display:inline-flex;align-items:center;gap:.3rem;background:{t['card']};border:1px solid {t['borde']};border-radius:20px;padding:.2rem .8rem;font-size:.65rem;color:{t['acento1']};font-weight:700;letter-spacing:.06em;margin-top:.3rem;box-shadow:0 2px 8px {t['sombra2']}}}

/* ── CARDS Y CONTENEDORES ── */
.file-badge{{display:flex;align-items:center;gap:.75rem;background:{t['card']};border:1px solid {t['borde']};border-radius:18px;padding:.85rem 1rem;margin:.5rem 0;animation:fadeUp .3s ease;box-shadow:0 4px 16px {t['sombra2']}}}
.file-icon{{font-size:1.6rem;flex-shrink:0}}
.file-info-name{{color:{t['texto']}!important;font-weight:700;font-size:.85rem;word-break:break-all;line-height:1.3}}
.file-info-stats{{color:{t['texto3']}!important;font-size:.7rem;margin-top:.15rem;display:flex;gap:.4rem;flex-wrap:wrap}}
.stat-chip{{background:{t['bg2']};border:1px solid {t['borde']};border-radius:8px;padding:.05rem .4rem;font-size:.65rem;color:{t['acento1']};font-weight:600}}

/* ── BOTONES ── */
.stButton>button{{background:{t['card']}!important;color:{t['texto2']}!important;border:1.5px solid {t['borde']}!important;border-radius:12px!important;font-weight:600!important;font-size:.84rem!important;min-height:3rem!important;width:100%!important;transition:all .15s ease!important;font-family:'Inter',sans-serif!important;box-shadow:0 2px 6px {t['sombra2']}!important}}
.stButton>button:hover{{border-color:{t['acento1']}!important;color:{t['acento2']}!important;background:{t['bg2']}!important;box-shadow:0 4px 14px {t['sombra']}!important;transform:translateY(-1px)!important}}
.stButton>button:active{{transform:scale(.97)!important;box-shadow:none!important}}
.btn-analizar>button{{background:linear-gradient(135deg,{t['acento1']},{t['acento3']})!important;color:white!important;border:none!important;font-weight:700!important;box-shadow:0 4px 14px {t['sombra']}!important}}
.btn-analizar>button:hover{{filter:brightness(1.08)!important;box-shadow:0 6px 20px {t['sombra']}!important}}
.btn-evaluar>button{{background:linear-gradient(135deg,#059669,#0891b2)!important;color:white!important;border:none!important;font-weight:700!important;box-shadow:0 4px 14px rgba(5,150,105,.2)!important}}
.btn-evaluar>button:hover{{filter:brightness(1.08)!important}}
.btn-peligro>button{{background:linear-gradient(135deg,#dc2626,#e11d48)!important;color:white!important;border:none!important;font-weight:600!important}}
[data-testid="stDownloadButton"]>button{{background:linear-gradient(135deg,{t['acento1']},{t['acento3']})!important;color:white!important;border:none!important;border-radius:12px!important;font-weight:700!important;height:2.8rem!important;width:100%!important;box-shadow:0 3px 10px {t['sombra']}!important;transition:all .15s!important}}
[data-testid="stDownloadButton"]>button:hover{{filter:brightness(1.08)!important;box-shadow:0 5px 16px {t['sombra']}!important;transform:translateY(-1px)!important}}

/* ── MODO TABS (archivo/imagen) ── */
.nav-tab-activo>button{{background:linear-gradient(135deg,{t['acento1']},{t['acento3']})!important;color:white!important;border:none!important;font-weight:700!important;box-shadow:0 3px 10px {t['sombra']}!important}}
.nav-tab-inactivo>button{{background:{t['card']}!important;color:{t['texto3']}!important;border:1.5px solid {t['borde']}!important;font-weight:500!important}}

/* ── UPLOADER ── */
[data-testid="stFileUploader"]{{background:transparent!important;border:none!important}}
[data-testid="stFileUploader"]>div{{background:{t['card2']}!important;border:2px dashed {t['borde2']}!important;border-radius:20px!important;padding:1.3rem!important;box-shadow:0 4px 16px {t['sombra']}!important;transition:border-color .3s!important}}
[data-testid="stFileUploader"]>div:hover{{border-color:{t['acento1']}!important}}
[data-testid="stFileUploader"] label{{color:{t['acento1']}!important;font-weight:600!important;font-size:.92rem!important}}

/* ── SUMMARY CARD ── */
.summary-card{{background:{t['card']};border:1px solid {t['borde']};border-left:4px solid {t['acento1']};border-radius:18px;padding:1.1rem 1.2rem;margin:.7rem 0;color:{t['texto2']}!important;line-height:1.75;font-size:.88rem;box-shadow:0 4px 20px {t['sombra2']};animation:fadeUp .4s ease}}
.summary-card-title{{color:{t['acento2']}!important;font-size:1rem;font-weight:800;margin-bottom:.5rem;display:flex;align-items:center;gap:.4rem}}

/* ── MÉTRICAS ── */
.metrics-grid{{display:grid;grid-template-columns:1fr 1fr;gap:.5rem;margin:.6rem 0}}
.metric-pill{{background:{t['card']};border:1px solid {t['borde']};border-radius:14px;padding:.75rem 1rem;text-align:center;box-shadow:0 2px 8px {t['sombra2']};transition:transform .2s,box-shadow .2s}}
.metric-pill:hover{{transform:translateY(-2px);box-shadow:0 6px 16px {t['sombra']}}}
.metric-pill-label{{color:{t['texto3']}!important;font-size:.62rem;font-weight:700;text-transform:uppercase;letter-spacing:.08em}}
.metric-pill-value{{color:{t['acento2']}!important;font-size:1.1rem;font-weight:800;margin-top:.2rem;font-family:'JetBrains Mono',monospace}}

/* ── TAGS ── */
.tags-wrap{{display:flex;flex-wrap:wrap;gap:.3rem;margin:.5rem 0}}
.tag{{background:{t['bg2']};color:{t['acento2']}!important;border:1px solid {t['borde']};border-radius:20px;padding:.28rem .7rem;font-size:.7rem;font-weight:600;transition:all .2s}}
.tag:hover{{background:{t['borde']};border-color:{t['acento1']}}}

/* ── HALLAZGO ── */
.hallazgo-card{{background:linear-gradient(135deg,{t['bg2']},{t['card']});border:1px solid {t['borde']};border-left:4px solid {t['acento1']};border-radius:14px;padding:.85rem 1rem;color:{t['texto2']}!important;font-size:.82rem;margin:.6rem 0;line-height:1.65;box-shadow:0 2px 8px {t['sombra2']}}}

/* ── INFO / WARN ── */
.info-box{{background:#f0fdf4;border:1px solid #86efac;border-radius:12px;padding:.75rem 1rem;color:#166534;font-size:.83rem;margin:.5rem 0;display:flex;align-items:center;gap:.5rem}}
.warn-box{{background:#fffbeb;border:1px solid #fcd34d;border-radius:12px;padding:.75rem 1rem;color:#92400e;font-size:.83rem;margin:.5rem 0;display:flex;align-items:center;gap:.5rem}}

/* ── CAMBIOS ── */
.cambio-item{{background:{t['bg2']};border:1px solid {t['borde']};border-radius:10px;padding:.5rem .8rem;margin:.2rem 0;font-size:.78rem;font-family:'JetBrains Mono',monospace;display:flex;align-items:center;gap:.45rem;color:{t['texto']}}}
.cambio-num{{color:{t['texto3']};font-size:.66rem;min-width:1rem}}.cambio-arrow{{color:#d97706}}

/* ── EXPANDER ── */
[data-testid="stExpander"]{{background:{t['card']}!important;border:1px solid {t['borde']}!important;border-radius:14px!important;box-shadow:0 2px 8px {t['sombra2']}!important;overflow:hidden!important}}

/* ── CHAT ── */
[data-testid="stChatInput"] textarea{{background:{t['card2']}!important;border:2px solid {t['borde2']}!important;border-radius:16px!important;color:{t['texto']}!important;font-family:'Inter',sans-serif!important;font-size:.92rem!important;box-shadow:0 2px 8px {t['sombra2']}!important}}
[data-testid="stChatInput"] textarea:focus{{border-color:{t['acento1']}!important;box-shadow:0 0 0 3px {t['sombra']}!important}}[data-testid="stBottom"]{{background:linear-gradient(transparent,{t['bg1']} 40%)!important;padding-top:1rem!important}}
[data-testid="stChatMessageContent"]{{font-size:.88rem!important;line-height:1.65!important}}
[data-testid="stChatMessage"]{{background:{t['card']}!important;border:1px solid {t['borde']}!important;border-radius:14px!important;padding:.7rem!important;margin:.3rem 0!important;box-shadow:0 2px 6px {t['sombra2']}!important}}

/* ── INPUTS ── */
.stTextInput>div>div>input{{background:{t['card']}!important;border:1.5px solid {t['borde']}!important;border-radius:12px!important;color:{t['texto']}!important;font-family:'Inter',sans-serif!important;font-size:.88rem!important;padding:.6rem .9rem!important}}
.stTextInput>div>div>input:focus{{border-color:{t['acento1']}!important;box-shadow:0 0 0 3px {t['sombra']}!important}}
.stSelectbox>div>div{{background:{t['card']}!important;border:1.5px solid {t['borde']}!important;border-radius:12px!important;color:{t['texto']}!important}}

/* ── DIVIDER ── */
.oro-divider{{height:1px;background:linear-gradient(90deg,transparent,{t['borde2']},transparent);margin:.9rem 0}}

/* ── SECTION TITLE ── */
.section-title{{color:{t['texto']}!important;font-size:.95rem;font-weight:700;margin:.9rem 0 .4rem;display:flex;align-items:center;gap:.4rem}}

/* ── CHAT PLACEHOLDER ── */
.chat-placeholder{{text-align:center;padding:1.2rem .8rem;background:{t['card']};border:1.5px dashed {t['borde2']};border-radius:18px;margin:.4rem 0;box-shadow:0 2px 8px {t['sombra2']}}}
.chip{{display:inline-block;background:{t['bg2']};border:1px solid {t['borde']};border-radius:20px;padding:.22rem .65rem;font-size:.68rem;color:{t['acento1']};font-weight:600;margin:.15rem;cursor:default;transition:all .2s}}
.chip:hover{{background:{t['borde']};border-color:{t['acento1']}}}[data-testid="stChatInputContainer"]{{padding:.6rem 0!important}}[data-testid="stChatInput"]{{margin-top:.4rem!important}}

/* ── EMPTY STATE ── */
.empty-state{{text-align:center;padding:3rem 1rem;animation:fadeUp .5s ease}}
.empty-icon{{font-size:3.5rem;margin-bottom:.8rem;filter:drop-shadow(0 4px 12px {t['sombra']})}}
.empty-title{{color:{t['texto']};font-size:1rem;font-weight:700}}
.empty-hint{{color:{t['texto3']};font-size:.78rem;margin-top:.4rem;line-height:1.7}}
.format-badges{{display:flex;justify-content:center;gap:.4rem;margin-top:.8rem;flex-wrap:wrap}}
.format-badge{{background:{t['bg2']};border:1px solid {t['borde']};border-radius:8px;padding:.22rem .6rem;color:{t['texto3']};font-size:.7rem;font-family:'JetBrains Mono',monospace;font-weight:600}}

/* ── FOOTER ── */
.oro-footer{{text-align:center;font-size:.68rem;color:{t['texto3']};padding:.5rem 0;opacity:.7}}

/* ── SIDEBAR ── */
[data-testid="stSidebar"]{{background:linear-gradient(180deg,{t['card']},{t['bg2']})!important;border-right:1px solid {t['borde']}!important}}
[data-testid="stSidebar"] .stMarkdown,
[data-testid="stSidebar"] .stCaption{{color:{t['texto2']}!important}}
</style>"""


# ══════════════════════════════════════════════════════════════
# SESSION STATE
# ══════════════════════════════════════════════════════════════
_defaults = {
    "texto_extraido":"","nombre_archivo":"","archivo_bytes":None,"resumen_data":None,
    "historial_chat":[],"cambios_aplicados":None,"archivo_tipo":"","lista_cambios":[],
    "texto_modificado":"","generando_resumen":False,"resumen_error":False,
    "preview_cambio":None,"edicion_counter":0,"texto_corregido":"",
    "guia_paso":0,"guia_vista":False,"ejecutar_evaluacion":False,
    "tema":"noche","modo_entrada":"archivo",
    "imagen_archivo_bytes":None,"imagen_archivo_nombre":"","imagen_archivo_mime":"",
    "historial_versiones":[],"buscar_query":"",
    "resultado_evaluacion":None,
    "scroll_to":None,
    "idioma":"es",
    "vista_activa":None,
}
for k,v in _defaults.items():
    if k not in st.session_state:
        st.session_state[k] = v

st.markdown(_get_all_css(st.session_state.get("tema","noche")), unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════
# TRADUCCIONES
# ══════════════════════════════════════════════════════════════
_TXT = {
    "es": {
        "analizar":"⚡ Analizar","evaluar":"🔍 Evaluar","ver_doc":"👁 Ver documento",
        "analizar_doc":"⚡ Analizar documento","evaluar_doc":"🔍 Evaluar documento ahora",
        "exportar":"📥 Exportar informe","regen":"🔄 Regenerar resumen",
        "confirmar":"✅ Confirmar","cancelar":"❌ Cancelar",
        "descargar_corregido":"📥 Descargar corregido",
        "limpiar_cambios":"🗑️ Limpiar todos los cambios",
        "chat_placeholder":"✍️ Escribe un cambio o una pregunta...",
        "chat_titulo":"Conversa sobre el documento",
        "chat_hint":"Edita, pregunta o pide cambios en lenguaje natural",
        "chip1":"✏️ cambia X por Y","chip2":"➕ agrega dato a persona",
        "chip3":"❓ ¿cuántos hay?","chip4":"📝 resume en 3 puntos",
        "analizando":"🧠 Analizando...","evaluando":"🔎 Evaluando calidad...",
        "procesando":"🔍 Procesando...","pensando":"🤔 Pensando...",
        "cargando":"📖 Cargando...","interpretando":"🧠 Leyendo...",
        "reintentar":"🔄 Reintentar","cerrar_eval":"✕ Cerrar evaluación",
        "hallazgo":"💡 Hallazgo:","recomendacion":"💡 Recomendación:",
        "antes":"Antes","despues":"Después",
        "version":"↩️ Deshacer último cambio",
        "idioma_doc":"El programa responde en español sin importar el idioma del documento.",
        "subir":"📎 Sube tu archivo","foto":"📷 Foto de documento",
        "interpretar":"🔍 Interpretar","archivo":"📎 Archivo",
        "palabras":"palabras","cambios":"cambio(s)","versiones":"versión(es)",
        "no_encontrado":"no encontrado en el documento",
        "listo_analizar":"Toca ⚡ Analizar para generar el resumen inteligente",
        "sin_problemas":"¡Sin problemas detectados! El documento está bien 🎉",
        "prompt_idioma":"Responde SIEMPRE en español, sin importar el idioma del documento.",
        "word":"📄 Word","excel":"📊 Excel","pdf":"📕 PDF",
        "word_corregido":"📄 Word corregido","excel_corregido":"📊 Excel corregido",
        "pdf_corregido":"📕 PDF corregido","exportar_word":"📄 Exportar como Word",
        "cambiar_tema":"🎨 Tema","cambiar_idioma":"🌐 EN",
    },
    "en": {
        "analizar":"⚡ Analyze","evaluar":"🔍 Evaluate","ver_doc":"👁 View document",
        "analizar_doc":"⚡ Analyze document","evaluar_doc":"🔍 Evaluate document now",
        "exportar":"📥 Export report","regen":"🔄 Regenerate summary",
        "confirmar":"✅ Confirm","cancelar":"❌ Cancel",
        "descargar_corregido":"📥 Download corrected",
        "limpiar_cambios":"🗑️ Clear all changes",
        "chat_placeholder":"✍️ Write a change or ask a question...",
        "chat_titulo":"Chat about the document",
        "chat_hint":"Edit, ask questions or request changes in natural language",
        "chip1":"✏️ change X to Y","chip2":"➕ add data to person",
        "chip3":"❓ how many are there?","chip4":"📝 summarize in 3 points",
        "analizando":"🧠 Analyzing...","evaluando":"🔎 Evaluating quality...",
        "procesando":"🔍 Processing...","pensando":"🤔 Thinking...",
        "cargando":"📖 Loading...","interpretando":"🧠 Reading...",
        "reintentar":"🔄 Retry","cerrar_eval":"✕ Close evaluation",
        "hallazgo":"💡 Finding:","recomendacion":"💡 Recommendation:",
        "antes":"Before","despues":"After",
        "version":"↩️ Undo last change",
        "idioma_doc":"The program responds in English regardless of the document language.",
        "subir":"📎 Upload your file","foto":"📷 Photo of document",
        "interpretar":"🔍 Interpret","archivo":"📎 File",
        "palabras":"words","cambios":"change(s)","versiones":"version(s)",
        "no_encontrado":"not found in document",
        "listo_analizar":"Tap ⚡ Analyze to generate the smart summary",
        "sin_problemas":"No issues detected! The document looks good 🎉",
        "prompt_idioma":"Always respond in English, regardless of the document language.",
        "word":"📄 Word","excel":"📊 Excel","pdf":"📕 PDF",
        "word_corregido":"📄 Corrected Word","excel_corregido":"📊 Corrected Excel",
        "pdf_corregido":"📕 Corrected PDF","exportar_word":"📄 Export as Word",
        "cambiar_tema":"🎨 Theme","cambiar_idioma":"🌐 ES",
    },
}

def T(key):
    """Get translated text."""
    lang = st.session_state.get("idioma","es")
    return _TXT.get(lang,_TXT["es"]).get(key, _TXT["es"].get(key,""))

# ══════════════════════════════════════════════════════════════
# GEMINI
# ══════════════════════════════════════════════════════════════
try:
    LLAVE_GEMINI = st.secrets["LLAVE_GEMINI"]
    genai.configure(api_key=LLAVE_GEMINI)
except Exception as e:
    st.error(f"🔑 Error configurando la IA: {e}"); st.stop()

MODELOS_FALLBACK = ["gemini-3.1-flash-lite-preview","gemini-3.1-flash-preview","gemini-3.1-pro-preview"]

def llamar_ia(prompt, es_json=False):
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

# ══════════════════════════════════════════════════════════════
# SIDEBAR
# ══════════════════════════════════════════════════════════════
with st.sidebar:
    st.markdown(f"<div style='text-align:center;padding:.5rem 0'><span style='font-size:1.5rem'>🏆</span><div style='font-size:.9rem;font-weight:800;color:#34d399'>Oro Asistente</div></div>", unsafe_allow_html=True)
    st.markdown("---")
    st.markdown("**🎨 Tema**")
    tema_opc = {"☀️ Claro":"claro","🌸 Aurora":"aurora","🌿 Menta":"menta","🌅 Sol":"sol","🌹 Rose":"rose","🌑 Noche":"noche","⬛ Carbón":"carbon","🌌 Cosmos":"cosmos"}
    sel = st.selectbox("Tema", list(tema_opc.keys()),
        index=list(tema_opc.values()).index(st.session_state.get("tema","noche")),
        label_visibility="collapsed")
    if tema_opc[sel] != st.session_state.tema:
        st.session_state.tema = tema_opc[sel]; st.rerun()
    st.markdown("---")
    paso_sb = st.session_state.get("guia_paso",0)
    if paso_sb > 0 and not st.session_state.get("guia_vista",False):
        guias_sb = {
            1:("🎉","Paso 1 — Analiza","Tu archivo está listo.\n\nToca **⚡ Analizar** para que la IA extraiga el resumen, métricas y puntos clave."),
            2:("📊","Paso 2 — Revisa","El resumen está arriba.\n\nDescárgalo como Word, Excel o PDF con los botones de exportación."),
            3:("💬","Paso 3 — Edita","Usa el chat para editar:\n• *cambia X por Y*\n• *agrega el teléfono a Juan*\n• *¿cuántas personas hay?*"),
        }
        if paso_sb in guias_sb:
            ico_sb,tit_sb,desc_sb = guias_sb[paso_sb]
            st.markdown(f"**{ico_sb} {tit_sb}**")
            st.info(desc_sb)
            c1sb,c2sb = st.columns(2)
            with c1sb:
                if st.button("👍 Ok",use_container_width=True,key="guia_ok"):
                    st.session_state.guia_paso = 0 if paso_sb>=3 else paso_sb+1
                    if paso_sb>=3: st.session_state.guia_vista=True
                    st.rerun()
            with c2sb:
                if st.button("✕ Saltar",use_container_width=True,key="guia_skip"):
                    st.session_state.guia_vista=True; st.session_state.guia_paso=0; st.rerun()
            st.caption(f"Paso {paso_sb} de 3")
    elif st.session_state.get("texto_extraido"):
        st.markdown("**💡 Comandos útiles**")
        st.markdown(f"""<div style='font-size:.78rem;color:#6ee7b7;line-height:2'>
            ✏️ cambia X por Y<br>
            ➕ agrega dato a persona<br>
            🔍 ¿dónde aparece X?<br>
            📝 resume en 3 puntos<br>
            📊 ¿cuántos registros hay?
        </div>""", unsafe_allow_html=True)
        if st.session_state.get("historial_versiones"):
            st.markdown("---")
            st.markdown("**⏮ Versiones**")
            st.caption(f"{len(st.session_state.historial_versiones)} versión(es) guardada(s)")
            if st.button(T("version"), use_container_width=True):
                v = st.session_state.historial_versiones.pop()
                st.session_state.texto_corregido = v["texto"]
                st.session_state.cambios_aplicados = v["bytes"]
                if st.session_state.lista_cambios:
                    st.session_state.lista_cambios.pop()
                st.session_state.resumen_data = None
                st.rerun()
    else:
        st.markdown("**👋 Bienvenido**")
        st.info("Sube un archivo Word, Excel o PDF — o una **foto** de un documento — para empezar.")
    st.markdown("---")
    st.caption("Oro Asistente v3.1")

# ══════════════════════════════════════════════════════════════
# TOP BAR — tema izquierda, idioma derecha
# ══════════════════════════════════════════════════════════════
_col_tema, _col_mid, _col_lang = st.columns([2, 3, 1])

with _col_tema:
    _temas_map = {"🌑 Noche":"noche","⬛ Carbón":"carbon","🌌 Cosmos":"cosmos"}
    _tema_actual = st.session_state.get("tema","noche")
    _tema_label = next((k for k,v in _temas_map.items() if v==_tema_actual), "☀️ Claro")
    _nuevo_tema = st.selectbox(
        "tema", list(_temas_map.keys()),
        index=list(_temas_map.keys()).index(_tema_label),
        label_visibility="collapsed", key="tema_sel")
    if _temas_map[_nuevo_tema] != _tema_actual:
        st.session_state.tema = _temas_map[_nuevo_tema]; st.rerun()

with _col_lang:
    _lang_actual = st.session_state.get("idioma","es")
    _lang_label = "🌐 EN" if _lang_actual=="es" else "🌐 ES"
    if st.button(_lang_label, key="btn_lang", use_container_width=True):
        st.session_state.idioma = "en" if _lang_actual=="es" else "es"
        st.rerun()

# ── Header ──
st.markdown("""
<div class="oro-header">
    <div class="oro-logo-wrap">
        <div class="oro-logo-ring"></div>
        <span class="oro-logo">🏆</span>
    </div>
    <div class="oro-title">Oro Asistente</div>
    <div class="oro-badge">✦ ANALIZA &nbsp;·&nbsp; EDITA &nbsp;·&nbsp; EXPORTA ✦</div>
</div>
""", unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════
# FUNCIONES UTILITARIAS
# ══════════════════════════════════════════════════════════════
def extraer_json_seguro(texto, es_lista=False):
    t = texto.replace("```json","").replace("```","").strip()
    c1,c2 = ("[","]") if es_lista else ("{","}")
    ini=t.find(c1); fin=t.rfind(c2)+1
    if ini!=-1 and fin>0:
        try: return json.loads(t[ini:fin],strict=False)
        except:
            try: return ast.literal_eval(t[ini:fin])
            except: pass
    return None


def _scroll_to(anchor_id):
    """Hace scroll suave al elemento con ese id."""
    st.markdown(f'''<script>
        window.parent.document.getElementById("{anchor_id}") &&
        window.parent.document.getElementById("{anchor_id}").scrollIntoView({{behavior:"smooth",block:"start"}});
    </script>''', unsafe_allow_html=True)

def guardar_version(texto, bytes_doc):
    """Guarda snapshot antes de aplicar un cambio."""
    st.session_state.historial_versiones.append({
        "texto": texto,
        "bytes": bytes_doc,
        "ts": datetime.now().strftime("%H:%M:%S")
    })
    # Máximo 10 versiones
    if len(st.session_state.historial_versiones) > 10:
        st.session_state.historial_versiones.pop(0)

# ══════════════════════════════════════════════════════════════
# FUNCIONES IA
# ══════════════════════════════════════════════════════════════
def solicitar_resumen_estructurado(texto):
    idioma_prompt = T("prompt_idioma")
    prompt = (
        f"Analista experto en documentos. {idioma_prompt} Devuelve SOLO JSON válido:\n"
        '{"titulo":"...","emoji_categoria":"📋","resumen_ejecutivo":"max 3 oraciones amigables",'
        '"metricas":{"Clave":"Valor"},"puntos_clave":["punto"],"hallazgo_destacado":"observación"}\n\n'
        f"DOCUMENTO:\n{texto[:12000]}"
    )
    r = llamar_ia(prompt)
    return extraer_json_seguro(r) if r else None

def extraer_cambio_con_regex(instruccion):
    patrones = [
        r"(?:cambia|reemplaza|sustituye|cambie)\s+['\"]?(.+?)['\"]?\s+(?:por|con|a)\s+['\"]?(.+?)['\"]?\s*$",
        r"['\"](.+?)['\"]\s*(?:→|->|=>|por|con)\s*['\"]?(.+?)['\"]?\s*$",
        r"(.+?)\s*(?:→|->|=>)\s*(.+)",
    ]
    for pat in patrones:
        m = re.search(pat,instruccion.strip(),re.IGNORECASE)
        if m:
            b=m.group(1).strip().strip("'\"")
            r2=m.group(2).strip().strip("'\"")
            if b and r2: return [{"buscar":b,"reemplazar":r2}]
    return []

def solicitar_cambios(instruccion, texto_doc=""):
    ctx = f"\n\nCONTENIDO DEL DOCUMENTO:\n{texto_doc[:4000]}" if texto_doc else ""
    idioma_prompt = T("prompt_idioma")
    prompt = (
        f"Asistente de edición de documentos. {idioma_prompt}\nINSTRUCCIÓN: \"{instruccion}\"{ctx}\n\n"
        "REGLAS:\n1. cambia X por Y → buscar=X, reemplazar=Y\n"
        "2. agrega DATO a PERSONA → buscar=PERSONA exacto, reemplazar='PERSONA DATO'\n"
        "3. completa campo de X con Y → buscar=X, reemplazar='X Y'\n"
        "4. SIEMPRE usa texto EXACTO del doc como buscar\n"
        "5. Si pide formato (negrita/mayúsculas/cursiva) → agrega campo 'formato'\n"
        "Responde SOLO JSON array:\n"
        '[{"buscar":"texto_exacto","reemplazar":"texto_nuevo"}]'
    )
    r = llamar_ia(prompt)
    if r:
        res = extraer_json_seguro(r,es_lista=True)
        if res and isinstance(res,list):
            v = [c for c in res if isinstance(c,dict) and "buscar" in c and "reemplazar" in c
                 and str(c["buscar"]).strip() and str(c["reemplazar"]).strip() and c["buscar"]!=c["reemplazar"]]
            if v: return v
    return extraer_cambio_con_regex(instruccion)

def preguntar_al_documento(pregunta, texto):
    ctx = "\n".join([f"{m['rol']}: {m['texto']}" for m in st.session_state.historial_chat[-6:]])
    idioma_prompt = T("prompt_idioma")
    prompt = (
        f"Asistente experto en documentos. {idioma_prompt}\nDOCUMENTO:\n{texto[:10000]}\n\n"
        f"CONVERSACIÓN:\n{ctx}\n\nPREGUNTA: {pregunta}\nResponde conciso y directo en español."
    )
    return llamar_ia(prompt) or "No pude procesar tu pregunta."

def detectar_anomalias(texto):
    idioma_prompt = T("prompt_idioma")
    prompt = (
        f"Analiza el documento. {idioma_prompt} Devuelve SOLO JSON:\n"
        '{"nivel_general":"Excelente/Bueno/Regular/Deficiente","puntaje":85,'
        '"criticos":["..."],"altos":["..."],"medios":["..."],"leves":["..."],'
        '"recomendacion":"..."}\n\n'
        f"DOCUMENTO:\n{texto[:12000]}"
    )
    r = llamar_ia(prompt)
    return extraer_json_seguro(r) if r else None

def detectar_tipo_imagen(texto_raw):
    """Detecta si el contenido extraído es tabla o documento de texto."""
    lineas = [l for l in texto_raw.split('\n') if l.strip()]
    if not lineas: return "word"
    # Detectar tabla: muchas líneas con 2+ columnas separadas por espacios/tabs
    lineas_con_cols = sum(1 for l in lineas if len(re.split(r'\s{2,}|\t', l.strip())) >= 2)
    ratio_tabla = lineas_con_cols / max(len(lineas), 1)
    # Si más del 50% de las líneas tienen múltiples columnas → probablemente tabla
    if ratio_tabla >= 0.5:
        return "excel"
    return "word"

def interpretar_imagen_documento(imagen_bytes, mime_type="image/jpeg", formato_salida="auto"):
    """
    Extrae texto de imagen con Tesseract (gratis) o Gemini Vision (fallback).
    Si formato_salida="auto", detecta automáticamente si es tabla o documento.
    Preserva el contenido tal como aparece en la imagen.
    """
    texto_raw = None

    # Intentar Tesseract primero (sin tokens)
    try:
        from PIL import Image, ImageEnhance, ImageFilter
        import pytesseract, io
        img = Image.open(io.BytesIO(imagen_bytes))
        # Preprocesar para mejor OCR
        img_gray = img.convert('L')
        img_contrast = ImageEnhance.Contrast(img_gray).enhance(2.0)
        img_sharp = img_contrast.filter(ImageFilter.SHARPEN)
        texto_raw = pytesseract.image_to_string(img_sharp, lang='spa+eng', config='--psm 6')
        if not texto_raw.strip():
            texto_raw = None
    except ImportError:
        texto_raw = None
    except Exception:
        texto_raw = None

    # Fallback: Gemini Vision solo con modelo lite
    if not texto_raw:
        try:
            import base64
            img_b64 = base64.b64encode(imagen_bytes).decode("utf-8")
            model = genai.GenerativeModel("gemini-3.1-flash-lite-preview")
            prompt = (
                "Eres un OCR experto. Extrae TODO el contenido de esta imagen EXACTAMENTE como aparece.\n"
                "REGLAS IMPORTANTES:\n"
                "- Copia el texto tal como está, sin cambiar nada\n"
                "- Si hay una tabla: reproduce cada fila con columnas separadas por | (pipe)\n"
                "- Si hay un documento/oficio: copia párrafo por párrafo\n"
                "- Preserva la estructura: encabezados, secciones, filas\n"
                "- Si hay texto ilegible escribe [ilegible]\n"
                "Devuelve SOLO el contenido, sin explicaciones."
            )
            resp = model.generate_content([
                {"mime_type": mime_type, "data": img_b64},
                prompt
            ])
            texto_raw = resp.text.strip()
        except Exception:
            return None, "word"

    if not texto_raw:
        return None, "word"

    # Auto-detectar tipo si no se especificó
    if formato_salida == "auto":
        tipo_detectado = detectar_tipo_imagen(texto_raw)
    else:
        tipo_detectado = formato_salida

    # Formatear según tipo
    lineas = [l for l in texto_raw.split('\n') if l.strip()]
    if tipo_detectado == "excel":
        # Normalizar columnas usando separación por espacios múltiples o pipes
        filas_formateadas = []
        for linea in lineas:
            if '|' in linea:
                cols = [c.strip() for c in linea.split('|') if c.strip()]
            else:
                cols = re.split(r'\s{2,}|\t', linea.strip())
                cols = [c.strip() for c in cols if c.strip()]
            if cols:
                filas_formateadas.append(' | '.join(cols))
        return '\n'.join(filas_formateadas), "excel"
    else:
        return '\n'.join(lineas), "word"

# ══════════════════════════════════════════════════════════════
# EXPORTADORES
# ══════════════════════════════════════════════════════════════
def exportar_word(texto, resumen_data=None, archivo_bytes=None, archivo_tipo=None, cambios=None):
    zona=pytz.timezone('America/Caracas'); fecha=datetime.now(zona).strftime('%d de %B de %Y, %I:%M %p')
    cambios=cambios or []
    if archivo_tipo=="docx" and archivo_bytes and cambios:
        r,_=reemplazar_docx_preservando_formato(archivo_bytes,cambios); return r
    if archivo_tipo=="xlsx" and archivo_bytes:
        doc=Document(); doc.styles['Normal'].font.name='Calibri'
        zona2=pytz.timezone('America/Caracas'); fecha2=datetime.now(zona2).strftime('%d de %B de %Y, %I:%M %p')
        th_x=doc.add_table(rows=1,cols=1); th_x.style='Table Grid'
        ch_x=th_x.cell(0,0); ch_x.paragraphs[0].clear()
        tit_x=resumen_data.get("titulo","Reporte desde Excel") if resumen_data else "Reporte desde Excel"
        rh_x=ch_x.paragraphs[0].add_run(tit_x)
        rh_x.font.bold=True; rh_x.font.size=Pt(16); rh_x.font.color.rgb=RGBColor(0xFF,0xFF,0xFF)
        ch_x.paragraphs[0].alignment=WD_ALIGN_PARAGRAPH.CENTER
        tcp_x=ch_x._tc.get_or_add_tcPr(); shd_x=OxmlElement('w:shd')
        shd_x.set(qn('w:val'),'clear'); shd_x.set(qn('w:color'),'auto'); shd_x.set(qn('w:fill'),'1E3A5F'); tcp_x.append(shd_x)
        doc.add_paragraph()
        pf_x=doc.add_paragraph(); rf_x=pf_x.add_run(f'Generado: {fecha2}')
        rf_x.font.size=Pt(9); rf_x.font.color.rgb=RGBColor(0x6B,0x72,0x80); pf_x.alignment=WD_ALIGN_PARAGRAPH.RIGHT
        if resumen_data and resumen_data.get("resumen_ejecutivo"):
            doc.add_paragraph()
            tr_x=doc.add_table(rows=1,cols=1); tr_x.style='Table Grid'
            cr_x=tr_x.cell(0,0); cr_x.paragraphs[0].clear()
            rr_x=cr_x.paragraphs[0].add_run(resumen_data["resumen_ejecutivo"])
            rr_x.font.size=Pt(10); rr_x.font.italic=True
            tp2_x=cr_x._tc.get_or_add_tcPr(); sh2_x=OxmlElement('w:shd')
            sh2_x.set(qn('w:val'),'clear'); sh2_x.set(qn('w:color'),'auto'); sh2_x.set(qn('w:fill'),'EFF6FF'); tp2_x.append(sh2_x)
        doc.add_paragraph()
        bu=archivo_bytes
        if cambios: bu,_=reemplazar_xlsx_preservando_formato(archivo_bytes,cambios)
        wb=openpyxl.load_workbook(BytesIO(bu),data_only=True)
        for sheet in wb.worksheets:
            doc.add_heading(f'Hoja: {sheet.title}',level=1)
            filas=[f for f in sheet.iter_rows(values_only=True) if any(c is not None for c in f)]
            if not filas: doc.add_paragraph('(vacía)'); continue
            nc=max(len(f) for f in filas); tb=doc.add_table(rows=len(filas),cols=nc); tb.style='Table Grid'
            for i,fila in enumerate(filas):
                for j in range(nc):
                    v=fila[j] if j<len(fila) else ""
                    cell=tb.cell(i,j); cell.text=str(v) if v is not None else ""
                    if i==0:
                        for run in cell.paragraphs[0].runs: run.font.bold=True
            doc.add_paragraph()
        buf=BytesIO(); doc.save(buf); return buf.getvalue()
    doc=Document()
    for sec in doc.sections:
        sec.top_margin=Inches(0.8); sec.bottom_margin=Inches(0.8); sec.left_margin=Inches(1.0); sec.right_margin=Inches(1.0)
    doc.styles['Normal'].font.name='Calibri'; doc.styles['Normal'].font.size=Pt(11)
    th=doc.add_table(rows=1,cols=1); th.style='Table Grid'; ch=th.cell(0,0); ch.paragraphs[0].clear()
    rh=ch.paragraphs[0].add_run(resumen_data.get("titulo","INFORME") if resumen_data else "INFORME")
    rh.font.bold=True; rh.font.size=Pt(16); rh.font.color.rgb=RGBColor(0xFF,0xFF,0xFF)
    ch.paragraphs[0].alignment=WD_ALIGN_PARAGRAPH.CENTER
    tcp=ch._tc.get_or_add_tcPr(); shd=OxmlElement('w:shd')
    shd.set(qn('w:val'),'clear'); shd.set(qn('w:color'),'auto'); shd.set(qn('w:fill'),'1E3A5F'); tcp.append(shd)
    doc.add_paragraph()
    pf=doc.add_paragraph(); rf=pf.add_run(f'Generado: {fecha}')
    rf.font.size=Pt(9); rf.font.color.rgb=RGBColor(0x6B,0x72,0x80); pf.alignment=WD_ALIGN_PARAGRAPH.RIGHT
    doc.add_paragraph()
    if resumen_data:
        if resumen_data.get("resumen_ejecutivo"):
            tr=doc.add_table(rows=1,cols=1); tr.style='Table Grid'; cr=tr.cell(0,0); cr.paragraphs[0].clear()
            rr=cr.paragraphs[0].add_run(resumen_data["resumen_ejecutivo"]); rr.font.size=Pt(10); rr.font.italic=True
            tp2=cr._tc.get_or_add_tcPr(); sh2=OxmlElement('w:shd')
            sh2.set(qn('w:val'),'clear'); sh2.set(qn('w:color'),'auto'); sh2.set(qn('w:fill'),'EFF6FF'); tp2.append(sh2)
            doc.add_paragraph()
        if resumen_data.get("metricas"):
            h2=doc.add_heading('Métricas Clave',level=1); h2.runs[0].font.color.rgb=RGBColor(0x1E,0x40,0xAF)
            tm=doc.add_table(rows=1,cols=2); tm.style='Table Grid'; hdr=tm.rows[0].cells
            for ci,txt in enumerate(['Indicador','Valor']):
                hdr[ci].paragraphs[0].clear(); r2=hdr[ci].paragraphs[0].add_run(txt)
                r2.font.bold=True; r2.font.color.rgb=RGBColor(0xFF,0xFF,0xFF)
                hdr[ci].paragraphs[0].alignment=WD_ALIGN_PARAGRAPH.CENTER
                tph=hdr[ci]._tc.get_or_add_tcPr(); shh=OxmlElement('w:shd')
                shh.set(qn('w:val'),'clear'); shh.set(qn('w:color'),'auto'); shh.set(qn('w:fill'),'1E40AF'); tph.append(shh)
            for idx,(k,v) in enumerate(resumen_data["metricas"].items()):
                rm=tm.add_row().cells; rm[0].text=str(k); rm[1].text=str(v)
                fill='F8FAFC' if idx%2==0 else 'FFFFFF'
                for ci2 in range(2):
                    tpd=rm[ci2]._tc.get_or_add_tcPr(); shd2=OxmlElement('w:shd')
                    shd2.set(qn('w:val'),'clear'); shd2.set(qn('w:color'),'auto'); shd2.set(qn('w:fill'),fill); tpd.append(shd2)
            doc.add_paragraph()
        if resumen_data.get("puntos_clave"):
            h3=doc.add_heading('Puntos Clave',level=1); h3.runs[0].font.color.rgb=RGBColor(0x1E,0x40,0xAF)
            for p in resumen_data["puntos_clave"]: doc.add_paragraph(style='List Bullet').add_run(p).font.size=Pt(11)
        if resumen_data.get("hallazgo_destacado"):
            doc.add_paragraph()
            h4=doc.add_heading('💡 Hallazgo',level=1); h4.runs[0].font.color.rgb=RGBColor(0x1E,0x40,0xAF)
            th2=doc.add_table(rows=1,cols=1); th2.style='Table Grid'; ch2=th2.cell(0,0); ch2.paragraphs[0].clear()
            rh2=ch2.paragraphs[0].add_run(resumen_data["hallazgo_destacado"]); rh2.font.italic=True; rh2.font.size=Pt(10)
            tph2=ch2._tc.get_or_add_tcPr(); shh2=OxmlElement('w:shd')
            shh2.set(qn('w:val'),'clear'); shh2.set(qn('w:color'),'auto'); shh2.set(qn('w:fill'),'F0FDF4'); tph2.append(shh2)
        doc.add_page_break()
    hc=doc.add_heading('Contenido del Documento',level=1); hc.runs[0].font.color.rgb=RGBColor(0x1E,0x40,0xAF)
    for linea in texto.split('\n'):
        ll=linea.strip().replace('*','').replace('#','')
        if ll: p=doc.add_paragraph(ll); p.paragraph_format.space_after=Pt(2)
    buf=BytesIO(); doc.save(buf); return buf.getvalue()

def exportar_excel(texto, resumen_data=None, archivo_bytes=None, archivo_tipo=None, cambios=None):
    cambios=cambios or []
    if archivo_tipo=="xlsx" and archivo_bytes:
        if cambios:
            r,_=reemplazar_xlsx_preservando_formato(archivo_bytes,cambios); return r
        return archivo_bytes  # devolver original si no hay cambios
    if archivo_tipo=="docx" and archivo_bytes:
        bu=archivo_bytes
        if cambios: bu,_=reemplazar_docx_preservando_formato(archivo_bytes,cambios)
        wb=openpyxl.Workbook(); wb.remove(wb.active); doc_src=Document(BytesIO(bu))
        for i,tabla in enumerate(doc_src.tables):
            ws=wb.create_sheet(title=f"Tabla_{i+1}"); fl=[]
            for row in tabla.rows:
                vis=set(); fila=[]
                for cell in row.cells:
                    if cell._tc not in vis: vis.add(cell._tc); fila.append(cell.text.strip())
                if any(fila): fl.append(fila)
            for ri,fila in enumerate(fl,1):
                for ci,val in enumerate(fila,1):
                    c=ws.cell(row=ri,column=ci,value=val)
                    if ri==1: c.fill=PatternFill("solid",fgColor="1E3A5F"); c.font=Font(color="FFFFFF",bold=True,size=10)
                    else: c.fill=PatternFill("solid",fgColor="F8FAFC" if ri%2==0 else "FFFFFF"); c.font=Font(size=10)
                    c.alignment=Alignment(wrap_text=True,vertical="center")
            for col in ws.columns: ws.column_dimensions[col[0].column_letter].width=22
        if not wb.sheetnames: ws=wb.create_sheet("Datos"); ws.cell(1,1,"Sin tablas detectadas.")
        buf=BytesIO(); wb.save(buf); return buf.getvalue()
    wb=openpyxl.Workbook(); AO="1E3A5F"; AM="2563EB"; AC="DBEAFE"; BL="FFFFFF"; GC="F8FAFC"
    def hc(ws,row,col,txt,bg=AO,fg=BL,sz=12,bold=True):
        c=ws.cell(row=row,column=col,value=txt); c.fill=PatternFill("solid",fgColor=bg)
        c.font=Font(color=fg,bold=bold,size=sz); c.alignment=Alignment(horizontal="center",vertical="center",wrap_text=True); return c
    def dc(ws,row,col,txt,bg=BL,bold=False,align="left"):
        c=ws.cell(row=row,column=col,value=txt); c.fill=PatternFill("solid",fgColor=bg)
        c.font=Font(bold=bold,size=11); c.alignment=Alignment(horizontal=align,vertical="center",wrap_text=True); return c
    thin=Border(left=Side(style='thin',color='CBD5E1'),right=Side(style='thin',color='CBD5E1'),top=Side(style='thin',color='CBD5E1'),bottom=Side(style='thin',color='CBD5E1'))
    ws=wb.active; ws.title="Resumen"
    zona=pytz.timezone('America/Caracas'); fecha=datetime.now(zona).strftime('%d/%m/%Y %I:%M %p')
    ws.merge_cells("A1:D1"); hc(ws,1,1,"ORO ASISTENTE - REPORTE",bg=AO,sz=14); ws.row_dimensions[1].height=40
    ws.merge_cells("A2:D2"); dc(ws,2,1,f"Generado: {fecha}",bg=AC,align="center")
    fila=4
    if resumen_data:
        td=resumen_data.get("titulo","Sin título")
        ws.merge_cells(f"A{fila}:D{fila}"); hc(ws,fila,1,td,bg=AM,sz=12); ws.row_dimensions[fila].height=30; fila+=1
        re2=resumen_data.get("resumen_ejecutivo","")
        if re2:
            ws.merge_cells(f"A{fila}:D{fila+2}"); c2=ws.cell(row=fila,column=1,value=re2)
            c2.fill=PatternFill("solid",fgColor="EFF6FF"); c2.alignment=Alignment(horizontal="left",vertical="center",wrap_text=True)
            c2.font=Font(italic=True,size=11); ws.row_dimensions[fila].height=60; fila+=3
        if resumen_data.get("metricas"):
            fila+=1; ws.merge_cells(f"A{fila}:D{fila}"); hc(ws,fila,1,"MÉTRICAS",bg="1E40AF",sz=11); fila+=1
            hc(ws,fila,1,"Indicador",bg=AC,fg="1E3A5F",sz=10); hc(ws,fila,2,"Valor",bg=AC,fg="1E3A5F",sz=10)
            ws.merge_cells(f"C{fila}:D{fila}"); fila+=1
            for idx,(k,v) in enumerate(resumen_data["metricas"].items()):
                bg=GC if idx%2==0 else BL; dc(ws,fila,1,k,bg=bg,bold=True)
                ws.merge_cells(f"B{fila}:C{fila}"); dc(ws,fila,2,str(v),bg=bg,align="center")
                for c3 in range(1,4): ws.cell(row=fila,column=c3).border=thin
                fila+=1
        if resumen_data.get("puntos_clave"):
            fila+=1; ws.merge_cells(f"A{fila}:D{fila}"); hc(ws,fila,1,"PUNTOS CLAVE",bg="1E40AF",sz=11); fila+=1
            for i,p in enumerate(resumen_data["puntos_clave"],1):
                ws.merge_cells(f"A{fila}:D{fila}"); c4=ws.cell(row=fila,column=1,value=f"{i}. {p}")
                c4.fill=PatternFill("solid",fgColor=GC if i%2==0 else BL); c4.font=Font(size=11)
                c4.alignment=Alignment(horizontal="left",vertical="center",wrap_text=True); c4.border=thin; ws.row_dimensions[fila].height=22; fila+=1
        if resumen_data.get("hallazgo_destacado"):
            fila+=1; ws.merge_cells(f"A{fila}:D{fila}"); hc(ws,fila,1,"HALLAZGO",bg="F59E0B",fg=BL,sz=11); fila+=1
            ws.merge_cells(f"A{fila}:D{fila+1}"); c5=ws.cell(row=fila,column=1,value=resumen_data["hallazgo_destacado"])
            c5.fill=PatternFill("solid",fgColor="FFFBEB"); c5.font=Font(italic=True,size=11,color="92400E")
            c5.alignment=Alignment(horizontal="left",vertical="center",wrap_text=True); ws.row_dimensions[fila].height=45
    for col in ['A','B','C','D']: ws.column_dimensions[col].width=28
    wd=wb.create_sheet("Datos"); hc(wd,1,1,"Contenido",bg=AO,sz=12); wd.merge_cells("A1:B1"); wd.column_dimensions['A'].width=120
    for i,linea in enumerate(texto.split('\n'),start=2):
        if linea.strip():
            c6=wd.cell(row=i,column=1,value=linea.strip()); c6.alignment=Alignment(wrap_text=True,vertical="center")
            c6.fill=PatternFill("solid",fgColor=GC if i%2==0 else BL); wd.row_dimensions[i].height=18
    buf=BytesIO(); wb.save(buf); return buf.getvalue()

def safe_text(t): return str(t).encode('latin-1','replace').decode('latin-1')

def exportar_pdf(texto, resumen_data=None):
    pdf=FPDF(); pdf.set_margins(10,10,10); pdf.add_page(); pdf.set_auto_page_break(auto=True,margin=15)
    pdf.set_fill_color(30,58,95); pdf.rect(0,0,210,32,'F')
    pdf.set_text_color(255,255,255); pdf.set_font("Helvetica",'B',16); pdf.set_xy(10,8)
    pdf.cell(190,10,"INFORME - ORO ASISTENTE",new_x="LMARGIN",new_y="NEXT",align='C')
    zona=pytz.timezone('America/Caracas'); fecha=datetime.now(zona).strftime('%d/%m/%Y %I:%M %p')
    pdf.set_font("Helvetica",'',9); pdf.set_xy(10,20)
    pdf.cell(190,8,safe_text(f"Generado: {fecha}"),new_x="LMARGIN",new_y="NEXT",align='C')
    pdf.set_xy(10,35); pdf.set_text_color(30,30,30)
    if resumen_data:
        td=resumen_data.get("titulo","")
        if td:
            pdf.set_fill_color(37,99,235); pdf.set_text_color(255,255,255); pdf.set_font("Helvetica",'B',12)
            pdf.cell(190,10,safe_text(td[:90]),border=0,new_x="LMARGIN",new_y="NEXT",align='C',fill=True); pdf.ln(3); pdf.set_text_color(30,30,30)
        re2=resumen_data.get("resumen_ejecutivo","")
        if re2: pdf.set_font("Helvetica",'I',10); pdf.multi_cell(190,6,safe_text(re2)); pdf.ln(4)
        met=resumen_data.get("metricas",{})
        if met:
            pdf.set_fill_color(30,58,95); pdf.set_text_color(255,255,255); pdf.set_font("Helvetica",'B',11)
            pdf.cell(190,8,safe_text("  MÉTRICAS"),border=0,new_x="LMARGIN",new_y="NEXT",fill=True); pdf.ln(2); pdf.set_text_color(30,30,30)
            tog=False
            for k,v in met.items():
                rb,gb,bb=(245,247,250) if tog else (255,255,255); pdf.set_fill_color(rb,gb,bb)
                pdf.set_font("Helvetica",'B',10); pdf.cell(85,8,safe_text(f"  {k}"),border=0,fill=True)
                pdf.set_font("Helvetica",'',10); pdf.cell(105,8,safe_text(str(v)),border=0,new_x="LMARGIN",new_y="NEXT",fill=True); tog=not tog
            pdf.ln(4)
        pts=resumen_data.get("puntos_clave",[])
        if pts:
            pdf.set_fill_color(30,64,175); pdf.set_text_color(255,255,255); pdf.set_font("Helvetica",'B',11)
            pdf.cell(190,8,safe_text("  PUNTOS CLAVE"),border=0,new_x="LMARGIN",new_y="NEXT",fill=True); pdf.ln(2); pdf.set_text_color(30,30,30)
            pdf.set_font("Helvetica",'',10)
            for i,p in enumerate(pts,1): pdf.multi_cell(190,7,safe_text(f"  {i}. {p}")); pdf.ln(4)
        hall=resumen_data.get("hallazgo_destacado","")
        if hall:
            pdf.set_fill_color(180,120,10); pdf.set_text_color(255,255,255); pdf.set_font("Helvetica",'B',11)
            pdf.cell(190,8,safe_text("  HALLAZGO"),border=0,new_x="LMARGIN",new_y="NEXT",fill=True); pdf.ln(2); pdf.set_text_color(30,30,30)
            pdf.set_font("Helvetica",'I',10); pdf.multi_cell(190,7,safe_text(f"  {hall}")); pdf.ln(4)
        pdf.add_page()
    pdf.set_fill_color(30,58,95); pdf.set_text_color(255,255,255); pdf.set_font("Helvetica",'B',11)
    pdf.cell(190,8,safe_text("  CONTENIDO"),border=0,new_x="LMARGIN",new_y="NEXT",fill=True); pdf.ln(2); pdf.set_text_color(30,30,30)
    pdf.set_font("Helvetica",'',9)
    for linea in texto.split('\n'):
        if linea.strip(): pdf.multi_cell(190,5,safe_text(linea.strip()))
    raw=pdf.output()
    return bytes(raw) if isinstance(raw,(bytes,bytearray)) else raw.encode('latin-1') if isinstance(raw,str) else bytes(raw)


def exportar_excel_como_pdf(archivo_bytes, cambios=None):
    """Convierte Excel a PDF tabla real usando fpdf."""
    cambios = cambios or []
    try:
        # Aplicar cambios si hay
        bytes_usar = archivo_bytes
        if cambios:
            bytes_usar, _ = reemplazar_xlsx_preservando_formato(archivo_bytes, cambios)
        wb = openpyxl.load_workbook(BytesIO(bytes_usar), data_only=True, read_only=True)
        pdf = FPDF(orientation='L')  # landscape para tablas anchas
        pdf.set_margins(8, 8, 8)
        pdf.set_auto_page_break(auto=True, margin=12)
        zona = pytz.timezone('America/Caracas')
        fecha = datetime.now(zona).strftime('%d/%m/%Y %I:%M %p')
        for sheet in wb.worksheets:
            pdf.add_page()
            # Encabezado de hoja
            pdf.set_fill_color(30, 58, 95)
            pdf.set_text_color(255, 255, 255)
            pdf.set_font("Helvetica", 'B', 11)
            pdf.cell(0, 8, safe_text(f"Hoja: {sheet.title}  |  {fecha}"),
                     new_x="LMARGIN", new_y="NEXT", fill=True)
            pdf.ln(2)
            # Recoger filas
            filas = []
            for row in sheet.iter_rows(values_only=True):
                fila = [str(c) if c is not None else "" for c in row]
                if any(v.strip() for v in fila):
                    filas.append(fila)
            if not filas:
                pdf.set_text_color(100, 100, 100)
                pdf.set_font("Helvetica", 'I', 9)
                pdf.cell(0, 6, "(Hoja vacía)", new_x="LMARGIN", new_y="NEXT")
                continue
            n_cols = max(len(f) for f in filas)
            page_w = pdf.w - 16  # margen total
            col_w = min(page_w / n_cols, 55)
            # Filas
            for ri, fila in enumerate(filas):
                if ri == 0:
                    pdf.set_fill_color(30, 58, 95)
                    pdf.set_text_color(255, 255, 255)
                    pdf.set_font("Helvetica", 'B', 8)
                elif ri % 2 == 0:
                    pdf.set_fill_color(245, 248, 255)
                    pdf.set_text_color(20, 20, 40)
                    pdf.set_font("Helvetica", '', 8)
                else:
                    pdf.set_fill_color(255, 255, 255)
                    pdf.set_text_color(20, 20, 40)
                    pdf.set_font("Helvetica", '', 8)
                for ci in range(n_cols):
                    val = fila[ci] if ci < len(fila) else ""
                    pdf.cell(col_w, 6, safe_text(val[:30]),
                             border=1, fill=True)
                pdf.ln()
        wb.close()
        raw = pdf.output()
        return bytes(raw) if isinstance(raw, (bytes, bytearray)) else raw.encode('latin-1')
    except Exception as e:
        return None

def exportar_word_como_pdf(archivo_bytes, cambios=None):
    """Convierte DOCX a PDF preservando texto via fpdf."""
    cambios = cambios or []
    try:
        bytes_usar = archivo_bytes
        if cambios:
            bytes_usar, _ = reemplazar_docx_preservando_formato(archivo_bytes, cambios)
        doc = Document(BytesIO(bytes_usar))
        pdf = FPDF()
        pdf.set_margins(15, 15, 15)
        pdf.add_page()
        pdf.set_auto_page_break(auto=True, margin=15)
        zona = pytz.timezone('America/Caracas')
        fecha = datetime.now(zona).strftime('%d/%m/%Y %I:%M %p')
        # Encabezado
        pdf.set_fill_color(30, 58, 95)
        pdf.set_text_color(255, 255, 255)
        pdf.set_font("Helvetica", 'B', 12)
        pdf.cell(0, 8, safe_text("DOCUMENTO WORD — ORO ASISTENTE"),
                 new_x="LMARGIN", new_y="NEXT", fill=True)
        pdf.set_font("Helvetica", '', 8)
        pdf.cell(0, 6, safe_text(f"Generado: {fecha}"),
                 new_x="LMARGIN", new_y="NEXT", fill=True)
        pdf.ln(4)
        pdf.set_text_color(20, 20, 40)
        # Paragraphs
        for para in doc.paragraphs:
            txt = para.text.strip()
            if not txt:
                pdf.ln(2); continue
            style = para.style.name if para.style else ""
            if "Heading 1" in style or "Título 1" in style:
                pdf.set_font("Helvetica", 'B', 13)
                pdf.set_text_color(30, 58, 95)
                pdf.multi_cell(0, 7, safe_text(txt)); pdf.ln(1)
                pdf.set_text_color(20, 20, 40)
            elif "Heading 2" in style or "Título 2" in style:
                pdf.set_font("Helvetica", 'B', 11)
                pdf.set_text_color(50, 80, 150)
                pdf.multi_cell(0, 6, safe_text(txt)); pdf.ln(1)
                pdf.set_text_color(20, 20, 40)
            else:
                pdf.set_font("Helvetica", '', 10)
                pdf.multi_cell(0, 5, safe_text(txt))
        raw = pdf.output()
        return bytes(raw) if isinstance(raw, (bytes, bytearray)) else raw.encode('latin-1')
    except Exception as e:
        return None


def ordenar_excel(archivo_bytes, instruccion, texto_doc=""):
    """
    Ordena filas de Excel. Lee solo valores, crea archivo nuevo ordenado.
    Usa el contenido del documento para identificar la columna correcta.
    """
    try:
        # Leer todas las hojas como valores primero
        wb_read = openpyxl.load_workbook(BytesIO(archivo_bytes), data_only=True, read_only=True)
        sheets_data = {}
        for ws in wb_read.worksheets:
            filas = [list(row) for row in ws.iter_rows(values_only=True)]
            # Filtrar filas completamente vacías
            filas = [f for f in filas if any(v is not None and str(v).strip() for v in f)]
            if filas:
                sheets_data[ws.title] = filas
        wb_read.close()

        if not sheets_data:
            return None, "No se encontraron datos en el archivo."

        # Construir contexto de encabezados para la IA
        enc_ctx = ""
        for title, filas in sheets_data.items():
            if filas:
                hdrs = [str(v) if v is not None else "" for v in filas[0]]
                enc_ctx += f"Hoja '{title}' — Columnas: " + ", ".join([f"{chr(65+i)}={h}" for i,h in enumerate(hdrs)]) + "\n"

        # IA con contexto real del documento
        prompt = (
            f"El usuario tiene un Excel con esta estructura:\n{enc_ctx}\n"
            f"Instrucción del usuario: \"{instruccion}\"\n"
            "Identifica qué columna ordenar y en qué dirección.\n"
            "La primera fila ES el encabezado (no se ordena).\n"
            "Devuelve SOLO JSON:\n"
            '{"hoja":"nombre de la hoja o ALL para todas","col_letra":"A","col_nombre":"nombre del encabezado","direccion":"asc","tiene_encabezado":true}'
        )
        r = llamar_ia(prompt)
        params = extraer_json_seguro(r) if r else {}
        if not params: params = {}

        col_letra = str(params.get("col_letra","A")).strip().upper()
        col_nombre = params.get("col_nombre", col_letra)
        direccion_asc = str(params.get("direccion","asc")).lower() == "asc"
        tiene_enc = params.get("tiene_encabezado", True)

        from openpyxl.utils import column_index_from_string
        try:
            col_idx = column_index_from_string(col_letra) - 1
        except:
            col_idx = 0

        # Crear nuevo workbook con datos ordenados
        wb_new = openpyxl.Workbook()
        wb_new.remove(wb_new.active)

        for sheet_title, filas in sheets_data.items():
            ws_new = wb_new.create_sheet(title=sheet_title)

            encabezado = filas[0] if tiene_enc else None
            datos = filas[1:] if tiene_enc else filas

            if not datos:
                # Hoja solo con encabezado
                if encabezado:
                    for ci, val in enumerate(encabezado, 1):
                        ws_new.cell(row=1, column=ci, value=val)
                continue

            # Asegurar col_idx válido
            max_cols = max(len(f) for f in filas)
            col_idx_real = min(col_idx, max_cols - 1)

            def sort_key(fila):
                val = fila[col_idx_real] if col_idx_real < len(fila) else None
                if val is None: return (2, "")
                # Número → ordenar numéricamente
                try: return (0, float(val))
                except:
                    return (1, str(val).strip().lower())

            datos_ord = sorted(datos, key=sort_key, reverse=not direccion_asc)

            # Escribir encabezado
            ri = 1
            if encabezado is not None:
                for ci, val in enumerate(encabezado, 1):
                    cell = ws_new.cell(row=ri, column=ci, value=val)
                    cell.font = Font(bold=True, color="FFFFFF", size=10)
                    cell.fill = PatternFill("solid", fgColor="1E3A5F")
                    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                ri += 1

            # Escribir datos ordenados
            for fila in datos_ord:
                bg = "EEF2FF" if ri % 2 == 0 else "FFFFFF"
                for ci, val in enumerate(fila, 1):
                    cell = ws_new.cell(row=ri, column=ci, value=val)
                    cell.fill = PatternFill("solid", fgColor=bg)
                    cell.font = Font(size=10)
                    cell.alignment = Alignment(vertical="center", wrap_text=True)
                # Rellenar celdas vacías si la fila tiene menos columnas
                for ci in range(len(fila)+1, max_cols+1):
                    ws_new.cell(row=ri, column=ci, value=None)
                ri += 1

            # Ajustar anchos de columna
            for col in ws_new.columns:
                max_w = max((len(str(c.value or "")) for c in col if c.value), default=8)
                ws_new.column_dimensions[col[0].column_letter].width = min(max_w + 4, 50)

        buf = BytesIO()
        wb_new.save(buf)
        dir_txt = "A → Z" if direccion_asc else "Z → A"
        return buf.getvalue(), f"✅ Ordenado por **{col_nombre}** ({col_letra}) {dir_txt}"

    except Exception as e:
        return None, f"No pude ordenar: {str(e)}"


# ══════════════════════════════════════════════════════════════
# REEMPLAZOS PRESERVANDO FORMATO
# ══════════════════════════════════════════════════════════════
def _aplicar_formato_run(run, fmt):
    if not fmt: return
    if fmt.get("bold") is not None: run.font.bold=fmt["bold"]
    if fmt.get("italic") is not None: run.font.italic=fmt["italic"]
    if fmt.get("underline") is not None: run.font.underline=fmt["underline"]
    if fmt.get("size"): run.font.size=Pt(fmt["size"])
    if fmt.get("color"):
        try:
            h=fmt["color"].lstrip("#"); run.font.color.rgb=RGBColor(int(h[0:2],16),int(h[2:4],16),int(h[4:6],16))
        except: pass
    if fmt.get("upper"): run.text=run.text.upper()
    if fmt.get("lower"): run.text=run.text.lower()

def reemplazar_docx_preservando_formato(archivo_bytes, cambios):
    doc=Document(BytesIO(archivo_bytes)); conteo=0
    for c in cambios:
        buscar=str(c["buscar"]); reemplazar=str(c["reemplazar"]); fmt=c.get("formato",{})
        if not buscar or buscar.lower()==reemplazar.lower(): continue
        regex=re.compile(re.escape(buscar),re.IGNORECASE)
        def rep_p(p):
            nonlocal conteo
            if not regex.search(p.text): return
            for run in p.runs:
                if regex.search(run.text):
                    nt,n=regex.subn(reemplazar,run.text)
                    if n>0:
                        run.text=nt; conteo+=n
                        if fmt: _aplicar_formato_run(run,fmt)
                    return
            nt,n=regex.subn(reemplazar,p.text)
            if n==0: return
            conteo+=n
            run_ref=next((r for r in p.runs if buscar.lower() in r.text.lower()),p.runs[0] if p.runs else None)
            if p.runs:
                r0=p.runs[0]
                if run_ref and run_ref!=r0:
                    r0.font.bold=run_ref.font.bold; r0.font.italic=run_ref.font.italic
                    r0.font.underline=run_ref.font.underline; r0.font.size=run_ref.font.size
                    try:
                        if run_ref.font.color and run_ref.font.color.rgb: r0.font.color.rgb=run_ref.font.color.rgb
                    except: pass
                r0.text=nt; [setattr(r,'text','') for r in p.runs[1:]]
                if fmt: _aplicar_formato_run(r0,fmt)
        [rep_p(p) for p in doc.paragraphs]
        [rep_p(p) for t in doc.tables for row in t.rows for cell in row.cells for p in cell.paragraphs]
    buf=BytesIO(); doc.save(buf); return buf.getvalue(),conteo

def reemplazar_xlsx_preservando_formato(archivo_bytes, cambios):
    wb=openpyxl.load_workbook(BytesIO(archivo_bytes)); conteo=0
    for c in cambios:
        buscar=str(c["buscar"]); rv=str(c["reemplazar"])
        if not buscar or buscar.lower()==rv.lower(): continue
        regex=re.compile(re.escape(buscar),re.IGNORECASE)
        for sheet in wb.worksheets:
            for row in sheet.iter_rows():
                for cell in row:
                    if cell.value and isinstance(cell.value,str) and regex.search(cell.value):
                        nv,n=regex.subn(rv,cell.value); cell.value=nv; conteo+=n
    buf=BytesIO(); wb.save(buf); return buf.getvalue(),conteo

def reemplazar_pdf_original(archivo_bytes, cambios):
    if not PYMUPDF_OK: return archivo_bytes,0
    doc=fitz.open(stream=archivo_bytes,filetype="pdf"); conteo=0
    for c in cambios:
        buscar=str(c["buscar"]).strip(); reemplazar=str(c["reemplazar"]).strip()
        if not buscar or buscar.lower()==reemplazar.lower(): continue
        for pagina in doc:
            instancias=pagina.search_for(buscar,quads=False)
            if not instancias: continue
            bloques=pagina.get_text("dict")["blocks"]
            for rect in instancias:
                font_size=11.0; font_name="helv"; color=(0.,0.,0.); bold=False; italic=False
                for bloque in bloques:
                    for linea in bloque.get("lines",[]):
                        for span in linea.get("spans",[]):
                            if buscar.lower() in span["text"].lower():
                                font_size=span.get("size",11.0); font_name=span.get("font","helv")
                                ci=span.get("color",0); color=(((ci>>16)&0xFF)/255,((ci>>8)&0xFF)/255,(ci&0xFF)/255)
                                flags=span.get("flags",0); bold=bool(flags&2**4); italic=bool(flags&2**1); break
                fn=font_name.lower()
                use_font=("Times-BoldItalic" if "bold" in fn and "italic" in fn else
                          "Helvetica-Bold" if "bold" in fn or bold else
                          "Helvetica-Oblique" if "italic" in fn or italic else
                          "Times-Roman" if "times" in fn or "serif" in fn else
                          "Courier" if "courier" in fn or "mono" in fn else "Helvetica")
                try:
                    # CORRECCIÓN: Tomamos el color de la esquina (0,0) para no chocar con la tinta
                    pix=pagina.get_pixmap(clip=rect,dpi=72); s=pix.pixel(0,0); bg=(s[0]/255,s[1]/255,s[2]/255)
                except: bg=(1.,1.,1.)
                pagina.add_redact_annot(rect,fill=bg); pagina.apply_redactions()
                pagina.insert_text(fitz.Point(rect.x0,rect.y1-1.5),reemplazar,fontname=use_font,fontsize=font_size,color=color)
                conteo+=1
    buf=BytesIO(); doc.save(buf); doc.close(); return buf.getvalue(),conteo

# ══════════════════════════════════════════════════════════════
# UPLOADER — modo archivo o imagen
# ══════════════════════════════════════════════════════════════
st.markdown('<div class="oro-divider"></div>', unsafe_allow_html=True)
st.markdown("""<style>
[data-testid="stFileUploaderDropzoneInstructions"]>div>span::after{content:"Arrastra aquí o toca para subir"}
[data-testid="stFileUploaderDropzoneInstructions"]>div>span{font-size:0!important}
[data-testid="stFileUploaderDropzoneInstructions"]>div>small::after{content:"Máx 200MB • DOCX · XLSX · PDF · JPG · PNG"}
[data-testid="stFileUploaderDropzoneInstructions"]>div>small{font-size:0!important}
[data-testid="stFileUploadDropzone"]>div>button{visibility:hidden;position:relative}
[data-testid="stFileUploadDropzone"]>div>button::after{content:"Seleccionar archivo o foto";visibility:visible;position:absolute;left:0;right:0;text-align:center}
</style>""", unsafe_allow_html=True)

# Uploader unificado — acepta documentos E imágenes
archivo_unificado = st.file_uploader(
    "📎 Sube tu archivo o foto de documento",
    type=["docx","xlsx","pdf","jpg","jpeg","png","webp"],
    help="Word, Excel, PDF o foto de documento — máx 200MB",
    label_visibility="visible"
)

# Separar según tipo
archivo = None
img_sub = None
if archivo_unificado:
    ext = archivo_unificado.name.split(".")[-1].lower()
    if ext in ("jpg","jpeg","png","webp"):
        img_sub = archivo_unificado
    else:
        archivo = archivo_unificado
    if img_sub:
        st.image(img_sub, use_container_width=True)

        # Modo: auto-detectar o manual
        _modo_img = st.radio("",
            ["🤖 Auto-detectar tipo","📄 Forzar Word","📊 Forzar Excel"],
            horizontal=True, label_visibility="collapsed", key="img_modo")

        if st.button("🔍 Interpretar y convertir", use_container_width=True, key="btn_interpretar"):
            with st.spinner(T("interpretando")):
                img_bytes = img_sub.read()
                mime = img_sub.type or "image/jpeg"
                fmt_force = "auto" if "Auto" in _modo_img else ("word" if "Word" in _modo_img else "excel")
                resultado_img = interpretar_imagen_documento(img_bytes, mime, fmt_force)
                                                                                                                                                                                                                                
            if resultado_img and resultado_img[0]:
                texto_img, tipo_img = resultado_img
                tipo_archivo = tipo_img  # "word" o "excel"

                # Mostrar tipo detectado
                tipo_label = "📊 Tabla/Excel" if tipo_img == "excel" else "📄 Documento/Word"
                st.markdown(f'<div class="info-box">🤖 Tipo detectado: <strong>{tipo_label}</strong></div>', unsafe_allow_html=True)

                # Guardar en session state
                for k,v in [("texto_extraido",texto_img),("nombre_archivo",img_sub.name),
                            ("archivo_tipo",tipo_img),("archivo_bytes",img_bytes),
                            ("resumen_data",None),("historial_chat",[]),("lista_cambios",[]),
                            ("cambios_aplicados",None),("texto_corregido",""),("preview_cambio",None),
                            ("resumen_error",False),("generando_resumen",False),
                            ("guia_paso",1),("guia_vista",False),("historial_versiones",[]),
                            ("resultado_evaluacion",None)]:
                    st.session_state[k] = v

                # Generar archivo del tipo correcto
                if tipo_img == "excel":
                    # Tabla → Excel con formato
                    wi = openpyxl.Workbook(); wsi = wi.active; wsi.title = "Datos"
                    ri_real = 0
                    for l in texto_img.split('\n'):
                        if not l.strip(): continue
                        ri_real += 1
                        cols = [c.strip() for c in l.split('|') if c.strip()]
                        for ci2, val in enumerate(cols, 1):
                            cc = wsi.cell(row=ri_real, column=ci2, value=val)
                            if ri_real == 1:
                                cc.font = Font(bold=True, color="FFFFFF", size=10)
                                cc.fill = PatternFill("solid", fgColor="1E3A5F")
                                cc.alignment = Alignment(horizontal="center", vertical="center")
                            else:
                                bg = "F0F4FF" if ri_real % 2 == 0 else "FFFFFF"
                                cc.fill = PatternFill("solid", fgColor=bg)
                                cc.font = Font(size=10)
                                cc.alignment = Alignment(wrap_text=True, vertical="center")
                    # Ajustar anchos
                    for col in wsi.columns:
                        max_w = max((len(str(cell.value or "")) for cell in col), default=10)
                        wsi.column_dimensions[col[0].column_letter].width = min(max_w + 4, 40)
                    bxi = BytesIO(); wi.save(bxi)
                    st.session_state.imagen_archivo_bytes = bxi.getvalue()
                    st.session_state.imagen_archivo_nombre = img_sub.name.rsplit(".",1)[0] + ".xlsx"
                    st.session_state.imagen_archivo_mime = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                else:
                    # Documento → Word preservando estructura
                    di = Document(); di.styles['Normal'].font.name = 'Calibri'
                    di.styles['Normal'].font.size = Pt(11)
                    for section in di.sections:
                        section.top_margin = Inches(1.0); section.bottom_margin = Inches(1.0)
                        section.left_margin = Inches(1.2); section.right_margin = Inches(1.2)
                    for l in texto_img.split('\n'):
                        if not l.strip(): di.add_paragraph(); continue
                        # Detectar si parece un título (línea corta, toda mayúsculas o con patrón de encabezado)
                        if len(l.strip()) < 60 and (l.strip().isupper() or l.strip().startswith(('REPÚBLICA','MINISTERIO','OFICIO','MEMO','CIRCULAR','ACTA','RESOLUCIÓN','DECRETO'))):
                            h = di.add_heading(l.strip(), level=1)
                            h.runs[0].font.color.rgb = RGBColor(0x1E, 0x3A, 0x8A)
                        else:
                            p = di.add_paragraph(l.strip())
                            p.paragraph_format.space_after = Pt(4)
                    bi = BytesIO(); di.save(bi)
                    st.session_state.imagen_archivo_bytes = bi.getvalue()
                    st.session_state.imagen_archivo_nombre = img_sub.name.rsplit(".",1)[0] + ".docx"
                    st.session_state.imagen_archivo_mime = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"

                st.rerun()
            else:
                st.markdown('<div class="warn-box">⚠️ No pude leer la imagen. Intenta con una foto más nítida y con buena iluminación.</div>', unsafe_allow_html=True)

    if st.session_state.get("imagen_archivo_bytes"):
        _nom = st.session_state.get('imagen_archivo_nombre','Extraido')
        _ico = "📊" if _nom.endswith(".xlsx") else "📄"
        st.download_button(f"{_ico} Descargar {_nom}",
            st.session_state.imagen_archivo_bytes, _nom,
            mime=st.session_state.get('imagen_archivo_mime','application/octet-stream'),
            use_container_width=True)

# Procesar archivo subido
if archivo and archivo.name != st.session_state.nombre_archivo:
    with st.spinner(T("cargando")):
        contenido=archivo.read()
        for k,v in [("archivo_bytes",contenido),("nombre_archivo",archivo.name),
                    ("archivo_tipo",archivo.name.split(".")[-1].lower()),
                    ("resumen_data",None),("historial_chat",[]),("lista_cambios",[]),
                    ("cambios_aplicados",None),("texto_corregido",""),("preview_cambio",None),
                    ("resumen_error",False),("generando_resumen",False),
                    ("guia_paso",1),("guia_vista",False),("historial_versiones",[]),
                    ("resultado_evaluacion",None),("buscar_query","")]:
            st.session_state[k]=v
        texto=""
        try:
            if archivo.name.endswith(".docx"):
                doc=Document(BytesIO(contenido))
                partes=[p.text for p in doc.paragraphs if p.text.strip()]
                for t in doc.tables:
                    for row in t.rows:
                        celdas=list(dict.fromkeys([c.text.strip() for c in row.cells]))
                        if any(celdas): partes.append(" | ".join(celdas))
                texto="\n".join(partes)
            elif archivo.name.endswith(".xlsx"):
                wb=openpyxl.load_workbook(BytesIO(contenido),data_only=True,read_only=True)
                for s in wb.worksheets:
                    for r in s.iter_rows(values_only=True):
                        linea=" | ".join([str(c) for c in r if c is not None and str(c).strip()])
                        if linea.strip(): texto+=linea+"\n"
                wb.close()
            elif archivo.name.endswith(".pdf"):
                reader=PyPDF2.PdfReader(BytesIO(contenido))
                for p in reader.pages:
                    t=p.extract_text()
                    if t: texto+=t+"\n"
            st.session_state.texto_extraido=texto
        except Exception as e:
            st.error(f"Error leyendo archivo: {e}")

if not st.session_state.get("texto_extraido") and st.session_state.get("generando_resumen"):
    st.session_state.generando_resumen=False

# ══════════════════════════════════════════════════════════════
# PANEL PRINCIPAL
# ══════════════════════════════════════════════════════════════
if st.session_state.texto_extraido:
    texto=st.session_state.texto_extraido
    tipo=st.session_state.archivo_tipo
    texto_activo=st.session_state.texto_corregido if st.session_state.texto_corregido else texto

    # ── Scroll automático ──
    if st.session_state.get("scroll_to"):
        _scroll_to(st.session_state.scroll_to)
        st.session_state.scroll_to = None

    # ── File badge ──
    palabras=len(texto.split())
    ext_icon={"docx":"📄","xlsx":"📊","pdf":"📕","word":"📄","excel":"📊","texto":"📝"}.get(tipo,"📎")
    cn=len(st.session_state.lista_cambios)
    nv=len(st.session_state.historial_versiones)
    cambios_chip=f'<span class="stat-chip">✏️ {cn} cambio(s)</span>' if cn else ""
    version_chip=f'<span class="stat-chip">⏮ {nv} versión(es)</span>' if nv else ""
    st.markdown(f"""<div class="file-badge">
        <div class="file-icon">{ext_icon}</div>
        <div style="flex:1;min-width:0">
            <div class="file-info-name">{st.session_state.nombre_archivo}</div>
            <div class="file-info-stats">
                <span>📝 {palabras:,} palabras</span>{cambios_chip}{version_chip}
            </div>
        </div>
    </div>""", unsafe_allow_html=True)

    # ── Botones Analizar + Evaluar ──
    ba,be = st.columns(2)
    with ba:
        st.markdown('<div class="btn-analizar">', unsafe_allow_html=True)
        if st.button(T("analizar"),use_container_width=True,key="btn_analizar"):
            st.session_state.generando_resumen=True
            st.session_state.resumen_data=None
            st.session_state.vista_activa="resumen"
            st.session_state.resultado_evaluacion=None
            st.session_state.scroll_to="seccion-resumen"
            st.rerun()
        st.markdown('</div>', unsafe_allow_html=True)
    with be:
        st.markdown('<div class="btn-evaluar">', unsafe_allow_html=True)
        if st.button(T("evaluar"),use_container_width=True,key="btn_evaluar"):
            st.session_state.ejecutar_evaluacion=True
            st.session_state.vista_activa="evaluacion"
            st.session_state.resumen_data=None
            st.session_state.generando_resumen=False
            st.session_state.scroll_to="seccion-evaluacion"
            st.rerun()
        st.markdown('</div>', unsafe_allow_html=True)

    # ── Visor del documento ──
    with st.expander(T("ver_doc"), expanded=False):
        txt_mostrar=st.session_state.texto_corregido if st.session_state.texto_corregido else texto
        # Búsqueda dentro del documento
        bq=st.text_input("🔍 Buscar en el documento",
            value=st.session_state.get("buscar_query",""),
            placeholder="Escribe una palabra para resaltar...",
            label_visibility="collapsed",key="buscar_input")
        if bq != st.session_state.get("buscar_query",""):
            st.session_state.buscar_query=bq
        lineas=[l for l in txt_mostrar.split('\n') if l.strip()]
        if tipo=="xlsx":
            filas_t=[[c.strip() for c in l.split('|') if c.strip()] for l in lineas if '|' in l]
            if filas_t:
                max_c=max(len(f) for f in filas_t)
                hdr="".join([f'<th style="background:#0d4a1f;color:#34d399;padding:.4rem .6rem;font-size:.73rem;text-align:left;white-space:nowrap;border-right:1px solid #145c28">{i+1}</th>' for i in range(max_c)])
                tbody=""
                for ri,fila in enumerate(filas_t):
                    bg="#041c0a" if ri%2==0 else "#062510"
                    is_header=ri==0
                    cds="".join([
                        f'<td style="padding:.32rem .6rem;font-size:.76rem;{"color:#6ee7b7;font-weight:700" if is_header else "color:#d1fae5"};border-bottom:1px solid #0d4a1f;border-right:1px solid #0d4a1f;white-space:nowrap">'
                        f'{("<mark style=\"background:#854d0e;color:#fef3c7;border-radius:3px;padding:0 2px\">"+v+"</mark>" if bq and bq.lower() in v.lower() else v)}</td>'
                        for v in fila])
                    cds+="".join([f'<td style="background:{bg};padding:.32rem .6rem"></td>' for _ in range(max_c-len(fila))])
                    tbody+=f'<tr style="background:{bg}">{cds}</tr>'
                n_resultados=sum(1 for f in filas_t for v in f if bq and bq.lower() in v.lower())
                if bq and n_resultados:
                    st.markdown(f'<div class="info-box">🔍 {n_resultados} ocurrencia(s) de "{bq}"</div>', unsafe_allow_html=True)
                elif bq:
                    st.markdown(f'<div class="warn-box">⚠️ "{bq}" no encontrado</div>', unsafe_allow_html=True)
                st.markdown(
                    f'<div style="overflow-x:auto;border-radius:12px;border:1px solid #0d4a1f;max-height:320px;overflow-y:auto">'
                    f'<table style="width:100%;border-collapse:collapse"><thead><tr>{hdr}</tr></thead><tbody>{tbody}</tbody></table></div>',
                    unsafe_allow_html=True)
            else:
                st.text(txt_mostrar[:3000])
        else:
            if bq:
                ocurrencias=len(re.findall(re.escape(bq),txt_mostrar,re.IGNORECASE))
                if ocurrencias:
                    st.markdown(f'<div class="info-box">🔍 {ocurrencias} ocurrencia(s) de "{bq}"</div>', unsafe_allow_html=True)
                else:
                    st.markdown(f'<div class="warn-box">⚠️ "{bq}" no encontrado en el documento</div>', unsafe_allow_html=True)
            doc_html=""
            for i,linea in enumerate(lineas[:120],1):
                bg="#041c0a" if i%2==0 else "#062510"
                ls=linea.replace("<","&lt;").replace(">","&gt;")
                if bq and bq.lower() in ls.lower():
                    ls=re.sub(f'(?i){re.escape(bq)}',lambda m:f'<mark style="background:#854d0e;color:#fef3c7;border-radius:3px;padding:0 2px">{m.group()}</mark>',ls)
                doc_html+=(f'<div style="display:flex;gap:.5rem;padding:.28rem .5rem;background:{bg};border-bottom:1px solid #0d4a1f">'
                    f'<span style="color:#166534;font-size:.62rem;min-width:1.6rem;text-align:right;padding-top:.05rem;flex-shrink:0">{i}</span>'
                    f'<span style="color:#d1fae5;font-size:.76rem;word-break:break-word;line-height:1.5">{ls}</span></div>')
            if len(lineas)>120:
                doc_html+=f'<div style="color:#166534;font-size:.7rem;padding:.4rem;text-align:center;background:#041208">... {len(lineas)-120} líneas más</div>'
            st.markdown(f'<div style="border-radius:12px;border:1px solid #0d4a1f;overflow:hidden;max-height:340px;overflow-y:auto">{doc_html}</div>', unsafe_allow_html=True)
        if st.session_state.texto_corregido:
            st.markdown('<div class="info-box" style="margin-top:.4rem">✏️ Mostrando versión con cambios aplicados</div>', unsafe_allow_html=True)

    st.markdown('<div class="oro-divider"></div>', unsafe_allow_html=True)

    # ── Evaluación ──
    if st.session_state.get("ejecutar_evaluacion"):
        st.session_state.ejecutar_evaluacion=False
        with st.spinner(T("evaluando")):
            st.session_state.resultado_evaluacion=detectar_anomalias(texto_activo)

    if st.session_state.get("resultado_evaluacion"):
        st.markdown('<div id="seccion-evaluacion"></div>', unsafe_allow_html=True)
        resultado=st.session_state.resultado_evaluacion
        niv=resultado.get("nivel_general","Regular"); puntaje=resultado.get("puntaje",0)
        ncfg={"Excelente":("#10b981","🟢"),"Bueno":("#34d399","🟢"),"Regular":("#f59e0b","🟡"),"Deficiente":("#ef4444","🔴")}
        cfg=ncfg.get(niv,ncfg["Regular"])
        st.markdown(f"""<div style="background:linear-gradient(135deg,#041208,#062510);border:1px solid #0d4a1f;
            border-radius:16px;padding:1rem;margin:.5rem 0;text-align:center">
            <div style="font-size:2.2rem">{cfg[1]}</div>
            <div style="color:{cfg[0]};font-size:1.15rem;font-weight:800">{niv}</div>
            <div style="color:#166534;font-size:.75rem;margin-top:.2rem">Puntaje: <strong style="color:{cfg[0]}">{puntaje}/100</strong></div>
        </div>""", unsafe_allow_html=True)
        ne=[("criticos","🔴 Crítico","#ef4444","#1f0707","#450a0a"),
            ("altos","🟠 Alto","#f97316","#1c0a03","#431407"),
            ("medios","🟡 Medio","#f59e0b","#1c1003","#451a03"),
            ("leves","🟢 Leve","#22c55e","#052e16","#14532d")]
        hay=False
        for key,label,cfg2,cbg,cbrd in ne:
            items_e=resultado.get(key,[])
            if items_e:
                hay=True
                rows="".join([f'<div style="color:#d1fae5;font-size:.77rem;padding:.2rem 0;border-bottom:1px solid {cbrd}">• {it}</div>' for it in items_e])
                st.markdown(f'<div style="background:{cbg};border:1px solid {cbrd};border-left:4px solid {cfg2};border-radius:12px;padding:.7rem .9rem;margin:.35rem 0"><div style="color:{cfg2};font-weight:700;font-size:.82rem;margin-bottom:.3rem">{label}</div>{rows}</div>', unsafe_allow_html=True)
        if not hay:
            st.markdown(f'<div class="info-box">✅ {T("sin_problemas")}</div>', unsafe_allow_html=True)
        rec=resultado.get("recomendacion","")
        if rec: st.markdown(f'<div class="hallazgo-card">💡 <strong>Recomendación:</strong> {rec}</div>', unsafe_allow_html=True)
        if st.button(T("cerrar_eval"),use_container_width=True):
            st.session_state.resultado_evaluacion=None; st.rerun()
        st.markdown('<div class="oro-divider"></div>', unsafe_allow_html=True)

    # ── Resumen ──
    if st.session_state.generando_resumen:
        with st.spinner(T("analizando")):
            txt_r=st.session_state.texto_corregido if st.session_state.texto_corregido else texto
            data_nueva=solicitar_resumen_estructurado(txt_r)
        st.session_state.generando_resumen=False
        if data_nueva:
            st.session_state.resumen_data=data_nueva; st.session_state.resumen_error=False
            if st.session_state.get("guia_paso")==1: st.session_state.guia_paso=2
        else:
            st.session_state.resumen_error=True
        st.rerun()

    if st.session_state.get("resumen_error"):
        st.markdown('<div class="warn-box">⚠️ No se pudo generar el resumen. Intenta de nuevo.</div>', unsafe_allow_html=True)
        if st.button(T("reintentar"),use_container_width=True):
            st.session_state.resumen_error=False; st.session_state.generando_resumen=True; st.rerun()

    data=st.session_state.resumen_data
    if not data and not st.session_state.generando_resumen and not st.session_state.get("resumen_error"):
        st.markdown(f"""<div style="text-align:center;padding:1.2rem 0 .6rem">
            <div style="font-size:2rem">🧠</div>
            <div style="font-weight:700;font-size:.88rem;margin-top:.3rem;color:#6b83f8">{T('listo_analizar')}</div>
        </div>""", unsafe_allow_html=True)

    if data:
        st.markdown('<div id="seccion-resumen"></div>', unsafe_allow_html=True)
        emoji=data.get("emoji_categoria","📋"); titulo_doc=data.get("titulo","Documento analizado")
        st.markdown(f"""<div class="summary-card">
            <div class="summary-card-title"><span>{emoji}</span>{titulo_doc}</div>
            {data.get("resumen_ejecutivo","")}
        </div>""", unsafe_allow_html=True)
        metricas=data.get("metricas",{})
        if metricas:
            pills='<div class="metrics-grid">'
            for k,v in list(metricas.items())[:4]:
                pills+=f'<div class="metric-pill"><div class="metric-pill-label">{k}</div><div class="metric-pill-value">{v}</div></div>'
            pills+='</div>'; st.markdown(pills,unsafe_allow_html=True)
        puntos=data.get("puntos_clave",[])
        if puntos:
            tags='<div class="tags-wrap">'+"".join([f'<span class="tag">✓ {p}</span>' for p in puntos])+'</div>'
            st.markdown(tags,unsafe_allow_html=True)
        hall=data.get("hallazgo_destacado","")
        if hall: st.markdown(f'<div class="hallazgo-card">💡 <strong>Hallazgo:</strong> {hall}</div>',unsafe_allow_html=True)

        st.markdown('<div class="oro-divider"></div>', unsafe_allow_html=True)
        st.markdown(f'<div class="section-title">{T("exportar")}</div>', unsafe_allow_html=True)
        ab=st.session_state.archivo_bytes; ca=st.session_state.lista_cambios
        ec1,ec2,ec3=st.columns(3)

        if tipo == "xlsx":
            # EXCEL → Excel corregido | Word informe | PDF tabla
            with ec1:
                _xls_out = reemplazar_xlsx_preservando_formato(ab, ca)[0] if ca else ab
                st.download_button("📊 Excel", _xls_out, "Documento.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True)
            with ec2:
                st.download_button("📄 Word", exportar_word(texto_activo, data, archivo_bytes=ab, archivo_tipo=tipo, cambios=ca),
                    "Informe.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True)
            with ec3:
                _pdf_xls = exportar_excel_como_pdf(ab, ca)
                if _pdf_xls:
                    st.download_button("📕 PDF", _pdf_xls, "Documento.pdf",
                        mime="application/pdf", use_container_width=True)
                else:
                    st.download_button("📕 PDF", exportar_pdf(texto_activo, data), "Informe.pdf",
                        mime="application/pdf", use_container_width=True)

        elif tipo == "docx":
            # WORD → Word corregido | Excel tablas | PDF del Word
            with ec1:
                _doc_out = reemplazar_docx_preservando_formato(ab, ca)[0] if ca else ab
                st.download_button("📄 Word", _doc_out, "Documento.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True)
            with ec2:
                st.download_button("📊 Excel", exportar_excel(texto_activo, data, archivo_bytes=ab, archivo_tipo=tipo, cambios=ca),
                    "Tablas.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True)
            with ec3:
                _pdf_doc = exportar_word_como_pdf(ab, ca)
                if _pdf_doc:
                    st.download_button("📕 PDF", _pdf_doc, "Documento.pdf",
                        mime="application/pdf", use_container_width=True)
                else:
                    st.download_button("📕 PDF", exportar_pdf(texto_activo, data), "Informe.pdf",
                        mime="application/pdf", use_container_width=True)

        else:
            # PDF u otros → Word informe | Excel datos | PDF original/corregido
            with ec1:
                st.download_button("📄 Word", exportar_word(texto_activo, data, archivo_bytes=ab, archivo_tipo=tipo, cambios=ca),
                    "Informe.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True)
            with ec2:
                st.download_button("📊 Excel", exportar_excel(texto_activo, data, archivo_bytes=ab, archivo_tipo=tipo, cambios=ca),
                    "Datos.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True)
            with ec3:
                if tipo == "pdf" and PYMUPDF_OK and ca:
                    _pdf_corr, _ = reemplazar_pdf_original(ab, ca)
                    st.download_button("📕 PDF", _pdf_corr, "Documento.pdf",
                        mime="application/pdf", use_container_width=True)
                elif tipo == "pdf" and ab:
                    st.download_button("📕 PDF", ab, "Documento.pdf",
                        mime="application/pdf", use_container_width=True)
                else:
                    st.download_button("📕 PDF", exportar_pdf(texto_activo, data), "Informe.pdf",
                        mime="application/pdf", use_container_width=True)

        if st.button(T("regen"),use_container_width=True):
            st.session_state.generando_resumen=True; st.session_state.resumen_data=None; st.rerun()

    st.markdown('<div class="oro-divider"></div>', unsafe_allow_html=True)

    # ── Preview cambio ──
    if st.session_state.preview_cambio:
        st.markdown('<div id="seccion-preview"></div>', unsafe_allow_html=True)
        preview=st.session_state.preview_cambio
        st.markdown("""<div style="background:#f0fdf4;border:1px solid #86efac;
            border-radius:16px;padding:.9rem 1rem;margin:.4rem 0;box-shadow:0 2px 8px rgba(5,150,105,.1)">
            <div style="color:#065f46;font-weight:700;font-size:.85rem;margin-bottom:.5rem">👁 Vista previa del cambio</div>""",
            unsafe_allow_html=True)
        for c in preview:
            bq2=c["buscar"][:50]+("..." if len(c["buscar"])>50 else "")
            rq2=c["reemplazar"][:50]+("..." if len(c["reemplazar"])>50 else "")
            idx=texto_activo.lower().find(c["buscar"].lower())
            if idx!=-1:
                ini=max(0,idx-35); fin=min(len(texto_activo),idx+len(c["buscar"])+35)
                ca2=texto_activo[ini:idx].replace("<","&lt;"); cd=texto_activo[idx+len(c["buscar"]):fin].replace("<","&lt;")
                st.markdown(
                    f'<div style="background:#fff1f1;border:1px solid #fecaca;border-radius:8px;padding:.5rem .7rem;margin:.3rem 0;font-size:.76rem">'
                    f'<span style="color:#9ca3af;font-size:.6rem;text-transform:uppercase;display:block;margin-bottom:.2rem">{T("antes")}</span>'
                    f'<span style="color:#7f1d1d;font-family:JetBrains Mono,monospace">...{ca2}<mark style="background:#fee2e2;color:#991b1b;border-radius:3px;padding:0 3px;font-weight:700">{bq2}</mark>{cd}...</span></div>'
                    f'<div style="background:#f0fdf4;border:1px solid #bbf7d0;border-radius:8px;padding:.5rem .7rem;margin:.3rem 0;font-size:.76rem">'
                    f'<span style="color:#9ca3af;font-size:.6rem;text-transform:uppercase;display:block;margin-bottom:.2rem">{T("despues")}</span>'
                    f'<span style="color:#14532d;font-family:JetBrains Mono,monospace">...{ca2}<mark style="background:#dcfce7;color:#166534;border-radius:3px;padding:0 3px;font-weight:700">{rq2}</mark>{cd}...</span></div>',
                    unsafe_allow_html=True)
            else:
                st.markdown(f'<div class="warn-box">⚠️ "{bq2}" no encontrado en el documento</div>', unsafe_allow_html=True)
        st.markdown('</div>', unsafe_allow_html=True)
        cs,cn3=st.columns(2)
        with cs:
            if st.button(T("confirmar"),use_container_width=True,key="confirm_cambio"):
                # Guardar versión antes de aplicar
                guardar_version(texto_activo, st.session_state.cambios_aplicados or st.session_state.archivo_bytes)
                st.session_state.lista_cambios.extend(preview)
                st.session_state.preview_cambio=None
                # Aplicar SOLO el nuevo cambio sobre el estado actual del archivo
                base_bytes = st.session_state.cambios_aplicados or st.session_state.archivo_bytes
                if tipo=="docx": final_bytes,n=reemplazar_docx_preservando_formato(base_bytes,preview)
                elif tipo=="xlsx": final_bytes,n=reemplazar_xlsx_preservando_formato(base_bytes,preview)
                elif tipo in ("pdf","application/pdf") and PYMUPDF_OK: final_bytes,n=reemplazar_pdf_original(base_bytes,preview)
                else:
                    txt_m=texto_activo; n=0
                    for c2 in preview:
                        txt_m,cnt=re.compile(re.escape(c2["buscar"]),re.IGNORECASE).subn(c2["reemplazar"],txt_m); n+=cnt
                    final_bytes=txt_m.encode()
                # Actualizar texto corregido aplicando solo el nuevo cambio
                txt_c=texto_activo
                for c2 in preview: txt_c=re.compile(re.escape(c2["buscar"]),re.IGNORECASE).sub(c2["reemplazar"],txt_c)
                st.session_state.texto_corregido=txt_c
                st.session_state.cambios_aplicados=final_bytes
                # NO regenerar resumen automáticamente — evita gastar tokens
                st.session_state.edicion_counter+=1
                st.session_state.historial_chat.append({"rol":"Asistente",
                    "texto":f"✅ Listo — cambié **{preview[0]['buscar']}** → **{preview[0]['reemplazar']}**. ¿Algo más?"})
                st.rerun()
        with cn3:
            if st.button(T("cancelar"),use_container_width=True,key="cancel_cambio"):
                st.session_state.preview_cambio=None; st.session_state.edicion_counter+=1; st.rerun()

    # ── Descargar corregido ──
    if st.session_state.cambios_aplicados:
        with st.expander(f"📥 Descargar corregido · {len(st.session_state.lista_cambios)} cambio(s)"):
            fb=st.session_state.cambios_aplicados
            if tipo=="docx":
                st.download_button(T("word_corregido"),fb,"Corregido.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",use_container_width=True)
            elif tipo=="xlsx":
                st.download_button(T("excel_corregido"),fb,"Corregido.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",use_container_width=True)
            elif tipo=="pdf" and PYMUPDF_OK:
                st.download_button(T("pdf_corregido"),fb,"Corregido.pdf",mime="application/pdf",use_container_width=True)
            wc=exportar_word(st.session_state.texto_corregido or texto,None)
            st.download_button(T("exportar_word"),wc,"Exportado.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",use_container_width=True)
            st.markdown('<div class="btn-peligro">', unsafe_allow_html=True)
            if st.button(T("limpiar_cambios"),use_container_width=True):
                st.session_state.lista_cambios=[]; st.session_state.cambios_aplicados=None
                st.session_state.texto_corregido=""; st.session_state.preview_cambio=None
                st.session_state.historial_versiones=[]; st.rerun()
            st.markdown('</div>', unsafe_allow_html=True)

    # ══════════════════════════════════════════
    # CHAT — siempre visible
    # ══════════════════════════════════════════
    for msg in st.session_state.historial_chat:
        with st.chat_message("user" if msg["rol"]=="Usuario" else "assistant"):
            st.write(msg["texto"])

    if not st.session_state.historial_chat and not st.session_state.preview_cambio:
        st.markdown(f"""<div class="chat-placeholder">
            <div style="font-size:1.6rem">💬</div>
            <div style="font-weight:700;font-size:.95rem;margin:.25rem 0;color:#1a1f36">{T('chat_titulo')}</div>
            <div style="font-size:.76rem;margin-bottom:.7rem;color:#8896b3">{T('chat_hint')}</div>
            <div>
                <span class="chip">{T('chip1')}</span>
                <span class="chip">{T('chip2')}</span>
                <span class="chip">{T('chip3')}</span>
                <span class="chip">{T('chip4')}</span>
            </div>
        </div>""", unsafe_allow_html=True)

    entrada=st.chat_input(T("chat_placeholder"), key=f"chat_{st.session_state.edicion_counter}")
    if entrada:
        st.session_state.historial_chat.append({"rol":"Usuario","texto":entrada})
        if st.session_state.get("guia_paso")==2: st.session_state.guia_paso=3
        palabras_cambio_es=["cambia","reemplaza","sustituye","corrige","agrega","añade","borra","elimina","pon","escribe","modifica","quita","actualiza","completa"]
        palabras_cambio_en=["change","replace","substitute","correct","add","delete","remove","put","write","modify","update","complete"]
        palabras_ordenar=["ordena","ordenar","ordena por","sort","organiza","organizar","alfabético","alfabeticamente","de mayor a menor","de menor a mayor","ascendente","descendente","a-z","z-a","a la z","la z"]
        es_ordenar = any(p in entrada.lower() for p in palabras_ordenar) and tipo=="xlsx"
        es_cambio=any(p in entrada.lower() for p in palabras_cambio_es+palabras_cambio_en) and not es_ordenar
        if es_ordenar:
            with st.spinner("📊 Ordenando..."):
                resultado_ord = ordenar_excel(st.session_state.archivo_bytes, entrada, texto_activo)
            if resultado_ord[0]:
                st.session_state.cambios_aplicados = resultado_ord[0]
                # Actualizar texto extraído con el nuevo orden
                wb_ord = openpyxl.load_workbook(BytesIO(resultado_ord[0]), data_only=True, read_only=True)
                txt_ord = ""
                for s in wb_ord.worksheets:
                    for r in s.iter_rows(values_only=True):
                        linea = " | ".join([str(c) for c in r if c is not None and str(c).strip()])
                        if linea.strip(): txt_ord += linea + "\n"
                wb_ord.close()
                st.session_state.texto_extraido = txt_ord
                st.session_state.texto_corregido = txt_ord
                st.session_state.archivo_bytes = resultado_ord[0]
                st.session_state.edicion_counter += 1
                guardar_version(texto_activo, st.session_state.archivo_bytes)
                st.session_state.historial_chat.append({"rol":"Asistente","texto": resultado_ord[1] + "\n\nYa puedes descargarlo con los cambios."})
            else:
                st.session_state.historial_chat.append({"rol":"Asistente","texto": "No pude ordenar el archivo. " + resultado_ord[1]})

        elif es_cambio:
            with st.spinner(T("procesando")):
                nuevos=solicitar_cambios(entrada,texto_activo)
            if nuevos:
                st.session_state.preview_cambio=nuevos
                st.session_state.vista_activa="preview"
                st.session_state.resultado_evaluacion=None
                st.session_state.scroll_to="seccion-preview"
                msg = "Found the change 👆 Review the preview above and confirm." if st.session_state.get("idioma")=="en" else "Encontré el cambio 👆 Revisa la vista previa arriba y confirma."
                st.session_state.historial_chat.append({"rol":"Asistente","texto":msg})
            else:
                msg = "Couldn't find what to change. Try: *change 'original word' to 'new word'*" if st.session_state.get("idioma")=="en" else "No encontré qué cambiar. Intenta: *cambia 'palabra original' por 'palabra nueva'*"
                st.session_state.historial_chat.append({"rol":"Asistente","texto":msg})
        else:
            with st.spinner(T("pensando")):
                resp=preguntar_al_documento(entrada,texto_activo)
            st.session_state.historial_chat.append({"rol":"Asistente","texto":resp})
        st.rerun()

else:
    st.markdown("""<div class="empty-state">
        <div class="empty-icon">📂</div>
        <div class="empty-title">Sube un archivo para empezar</div>
        <div class="empty-hint">
            Analiza documentos con IA · Edita con lenguaje natural<br>
            Exporta a Word, Excel o PDF
        </div>
        <div class="format-badges">
            <span class="format-badge">.docx</span>
            <span class="format-badge">.xlsx</span>
            <span class="format-badge">.pdf</span>
            <span class="format-badge">📷 foto</span>
        </div>
    </div>""", unsafe_allow_html=True)

zona_horaria=pytz.timezone('America/Caracas')
hora=datetime.now(zona_horaria).strftime('%I:%M %p')
st.markdown(f"<p class='oro-footer'>🏆 Oro Asistente v3.1 · {hora} VET</p>", unsafe_allow_html=True)
