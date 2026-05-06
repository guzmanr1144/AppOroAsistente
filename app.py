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

# Manteniendo el fallback 2.5 solicitado por el usuario
MODELOS_FALLBACK = ["gemini-2.5-flash", "gemini-2.0-flash", "gemini-1.5-flash"]

def llamar_ia(prompt, es_json=False):
    errores = []
    for modelo in MODELOS_FALLBACK:
        try:
            model = genai.GenerativeModel(modelo)
            resp = model.generate_content(prompt)
            texto = resp.text
            if es_json:
                return extraer_json_seguro(texto, es_lista=texto.strip().startswith("["))
            return texto
        except Exception as e:
            errores.append(f"{modelo}: {str(e)}")
            continue
            
    st.error(f"Error de conexión con IA: {errores}")
    return None

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

# CORRECCIÓN 6: Regex más estricto para evitar capturas accidentales
def extraer_cambio_con_regex(instruccion):
    patrones = [
        r"(?i)(?:cambia|reemplaza|sustituye)\s+['\"]?(.+?)['\"]?\s+(?:por|con|a)\s+['\"]?(.+?)['\"]?$",
        r"(?i)['\"](.+?)['\"]\s*(?:→|->|=>|por|con)\s*['\"]?(.+?)['\"]?$",
        r"(?i)^(.+?)\s*(?:→|->|=>)\s*(.+)$",
    ]
    for pat in patrones:
        m = re.search(pat,instruccion.strip())
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
        "Responde SOLO JSON array:\n"
        '[{"buscar":"texto_exacto","reemplazar":"texto_nuevo"}]'
    )
    r = llamar_ia(prompt)
    if r:
        res = extraer_json_seguro(r,es_lista=True)
        if res and isinstance(res,list):
            v = [c for c in res if isinstance(c,dict) and "buscar" in c and "reemplazar" in c]
            if v: return v
    return extraer_cambio_con_regex(instruccion)

def preguntar_al_documento(pregunta, texto):
    ctx = "\n".join([f"{m['rol']}: {m['texto']}" for m in st.session_state.historial_chat[-6:]])
    idioma_prompt = T("prompt_idioma")
    prompt = (
        f"Asistente experto en documentos. {idioma_prompt}\nDOCUMENTO:\n{texto[:10000]}\n\n"
        f"PREGUNTA: {pregunta}\nResponde conciso."
    )
    return llamar_ia(prompt) or "No pude procesar tu pregunta."

def detectar_anomalias(texto):
    idioma_prompt = T("prompt_idioma")
    prompt = (
        f"Analiza el documento. {idioma_prompt} Devuelve SOLO JSON:\n"
        '{"nivel_general":"Bueno","puntaje":85,"criticos":[],"altos":[],"medios":[],"leves":[],"recomendacion":"..."}\n\n'
        f"DOCUMENTO:\n{texto[:12000]}"
    )
    r = llamar_ia(prompt)
    return extraer_json_seguro(r) if r else None

def detectar_tipo_imagen(texto_raw):
    lineas = [l for l in texto_raw.split('\n') if l.strip()]
    if not lineas: return "word"
    lineas_con_cols = sum(1 for l in lineas if len(re.split(r'\s{2,}|\t|\|', l.strip())) >= 2)
    ratio_tabla = lineas_con_cols / max(len(lineas), 1)
    return "excel" if ratio_tabla >= 0.4 else "word"

def interpretar_imagen_documento(imagen_bytes, mime_type="image/jpeg", formato_salida="auto"):
    img_b64 = None
    try:
        import base64
        img_b64 = base64.b64encode(imagen_bytes).decode("utf-8")
        model = genai.GenerativeModel("gemini-2.0-flash")
        prompt = "OCR exacto. Si hay tabla usa pipes |. Si es texto usa párrafos. Devuelve SOLO el contenido."
        resp = model.generate_content([{"mime_type": mime_type, "data": img_b64}, prompt])
        texto_raw = resp.text.strip()
    except:
        return None, "word"

    tipo_detectado = detectar_tipo_imagen(texto_raw) if formato_salida == "auto" else formato_salida
    return texto_raw, tipo_detectado

# ══════════════════════════════════════════════════════════════
# EXPORTADORES
# ══════════════════════════════════════════════════════════════
# CORRECCIÓN 5: safe_text mejorado para evitar errores de latin-1
def safe_text(t):
    if not t: return ""
    return str(t).replace('\u2013', '-').replace('\u2014', '-').replace('\u2018', "'").replace('\u2019', "'").replace('\u201c', '"').replace('\u201d', '"').encode('latin-1','replace').decode('latin-1')

def exportar_word(texto, resumen_data=None, archivo_bytes=None, archivo_tipo=None, cambios=None):
    zona=pytz.timezone('America/Caracas'); fecha=datetime.now(zona).strftime('%d/%m/%Y %I:%M %p')
    if archivo_tipo=="docx" and archivo_bytes and cambios:
        r,_=reemplazar_docx_preservando_formato(archivo_bytes,cambios); return r
    
    doc=Document()
    doc.styles['Normal'].font.name='Calibri'
    doc.add_heading(resumen_data.get("titulo","INFORME") if resumen_data else "INFORME", 0)
    doc.add_paragraph(f"Generado: {fecha}")
    
    if resumen_data:
        if resumen_data.get("resumen_ejecutivo"):
            p = doc.add_paragraph(resumen_data["resumen_ejecutivo"])
            p.italic = True
        if resumen_data.get("metricas"):
            doc.add_heading('Métricas', 1)
            for k,v in resumen_data["metricas"].items():
                doc.add_paragraph(f"{k}: {v}", style='List Bullet')
                
    doc.add_heading('Contenido', 1)
    for linea in texto.split('\n'):
        if linea.strip(): doc.add_paragraph(linea.strip())
        
    buf=BytesIO(); doc.save(buf); return buf.getvalue()

def exportar_excel(texto, resumen_data=None, archivo_bytes=None, archivo_tipo=None, cambios=None):
    if archivo_tipo=="xlsx" and archivo_bytes:
        if cambios:
            r,_=reemplazar_xlsx_preservando_formato(archivo_bytes,cambios); return r
        return archivo_bytes
    
    wb=openpyxl.Workbook()
    ws=wb.active
    ws.title="Datos"
    for i,linea in enumerate(texto.split('\n'), 1):
        if '|' in linea:
            for j,col in enumerate(linea.split('|'), 1):
                ws.cell(row=i, column=j, value=col.strip())
        else:
            ws.cell(row=i, column=1, value=linea.strip())
            
    buf=BytesIO(); wb.save(buf); return buf.getvalue()

def exportar_pdf(texto, resumen_data=None):
    pdf=FPDF(); pdf.add_page(); pdf.set_font("Helvetica", size=10)
    pdf.cell(200, 10, txt="INFORME ORO ASISTENTE", ln=True, align='C')
    for linea in texto.split('\n'):
        if linea.strip(): pdf.multi_cell(0, 5, safe_text(linea.strip()))
    return bytes(pdf.output(dest='S'))

# CORRECCIÓN 1: Mejor manejo de ordenación (avisando sobre fórmulas)
def ordenar_excel(archivo_bytes, instruccion, texto_doc=""):
    try:
        # Cargamos el archivo manteniendo fórmulas para la base
        wb = openpyxl.load_workbook(BytesIO(archivo_bytes))
        ws = wb.active
        
        # Leemos datos para lógica (esto no afecta las fórmulas del archivo original)
        filas = list(ws.values)
        if not filas or len(filas) < 2: return None, "No hay suficientes datos."
        
        headers = [str(h) for h in filas[0]]
        prompt = f"Excel headers: {headers}. Instruction: {instruccion}. JSON: {{'col_idx': 0, 'asc': true}}"
        r = llamar_ia(prompt)
        params = extraer_json_seguro(r) or {"col_idx": 0, "asc": True}
        
        idx = params.get("col_idx", 0)
        # Ordenar preservando la fila de encabezado
        header_row = filas[0]
        data_rows = filas[1:]
        data_rows.sort(key=lambda x: (x[idx] is None, x[idx]), reverse=not params.get("asc", True))
        
        # Escribir de vuelta (esto puede romper fórmulas que dependan de filas específicas)
        for r_idx, row_data in enumerate(data_rows, 2):
            for c_idx, value in enumerate(row_data, 1):
                ws.cell(row=r_idx, column=c_idx, value=value)
                
        buf = BytesIO(); wb.save(buf)
        return buf.getvalue(), f"✅ Ordenado por {headers[idx]}. Nota: Las fórmulas de fila podrían requerir revisión."
    except Exception as e:
        return None, str(e)

# ══════════════════════════════════════════════════════════════
# REEMPLAZOS PRESERVANDO FORMATO
# ══════════════════════════════════════════════════════════════
def reemplazar_docx_preservando_formato(archivo_bytes, cambios):
    doc=Document(BytesIO(archivo_bytes)); conteo=0
    for c in cambios:
        buscar=str(c["buscar"]); reemplazar=str(c["reemplazar"])
        regex=re.compile(re.escape(buscar),re.IGNORECASE)
        
        def rep_p(p):
            nonlocal conteo
            if regex.search(p.text):
                for run in p.runs:
                    if regex.search(run.text):
                        nt,n=regex.subn(reemplazar,run.text)
                        run.text=nt; conteo+=n
                        
        [rep_p(p) for p in doc.paragraphs]
        [rep_p(p) for t in doc.tables for row in t.rows for cell in row.cells for p in cell.paragraphs]
    buf=BytesIO(); doc.save(buf); return buf.getvalue(),conteo

def reemplazar_xlsx_preservando_formato(archivo_bytes, cambios):
    # CORRECCIÓN: data_only=False por defecto para preservar fórmulas al guardar
    wb=openpyxl.load_workbook(BytesIO(archivo_bytes), data_only=False); conteo=0
    for c in cambios:
        buscar=str(c["buscar"]); rv=str(c["reemplazar"])
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
        for pagina in doc:
            instancias=pagina.search_for(buscar)
            for rect in instancias:
                pagina.add_redact_annot(rect, text=reemplazar, fontname="helv", fontsize=10)
                pagina.apply_redactions()
                conteo+=1
    buf=BytesIO(); doc.save(buf); doc.close(); return buf.getvalue(),conteo

# ══════════════════════════════════════════════════════════════
# UPLOADER
# ══════════════════════════════════════════════════════════════
st.markdown('<div class="oro-divider"></div>', unsafe_allow_html=True)
archivo_unificado = st.file_uploader("📎 Sube tu archivo o foto", type=["docx","xlsx","pdf","jpg","jpeg","png","webp"])

if archivo_unificado:
    if archivo_unificado.name != st.session_state.nombre_archivo:
        with st.spinner(T("cargando")):
            contenido=archivo_unificado.read()
            ext = archivo_unificado.name.split(".")[-1].lower()
            
            if ext in ("jpg","jpeg","png","webp"):
                res, tipo = interpretar_imagen_documento(contenido, archivo_unificado.type)
                st.session_state.texto_extraido = res
                st.session_state.archivo_tipo = tipo
            else:
                st.session_state.archivo_bytes = contenido
                st.session_state.archivo_tipo = ext
                texto = ""
                if ext == "docx":
                    doc = Document(BytesIO(contenido))
                    # CORRECCIÓN 2: Eliminado dict.fromkeys para no perder datos duplicados legítimos
                    partes = [p.text for p in doc.paragraphs if p.text.strip()]
                    for t in doc.tables:
                        for row in t.rows:
                            celdas = [c.text.strip() for c in row.cells]
                            if any(celdas): partes.append(" | ".join(celdas))
                    texto = "\n".join(partes)
                elif ext == "xlsx":
                    wb = openpyxl.load_workbook(BytesIO(contenido), data_only=True)
                    for s in wb.worksheets:
                        for r in s.iter_rows(values_only=True):
                            linea = " | ".join([str(c) for c in r if c is not None])
                            if linea.strip(): texto += linea + "\n"
                elif ext == "pdf":
                    reader = PyPDF2.PdfReader(BytesIO(contenido))
                    for p in reader.pages: texto += (p.extract_text() or "") + "\n"
                st.session_state.texto_extraido = texto
                
            st.session_state.nombre_archivo = archivo_unificado.name
            st.session_state.resumen_data = None
            st.rerun()

# ══════════════════════════════════════════════════════════════
# PANEL PRINCIPAL
# ══════════════════════════════════════════════════════════════
if st.session_state.texto_extraido:
    texto_activo = st.session_state.texto_corregido if st.session_state.texto_corregido else st.session_state.texto_extraido
    
    st.markdown(f'<div class="file-badge"><b>{st.session_state.nombre_archivo}</b></div>', unsafe_allow_html=True)
    
    col1, col2 = st.columns(2)
    with col1:
        if st.button(T("analizar"), use_container_width=True):
            st.session_state.generando_resumen = True
            st.rerun()
    with col2:
        if st.button(T("evaluar"), use_container_width=True):
            st.session_state.ejecutar_evaluacion = True
            st.rerun()

    if st.session_state.generando_resumen:
        with st.spinner(T("analizando")):
            st.session_state.resumen_data = solicitar_resumen_estructurado(texto_activo)
            st.session_state.generando_resumen = False
            st.rerun()

    if st.session_state.resumen_data:
        d = st.session_state.resumen_data
        st.markdown(f'<div class="summary-card"><b>{d.get("titulo")}</b><br>{d.get("resumen_ejecutivo")}</div>', unsafe_allow_html=True)

    # CHAT
    for m in st.session_state.historial_chat:
        with st.chat_message("user" if m["rol"]=="Usuario" else "assistant"):
            st.write(m["texto"])

    entrada = st.chat_input(T("chat_placeholder"))
    if entrada:
        st.session_state.historial_chat.append({"rol":"Usuario", "texto":entrada})
        if "ordena" in entrada.lower() and st.session_state.archivo_tipo == "xlsx":
            res, msg = ordenar_excel(st.session_state.archivo_bytes, entrada)
            if res: 
                st.session_state.archivo_bytes = res
                st.session_state.cambios_aplicados = res
            st.session_state.historial_chat.append({"rol":"Asistente", "texto":msg})
        else:
            cambios = solicitar_cambios(entrada, texto_activo)
            if cambios:
                st.session_state.preview_cambio = cambios
                st.session_state.historial_chat.append({"rol":"Asistente", "texto":"He encontrado los cambios. ¿Confirmas?"})
            else:
                resp = preguntar_al_documento(entrada, texto_activo)
                st.session_state.historial_chat.append({"rol":"Asistente", "texto":resp})
        st.rerun()

    if st.session_state.preview_cambio:
        if st.button(T("confirmar")):
            guardar_version(texto_activo, st.session_state.archivo_bytes)
            # Aplicar
            bytes_org = st.session_state.archivo_bytes
            tipo = st.session_state.archivo_tipo
            if tipo == "docx": res, _ = reemplazar_docx_preservando_formato(bytes_org, st.session_state.preview_cambio)
            elif tipo == "xlsx": res, _ = reemplazar_xlsx_preservando_formato(bytes_org, st.session_state.preview_cambio)
            elif tipo == "pdf": res, _ = reemplazar_pdf_original(bytes_org, st.session_state.preview_cambio)
            else: res = bytes_org
            
            st.session_state.archivo_bytes = res
            st.session_state.cambios_aplicados = res
            st.session_state.preview_cambio = None
            st.success("Cambio aplicado.")
            st.rerun()

    if st.session_state.cambios_aplicados:
        st.download_button("📥 Descargar archivo corregido", st.session_state.cambios_aplicados, "corregido."+st.session_state.archivo_tipo)

else:
    st.markdown('<div class="empty-state"><h3>Sube un documento para comenzar</h3></div>', unsafe_allow_html=True)

st.markdown(f"<p class='oro-footer'>🏆 Oro Asistente v3.2 · IDANZ - Deporte Popular</p>", unsafe_allow_html=True)
