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
.file-badge{{display:flex;align-items:center;gap:.75rem;background:{t['card']};border:1px solid {t['borde']};border-radius:18px;padding:.85rem 1rem;margin:.5rem 0;animation:fadeUp .3s ease;box-shadow:0 4px 16px {t['sombra2']}}}
.file-icon{{font-size:1.6rem;flex-shrink:0}}
.file-info-name{{color:{t['texto']}!important;font-weight:700;font-size:.85rem;word-break:break-all;line-height:1.3}}
.file-info-stats{{color:{t['texto3']}!important;font-size:.7rem;margin-top:.15rem;display:flex;gap:.4rem;flex-wrap:wrap}}
.stat-chip{{background:{t['bg2']};border:1px solid {t['borde']};border-radius:8px;padding:.05rem .4rem;font-size:.65rem;color:{t['acento1']};font-weight:600}}
.stButton>button{{background:{t['card']}!important;color:{t['texto2']}!important;border:1.5px solid {t['borde']}!important;border-radius:12px!important;font-weight:600!important;font-size:.84rem!important;min-height:3rem!important;width:100%!important;transition:all .15s ease!important;font-family:'Inter',sans-serif!important;box-shadow:0 2px 6px {t['sombra2']}!important}}
.stButton>button:hover{{border-color:{t['acento1']}!important;color:{t['acento2']}!important;background:{t['bg2']}!important;box-shadow:0 4px 14px {t['sombra']}!important;transform:translateY(-1px)!important}}
.btn-analizar>button{{background:linear-gradient(135deg,{t['acento1']},{t['acento3']})!important;color:white!important;border:none!important;font-weight:700!important;box-shadow:0 4px 14px {t['sombra']}!important}}
.btn-evaluar>button{{background:linear-gradient(135deg,#059669,#0891b2)!important;color:white!important;border:none!important;font-weight:700!important;box-shadow:0 4px 14px rgba(5,150,105,.2)!important}}
.summary-card{{background:{t['card']};border:1px solid {t['borde']};border-left:4px solid {t['acento1']};border-radius:18px;padding:1.1rem 1.2rem;margin:.7rem 0;color:{t['texto2']}!important;line-height:1.75;font-size:.88rem;box-shadow:0 4px 20px {t['sombra2']};animation:fadeUp .4s ease}}
.metrics-grid{{display:grid;grid-template-columns:1fr 1fr;gap:.5rem;margin:.6rem 0}}
.metric-pill{{background:{t['card']};border:1px solid {t['borde']};border-radius:14px;padding:.75rem 1rem;text-align:center;box-shadow:0 2px 8px {t['sombra2']}}}
.metric-pill-value{{color:{t['acento2']}!important;font-size:1.1rem;font-weight:800;margin-top:.2rem;font-family:'JetBrains Mono',monospace}}
.hallazgo-card{{background:linear-gradient(135deg,{t['bg2']},{t['card']});border:1px solid {t['borde']};border-left:4px solid {t['acento1']};border-radius:14px;padding:.85rem 1rem;color:{t['texto2']}!important;font-size:.82rem;margin:.6rem 0;line-height:1.65;box-shadow:0 2px 8px {t['sombra2']}}}
.info-box{{background:#f0fdf4;border:1px solid #86efac;border-radius:12px;padding:.75rem 1rem;color:#166534;font-size:.83rem;margin:.5rem 0;display:flex;align-items:center;gap:.5rem}}
.warn-box{{background:#fffbeb;border:1px solid #fcd34d;border-radius:12px;padding:.75rem 1rem;color:#92400e;font-size:.83rem;margin:.5rem 0;display:flex;align-items:center;gap:.5rem}}
.oro-divider{{height:1px;background:linear-gradient(90deg,transparent,{t['borde2']},transparent);margin:.9rem 0}}
.oro-footer{{text-align:center;font-size:.68rem;color:{t['texto3']};padding:.5rem 0;opacity:.7}}
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
        "exportar":"📥 Exportar informe","regen":"🔄 Regenerar resumen",
        "confirmar":"✅ Confirmar","cancelar":"❌ Cancelar",
        "chat_placeholder":"✍️ Escribe un cambio o una pregunta...",
        "chat_titulo":"Conversa sobre el documento",
        "chat_hint":"Edita, pregunta o pide cambios en lenguaje natural",
        "analizando":"🧠 Analizando...","evaluando":"🔎 Evaluando calidad...",
        "reintentar":"🔄 Reintentar","cerrar_eval":"✕ Cerrar evaluación",
        "antes":"Antes","despues":"Después",
        "version":"↩️ Deshacer último cambio",
        "sin_problemas":"¡Sin problemas detectados! 🎉",
        "prompt_idioma":"Responde SIEMPRE en español.",
    },
    "en": {
        "analizar":"⚡ Analyze","evaluar":"🔍 Evaluate","ver_doc":"👁 View document",
        "exportar":"📥 Export report","regen":"🔄 Regenerate summary",
        "confirmar":"✅ Confirm","cancelar":"❌ Cancel",
        "chat_placeholder":"✍️ Write a change or ask a question...",
        "chat_titulo":"Chat about document",
        "chat_hint":"Edit or ask questions in natural language",
        "analizando":"🧠 Analyzing...","evaluando":"🔎 Evaluating...",
        "reintentar":"🔄 Retry","cerrar_eval":"✕ Close",
        "antes":"Before","despues":"After",
        "version":"↩️ Undo last change",
        "sin_problemas":"No issues detected! 🎉",
        "prompt_idioma":"Always respond in English.",
    },
}

def T(key):
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

MODELOS_FALLBACK = ["gemini-1.5-flash","gemini-1.5-pro"]

def llamar_ia(prompt, es_json=False):
    for modelo in MODELOS_FALLBACK:
        try:
            model = genai.GenerativeModel(modelo)
            resp = model.generate_content(prompt)
            texto = resp.text
            if es_json:
                return extraer_json_seguro(texto, es_lista=texto.strip().startswith("["))
            return texto
        except:
            continue
    return None

# ══════════════════════════════════════════════════════════════
# FUNCIONES UTILITARIAS
# ══════════════════════════════════════════════════════════════
def extraer_json_seguro(texto, es_lista=False):
    t = texto.replace("```json","").replace("
```","").strip()
    c1,c2 = ("[","]") if es_lista else ("{","}")
    ini=t.find(c1); fin=t.rfind(c2)+1
    if ini!=-1 and fin>0:
        try: return json.loads(t[ini:fin],strict=False)
        except:
            try: return ast.literal_eval(t[ini:fin])
            except: pass
    return None

def _scroll_to(anchor_id):
    st.markdown(f'''<script>
        window.parent.document.getElementById("{anchor_id}") &&
        window.parent.document.getElementById("{anchor_id}").scrollIntoView({{behavior:"smooth",block:"start"}});
    </script>''', unsafe_allow_html=True)

def guardar_version(texto, bytes_doc):
    st.session_state.historial_versiones.append({
        "texto": texto,
        "bytes": bytes_doc,
        "ts": datetime.now().strftime("%H:%M:%S")
    })
    if len(st.session_state.historial_versiones) > 10:
        st.session_state.historial_versiones.pop(0)

# ══════════════════════════════════════════════════════════════
# FUNCIONES IA LOGIC
# ══════════════════════════════════════════════════════════════
def solicitar_resumen_estructurado(texto):
    idioma_prompt = T("prompt_idioma")
    prompt = (
        f"Analista experto. {idioma_prompt} Devuelve JSON:\n"
        '{"titulo":"...","emoji_categoria":"📋","resumen_ejecutivo":"max 3 oraciones",'
        '"metricas":{"Clave":"Valor"},"puntos_clave":["punto"],"hallazgo_destacado":"observación"}\n'
        f"DOC: {texto[:10000]}"
    )
    r = llamar_ia(prompt)
    return extraer_json_seguro(r) if r else None

def solicitar_cambios(instruccion, texto_doc=""):
    idioma_prompt = T("prompt_idioma")
    prompt = (
        f"Editor experto. {idioma_prompt}\nINST: \"{instruccion}\"\n"
        "Regla: cambia X por Y -> buscar=X, reemplazar=Y.\n"
        'JSON array: [{"buscar":"texto_exacto","reemplazar":"texto_nuevo"}]'
    )
    r = llamar_ia(prompt)
    return extraer_json_seguro(r, es_lista=True) if r else []

def preguntar_al_documento(pregunta, texto):
    prompt = f"Analiza: {texto[:8000]}\nPregunta: {pregunta}\nResponde directo en español."
    return llamar_ia(prompt) or "No pude procesar tu pregunta."

def detectar_anomalias(texto):
    prompt = (
        f"Analiza calidad. Devuelve JSON:\n"
        '{"nivel_general":"Excelente/Bueno/Regular/Deficiente","puntaje":85,'
        '"criticos":["..."],"altos":["..."],"medios":[],"leves":[],"recomendacion":"..."}\n'
        f"DOC: {texto[:10000]}"
    )
    r = llamar_ia(prompt)
    return extraer_json_seguro(r) if r else None

def interpretar_imagen_documento(imagen_bytes, mime_type="image/jpeg"):
    try:
        import base64
        img_b64 = base64.b64encode(imagen_bytes).decode("utf-8")
        model = genai.GenerativeModel("gemini-1.5-flash")
        prompt = "Extrae el texto de esta imagen. Si hay tablas usa | como separador."
        resp = model.generate_content([{"mime_type": mime_type, "data": img_b64}, prompt])
        texto_raw = resp.text.strip()
        lineas = [l for l in texto_raw.split('\n') if l.strip()]
        lineas_con_cols = sum(1 for l in lineas if len(re.split(r'\s{2,}|\t', l.strip())) >= 2)
        tipo = "excel" if (lineas_con_cols / max(len(lineas), 1)) >= 0.5 else "word"
        return texto_raw, tipo
    except:
        return None, "word"

# ══════════════════════════════════════════════════════════════
# REEMPLAZOS & EXPORTADORES
# ══════════════════════════════════════════════════════════════
def reemplazar_docx_preservando_formato(archivo_bytes, cambios):
    doc=Document(BytesIO(archivo_bytes)); conteo=0
    for c in cambios:
        buscar=str(c["buscar"]); reemplazar=str(c["reemplazar"])
        if not buscar or buscar.lower()==reemplazar.lower(): continue
        regex=re.compile(re.escape(buscar),re.IGNORECASE)
        for p in doc.paragraphs:
            if regex.search(p.text):
                nt, n = regex.subn(reemplazar, p.text)
                if n > 0: p.text = nt; conteo += n
    buf=BytesIO(); doc.save(buf); return buf.getvalue(), conteo

def reemplazar_xlsx_preservando_formato(archivo_bytes, cambios):
    wb=openpyxl.load_workbook(BytesIO(archivo_bytes)); conteo=0
    for c in cambios:
        buscar=str(c["buscar"]); rv=str(c["reemplazar"])
        regex=re.compile(re.escape(buscar),re.IGNORECASE)
        for sheet in wb.worksheets:
            for row in sheet.iter_rows():
                for cell in row:
                    if cell.value and isinstance(cell.value,str) and regex.search(cell.value):
                        nv,n=regex.subn(rv,cell.value); cell.value=nv; conteo+=n
    buf=BytesIO(); wb.save(buf); return buf.getvalue(),conteo

# ══════════════════════════════════════════════════════════════
# PANEL PRINCIPAL UI
# ══════════════════════════════════════════════════════════════
archivo_unificado = st.file_uploader("📎 Sube tu archivo o foto de documento", type=["docx","xlsx","pdf","jpg","jpeg","png"])

if archivo_unificado and archivo_unificado.name != st.session_state.nombre_archivo:
    with st.spinner("Cargando..."):
        contenido=archivo_unificado.read()
        ext=archivo_unificado.name.split(".")[-1].lower()
        st.session_state.archivo_bytes=contenido
        st.session_state.nombre_archivo=archivo_unificado.name
        st.session_state.archivo_tipo=ext
        
        if ext in ("jpg","jpeg","png"):
            txt_img, tipo_img = interpretar_imagen_documento(contenido, archivo_unificado.type)
            st.session_state.texto_extraido=txt_img
            st.session_state.archivo_tipo=tipo_img
        elif ext == "docx":
            doc=Document(BytesIO(contenido))
            st.session_state.texto_extraido="\n".join([p.text for p in doc.paragraphs])
        elif ext == "xlsx":
            wb=openpyxl.load_workbook(BytesIO(contenido), data_only=True)
            txt=""
            for s in wb.worksheets:
                for r in s.iter_rows(values_only=True):
                    txt += " | ".join([str(c) for c in r if c]) + "\n"
            st.session_state.texto_extraido=txt
        st.rerun()

if st.session_state.texto_extraido:
    texto_activo = st.session_state.texto_corregido or st.session_state.texto_extraido
    st.markdown(f'<div class="file-badge">🏆 <b>{st.session_state.nombre_archivo}</b></div>', unsafe_allow_html=True)
    
    col_a, col_e = st.columns(2)
    with col_a:
        if st.button(T("analizar"), use_container_width=True):
            st.session_state.generando_resumen=True; st.rerun()
    with col_e:
        if st.button(T("evaluar"), use_container_width=True):
            st.session_state.ejecutar_evaluacion=True; st.rerun()

    # Evaluación
    if st.session_state.ejecutar_evaluacion:
        with st.spinner(T("evaluando")):
            st.session_state.resultado_evaluacion=detectar_anomalias(texto_activo)
            st.session_state.ejecutar_evaluacion=False

    if st.session_state.resultado_evaluacion:
        res = st.session_state.resultado_evaluacion
        st.markdown(f"### Calidad: {res.get('nivel_general')}")
        for k in ["criticos", "altos"]:
            for item in res.get(k, []):
                st.warning(item)
        if st.button(T("cerrar_eval")): st.session_state.resultado_evaluacion=None; st.rerun()

    # Chat & Edición
    for m in st.session_state.historial_chat:
        with st.chat_message("user" if m["rol"]=="Usuario" else "assistant"): st.write(m["texto"])

    entrada = st.chat_input(T("chat_placeholder"))
    if entrada:
        st.session_state.historial_chat.append({"rol":"Usuario","texto":entrada})
        if any(p in entrada.lower() for p in ["cambia","reemplaza","pon"]):
            cambios = solicitar_cambios(entrada, texto_activo)
            if cambios:
                st.session_state.preview_cambio = cambios
                st.info("Revisa la vista previa y confirma.")
        else:
            resp = preguntar_al_documento(entrada, texto_activo)
            st.session_state.historial_chat.append({"rol":"Asistente","texto":resp})
        st.rerun()

    if st.session_state.preview_cambio:
        if st.button(T("confirmar")):
            guardar_version(texto_activo, st.session_state.archivo_bytes)
            # Aplicar cambios lógicos...
            st.session_state.preview_cambio = None; st.rerun()

zona_horaria=pytz.timezone('America/Caracas')
hora=datetime.now(zona_horaria).strftime('%I:%M %p')
st.markdown(f"<p class='oro-footer'>🏆 Oro Asistente v3.1 | Richard Guzman | {hora}</p>", unsafe_allow_html=True)
