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
# CSS — RESTAURADO COMPLETAMENTE
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
.summary-card-title{{color:{t['acento2']}!important;font-size:1rem;font-weight:800;margin-bottom:.5rem;display:flex;align-items:center;gap:.4rem}}
.metrics-grid{{display:grid;grid-template-columns:1fr 1fr;gap:.5rem;margin:.6rem 0}}
.metric-pill{{background:{t['card']};border:1px solid {t['borde']};border-radius:14px;padding:.75rem 1rem;text-align:center;box-shadow:0 2px 8px {t['sombra2']};transition:transform .2s,box-shadow .2s}}
.metric-pill:hover{{transform:translateY(-2px);box-shadow:0 6px 16px {t['sombra']}}}
.metric-pill-label{{color:{t['texto3']}!important;font-size:.62rem;font-weight:700;text-transform:uppercase;letter-spacing:.08em}}
.metric-pill-value{{color:{t['acento2']}!important;font-size:1.1rem;font-weight:800;margin-top:.2rem;font-family:'JetBrains Mono',monospace}}
.tags-wrap{{display:flex;flex-wrap:wrap;gap:.3rem;margin:.5rem 0}}
.tag{{background:{t['bg2']};color:{t['acento2']}!important;border:1px solid {t['borde']};border-radius:20px;padding:.28rem .7rem;font-size:.7rem;font-weight:600;transition:all .2s}}
.hallazgo-card{{background:linear-gradient(135deg,{t['bg2']},{t['card']});border:1px solid {t['borde']};border-left:4px solid {t['acento1']};border-radius:14px;padding:.85rem 1rem;color:{t['texto2']}!important;font-size:.82rem;margin:.6rem 0;line-height:1.65;box-shadow:0 2px 8px {t['sombra2']}}}
.info-box{{background:#f0fdf4;border:1px solid #86efac;border-radius:12px;padding:.75rem 1rem;color:#166534;font-size:.83rem;margin:.5rem 0}}
.warn-box{{background:#fffbeb;border:1px solid #fcd34d;border-radius:12px;padding:.75rem 1rem;color:#92400e;font-size:.83rem;margin:.5rem 0}}
.chat-placeholder{{text-align:center;padding:1.2rem .8rem;background:{t['card']};border:1.5px dashed {t['borde2']};border-radius:18px;margin:.4rem 0}}
.chip{{display:inline-block;background:{t['bg2']};border:1px solid {t['borde']};border-radius:20px;padding:.22rem .65rem;font-size:.68rem;color:{t['acento1']};font-weight:600;margin:.15rem}}
.oro-divider{{height:1px;background:linear-gradient(90deg,transparent,{t['borde2']},transparent);margin:.9rem 0}}
.oro-footer{{text-align:center;font-size:.68rem;color:{t['texto3']};padding:.5rem 0;opacity:.7}}
</style>"""

# ══════════════════════════════════════════════════════════════
# SESSION STATE & TRADUCCIONES
# ══════════════════════════════════════════════════════════════
_defaults = {
    "texto_extraido":"","nombre_archivo":"","archivo_bytes":None,"resumen_data":None,
    "historial_chat":[],"cambios_aplicados":None,"archivo_tipo":"","lista_cambios":[],
    "texto_corregido":"","generando_resumen":False,"resumen_error":False,
    "preview_cambio":None,"edicion_counter":0,"tema":"noche","idioma":"es",
    "historial_versiones":[], "buscar_query":"", "resultado_evaluacion":None, "scroll_to":None
}
for k,v in _defaults.items():
    if k not in st.session_state: st.session_state[k] = v

st.markdown(_get_all_css(st.session_state.tema), unsafe_allow_html=True)

_TXT = {
    "es": {
        "analizar":"⚡ Analizar","evaluar":"🔍 Evaluar","ver_doc":"👁 Ver documento",
        "chat_placeholder":"✍️ Escribe un cambio o una pregunta...",
        "chip1":"✏️ cambia X por Y","chip2":"➕ agrega dato a persona","chip3":"❓ ¿cuántos hay?","chip4":"📝 resume en 3 puntos",
        "confirmar":"✅ Confirmar","cancelar":"❌ Cancelar","antes":"Antes","despues":"Después",
        "analizando":"🧠 Analizando...","procesando":"🧠 Procesando...","cargando":"📖 Cargando...","pensando":"🤔 Pensando...",
        "regen":"🔄 Regenerar resumen", "exportar":"📥 Exportar informe", "prompt_idioma":"Responde SIEMPRE en español.",
        "listo_analizar":"Toca ⚡ Analizar para generar el resumen inteligente"
    }
}
def T(key): return _TXT.get(st.session_state.idioma, _TXT["es"]).get(key, "")

# ══════════════════════════════════════════════════════════════
# MOTORES DE IA & REEMPLAZO
# ══════════════════════════════════════════════════════════════
try:
    genai.configure(api_key=st.secrets["LLAVE_GEMINI"])
except: st.error("🔑 Error: Falta LLAVE_GEMINI"); st.stop()

def llamar_ia(prompt):
    for modelo in ["gemini-1.5-flash", "gemini-2.0-flash"]:
        try: return genai.GenerativeModel(modelo).generate_content(prompt).text
        except: continue
    return None

def extraer_json_seguro(texto, es_lista=False):
    if not texto: return None
    t = str(texto).replace("```json","").replace("```","").strip()
    ini=t.find("[" if es_lista else "{"); fin=t.rfind("]" if es_lista else "}")+1
    if ini!=-1 and fin>0:
        try: return json.loads(t[ini:fin], strict=False)
        except: return None
    return None

def solicitar_resumen_estructurado(texto):
    idioma_prompt = T("prompt_idioma")
    prompt = (
        f"Analista experto. {idioma_prompt} Devuelve SOLO JSON:\n"
        '{"titulo":"...","emoji_categoria":"📋","resumen_ejecutivo":"max 3 oraciones",'
        '"metricas":{"Clave":"Valor"},"puntos_clave":["punto"],"hallazgo_destacado":"observación"}\n\n'
        f"DOCUMENTO:\n{texto[:10000]}"
    )
    r = llamar_ia(prompt)
    return extraer_json_seguro(r)

# MOTORES DE REEMPLAZO QUIRÚRGICO
def reemplazar_docx_preservando_formato(archivo_bytes, cambios):
    doc=Document(BytesIO(archivo_bytes)); conteo=0
    for c in cambios:
        buscar, reemplazar = str(c["buscar"]), str(c["reemplazar"])
        regex=re.compile(re.escape(buscar), re.IGNORECASE)
        for p in list(doc.paragraphs) + [p for t in doc.tables for r in t.rows for cell in r.cells for p in cell.paragraphs]:
            if regex.search(p.text):
                for run in p.runs:
                    if regex.search(run.text):
                        run.text, n = regex.subn(reemplazar, run.text); conteo += n
    buf=BytesIO(); doc.save(buf); return buf.getvalue(), conteo

def reemplazar_xlsx_preservando_formato(archivo_bytes, cambios):
    wb=openpyxl.load_workbook(BytesIO(archivo_bytes), data_only=False); conteo=0
    for c in cambios:
        buscar, rv = str(c["buscar"]), str(c["reemplazar"])
        regex=re.compile(re.escape(buscar), re.IGNORECASE)
        for sheet in wb.worksheets:
            for row in sheet.iter_rows():
                for cell in row:
                    if cell.value and isinstance(cell.value, str) and regex.search(cell.value):
                        cell.value, n = regex.subn(rv, cell.value); conteo += n
    buf=BytesIO(); wb.save(buf); return buf.getvalue(), conteo

# ══════════════════════════════════════════════════════════════
# INTERFAZ PRINCIPAL
# ══════════════════════════════════════════════════════════════
st.markdown('<div class="oro-header"><div class="oro-title">Oro Asistente</div></div>', unsafe_allow_html=True)
archivo_subido = st.file_uploader("📎 Sube tu archivo", type=["docx","xlsx","pdf"])

if archivo_subido and archivo_subido.name != st.session_state.nombre_archivo:
    with st.spinner(T("cargando")):
        st.session_state.archivo_bytes = archivo_subido.read()
        st.session_state.nombre_archivo = archivo_subido.name
        st.session_state.archivo_tipo = archivo_subido.name.split(".")[-1].lower()
        texto = ""
        if st.session_state.archivo_tipo == "docx":
            doc = Document(BytesIO(st.session_state.archivo_bytes))
            texto = "\n".join([p.text for p in doc.paragraphs if p.text.strip()])
            for t in doc.tables:
                for row in t.rows: texto += "\n" + " | ".join([c.text.strip() for c in row.cells])
        elif st.session_state.archivo_tipo == "xlsx":
            wb = openpyxl.load_workbook(BytesIO(st.session_state.archivo_bytes), data_only=True)
            for s in wb.worksheets:
                for r in s.iter_rows(values_only=True):
                    l = " | ".join([str(c) for c in r if c is not None])
                    if l.strip(): texto += l + "\n"
        st.session_state.texto_extraido = texto
        st.rerun()

if st.session_state.texto_extraido:
    texto_activo = st.session_state.texto_corregido if st.session_state.texto_corregido else st.session_state.texto_extraido
    st.markdown(f'<div class="file-badge"><b>{st.session_state.nombre_archivo}</b></div>', unsafe_allow_html=True)

    # BOTONES DE ACCIÓN
    ba, be = st.columns(2)
    with ba:
        if st.button(T("analizar")):
            with st.spinner(T("analizando")):
                st.session_state.resumen_data = solicitar_resumen_estructurado(texto_activo)
                st.rerun()
    with be:
        if st.button(T("evaluar")): st.info("🔍 Próximamente evaluación avanzada.")

    # RESUMEN INTELIGENTE (RESTAURADO CON MÉTRICAS Y TAGS)
    if st.session_state.resumen_data:
        data = st.session_state.resumen_data
        st.markdown(f'<div class="summary-card"><div class="summary-card-title">{data.get("titulo")}</div>{data.get("resumen_ejecutivo")}</div>', unsafe_allow_html=True)
        
        # MÉTRICAS NUMÉRICAS (PILLS)
        met = data.get("metricas", {})
        if met:
            pills = '<div class="metrics-grid">'
            for k, v in list(met.items())[:4]:
                pills += f'<div class="metric-pill"><div class="metric-pill-label">{k}</div><div class="metric-pill-value">{v}</div></div>'
            pills += '</div>'
            st.markdown(pills, unsafe_allow_html=True)
            
        # PUNTOS CLAVE (TAGS)
        pts = data.get("puntos_clave", [])
        if pts:
            st.markdown('<div class="tags-wrap">' + "".join([f'<span class="tag">✓ {p}</span>' for p in pts]) + '</div>', unsafe_allow_html=True)
            
        # HALLAZGO
        h = data.get("hallazgo_destacado", "")
        if h: st.markdown(f'<div class="hallazgo-card">💡 <b>Hallazgo:</b> {h}</div>', unsafe_allow_html=True)

    # VISTA PREVIA DE CAMBIO
    if st.session_state.preview_cambio:
        preview = st.session_state.preview_cambio
        st.markdown('<div class="info-box">👁️ Vista previa</div>', unsafe_allow_html=True)
        for c in preview:
            idx = texto_activo.lower().find(c["buscar"].lower())
            if idx != -1:
                ini, fin = max(0, idx-40), min(len(texto_activo), idx+len(c["buscar"])+40)
                st.markdown(f"""<div style="background:#fff1f1;padding:.5rem;border-radius:8px;font-size:.75rem;margin-bottom:.2rem">
                    <small>Antes</small><br>...{texto_activo[ini:idx]}<mark style="background:#fee2e2;color:#991b1b;font-weight:bold">{c['buscar']}</mark>{texto_activo[idx+len(c['buscar']):fin]}...
                </div>
                <div style="background:#f0fdf4;padding:.5rem;border-radius:8px;font-size:.75rem">
                    <small>Después</small><br>...{texto_activo[ini:idx]}<mark style="background:#dcfce7;color:#166534;font-weight:bold">{c['reemplazar']}</mark>{texto_activo[idx+len(c['buscar']):fin]}...
                </div>""", unsafe_allow_html=True)
        c1, c2 = st.columns(2)
        with c1:
            if st.button(T("confirmar")):
                b_in = st.session_state.cambios_aplicados or st.session_state.archivo_bytes
                if st.session_state.archivo_tipo == "docx": res, _ = reemplazar_docx_preservando_formato(b_in, preview)
                elif st.session_state.archivo_tipo == "xlsx": res, _ = reemplazar_xlsx_preservando_formato(b_in, preview)
                st.session_state.cambios_aplicados = res
                st.session_state.texto_corregido = texto_activo.replace(preview[0]["buscar"], preview[0]["reemplazar"])
                st.session_state.preview_cambio = None
                st.rerun()
        with c2:
            if st.button(T("cancelar")): st.session_state.preview_cambio = None; st.rerun()

    # CHAT & ENTRADA
    for m in st.session_state.historial_chat:
        with st.chat_message("user" if m["rol"]=="Usuario" else "assistant"): st.write(m["texto"])

    entrada = st.chat_input(T("chat_placeholder"))
    if entrada:
        st.session_state.historial_chat.append({"rol":"Usuario", "texto":entrada})
        with st.spinner(T("procesando")):
            p = f"Editorial. Devuelve SOLO JSON array: [{{'buscar':'', 'reemplazar':''}}]. Si no es edición, responde normal. Instrucción: {entrada}. Contexto: {texto_activo[:2000]}"
            resp = llamar_ia(p)
            nuevos = extraer_json_seguro(resp, es_lista=True)
            if nuevos: st.session_state.preview_cambio = nuevos
            else: st.session_state.historial_chat.append({"rol":"Asistente", "texto": resp if resp else "Sin respuesta."})
        st.rerun()

    if st.session_state.cambios_aplicados:
        st.download_button("📥 Descargar corregido", st.session_state.cambios_aplicados, f"corregido_{st.session_state.nombre_archivo}", use_container_width=True)
else:
    st.markdown('<div style="text-align:center;padding:4rem 1rem">🏆 Sube un archivo para comenzar</div>', unsafe_allow_html=True)

st.markdown(f"<p class='oro-footer'>🏆 Oro Asistente v3.4 · Restaurado</p>", unsafe_allow_html=True)
