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
# CSS — cacheado por tema (SE MANTIENE TU DISEÑO ORIGINAL)
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
.summary-card{{background:{t['card']};border:1px solid {t['borde']};border-left:4px solid {t['acento1']};border-radius:18px;padding:1.1rem 1.2rem;margin:.7rem 0;color:{t['texto2']}!important;line-height:1.75;font-size:.88rem;box-shadow:0 4px 20px {t['sombra2']};animation:fadeUp .4s ease}}
.summary-card-title{{color:{t['acento2']}!important;font-size:1rem;font-weight:800;margin-bottom:.5rem;display:flex;align-items:center;gap:.4rem}}
.metrics-grid{{display:grid;grid-template-columns:1fr 1fr;gap:.5rem;margin:.6rem 0}}
.metric-pill{{background:{t['card']};border:1px solid {t['borde']};border-radius:14px;padding:.75rem 1rem;text-align:center;box-shadow:0 2px 8px {t['sombra2']}}}
.metric-pill-label{{color:{t['texto3']}!important;font-size:.62rem;font-weight:700;text-transform:uppercase}}
.metric-pill-value{{color:{t['acento2']}!important;font-size:1.1rem;font-weight:800;margin-top:.2rem}}
.hallazgo-card{{background:linear-gradient(135deg,{t['bg2']},{t['card']});border:1px solid {t['borde']};border-left:4px solid {t['acento1']};border-radius:14px;padding:.85rem 1rem;color:{t['texto2']}!important;font-size:.82rem;margin:.6rem 0;line-height:1.65}}
.info-box{{background:#f0fdf4;border:1px solid #86efac;border-radius:12px;padding:.75rem 1rem;color:#166534;font-size:.83rem}}
.warn-box{{background:#fffbeb;border:1px solid #fcd34d;border-radius:12px;padding:.75rem 1rem;color:#92400e;font-size:.83rem}}
.chat-placeholder{{text-align:center;padding:1.2rem .8rem;background:{t['card']};border:1.5px dashed {t['borde2']};border-radius:18px;margin:.4rem 0}}
.chip{{display:inline-block;background:{t['bg2']};border:1px solid {t['borde']};border-radius:20px;padding:.22rem .65rem;font-size:.68rem;color:{t['acento1']};font-weight:600;margin:.15rem}}
.oro-footer{{text-align:center;font-size:.68rem;color:{t['texto3']};padding:.5rem 0;opacity:.7}}
</style>"""

# ══════════════════════════════════════════════════════════════
# SESSION STATE
# ══════════════════════════════════════════════════════════════
_defaults = {
    "texto_extraido":"","nombre_archivo":"","archivo_bytes":None,"resumen_data":None,
    "historial_chat":[],"cambios_aplicados":None,"archivo_tipo":"","lista_cambios":[],
    "texto_corregido":"","generando_resumen":False, "preview_cambio":None,
    "edicion_counter":0, "tema":"noche", "idioma":"es", "historial_versiones":[],
    "ejecutar_evaluacion":False, "resultado_evaluacion":None, "scroll_to":None
}
for k,v in _defaults.items():
    if k not in st.session_state: st.session_state[k] = v

st.markdown(_get_all_css(st.session_state.get("tema","noche")), unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════
# FUNCIONES IA & UTILITARIAS
# ══════════════════════════════════════════════════════════════
def T(key):
    lang = st.session_state.get("idioma","es")
    return _TXT.get(lang,_TXT["es"]).get(key, "")

# CORRECCIÓN 5: safe_text para evitar "muerte" de tildes en PDF
def safe_text(t):
    if not t: return ""
    # Mapeo manual de caracteres problemáticos para FPDF/MuPDF
    rep = {"\u2013":"-", "\u2014":"-", "\u201c":'"', "\u201d":'"', "\u2018":"'", "\u2019":"'"}
    for k, v in rep.items(): t = t.replace(k, v)
    return str(t).encode('latin-1','replace').decode('latin-1')

def llamar_ia(prompt):
    # Se mantiene el fallback solicitado por el usuario
    for modelo in ["gemini-2.5-flash", "gemini-2.0-flash", "gemini-1.5-flash"]:
        try:
            model = genai.GenerativeModel(modelo)
            return model.generate_content(prompt).text
        except: continue
    return None

def extraer_json_seguro(texto, es_lista=False):
    t = texto.replace("```json","").replace("```","").strip()
    ini=t.find("[" if es_lista else "{"); fin=t.rfind("]" if es_lista else "}")+1
    if ini!=-1 and fin>0:
        try: return json.loads(t[ini:fin], strict=False)
        except: pass
    return None

# CORRECCIÓN 6: Regex más robusto para cambios manuales
def solicitar_cambios(instruccion, texto_doc=""):
    ctx = f"\n\nDOC:\n{texto_doc[:3000]}"
    prompt = (f"Asistente editorial. Devuelve SOLO JSON array:\n"
              f"[{{\"buscar\":\"texto exacto\", \"reemplazar\":\"texto nuevo\"}}]\n"
              f"Instrucción: {instruccion}{ctx}")
    r = llamar_ia(prompt)
    if r:
        res = extraer_json_seguro(r, es_lista=True)
        if res: return res
    # Fallback regex mejorado
    m = re.search(r"(?i)(?:cambia|reemplaza)\s+['\"]?(.+?)['\"]?\s+(?:por|a)\s+['\"]?(.+?)['\"]?$", instruccion.strip())
    if m: return [{"buscar": m.group(1).strip(), "reemplazar": m.group(2).strip()}]
    return None

# ══════════════════════════════════════════════════════════════
# MOTORES DE REEMPLAZO (EL "QUIRÓFANO")
# ══════════════════════════════════════════════════════════════
def reemplazar_docx_preservando_formato(archivo_bytes, cambios):
    doc=Document(BytesIO(archivo_bytes)); conteo=0
    for c in cambios:
        buscar=str(c["buscar"]); reemplazar=str(c["reemplazar"])
        regex=re.compile(re.escape(buscar), re.IGNORECASE)
        for p in doc.paragraphs:
            if regex.search(p.text):
                for run in p.runs:
                    if regex.search(run.text):
                        run.text, n = regex.subn(reemplazar, run.text)
                        conteo += n
        for t in doc.tables:
            for row in t.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        if regex.search(p.text):
                            for run in p.runs:
                                if regex.search(run.text):
                                    run.text, n = regex.subn(reemplazar, run.text)
                                    conteo += n
    buf=BytesIO(); doc.save(buf); return buf.getvalue(), conteo

def reemplazar_xlsx_preservando_formato(archivo_bytes, cambios):
    # CORRECCIÓN 1: data_only=False preserva fórmulas al guardar
    wb=openpyxl.load_workbook(BytesIO(archivo_bytes), data_only=False); conteo=0
    for c in cambios:
        buscar=str(c["buscar"]); rv=str(c["reemplazar"])
        regex=re.compile(re.escape(buscar), re.IGNORECASE)
        for sheet in wb.worksheets:
            for row in sheet.iter_rows():
                for cell in row:
                    if cell.value and isinstance(cell.value, str) and regex.search(cell.value):
                        cell.value, n = regex.subn(rv, cell.value)
                        conteo += n
    buf=BytesIO(); wb.save(buf); return buf.getvalue(), conteo

# CORRECCIÓN 4: Recuperación del motor quirúrgico de PDF
def reemplazar_pdf_original(archivo_bytes, cambios):
    if not PYMUPDF_OK: return archivo_bytes, 0
    doc=fitz.open(stream=archivo_bytes, filetype="pdf"); conteo=0
    for c in cambios:
        buscar=str(c["buscar"]).strip(); reemplazar=str(c["reemplazar"]).strip()
        for pagina in doc:
            instancias=pagina.search_for(buscar)
            if not instancias: continue
            bloques=pagina.get_text("dict")["blocks"]
            for rect in instancias:
                # Detectar estilo original para que no salga "feo"
                f_size=11.0; f_name="helv"; col=(0,0,0)
                for b in bloques:
                    for line in b.get("lines", []):
                        for span in line.get("spans", []):
                            if buscar.lower() in span["text"].lower():
                                f_size=span.get("size", 11.0)
                                ci=span.get("color", 0); col=(((ci>>16)&0xFF)/255, ((ci>>8)&0xFF)/255, (ci&0xFF)/255)
                                break
                # Borrar texto original sin dejar parche de color
                pagina.add_redact_annot(rect); pagina.apply_redactions(images=0, graphics=0)
                # Insertar texto nuevo ajustado a la línea base original
                pagina.insert_text(fitz.Point(rect.x0, rect.y1 - (rect.height*0.15)), reemplazar, fontname=f_name, fontsize=f_size, color=col)
                conteo+=1
    buf=BytesIO(); doc.save(buf); doc.close(); return buf.getvalue(), conteo

# ══════════════════════════════════════════════════════════════
# PANEL DE CONTROL & UI
# ══════════════════════════════════════════════════════════════
# (Se asume la configuración previa de uploader...)
st.markdown('<div class="oro-header"><div class="oro-title">Oro Asistente</div></div>', unsafe_allow_html=True)
archivo_subido = st.file_uploader("📎 Sube tu archivo", type=["docx","xlsx","pdf"])

if archivo_subido:
    if archivo_subido.name != st.session_state.nombre_archivo:
        # Lógica de carga...
        st.session_state.archivo_bytes = archivo_subido.read()
        st.session_state.nombre_archivo = archivo_subido.name
        st.session_state.archivo_tipo = archivo_subido.name.split(".")[-1].lower()
        # CORRECCIÓN 2: No usar dict.fromkeys en tablas para no borrar datos legítimos repetidos
        texto = ""
        if st.session_state.archivo_tipo == "docx":
            doc = Document(BytesIO(st.session_state.archivo_bytes))
            partes = [p.text for p in doc.paragraphs if p.text.strip()]
            for t in doc.tables:
                for row in t.rows:
                    celdas = [c.text.strip() for c in row.cells] # Sin dict.fromkeys
                    if any(celdas): partes.append(" | ".join(celdas))
            texto = "\n".join(partes)
        elif st.session_state.archivo_tipo == "xlsx":
            wb = openpyxl.load_workbook(BytesIO(st.session_state.archivo_bytes), data_only=True)
            for s in wb.worksheets:
                for r in s.iter_rows(values_only=True):
                    linea = " | ".join([str(c) for c in r if c is not None])
                    if linea.strip(): texto += linea + "\n"
        st.session_state.texto_extraido = texto
        st.rerun()

# ══════════════════════════════════════════════════════════════
# VISTA PREVIA (RECUPERADA)
# ══════════════════════════════════════════════════════════════
if st.session_state.texto_extraido:
    texto_activo = st.session_state.texto_corregido if st.session_state.texto_corregido else st.session_state.texto_extraido
    
    # ── Muestra la Vista Previa si hay un cambio pendiente ──
    if st.session_state.preview_cambio:
        st.markdown('<div id="seccion-preview"></div>', unsafe_allow_html=True)
        preview = st.session_state.preview_cambio
        st.markdown('<div class="info-box">👁️ Vista previa del cambio</div>', unsafe_allow_html=True)
        
        for c in preview:
            idx = texto_activo.lower().find(c["buscar"].lower())
            if idx != -1:
                ini=max(0, idx-40); fin=min(len(texto_activo), idx+len(c["buscar"])+40)
                antes = texto_activo[ini:idx]; despues = texto_activo[idx+len(c["buscar"]):fin]
                st.markdown(f"""<div style="background:#fff1f1; border:1px solid #fecaca; padding:.5rem; border-radius:8px; font-size:.75rem; margin-bottom:.3rem">
                    <span style="color:#991b1b; font-family:monospace">...{antes}<mark style="background:#fee2e2; color:#991b1b; font-weight:bold">{c['buscar']}</mark>{despues}...</span>
                </div>
                <div style="background:#f0fdf4; border:1px solid #bbf7d0; padding:.5rem; border-radius:8px; font-size:.75rem">
                    <span style="color:#166534; font-family:monospace">...{antes}<mark style="background:#dcfce7; color:#166534; font-weight:bold">{c['reemplazar']}</mark>{despues}...</span>
                </div>""", unsafe_allow_html=True)
        
        c1, c2 = st.columns(2)
        with c1:
            if st.button("✅ Confirmar", use_container_width=True):
                # Aplicar cambio con el motor correspondiente
                bytes_org = st.session_state.cambios_aplicados or st.session_state.archivo_bytes
                tipo = st.session_state.archivo_tipo
                if tipo == "docx": res, _ = reemplazar_docx_preservando_formato(bytes_org, preview)
                elif tipo == "xlsx": res, _ = reemplazar_xlsx_preservando_formato(bytes_org, preview)
                elif tipo == "pdf": res, _ = reemplazar_pdf_original(bytes_org, preview)
                
                # Actualizar estados
                st.session_state.cambios_aplicados = res
                st.session_state.texto_corregido = texto_activo.replace(preview[0]["buscar"], preview[0]["reemplazar"])
                st.session_state.lista_cambios.append(preview[0])
                st.session_state.preview_cambio = None
                st.rerun()
        with c2:
            if st.button("❌ Cancelar", use_container_width=True):
                st.session_state.preview_cambio = None; st.rerun()

    # Chat y entrada de comandos...
    for m in st.session_state.historial_chat:
        with st.chat_message("user" if m["rol"]=="Usuario" else "assistant"): st.write(m["texto"])

    entrada = st.chat_input("Escribe un cambio...")
    if entrada:
        st.session_state.historial_chat.append({"rol":"Usuario", "texto":entrada})
        cambios = solicitar_cambios(entrada, texto_activo)
        if cambios: st.session_state.preview_cambio = cambios
        else: st.session_state.historial_chat.append({"rol":"Asistente", "texto":"No entendí el cambio."})
        st.rerun()

    if st.session_state.cambios_aplicados:
        st.download_button("📥 Descargar corregido", st.session_state.cambios_aplicados, "corregido."+st.session_state.archivo_tipo, use_container_width=True)

_TXT = {"es": {"analizar":"⚡ Analizar", "evaluar":"🔍 Evaluar", "chat_placeholder":"Escribe un cambio..."}, "en": {}}
genai.configure(api_key=st.secrets["LLAVE_GEMINI"])
st.markdown(f"<p class='oro-footer'>🏆 Oro Asistente v3.3 · Mejorado</p>", unsafe_allow_html=True)
