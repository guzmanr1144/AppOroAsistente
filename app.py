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
# CSS — COMPLETO POR TEMA
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
html,body,[class*="css"]{{font-family:'Inter',sans-serif!important}}
.stApp{{background:linear-gradient(145deg,{t['bg2']} 0%,{t['bg1']} 40%,{t['bg3']} 100%)!important}}
.main .block-container{{padding:.8rem .9rem 5rem .9rem!important;max-width:460px!important;margin:0 auto!important}}
.oro-header{{text-align:center;padding:1.6rem 0 .5rem}}
.oro-title{{font-size:1.9rem;font-weight:900;background:{t['titulo_grad']}!important;-webkit-background-clip:text!important;-webkit-text-fill-color:transparent!important;background-clip:text!important}}
.file-badge{{display:flex;align-items:center;gap:.75rem;background:{t['card']};border:1px solid {t['borde']};border-radius:18px;padding:.85rem 1rem;margin:.5rem 0}}
.summary-card{{background:{t['card']};border-left:4px solid {t['acento1']};border-radius:18px;padding:1.1rem 1.2rem;margin:.7rem 0;color:{t['texto2']}}}
.info-box{{background:#f0fdf4;border:1px solid #86efac;border-radius:12px;padding:.75rem 1rem;color:#166534;margin:.5rem 0}}
.warn-box{{background:#fffbeb;border:1px solid #fcd34d;border-radius:12px;padding:.75rem 1rem;color:#92400e;margin:.5rem 0}}
.stButton>button{{background:{t['card']}!important;color:{t['texto2']}!important;border:1.5px solid {t['borde']}!important;border-radius:12px!important;width:100%!important;min-height:3rem!important}}
.chip{{display:inline-block;background:{t['bg2']};border:1px solid {t['borde']};border-radius:20px;padding:.22rem .65rem;font-size:.68rem;color:{t['acento1']};margin:.15rem}}
.oro-footer{{text-align:center;font-size:.68rem;color:{t['texto3']};padding:1rem 0}}
</style>"""

# ══════════════════════════════════════════════════════════════
# TRADUCCIONES Y SESSION STATE
# ══════════════════════════════════════════════════════════════
_TXT = {
    "es": {
        "analizar":"⚡ Analizar","evaluar":"🔍 Evaluar","ver_doc":"👁 Ver documento","analizando":"🧠 Analizando...",
        "procesando":"🧠 Procesando...","cargando":"📖 Cargando...","confirmar":"✅ Confirmar","cancelar":"❌ Cancelar",
        "antes":"Antes","despues":"Después","chat_placeholder":"✍️ Escribe un cambio o pregunta...",
        "descargar_corregido":"📥 Descargar corregido"
    }
}

def T(key): return _TXT.get(st.session_state.idioma, _TXT["es"]).get(key, "")

_defaults = {
    "texto_extraido":"","nombre_archivo":"","archivo_bytes":None,"resumen_data":None,
    "historial_chat":[],"cambios_aplicados":None,"archivo_tipo":"","lista_cambios":[],
    "texto_corregido":"","generando_resumen":False, "preview_cambio":None,
    "tema":"noche", "idioma":"es", "historial_versiones":[]
}
for k,v in _defaults.items():
    if k not in st.session_state: st.session_state[k] = v

st.markdown(_get_all_css(st.session_state.tema), unsafe_allow_html=True)

# ══════════════════════════════════════════════════════════════
# MOTORES DE IA Y TRATAMIENTO DE DATOS
# ══════════════════════════════════════════════════════════════
try:
    genai.configure(api_key=st.secrets["LLAVE_GEMINI"])
except: st.error("🔑 Error: Falta LLAVE_GEMINI en secrets")

def llamar_ia(prompt):
    for modelo in ["gemini-1.5-flash", "gemini-2.0-flash"]:
        try:
            return genai.GenerativeModel(modelo).generate_content(prompt).text
        except: continue
    return None

def extraer_json_seguro(texto, es_lista=False):
    t = texto.replace("```json","").replace("```","").strip()
    ini=t.find("[" if es_lista else "{"); fin=t.rfind("]" if es_lista else "}")+1
    if ini!=-1 and fin>0:
        try: return json.loads(t[ini:fin], strict=False)
        except: return None
    return None

def safe_text(t):
    if not t: return ""
    rep = {"\u2013":"-", "\u2014":"-", "\u201c":'"', "\u201d":'"', "\u2018":"'", "\u2019":"'"}
    for k, v in rep.items(): t = t.replace(k, v)
    return str(t).encode('latin-1','replace').decode('latin-1')

# ══════════════════════════════════════════════════════════════
# MOTORES DE REEMPLAZO (EL QUIRÓFANO)
# ══════════════════════════════════════════════════════════════
def reemplazar_docx_preservando_formato(archivo_bytes, cambios):
    doc=Document(BytesIO(archivo_bytes)); conteo=0
    for c in cambios:
        buscar, reemplazar = str(c["buscar"]), str(c["reemplazar"])
        regex=re.compile(re.escape(buscar), re.IGNORECASE)
        # Procesar párrafos y tablas
        elementos = list(doc.paragraphs)
        for t in doc.tables:
            for r in t.rows:
                for cell in r.cells: elementos += list(cell.paragraphs)
        for p in elementos:
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

def reemplazar_pdf_original(archivo_bytes, cambios):
    if not PYMUPDF_OK: return archivo_bytes, 0
    doc=fitz.open(stream=archivo_bytes, filetype="pdf"); conteo=0
    for c in cambios:
        buscar, reemplazar = str(c["buscar"]).strip(), str(c["reemplazar"]).strip()
        for pagina in doc:
            instancias=pagina.search_for(buscar)
            if not instancias: continue
            bloques=pagina.get_text("dict")["blocks"]
            for rect in instancias:
                f_size=11.0; f_name="helv"; col=(0,0,0)
                for b in bloques:
                    for line in b.get("lines", []):
                        for span in line.get("spans", []):
                            if buscar.lower() in span["text"].lower():
                                f_size=span.get("size", 11.0)
                                ci=span.get("color", 0); col=(((ci>>16)&0xFF)/255, ((ci>>8)&0xFF)/255, (ci&0xFF)/255)
                                break
                pagina.add_redact_annot(rect); pagina.apply_redactions(images=0, graphics=0)
                pagina.insert_text(fitz.Point(rect.x0, rect.y1 - (rect.height*0.15)), reemplazar, fontname=f_name, fontsize=f_size, color=col)
                conteo+=1
    buf=BytesIO(); doc.save(buf); doc.close(); return buf.getvalue(), conteo

# ══════════════════════════════════════════════════════════════
# INTERFAZ Y LÓGICA PRINCIPAL
# ══════════════════════════════════════════════════════════════
st.markdown('<div class="oro-header"><div class="oro-title">Oro Asistente</div></div>', unsafe_allow_html=True)
archivo_subido = st.file_uploader("📎 Sube tu archivo (Word, Excel o PDF)", type=["docx","xlsx","pdf"])

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
                for row in t.rows:
                    celdas = [c.text.strip() for c in row.cells] # Sin borrar duplicados
                    if any(celdas): texto += "\n" + " | ".join(celdas)
        elif st.session_state.archivo_tipo == "xlsx":
            wb = openpyxl.load_workbook(BytesIO(st.session_state.archivo_bytes), data_only=True)
            for s in wb.worksheets:
                for r in s.iter_rows(values_only=True):
                    l = " | ".join([str(c) for c in r if c is not None])
                    if l.strip(): texto += l + "\n"
        elif st.session_state.archivo_tipo == "pdf":
            reader = PyPDF2.PdfReader(BytesIO(st.session_state.archivo_bytes))
            for p in reader.pages: texto += (p.extract_text() or "") + "\n"
        
        st.session_state.texto_extraido = texto
        st.session_state.texto_corregido = ""
        st.session_state.cambios_aplicados = None
        st.session_state.resumen_data = None
        st.session_state.historial_chat = []
        st.rerun()

# ══════════════════════════════════════════════════════════════
# FLUJO DE TRABAJO CON EL DOCUMENTO
# ══════════════════════════════════════════════════════════════
if st.session_state.texto_extraido:
    texto_activo = st.session_state.texto_corregido if st.session_state.texto_corregido else st.session_state.texto_extraido
    
    st.markdown(f'<div class="file-badge"><b>{st.session_state.nombre_archivo}</b></div>', unsafe_allow_html=True)
    
    # Muestra Vista Previa (Restaurada)
    if st.session_state.preview_cambio:
        preview = st.session_state.preview_cambio
        st.markdown('<div class="info-box">👁️ Vista previa del cambio</div>', unsafe_allow_html=True)
        for c in preview:
            idx = texto_activo.lower().find(c["buscar"].lower())
            if idx != -1:
                ini, fin = max(0, idx-40), min(len(texto_activo), idx+len(c["buscar"])+40)
                st.markdown(f"""<div style="background:#fff1f1;padding:.5rem;border-radius:8px;font-size:.75rem;margin-bottom:.2rem">
                    <small>{T('antes')}</small><br>...{texto_activo[ini:idx]}<mark style="background:#fee2e2;color:#991b1b;font-weight:bold">{c['buscar']}</mark>{texto_activo[idx+len(c['buscar']):fin]}...
                </div>
                <div style="background:#f0fdf4;padding:.5rem;border-radius:8px;font-size:.75rem">
                    <small>{T('despues')}</small><br>...{texto_activo[ini:idx]}<mark style="background:#dcfce7;color:#166534;font-weight:bold">{c['reemplazar']}</mark>{texto_activo[idx+len(c['buscar']):fin]}...
                </div>""", unsafe_allow_html=True)
        
        c1, c2 = st.columns(2)
        with c1:
            if st.button(T("confirmar")):
                bytes_in = st.session_state.cambios_aplicados or st.session_state.archivo_bytes
                tipo = st.session_state.archivo_tipo
                if tipo == "docx": res, _ = reemplazar_docx_preservando_formato(bytes_in, preview)
                elif tipo == "xlsx": res, _ = reemplazar_xlsx_preservando_formato(bytes_in, preview)
                elif tipo == "pdf": res, _ = reemplazar_pdf_original(bytes_in, preview)
                
                st.session_state.cambios_aplicados = res
                st.session_state.texto_corregido = texto_activo.replace(preview[0]["buscar"], preview[0]["reemplazar"])
                st.session_state.preview_cambio = None
                st.rerun()
        with c2:
            if st.button(T("cancelar")): st.session_state.preview_cambio = None; st.rerun()

    # Botones Analizar / Evaluar
    ba, be = st.columns(2)
    with ba:
        if st.button(T("analizar")):
            with st.spinner(T("analizando")):
                p = f"Analiza este doc y devuelve SOLO JSON: {{'titulo':'','resumen_ejecutivo':''}}\n\nDoc: {texto_activo[:3000]}"
                st.session_state.resumen_data = extraer_json_seguro(llamar_ia(p))
                st.rerun()
    with be:
        if st.button(T("evaluar")): st.info("🔍 Próximamente: Evaluación profunda de calidad.")

    if st.session_state.resumen_data:
        d = st.session_state.resumen_data
        st.markdown(f'<div class="summary-card"><b>{d.get("titulo")}</b><br>{d.get("resumen_ejecutivo")}</div>', unsafe_allow_html=True)

    # CHAT INTERACTIVO
    for m in st.session_state.historial_chat:
        with st.chat_message("user" if m["rol"]=="Usuario" else "assistant"): st.write(m["texto"])

    entrada = st.chat_input(T("chat_placeholder"))
    if entrada:
        st.session_state.historial_chat.append({"rol":"Usuario", "texto":entrada})
        with st.spinner(T("procesando")):
            p = f"Asistente editorial. Devuelve SOLO JSON array: [{{'buscar':'', 'reemplazar':''}}]. Si no es edición, responde normal. Instrucción: {entrada}. Contexto: {texto_activo[:2000]}"
            resp = llamar_ia(p)
            nuevos = extraer_json_seguro(resp, es_lista=True)
            if nuevos: st.session_state.preview_cambio = nuevos
            else: st.session_state.historial_chat.append({"rol":"Asistente", "texto": resp if resp else "No pude procesar la solicitud."})
        st.rerun()

    if st.session_state.cambios_aplicados:
        st.download_button("📥 Descargar documento corregido", st.session_state.cambios_aplicados, 
                           f"corregido_{st.session_state.nombre_archivo}", use_container_width=True)

else:
    st.markdown('<div style="text-align:center;padding:4rem 1rem">🏆 <b>Oro Asistente</b><br>Sube un archivo para analizarlo, editarlo o preguntar sobre él.</div>', unsafe_allow_html=True)

st.markdown(f"<p class='oro-footer'>🏆 Oro Asistente v3.4 · IDANZ Deporte Popular</p>", unsafe_allow_html=True)
