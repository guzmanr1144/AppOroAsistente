import os, json, ast, re, warnings
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

@st.cache_data(show_spinner=False)
def _get_all_css(tema_key="verde"):
    _T = {
        "verde": {"bg1":"#010c06","bg2":"#021008","card":"#041208","card2":"#051a0c","borde":"#0a3d1a","borde2":"#0f5225","acento1":"#10b981","acento2":"#34d399","titulo_grad":"linear-gradient(135deg,#10b981,#34d399,#6ee7b7,#10b981)","texto":"#d1fae5","texto2":"#34d399","texto3":"#065f46"},
        "oscuro":{"bg1":"#0a0e1a","bg2":"#0d1525","card":"#111827","card2":"#162032","borde":"#1e3a5f","borde2":"#2a4a6b","acento1":"#3b82f6","acento2":"#60a5fa","titulo_grad":"linear-gradient(135deg,#fbbf24,#f59e0b,#fde68a,#f59e0b)","texto":"#e2e8f0","texto2":"#93c5fd","texto3":"#4b6080"},
        "azul":  {"bg1":"#020818","bg2":"#030d24","card":"#041230","card2":"#061840","borde":"#0c3a7a","borde2":"#1050aa","acento1":"#38bdf8","acento2":"#7dd3fc","titulo_grad":"linear-gradient(135deg,#38bdf8,#7dd3fc,#bae6fd)","texto":"#e0f2fe","texto2":"#7dd3fc","texto3":"#1e5a8a"},
        "rosa":  {"bg1":"#120008","bg2":"#1a000f","card":"#1a000f","card2":"#280018","borde":"#7c0040","borde2":"#9d0050","acento1":"#f472b6","acento2":"#f9a8d4","titulo_grad":"linear-gradient(135deg,#f472b6,#f9a8d4,#fce7f3)","texto":"#fce7f3","texto2":"#f9a8d4","texto3":"#7c0040"},
        "ambar": {"bg1":"#0f0800","bg2":"#180d00","card":"#1a0e00","card2":"#251500","borde":"#78350f","borde2":"#92400e","acento1":"#f59e0b","acento2":"#fbbf24","titulo_grad":"linear-gradient(135deg,#fbbf24,#fde68a,#f59e0b)","texto":"#fef3c7","texto2":"#fbbf24","texto3":"#78350f"},
    }
    t = _T.get(tema_key, _T["verde"])
    return f"""<style>
@import url('https://fonts.googleapis.com/css2?family=Outfit:wght@300;400;500;600;700;800&family=JetBrains+Mono:wght@400;600&display=swap');
html,body,[class*="css"]{{font-family:'Outfit',sans-serif!important;-webkit-tap-highlight-color:transparent}}
.stApp{{background:linear-gradient(160deg,{t['bg1']} 0%,{t['bg2']} 50%,{t['bg1']} 100%)!important;min-height:100vh}}
.main .block-container{{padding:1rem 1rem 5rem 1rem!important;max-width:480px!important;margin:0 auto!important;background:transparent!important}}
#MainMenu,footer,header{{visibility:hidden}}[data-testid="stToolbar"]{{display:none}}
.oro-header{{text-align:center;padding:1.2rem 0 0.3rem}}
.oro-logo{{font-size:2.8rem;line-height:1;filter:drop-shadow(0 0 20px rgba(16,185,129,0.6));animation:pulse-glow 3s ease-in-out infinite}}
@keyframes pulse-glow{{0%,100%{{filter:drop-shadow(0 0 12px rgba(16,185,129,0.4))}}50%{{filter:drop-shadow(0 0 28px rgba(16,185,129,0.9))}}}}
.oro-title{{font-size:1.8rem;font-weight:800;background:{t['titulo_grad']}!important;-webkit-background-clip:text!important;-webkit-text-fill-color:transparent!important;background-clip:text!important;letter-spacing:-0.02em;margin:.1rem 0}}
.oro-subtitle{{color:{t['texto3']};font-size:.78rem;font-weight:400;letter-spacing:.08em;text-transform:uppercase}}
[data-testid="stFileUploader"]{{background:transparent!important;border:none!important}}
[data-testid="stFileUploader"]>div{{background:{t['card']}!important;border:2px dashed {t['borde']}!important;border-radius:20px!important;padding:1.2rem!important}}
[data-testid="stFileUploader"] label{{color:{t['acento2']}!important;font-weight:600!important;font-size:.95rem!important}}
.file-badge{{display:flex;align-items:center;gap:.8rem;background:{t['card']}!important;border:1px solid {t['borde']}!important;border-radius:16px;padding:.9rem 1.1rem;margin:.6rem 0}}
.file-icon{{font-size:1.8rem}}.file-info-name{{color:{t['texto']}!important;font-weight:600;font-size:.88rem;word-break:break-all}}
.file-info-stats{{color:{t['texto3']}!important;font-size:.72rem;margin-top:.15rem}}
.summary-card{{background:{t['card2']}!important;border:1px solid {t['borde2']}!important;border-left:4px solid {t['acento1']}!important;border-radius:18px;padding:1.2rem;margin:.8rem 0;color:{t['texto2']}!important;line-height:1.7;font-size:.9rem}}
.summary-card-title{{color:{t['acento2']}!important;font-size:1rem;font-weight:700;margin-bottom:.6rem}}
.metrics-grid{{display:grid;grid-template-columns:1fr 1fr;gap:.5rem;margin:.7rem 0}}
.metric-pill{{background:{t['card']}!important;border:1px solid {t['borde']}!important;border-radius:14px;padding:.7rem .9rem;text-align:center}}
.metric-pill-label{{color:{t['texto3']}!important;font-size:.65rem;font-weight:600;text-transform:uppercase;letter-spacing:.05em}}
.metric-pill-value{{color:{t['texto']}!important;font-size:1.15rem;font-weight:800;margin-top:.15rem;font-family:'JetBrains Mono',monospace}}
.tags-wrap{{display:flex;flex-wrap:wrap;gap:.35rem;margin:.5rem 0}}
.tag{{background:{t['card']}!important;color:{t['acento2']}!important;border:1px solid {t['borde']}!important;border-radius:20px;padding:.3rem .7rem;font-size:.72rem;font-weight:500}}
.hallazgo-card{{background:{t['card']}!important;border:1px solid {t['borde']};border-left:4px solid {t['acento1']}!important;border-radius:14px;padding:.9rem 1rem;color:{t['texto2']}!important;font-size:.83rem;margin:.7rem 0;line-height:1.6}}
.stButton>button{{background:{t['card']}!important;color:{t['texto3']}!important;border:1px solid {t['borde']}!important;border-radius:14px!important;font-weight:600!important;font-size:.85rem!important;min-height:3.2rem!important;width:100%!important;transition:all .15s!important;font-family:'Outfit',sans-serif!important}}
.stButton>button:hover{{border-color:{t['acento1']}!important;color:{t['acento2']}!important;background:{t['card2']}!important}}
.stButton>button:active{{transform:scale(.96)!important}}
.btn-evaluar>button{{background:linear-gradient(135deg,#065f46,#10b981)!important;color:white!important;border:none!important;font-weight:700!important;font-size:.95rem!important;min-height:3.5rem!important;box-shadow:0 4px 18px rgba(16,185,129,.35)!important}}
.btn-analizar>button{{background:linear-gradient(135deg,#1e3a5f,#2563eb)!important;color:white!important;border:none!important;font-weight:700!important}}
[data-testid="stDownloadButton"]>button{{background:linear-gradient(135deg,#065f46,#059669)!important;color:white!important;border:none!important;border-radius:12px!important;font-weight:700!important;height:2.8rem!important;width:100%!important}}
.section-title{{color:{t['texto']}!important;font-size:1rem;font-weight:700;margin:1rem 0 .4rem;display:flex;align-items:center;gap:.4rem}}
.info-box{{background:#052e16;border:1px solid #15803d;border-radius:12px;padding:.8rem 1rem;color:#4ade80;font-size:.85rem;margin:.5rem 0}}
.warn-box{{background:#1c1003;border:1px solid #b45309;border-radius:12px;padding:.8rem 1rem;color:#fbbf24;font-size:.85rem;margin:.5rem 0}}
.cambio-item{{background:{t['card']}!important;border:1px solid {t['borde']}!important;border-radius:10px;padding:.55rem .85rem;margin:.25rem 0;font-size:.8rem;font-family:'JetBrains Mono',monospace;display:flex;align-items:center;gap:.5rem}}
.cambio-num{{color:#4b6080;font-size:.68rem;min-width:1.1rem}}.cambio-arrow{{color:#f59e0b}}
[data-testid="stExpander"]{{background:{t['card']}!important;border:1px solid {t['borde']}!important;border-radius:14px!important}}
[data-testid="stChatInput"] textarea{{background:{t['card']}!important;border:2px solid {t['borde']}!important;border-radius:16px!important;color:{t['texto']}!important;font-family:'Outfit',sans-serif!important;font-size:.95rem!important}}
[data-testid="stChatInput"] textarea:focus{{border-color:{t['acento1']}!important;box-shadow:0 0 0 3px rgba(16,185,129,.15)!important}}
.oro-divider{{height:1px;background:linear-gradient(90deg,transparent,{t['borde']},transparent);margin:1rem 0}}
.empty-state{{text-align:center;padding:2.5rem 1rem}}
.empty-icon{{font-size:3.5rem;margin-bottom:.8rem;opacity:.5}}.empty-title{{color:#374151;font-size:.95rem;font-weight:600}}
.empty-hint{{color:#1f2937;font-size:.78rem;margin-top:.35rem;line-height:1.6}}
.format-badges{{display:flex;justify-content:center;gap:.5rem;margin-top:.8rem}}
.format-badge{{background:#111827;border:1px solid #1f2937;border-radius:8px;padding:.25rem .65rem;color:#374151;font-size:.72rem;font-family:'JetBrains Mono',monospace}}
.oro-footer{{text-align:center;font-size:.7rem;color:#1f2937;padding:.4rem 0}}
[data-testid="stSidebar"]{{background:{t['bg1']}!important}}
.guia-tooltip{{background:linear-gradient(135deg,{t['card']},{t['card2']});border:1px solid {t['borde']};border-left:3px solid {t['acento1']};border-radius:14px;padding:1rem 1.1rem;margin:.6rem 0;animation:fadeIn .4s ease}}
@keyframes fadeIn{{from{{opacity:0;transform:translateY(-6px)}}to{{opacity:1;transform:translateY(0)}}}}
.guia-titulo{{color:{t['acento2']};font-weight:700;font-size:.9rem;margin-bottom:.3rem}}
.guia-texto{{color:{t['texto3']};font-size:.78rem;line-height:1.55}}
.chat-placeholder{{text-align:center;padding:1.2rem .5rem .8rem;background:{t['card']};border:1px dashed {t['borde']};border-radius:18px;margin:.5rem 0}}
</style>"""

for key, val in {
    "texto_extraido":"","nombre_archivo":"","archivo_bytes":None,"resumen_data":None,
    "historial_chat":[],"cambios_aplicados":None,"archivo_tipo":"","lista_cambios":[],
    "texto_modificado":"","generando_resumen":False,"resumen_error":False,
    "tab_activa":"resumen","tema":"verde","preview_cambio":None,"edicion_counter":0,
    "texto_corregido":"","guia_paso":0,"guia_vista":False,
}.items():
    if key not in st.session_state:
        st.session_state[key] = val

st.markdown(_get_all_css(st.session_state.get("tema","verde")), unsafe_allow_html=True)

try:
    LLAVE_GEMINI = st.secrets["LLAVE_GEMINI"]
    genai.configure(api_key=LLAVE_GEMINI)
except Exception as e:
    st.error(f"🔑 Error configurando la IA: {e}")
    st.stop()

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

with st.sidebar:
    st.markdown("### ⚙️ Ajustes")
    tema_opc = {"🌿 Verde":"verde","🌑 Oscuro":"oscuro","🌊 Azul":"azul","🌸 Rosa":"rosa","⚡ Ámbar":"ambar"}
    sel = st.selectbox("🎨 Tema", list(tema_opc.keys()),
        index=list(tema_opc.values()).index(st.session_state.get("tema","verde")),
        label_visibility="collapsed")
    if tema_opc[sel] != st.session_state.tema:
        st.session_state.tema = tema_opc[sel]
        st.rerun()
    st.markdown("---")
    # ── Guía asistida en sidebar ──
    paso_sb = st.session_state.get("guia_paso", 0)
    if paso_sb > 0 and not st.session_state.get("guia_vista", False):
        guias_sb = {
            1: ("🎉", "Paso 1 — Analiza",
                "Tu archivo está listo.\n\nToca **⚡ Analizar** para que la IA extraiga métricas, puntos clave y un resumen inteligente."),
            2: ("📊", "Paso 2 — Revisa",
                "El resumen ya está arriba.\n\nDescarga el informe en Word, Excel o PDF con los botones que aparecen debajo del resumen."),
            3: ("💬", "Paso 3 — El chat",
                "El chat de abajo es tu asistente.\n\nPuedes escribir:\n• *cambia X por Y*\n• *agrega el teléfono a Juan*\n• *¿cuántas personas hay?*"),
        }
        if paso_sb in guias_sb:
            ico_sb, tit_sb, desc_sb = guias_sb[paso_sb]
            st.markdown(f"### {ico_sb} {tit_sb}")
            st.info(desc_sb)
            if st.button("👍 Entendido", use_container_width=True, key="guia_ok_sb"):
                if paso_sb >= 3:
                    st.session_state.guia_vista = True
                    st.session_state.guia_paso = 0
                else:
                    st.session_state.guia_paso = paso_sb + 1
                st.rerun()
            if st.button("✕ Saltar guía", use_container_width=True, key="guia_skip_sb"):
                st.session_state.guia_vista = True
                st.session_state.guia_paso = 0
                st.rerun()
            st.markdown(f"*Paso {paso_sb} de 3*")
    elif not st.session_state.get("texto_extraido"):
        st.markdown("### 👋 Bienvenido")
        st.info("Sube un archivo Word, Excel o PDF para empezar.\n\nLa IA lo analizará y podrás editarlo con comandos en lenguaje natural.")
    else:
        st.markdown("### 💡 Recuerda")
        st.markdown("""
**Analizar** → Resumen con IA  
**Evaluar** → Detectar errores  
**Chat** → Editar y preguntar  
        """)
    st.markdown("---")
    st.caption("Oro Asistente v3")

st.markdown("""
<div class="oro-header">
    <div class="oro-logo">🏆</div>
    <div class="oro-title">Oro Asistente</div>
    <div class="oro-subtitle">Analiza · Edita · Exporta</div>
</div>
""", unsafe_allow_html=True)

def extraer_json_seguro(texto, es_lista=False):
    t = texto.replace("```json","").replace("```","").strip()
    c1,c2 = ("[","]") if es_lista else ("{","}")
    inicio = t.find(c1); fin = t.rfind(c2)+1
    if inicio != -1 and fin > 0:
        try: return json.loads(t[inicio:fin], strict=False)
        except:
            try: return ast.literal_eval(t[inicio:fin])
            except: pass
    return None

def solicitar_resumen_estructurado(texto):
    prompt = (
        "Eres un analista profesional experto en documentos de cualquier tipo. Analiza y devuelve SOLO un JSON.\n"
        "Identifica el tipo de documento. resumen_ejecutivo: amigable, máximo 3 oraciones.\n"
        '{"titulo":"...","emoji_categoria":"📋","resumen_ejecutivo":"...",'
        '"metricas":{"Clave1":"Valor1"},"puntos_clave":["punto 1"],'
        '"hallazgo_destacado":"observación importante"}\n\n'
        f"DOCUMENTO:\n{texto[:12000]}"
    )
    r = llamar_ia(prompt)
    return extraer_json_seguro(r) if r else None

def extraer_cambio_con_regex(instruccion):
    patrones = [
        r"(?:cambia|reemplaza|sustituye|cambie|reemplaz[ao])\s+['\"]?(.+?)['\"]?\s+(?:por|con|a)\s+['\"]?(.+?)['\"]?\s*$",
        r"['\"](.+?)['\"]\s*(?:→|->|=>|por|con)\s*['\"]?(.+?)['\"]?\s*$",
        r"(.+?)\s*(?:→|->|=>)\s*(.+)",
    ]
    for pat in patrones:
        m = re.search(pat, instruccion.strip(), re.IGNORECASE)
        if m:
            b = m.group(1).strip().strip("'\""); r = m.group(2).strip().strip("'\"")
            if b and r: return [{"buscar":b,"reemplazar":r}]
    return []

def solicitar_cambios(instruccion, texto_doc=""):
    ctx = f"\n\nCONTENIDO ACTUAL DEL DOCUMENTO (usa esto para encontrar el texto exacto):\n{texto_doc[:4000]}" if texto_doc else ""
    prompt = (
        "Eres un asistente experto en edición de documentos.\n"
        f"INSTRUCCIÓN DEL USUARIO: \"{instruccion}\"\n{ctx}\n\n"
        "REGLAS IMPORTANTES:\n"
        "1. REEMPLAZAR: Si dice \'cambia X por Y\' → buscar=X exacto del doc, reemplazar=Y\n"
        "2. AGREGAR DATO A PERSONA/FILA: Si dice \'agrega el número 04XX a Juan Pérez\'\n"
        "   → buscar=\'Juan Pérez\' (el texto exacto como aparece en el doc)\n"
        "   → reemplazar=\'Juan Pérez\t04XX\' (agrega el dato separado por tab o espacio según contexto)\n"
        "   Si la tabla tiene columnas, el número va en la misma fila después del nombre.\n"
        "3. Si dice \'agrega X al final de Y\' → buscar=Y, reemplazar=\'Y X\'\n"
        "4. Si dice \'completa el teléfono/correo/dato de Y con X\' → igual que regla 2\n"
        "5. SIEMPRE usa el texto TAL COMO APARECE en el documento como \'buscar\'\n"
        "6. Si hay múltiples personas/filas → incluye TODOS en el array\n"
        "7. Formato opcional: si pide negrita/cursiva/mayúsculas/color, agrega campo \'formato\'\n\n"
        "Responde SOLO con JSON array (sin explicaciones):\n"
        '[{"buscar":"texto_exacto_del_doc","reemplazar":"texto_nuevo_completo"}]\n'
        "Ejemplo agrega teléfono: buscar=\'María González\', reemplazar=\'María González  04241234567\'"
    )
    r = llamar_ia(prompt)
    if r:
        res = extraer_json_seguro(r, es_lista=True)
        if res and isinstance(res, list):
            v = [c for c in res if isinstance(c,dict) and "buscar" in c and "reemplazar" in c
                 and str(c["buscar"]).strip() and str(c["reemplazar"]).strip() and c["buscar"]!=c["reemplazar"]]
            if v: return v
    return extraer_cambio_con_regex(instruccion)

def preguntar_al_documento(pregunta, texto):
    ctx = "\n".join([f"{m['rol']}: {m['texto']}" for m in st.session_state.historial_chat[-6:]])
    prompt = (
        f"Eres un asistente experto en análisis de documentos.\n"
        f"DOCUMENTO:\n{texto[:10000]}\n\nCONVERSACIÓN:\n{ctx}\n\nPREGUNTA: {pregunta}\n"
        "Responde de forma concisa y directa en español."
    )
    return llamar_ia(prompt) or "No pude procesar tu pregunta."

def detectar_anomalias(texto):
    prompt = (
        "Analiza este documento. Clasifica problemas por gravedad:\n"
        "CRITICO: errores graves | ALTO: datos incorrectos | MEDIO: ortografía/formato | LEVE: mejoras\n"
        "Devuelve SOLO JSON:\n"
        '{"nivel_general":"Excelente/Bueno/Regular/Deficiente","puntaje":85,'
        '"criticos":["..."],"altos":["..."],"medios":["..."],"leves":["..."],'
        '"recomendacion":"..."}\n\n'
        f"DOCUMENTO:\n{texto[:12000]}"
    )
    r = llamar_ia(prompt)
    return extraer_json_seguro(r) if r else None

def exportar_word(texto, resumen_data=None, archivo_bytes=None, archivo_tipo=None, cambios=None):
    zona = pytz.timezone('America/Caracas')
    fecha = datetime.now(zona).strftime('%d de %B de %Y, %I:%M %p')
    cambios = cambios or []
    if archivo_tipo == "docx" and archivo_bytes and cambios:
        resultado, _ = reemplazar_docx_preservando_formato(archivo_bytes, cambios)
        return resultado
    if archivo_tipo == "xlsx" and archivo_bytes:
        doc = Document()
        doc.styles['Normal'].font.name = 'Calibri'
        h = doc.add_heading('', 0); r = h.add_run('Reporte desde Excel')
        r.font.color.rgb = RGBColor(0x1E,0x40,0xAF); r.font.size = Pt(20)
        doc.add_paragraph().add_run(f'Generado: {fecha}').font.size = Pt(9)
        doc.add_paragraph()
        bu = archivo_bytes
        if cambios: bu, _ = reemplazar_xlsx_preservando_formato(archivo_bytes, cambios)
        wb = openpyxl.load_workbook(BytesIO(bu), data_only=True)
        for sheet in wb.worksheets:
            doc.add_heading(f'Hoja: {sheet.title}', level=1)
            filas = [f for f in sheet.iter_rows(values_only=True) if any(c is not None for c in f)]
            if not filas: doc.add_paragraph('(Hoja vacía)'); continue
            nc = max(len(f) for f in filas)
            tb = doc.add_table(rows=len(filas), cols=nc); tb.style = 'Table Grid'
            for i,fila in enumerate(filas):
                for j in range(nc):
                    v = fila[j] if j < len(fila) else ""
                    cell = tb.cell(i,j); cell.text = str(v) if v is not None else ""
                    if i==0:
                        for run in cell.paragraphs[0].runs: run.font.bold = True
            doc.add_paragraph()
        buf = BytesIO(); doc.save(buf); return buf.getvalue()
    doc = Document()
    for section in doc.sections:
        section.top_margin=Inches(0.8); section.bottom_margin=Inches(0.8)
        section.left_margin=Inches(1.0); section.right_margin=Inches(1.0)
    sty=doc.styles['Normal']; sty.font.name='Calibri'; sty.font.size=Pt(11)
    th=doc.add_table(rows=1,cols=1); th.style='Table Grid'
    ch=th.cell(0,0); ch.paragraphs[0].clear()
    rh=ch.paragraphs[0].add_run(resumen_data.get("titulo","INFORME") if resumen_data else "INFORME")
    rh.font.bold=True; rh.font.size=Pt(16); rh.font.color.rgb=RGBColor(0xFF,0xFF,0xFF)
    ch.paragraphs[0].alignment=WD_ALIGN_PARAGRAPH.CENTER
    tc=ch._tc; tcPr=tc.get_or_add_tcPr(); shd=OxmlElement('w:shd')
    shd.set(qn('w:val'),'clear'); shd.set(qn('w:color'),'auto'); shd.set(qn('w:fill'),'1E3A5F')
    tcPr.append(shd); doc.add_paragraph()
    pf=doc.add_paragraph(); rf=pf.add_run(f'Generado: {fecha}')
    rf.font.size=Pt(9); rf.font.color.rgb=RGBColor(0x6B,0x72,0x80)
    pf.alignment=WD_ALIGN_PARAGRAPH.RIGHT; doc.add_paragraph()
    if resumen_data:
        if resumen_data.get("resumen_ejecutivo"):
            tr=doc.add_table(rows=1,cols=1); tr.style='Table Grid'
            cr=tr.cell(0,0); cr.paragraphs[0].clear()
            rr=cr.paragraphs[0].add_run(resumen_data["resumen_ejecutivo"])
            rr.font.size=Pt(10); rr.font.italic=True
            tp2=cr._tc.get_or_add_tcPr(); sh2=OxmlElement('w:shd')
            sh2.set(qn('w:val'),'clear'); sh2.set(qn('w:color'),'auto'); sh2.set(qn('w:fill'),'EFF6FF')
            tp2.append(sh2); doc.add_paragraph()
        if resumen_data.get("metricas"):
            h2=doc.add_heading('Métricas Clave',level=1); h2.runs[0].font.color.rgb=RGBColor(0x1E,0x40,0xAF)
            tm=doc.add_table(rows=1,cols=2); tm.style='Table Grid'
            hdr=tm.rows[0].cells
            for ci,txt in enumerate(['Indicador','Valor']):
                hdr[ci].paragraphs[0].clear(); r=hdr[ci].paragraphs[0].add_run(txt)
                r.font.bold=True; r.font.color.rgb=RGBColor(0xFF,0xFF,0xFF)
                hdr[ci].paragraphs[0].alignment=WD_ALIGN_PARAGRAPH.CENTER
                tph=hdr[ci]._tc.get_or_add_tcPr(); sh=OxmlElement('w:shd')
                sh.set(qn('w:val'),'clear'); sh.set(qn('w:color'),'auto'); sh.set(qn('w:fill'),'1E40AF')
                tph.append(sh)
            for idx,(k,v) in enumerate(resumen_data["metricas"].items()):
                rm=tm.add_row().cells; rm[0].text=str(k); rm[1].text=str(v)
                fill='F8FAFC' if idx%2==0 else 'FFFFFF'
                for ci2 in range(2):
                    tpd=rm[ci2]._tc.get_or_add_tcPr(); shd=OxmlElement('w:shd')
                    shd.set(qn('w:val'),'clear'); shd.set(qn('w:color'),'auto'); shd.set(qn('w:fill'),fill)
                    tpd.append(shd)
            doc.add_paragraph()
        if resumen_data.get("puntos_clave"):
            h3=doc.add_heading('Puntos Clave',level=1); h3.runs[0].font.color.rgb=RGBColor(0x1E,0x40,0xAF)
            for p in resumen_data["puntos_clave"]:
                pb=doc.add_paragraph(style='List Bullet'); pb.add_run(p).font.size=Pt(11)
        if resumen_data.get("hallazgo_destacado"):
            doc.add_paragraph()
            h4=doc.add_heading('💡 Hallazgo',level=1); h4.runs[0].font.color.rgb=RGBColor(0x1E,0x40,0xAF)
            th2=doc.add_table(rows=1,cols=1); th2.style='Table Grid'
            ch2=th2.cell(0,0); ch2.paragraphs[0].clear()
            rh2=ch2.paragraphs[0].add_run(resumen_data["hallazgo_destacado"])
            rh2.font.italic=True; rh2.font.size=Pt(10)
            tph2=ch2._tc.get_or_add_tcPr(); shh=OxmlElement('w:shd')
            shh.set(qn('w:val'),'clear'); shh.set(qn('w:color'),'auto'); shh.set(qn('w:fill'),'F0FDF4')
            tph2.append(shh)
        doc.add_page_break()
    hc=doc.add_heading('Contenido del Documento',level=1); hc.runs[0].font.color.rgb=RGBColor(0x1E,0x40,0xAF)
    for linea in texto.split('\n'):
        ll=linea.strip().replace('*','').replace('#','')
        if ll: p=doc.add_paragraph(ll); p.paragraph_format.space_after=Pt(2)
    buf=BytesIO(); doc.save(buf); return buf.getvalue()

def exportar_excel(texto, resumen_data=None, archivo_bytes=None, archivo_tipo=None, cambios=None):
    cambios = cambios or []
    if archivo_tipo=="xlsx" and archivo_bytes and cambios:
        resultado,_=reemplazar_xlsx_preservando_formato(archivo_bytes,cambios); return resultado
    if archivo_tipo=="docx" and archivo_bytes:
        bu=archivo_bytes
        if cambios: bu,_=reemplazar_docx_preservando_formato(archivo_bytes,cambios)
        wb=openpyxl.Workbook(); wb.remove(wb.active)
        doc_src=Document(BytesIO(bu))
        for i,tabla in enumerate(doc_src.tables):
            ws=wb.create_sheet(title=f"Tabla_{i+1}"); fl=[]
            for row in tabla.rows:
                vistas=set(); fila=[]
                for cell in row.cells:
                    if cell._tc not in vistas: vistas.add(cell._tc); fila.append(cell.text.strip())
                if any(fila): fl.append(fila)
            for ri,fila in enumerate(fl,1):
                for ci,val in enumerate(fila,1):
                    cell=ws.cell(row=ri,column=ci,value=val)
                    if ri==1: cell.fill=PatternFill("solid",fgColor="1E3A5F"); cell.font=Font(color="FFFFFF",bold=True,size=10)
                    else: cell.fill=PatternFill("solid",fgColor="F8FAFC" if ri%2==0 else "FFFFFF"); cell.font=Font(size=10)
                    cell.alignment=Alignment(wrap_text=True,vertical="center")
            for col in ws.columns: ws.column_dimensions[col[0].column_letter].width=22
        if not wb.sheetnames:
            ws=wb.create_sheet("Datos"); ws.cell(1,1,"No se encontraron tablas.")
        buf=BytesIO(); wb.save(buf); return buf.getvalue()
    wb=openpyxl.Workbook(); AO="1E3A5F"; AM="2563EB"; AC="DBEAFE"; BL="FFFFFF"; GC="F8FAFC"
    def hc(ws,row,col,txt,bg=AO,fg=BL,size=12,bold=True):
        c=ws.cell(row=row,column=col,value=txt); c.fill=PatternFill("solid",fgColor=bg)
        c.font=Font(color=fg,bold=bold,size=size); c.alignment=Alignment(horizontal="center",vertical="center",wrap_text=True); return c
    def dc(ws,row,col,txt,bg=BL,bold=False,align="left"):
        c=ws.cell(row=row,column=col,value=txt); c.fill=PatternFill("solid",fgColor=bg)
        c.font=Font(bold=bold,size=11); c.alignment=Alignment(horizontal=align,vertical="center",wrap_text=True); return c
    thin=Border(left=Side(style='thin',color='CBD5E1'),right=Side(style='thin',color='CBD5E1'),top=Side(style='thin',color='CBD5E1'),bottom=Side(style='thin',color='CBD5E1'))
    ws=wb.active; ws.title="Resumen"
    zona=pytz.timezone('America/Caracas'); fecha=datetime.now(zona).strftime('%d/%m/%Y %I:%M %p')
    ws.merge_cells("A1:D1"); hc(ws,1,1,"ORO ASISTENTE - REPORTE",bg=AO,size=14); ws.row_dimensions[1].height=40
    ws.merge_cells("A2:D2"); dc(ws,2,1,f"Generado: {fecha}",bg=AC,align="center")
    fila=4
    if resumen_data:
        td=resumen_data.get("titulo","Sin título")
        ws.merge_cells(f"A{fila}:D{fila}"); hc(ws,fila,1,td,bg=AM,size=12); ws.row_dimensions[fila].height=30; fila+=1
        re2=resumen_data.get("resumen_ejecutivo","")
        if re2:
            ws.merge_cells(f"A{fila}:D{fila+2}"); c2=ws.cell(row=fila,column=1,value=re2)
            c2.fill=PatternFill("solid",fgColor="EFF6FF"); c2.alignment=Alignment(horizontal="left",vertical="center",wrap_text=True)
            c2.font=Font(italic=True,size=11); ws.row_dimensions[fila].height=60; fila+=3
        if resumen_data.get("metricas"):
            fila+=1; ws.merge_cells(f"A{fila}:D{fila}"); hc(ws,fila,1,"MÉTRICAS",bg="1E40AF",size=11); fila+=1
            hc(ws,fila,1,"Indicador",bg=AC,fg="1E3A5F",size=10); hc(ws,fila,2,"Valor",bg=AC,fg="1E3A5F",size=10)
            ws.merge_cells(f"C{fila}:D{fila}"); fila+=1
            for idx,(k,v) in enumerate(resumen_data["metricas"].items()):
                bg=GC if idx%2==0 else BL; dc(ws,fila,1,k,bg=bg,bold=True)
                ws.merge_cells(f"B{fila}:C{fila}"); dc(ws,fila,2,str(v),bg=bg,align="center")
                for c3 in range(1,4): ws.cell(row=fila,column=c3).border=thin
                fila+=1
        if resumen_data.get("puntos_clave"):
            fila+=1; ws.merge_cells(f"A{fila}:D{fila}"); hc(ws,fila,1,"PUNTOS CLAVE",bg="1E40AF",size=11); fila+=1
            for i,p in enumerate(resumen_data["puntos_clave"],1):
                ws.merge_cells(f"A{fila}:D{fila}"); c4=ws.cell(row=fila,column=1,value=f"{i}. {p}")
                c4.fill=PatternFill("solid",fgColor=GC if i%2==0 else BL); c4.font=Font(size=11)
                c4.alignment=Alignment(horizontal="left",vertical="center",wrap_text=True); c4.border=thin; ws.row_dimensions[fila].height=22; fila+=1
        if resumen_data.get("hallazgo_destacado"):
            fila+=1; ws.merge_cells(f"A{fila}:D{fila}"); hc(ws,fila,1,"HALLAZGO",bg="F59E0B",fg=BL,size=11); fila+=1
            ws.merge_cells(f"A{fila}:D{fila+1}"); c5=ws.cell(row=fila,column=1,value=resumen_data["hallazgo_destacado"])
            c5.fill=PatternFill("solid",fgColor="FFFBEB"); c5.font=Font(italic=True,size=11,color="92400E")
            c5.alignment=Alignment(horizontal="left",vertical="center",wrap_text=True); ws.row_dimensions[fila].height=45
    for col in ['A','B','C','D']: ws.column_dimensions[col].width=28
    wd=wb.create_sheet("Datos"); hc(wd,1,1,"Contenido",bg=AO,size=12); wd.merge_cells("A1:B1"); wd.column_dimensions['A'].width=120
    for i,linea in enumerate(texto.split('\n'),start=2):
        if linea.strip():
            c6=wd.cell(row=i,column=1,value=linea.strip()); c6.alignment=Alignment(wrap_text=True,vertical="center")
            c6.fill=PatternFill("solid",fgColor=GC if i%2==0 else BL); wd.row_dimensions[i].height=18
    buf=BytesIO(); wb.save(buf); return buf.getvalue()

def safe_text(t): return str(t).encode('latin-1','replace').decode('latin-1')

def pdf_sh(pdf,titulo,r,g,b):
    pdf.set_fill_color(r,g,b); pdf.set_text_color(255,255,255); pdf.set_font("Helvetica",'B',11)
    pdf.cell(190,8,safe_text(titulo),border=0,new_x="LMARGIN",new_y="NEXT",fill=True); pdf.ln(2); pdf.set_text_color(30,30,30)

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
            pdf_sh(pdf,"  MÉTRICAS",30,58,95); toggle=False
            for k,v in met.items():
                rb,gb,bb=(245,247,250) if toggle else (255,255,255)
                pdf.set_fill_color(rb,gb,bb); pdf.set_font("Helvetica",'B',10)
                pdf.cell(85,8,safe_text(f"  {k}"),border=0,fill=True)
                pdf.set_font("Helvetica",'',10)
                pdf.cell(105,8,safe_text(str(v)),border=0,new_x="LMARGIN",new_y="NEXT",fill=True); toggle=not toggle
            pdf.ln(4)
        pts=resumen_data.get("puntos_clave",[])
        if pts:
            pdf_sh(pdf,"  PUNTOS CLAVE",30,64,175); pdf.set_font("Helvetica",'',10)
            for i,p in enumerate(pts,1): pdf.multi_cell(190,7,safe_text(f"  {i}. {p}")); pdf.ln(4)
        hall=resumen_data.get("hallazgo_destacado","")
        if hall:
            pdf_sh(pdf,"  HALLAZGO",180,120,10); pdf.set_font("Helvetica",'I',10)
            pdf.multi_cell(190,7,safe_text(f"  {hall}")); pdf.ln(4)
        pdf.add_page()
    pdf_sh(pdf,"  CONTENIDO",30,58,95); pdf.set_font("Helvetica",'',9)
    for linea in texto.split('\n'):
        linea=linea.strip()
        if linea: pdf.multi_cell(190,5,safe_text(linea))
    raw=pdf.output()
    return bytes(raw) if isinstance(raw,(bytes,bytearray)) else raw.encode('latin-1') if isinstance(raw,str) else bytes(raw)

def _aplicar_formato_run(run, fmt):
    """Aplica formato explícito a un run si se especificó en la instrucción."""
    if not fmt: return
    if fmt.get("bold") is not None: run.font.bold = fmt["bold"]
    if fmt.get("italic") is not None: run.font.italic = fmt["italic"]
    if fmt.get("underline") is not None: run.font.underline = fmt["underline"]
    if fmt.get("size"): run.font.size = Pt(fmt["size"])
    if fmt.get("color"):
        try:
            color_hex = fmt["color"].lstrip("#")
            run.font.color.rgb = RGBColor(int(color_hex[0:2],16), int(color_hex[2:4],16), int(color_hex[4:6],16))
        except: pass
    if fmt.get("upper"): run.text = run.text.upper()
    if fmt.get("lower"): run.text = run.text.lower()

def reemplazar_docx_preservando_formato(archivo_bytes, cambios):
    """
    Reemplaza texto en DOCX preservando el formato de cada run.
    Estrategia: reemplazar solo dentro del run que contiene el texto,
    manteniendo negrita, cursiva, color, tamaño, etc.
    Si el texto cruza varios runs, reconstruye solo los afectados.
    """
    doc = Document(BytesIO(archivo_bytes))
    conteo = 0

    for c in cambios:
        buscar = str(c["buscar"])
        reemplazar = str(c["reemplazar"])
        fmt_extra = c.get("formato", {})  # formato explícito opcional
        if not buscar or buscar.lower() == reemplazar.lower():
            continue
        regex = re.compile(re.escape(buscar), re.IGNORECASE)

        def rep_parrafo(p):
            nonlocal conteo
            if not regex.search(p.text):
                return

            # Caso 1: el texto está completamente dentro de un solo run
            for run in p.runs:
                if regex.search(run.text):
                    nuevo_texto, n = regex.subn(reemplazar, run.text)
                    if n > 0:
                        run.text = nuevo_texto
                        # Aplicar formato adicional si se pidió
                        if fmt_extra:
                            _aplicar_formato_run(run, fmt_extra)
                        conteo += n
                    return

            # Caso 2: el texto cruza múltiples runs
            # Reconstruir el párrafo preservando formato del run dominante
            texto_completo = p.text
            nuevo_completo, n = regex.subn(reemplazar, texto_completo)
            if n == 0:
                return
            conteo += n

            # Guardar el formato del run que más se parece al texto buscado
            run_ref = None
            for run in p.runs:
                if buscar.lower() in run.text.lower() or len(run.text) > 0:
                    run_ref = run
                    break
            if run_ref is None and p.runs:
                run_ref = p.runs[0]

            # Aplicar: poner todo en runs[0] con el formato del run de referencia
            if p.runs:
                # Copiar formato del run de referencia al run[0]
                r0 = p.runs[0]
                if run_ref and run_ref != r0:
                    r0.font.bold      = run_ref.font.bold
                    r0.font.italic    = run_ref.font.italic
                    r0.font.underline = run_ref.font.underline
                    r0.font.size      = run_ref.font.size
                    try:
                        if run_ref.font.color and run_ref.font.color.rgb:
                            r0.font.color.rgb = run_ref.font.color.rgb
                    except: pass
                r0.text = nuevo_completo
                for run in p.runs[1:]:
                    run.text = ""
                if fmt_extra:
                    _aplicar_formato_run(r0, fmt_extra)

        [rep_parrafo(p) for p in doc.paragraphs]
        [rep_parrafo(p) for t in doc.tables for row in t.rows for cell in row.cells for p in cell.paragraphs]

    buf = BytesIO()
    doc.save(buf)
    return buf.getvalue(), conteo

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
                use_font="Times-BoldItalic" if "bold" in fn and "italic" in fn else "Helvetica-Bold" if "bold" in fn or bold else "Helvetica-Oblique" if "italic" in fn or italic else "Times-Roman" if "times" in fn or "serif" in fn else "Courier" if "courier" in fn or "mono" in fn else "Helvetica"
                try:
                    pix=pagina.get_pixmap(clip=rect,dpi=72); cx=pix.width//2; cy=pix.height//2
                    s=pix.pixel(cx,cy); bg=(s[0]/255,s[1]/255,s[2]/255)
                except: bg=(1.,1.,1.)
                pagina.add_redact_annot(rect,fill=bg); pagina.apply_redactions()
                pagina.insert_text(fitz.Point(rect.x0,rect.y1-1.5),reemplazar,fontname=use_font,fontsize=font_size,color=color)
                conteo+=1
    buf=BytesIO(); doc.save(buf); doc.close(); return buf.getvalue(),conteo

# ─────────────────────────────────────────────────────────
# UPLOADER
# ─────────────────────────────────────────────────────────
st.markdown('<div class="oro-divider"></div>', unsafe_allow_html=True)
st.markdown("""<style>
[data-testid="stFileUploaderDropzoneInstructions"]>div>span::after{content:"Arrastra tu archivo aquí"}
[data-testid="stFileUploaderDropzoneInstructions"]>div>span{font-size:0!important}
[data-testid="stFileUploaderDropzoneInstructions"]>div>small::after{content:"Límite 200MB • DOCX, XLSX, PDF"}
[data-testid="stFileUploaderDropzoneInstructions"]>div>small{font-size:0!important}
[data-testid="stFileUploadDropzone"]>div>button{visibility:hidden;position:relative}
[data-testid="stFileUploadDropzone"]>div>button::after{content:"Seleccionar archivo";visibility:visible;position:absolute;left:0;right:0;text-align:center}
</style>""", unsafe_allow_html=True)

archivo = st.file_uploader("📎 Sube tu archivo", type=["docx","xlsx","pdf"], help="Word, Excel o PDF — máx 200MB", label_visibility="visible")

if archivo and archivo.name != st.session_state.nombre_archivo:
    with st.spinner("📖 Cargando..."):
        contenido = archivo.read()
        for k,v in [("archivo_bytes",contenido),("nombre_archivo",archivo.name),("archivo_tipo",archivo.name.split(".")[-1].lower()),
                    ("resumen_data",None),("historial_chat",[]),("lista_cambios",[]),("cambios_aplicados",None),
                    ("texto_corregido",""),("preview_cambio",None),("resumen_error",False),("generando_resumen",False),
                    ("guia_paso",1),("guia_vista",False)]:
            st.session_state[k] = v
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
            st.error(f"Error leyendo el archivo: {e}")

if not st.session_state.get("texto_extraido") and st.session_state.get("generando_resumen"):
    st.session_state.generando_resumen=False

# ─────────────────────────────────────────────────────────
# PANEL PRINCIPAL
# ─────────────────────────────────────────────────────────
if st.session_state.texto_extraido:
    texto=st.session_state.texto_extraido
    tipo=st.session_state.archivo_tipo
    texto_activo=st.session_state.texto_corregido if st.session_state.texto_corregido else texto

    # ── File badge ──
    palabras=len(texto.split())
    ext_icon={"docx":"📄","xlsx":"📊","pdf":"📕"}.get(tipo,"📎")
    cn=len(st.session_state.lista_cambios)
    badge_extra=f' &nbsp;·&nbsp; ✏️ <strong style="color:#10b981">{cn} cambio(s)</strong>' if cn else ""
    st.markdown(f"""<div class="file-badge">
        <div class="file-icon">{ext_icon}</div>
        <div><div class="file-info-name">{st.session_state.nombre_archivo}</div>
        <div class="file-info-stats">📝 {palabras:,} palabras{badge_extra}</div></div>
    </div>""", unsafe_allow_html=True)

    # ── Dos botones principales ──
    _ba, _be = st.columns(2)
    with _ba:
        st.markdown('<div class="btn-analizar">', unsafe_allow_html=True)
        if st.button("⚡ Analizar", use_container_width=True, key="btn_analizar_top"):
            st.session_state.generando_resumen = True
            st.session_state.resumen_data = None
            st.rerun()
        st.markdown('</div>', unsafe_allow_html=True)
    with _be:
        st.markdown('<div class="btn-evaluar">', unsafe_allow_html=True)
        if st.button("🔍 Evaluar", use_container_width=True, key="btn_evaluar_top"):
            st.session_state.ejecutar_evaluacion = True
            st.rerun()
        st.markdown('</div>', unsafe_allow_html=True)

    # Ejecutar evaluación si se pidió
    if st.session_state.get("ejecutar_evaluacion"):
        st.session_state.ejecutar_evaluacion = False
        with st.spinner("🔎 Analizando calidad..."):
            resultado = detectar_anomalias(texto_activo)
        if resultado:
            niv=resultado.get("nivel_general","Regular"); puntaje=resultado.get("puntaje",0)
            ncfg={"Excelente":("#10b981","#021008","🟢"),"Bueno":("#34d399","#021008","🟢"),"Regular":("#f59e0b","#1c1003","🟡"),"Deficiente":("#ef4444","#1f0707","🔴")}
            cfg=ncfg.get(niv,ncfg["Regular"])
            st.markdown(f"""<div style="text-align:center;padding:.8rem 0 .4rem">
                <div style="font-size:2.5rem">{cfg[2]}</div>
                <div style="color:{cfg[0]};font-size:1.2rem;font-weight:800">{niv}</div>
                <div style="color:#2d6a4f;font-size:.78rem">Puntaje: <strong style="color:{cfg[0]}">{puntaje}/100</strong></div>
            </div>""", unsafe_allow_html=True)
            ne=[("criticos","🔴 Crítico","#ef4444","#1f0707","#450a0a"),("altos","🟠 Alto","#f97316","#1c0a03","#431407"),
                ("medios","🟡 Medio","#f59e0b","#1c1003","#451a03"),("leves","🟢 Leve","#22c55e","#052e16","#14532d")]
            hay=False
            for key,label,cfg2,cbg,cbrd in ne:
                items_e=resultado.get(key,[])
                if items_e:
                    hay=True
                    rows="".join([f'<div style="color:#d1fae5;font-size:.78rem;padding:.2rem 0;border-bottom:1px solid {cbrd}">• {it}</div>' for it in items_e])
                    st.markdown(f'<div style="background:{cbg};border:1px solid {cbrd};border-left:4px solid {cfg2};border-radius:12px;padding:.75rem .9rem;margin:.4rem 0"><div style="color:{cfg2};font-weight:700;font-size:.83rem;margin-bottom:.3rem">{label}</div>{rows}</div>', unsafe_allow_html=True)
            if not hay:
                st.markdown('<div class="info-box">✅ ¡Sin problemas detectados! 🎉</div>', unsafe_allow_html=True)
            rec=resultado.get("recomendacion","")
            if rec: st.markdown(f'<div class="hallazgo-card">💡 <strong>Recomendación:</strong> {rec}</div>', unsafe_allow_html=True)
        else:
            st.markdown('<div class="warn-box">⚠️ No se pudo evaluar. Intenta de nuevo.</div>', unsafe_allow_html=True)

    st.markdown('<div class="oro-divider"></div>', unsafe_allow_html=True)

    # ══════════════════════════════════════════
    # RESUMEN
    # ══════════════════════════════════════════
    if st.session_state.generando_resumen:
        with st.spinner("🧠 Analizando documento..."):
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
        if st.button("🔄 Reintentar análisis", use_container_width=True):
            st.session_state.resumen_error=False; st.session_state.generando_resumen=True; st.rerun()

    data=st.session_state.resumen_data
    if not data and not st.session_state.generando_resumen and not st.session_state.get("resumen_error"):
        st.markdown("""<div style="text-align:center;padding:.8rem 0 .4rem">
            <div style="font-size:2rem">🧠</div>
            <div style="color:#34d399;font-weight:600;font-size:.88rem;margin-top:.3rem">Toca ⚡ Analizar para generar el resumen</div>
        </div>""", unsafe_allow_html=True)

    if data:
        emoji=data.get("emoji_categoria","📋"); titulo_doc=data.get("titulo","Documento analizado")
        st.markdown(f"""<div class="summary-card">
            <div class="summary-card-title">{emoji} {titulo_doc}</div>
            {data.get("resumen_ejecutivo","")}
        </div>""", unsafe_allow_html=True)

        metricas=data.get("metricas",{})
        if metricas:
            pills='<div class="metrics-grid">'
            for k,v in list(metricas.items())[:4]:
                pills+=f'<div class="metric-pill"><div class="metric-pill-label">{k}</div><div class="metric-pill-value">{v}</div></div>'
            pills+='</div>'; st.markdown(pills, unsafe_allow_html=True)

        puntos=data.get("puntos_clave",[])
        if puntos:
            tags='<div class="tags-wrap">'+"".join([f'<span class="tag">✓ {p}</span>' for p in puntos])+'</div>'
            st.markdown(tags, unsafe_allow_html=True)

        hall=data.get("hallazgo_destacado","")
        if hall:
            st.markdown(f'<div class="hallazgo-card">💡 <strong>Hallazgo:</strong> {hall}</div>', unsafe_allow_html=True)

        st.markdown('<div class="oro-divider"></div>', unsafe_allow_html=True)
        st.markdown('<div class="section-title">📥 Exportar informe</div>', unsafe_allow_html=True)
        ab=st.session_state.archivo_bytes; ca=st.session_state.lista_cambios
        c1,c2,c3=st.columns(3)
        with c1:
            st.download_button("📄 Word",exportar_word(texto_activo,data,archivo_bytes=ab,archivo_tipo=tipo,cambios=ca),
                "Informe.docx",mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",use_container_width=True)
        with c2:
            st.download_button("📊 Excel",exportar_excel(texto_activo,data,archivo_bytes=ab,archivo_tipo=tipo,cambios=ca),
                "Informe.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",use_container_width=True)
        with c3:
            st.download_button("📕 PDF",exportar_pdf(texto_activo,data),"Informe.pdf",mime="application/pdf",use_container_width=True)

        if st.button("🔄 Regenerar resumen", use_container_width=True):
            st.session_state.generando_resumen=True; st.session_state.resumen_data=None; st.rerun()

    # ══════════════════════════════════════════
    # PREVIEW cambio pendiente
    # ══════════════════════════════════════════
    if st.session_state.preview_cambio:
        preview=st.session_state.preview_cambio
        st.markdown("""<div style="background:#021008;border:1px solid #10b981;border-radius:14px;padding:.9rem 1rem;margin:.4rem 0">
            <div style="color:#34d399;font-weight:700;font-size:.85rem;margin-bottom:.5rem">👁 Vista previa del cambio</div>""", unsafe_allow_html=True)
        for c in preview:
            bq=c["buscar"][:50]+("..." if len(c["buscar"])>50 else "")
            rq=c["reemplazar"][:50]+("..." if len(c["reemplazar"])>50 else "")
            idx=texto_activo.lower().find(c["buscar"].lower())
            if idx!=-1:
                ini=max(0,idx-30); fin=min(len(texto_activo),idx+len(c["buscar"])+30)
                ca2=texto_activo[ini:idx]; cd=texto_activo[idx+len(c["buscar"]):fin]
                st.markdown(
                    f'<div style="font-size:.77rem;margin:.3rem 0">'
                    f'<span style="color:#6b7280;font-size:.62rem;text-transform:uppercase">Antes: </span>'
                    f'<span style="color:#fca5a5;font-family:monospace">...{ca2}<mark style="background:#7f1d1d;color:#fca5a5;border-radius:3px;padding:0 3px">{bq}</mark>{cd}...</span></div>'
                    f'<div style="font-size:.77rem">'
                    f'<span style="color:#6b7280;font-size:.62rem;text-transform:uppercase">Después: </span>'
                    f'<span style="color:#86efac;font-family:monospace">...{ca2}<mark style="background:#14532d;color:#86efac;border-radius:3px;padding:0 3px">{rq}</mark>{cd}...</span></div>',
                    unsafe_allow_html=True)
            else:
                st.markdown(f'<div style="color:#fbbf24;font-size:.78rem">⚠️ "{bq}" no encontrado</div>', unsafe_allow_html=True)
        st.markdown('</div>', unsafe_allow_html=True)
        cs,cn2=st.columns(2)
        with cs:
            if st.button("✅ Confirmar", use_container_width=True):
                st.session_state.lista_cambios.extend(preview); st.session_state.preview_cambio=None
                todos_c=st.session_state.lista_cambios; ab_orig=st.session_state.archivo_bytes
                if tipo=="docx": final_bytes,n=reemplazar_docx_preservando_formato(ab_orig,todos_c)
                elif tipo=="xlsx": final_bytes,n=reemplazar_xlsx_preservando_formato(ab_orig,todos_c)
                elif tipo=="pdf" and PYMUPDF_OK: final_bytes,n=reemplazar_pdf_original(ab_orig,todos_c)
                else:
                    txt_m=texto_activo; n=0
                    for c2 in todos_c:
                        txt_m,cnt=re.compile(re.escape(c2["buscar"]),re.IGNORECASE).subn(c2["reemplazar"],txt_m); n+=cnt
                    final_bytes=txt_m.encode()
                txt_c=texto_activo
                for c2 in todos_c: txt_c=re.compile(re.escape(c2["buscar"]),re.IGNORECASE).sub(c2["reemplazar"],txt_c)
                st.session_state.texto_corregido=txt_c; st.session_state.cambios_aplicados=final_bytes
                st.session_state.resumen_data=None; st.session_state.generando_resumen=True; st.session_state.edicion_counter+=1
                st.session_state.historial_chat.append({"rol":"Asistente","texto":f"✅ Listo — cambié **{preview[0]['buscar']}** → **{preview[0]['reemplazar']}**. ¿Algo más?"})
                st.rerun()
        with cn2:
            if st.button("❌ Cancelar", use_container_width=True):
                st.session_state.preview_cambio=None; st.session_state.edicion_counter+=1; st.rerun()

    if st.session_state.cambios_aplicados:
        with st.expander(f"📥 Descargar corregido ({len(st.session_state.lista_cambios)} cambio(s))"):
            fb=st.session_state.cambios_aplicados; todos_c=st.session_state.lista_cambios
            if tipo=="docx":
                st.download_button("📄 Word corregido",fb,"Corregido.docx",mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",use_container_width=True)
            elif tipo=="xlsx":
                st.download_button("📊 Excel corregido",fb,"Corregido.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",use_container_width=True)
            elif tipo=="pdf" and PYMUPDF_OK:
                st.download_button("📕 PDF corregido",fb,"Corregido.pdf",mime="application/pdf",use_container_width=True)
            wc=exportar_word(st.session_state.texto_corregido or texto,None)
            st.download_button("📄 Exportar como Word",wc,"Exportado.docx",mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",use_container_width=True)
            if st.button("🗑️ Limpiar cambios",use_container_width=True):
                st.session_state.lista_cambios=[]; st.session_state.cambios_aplicados=None
                st.session_state.texto_corregido=""; st.session_state.preview_cambio=None; st.rerun()

    # ══════════════════════════════════════════
    # CHAT — siempre visible
    # ══════════════════════════════════════════
    for msg in st.session_state.historial_chat:
        with st.chat_message("user" if msg["rol"]=="Usuario" else "assistant"):
            st.write(msg["texto"])

    if not st.session_state.historial_chat and not st.session_state.preview_cambio:
        st.markdown(f"""<div class="chat-placeholder">
            <div style="font-size:1.8rem">💬</div>
            <div style="color:#10b981;font-weight:700;font-size:.9rem;margin:.3rem 0">Conversa sobre el documento</div>
            <div style="color:#065f46;font-size:.75rem;margin-bottom:.6rem">Edita palabras o haz cualquier pregunta</div>
            <div style="display:flex;justify-content:center;gap:.35rem;flex-wrap:wrap">
                <span style="background:#010c06;border:1px solid #0a3d1a;border-radius:20px;padding:.2rem .6rem;font-size:.68rem;color:#10b981">✏️ cambia X por Y</span>
                <span style="background:#010c06;border:1px solid #0a3d1a;border-radius:20px;padding:.2rem .6rem;font-size:.68rem;color:#10b981">❓ ¿cuántas personas hay?</span>
                <span style="background:#010c06;border:1px solid #0a3d1a;border-radius:20px;padding:.2rem .6rem;font-size:.68rem;color:#10b981">📝 resume en 3 puntos</span>
            </div>
        </div>""", unsafe_allow_html=True)

    entrada=st.chat_input("✍️ Escribe un cambio o una pregunta...", key=f"chat_{st.session_state.edicion_counter}")
    if entrada:
        st.session_state.historial_chat.append({"rol":"Usuario","texto":entrada})
        if st.session_state.get("guia_paso")==2: st.session_state.guia_paso=3
        palabras_cambio=["cambia","reemplaza","sustituye","corrige","agrega","añade","borra","elimina","pon","escribe","modifica","quita","actualiza"]
        es_cambio=any(p in entrada.lower() for p in palabras_cambio)
        if es_cambio:
            with st.spinner("🔍 Procesando cambio..."):
                nuevos=solicitar_cambios(entrada,texto_activo)
            if nuevos:
                st.session_state.preview_cambio=nuevos
                st.session_state.historial_chat.append({"rol":"Asistente","texto":"Encontré el cambio 👆 Revisa la vista previa arriba y confirma si es correcto."})
            else:
                st.session_state.historial_chat.append({"rol":"Asistente","texto":"No encontré qué cambiar exactamente. Intenta: *cambia 'palabra original' por 'palabra nueva'*"})
        else:
            with st.spinner("🤔 Pensando..."):
                resp=preguntar_al_documento(entrada,texto_activo)
            st.session_state.historial_chat.append({"rol":"Asistente","texto":resp})
        st.rerun()

else:
    st.markdown("""<div class="empty-state">
        <div class="empty-icon">📂</div>
        <div class="empty-title">Sube un archivo para empezar</div>
        <div class="empty-hint">
            🏆 Analiza · ✍️ Edita · 📥 Exporta<br>
            Soporta Word, Excel y PDF
        </div>
        <div class="format-badges">
            <span class="format-badge">.docx</span>
            <span class="format-badge">.xlsx</span>
            <span class="format-badge">.pdf</span>
        </div>
    </div>""", unsafe_allow_html=True)

zona_horaria=pytz.timezone('America/Caracas')
hora=datetime.now(zona_horaria).strftime('%I:%M %p')
st.markdown(f"<p class='oro-footer'>🏆 Oro Asistente · {hora} VET</p>", unsafe_allow_html=True)
