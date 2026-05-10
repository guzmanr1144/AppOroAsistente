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
# ESTILOS CSS (cacheados para no regenerar en cada ciclo)
# ==========================================
@st.cache_data(show_spinner=False)
def _get_all_css(tema_key="verde"):
    _TEMAS_LOCAL = {
        "oscuro": {"bg1":"#0a0e1a","bg2":"#0d1525","card":"#111827","card2":"#162032","borde":"#1e3a5f","borde2":"#2a4a6b","acento1":"#3b82f6","acento2":"#60a5fa","titulo_grad":"linear-gradient(135deg,#fbbf24,#f59e0b,#fde68a,#f59e0b)","texto":"#e2e8f0","texto2":"#93c5fd","texto3":"#4b6080"},
        "azul":   {"bg1":"#020818","bg2":"#030d24","card":"#041230","card2":"#061840","borde":"#0c3a7a","borde2":"#1050aa","acento1":"#38bdf8","acento2":"#7dd3fc","titulo_grad":"linear-gradient(135deg,#38bdf8,#7dd3fc,#bae6fd)","texto":"#e0f2fe","texto2":"#7dd3fc","texto3":"#1e5a8a"},
        "verde":  {"bg1":"#010c06","bg2":"#021008","card":"#041208","card2":"#051a0c","borde":"#0a3d1a","borde2":"#0f5225","acento1":"#10b981","acento2":"#34d399","titulo_grad":"linear-gradient(135deg,#10b981,#34d399,#6ee7b7,#10b981)","texto":"#d1fae5","texto2":"#34d399","texto3":"#065f46"},
        "rosa":   {"bg1":"#120008","bg2":"#1a000f","card":"#1a000f","card2":"#280018","borde":"#7c0040","borde2":"#9d0050","acento1":"#f472b6","acento2":"#f9a8d4","titulo_grad":"linear-gradient(135deg,#f472b6,#f9a8d4,#fce7f3)","texto":"#fce7f3","texto2":"#f9a8d4","texto3":"#7c0040"},
        "ambar":  {"bg1":"#0f0800","bg2":"#180d00","card":"#1a0e00","card2":"#251500","borde":"#78350f","borde2":"#92400e","acento1":"#f59e0b","acento2":"#fbbf24","titulo_grad":"linear-gradient(135deg,#fbbf24,#fde68a,#f59e0b)","texto":"#fef3c7","texto2":"#fbbf24","texto3":"#78350f"},
    }
    t = _TEMAS_LOCAL.get(tema_key, _TEMAS_LOCAL["verde"])
    return f"""<style>
@import url('https://fonts.googleapis.com/css2?family=Outfit:wght@300;400;500;600;700;800&family=JetBrains+Mono:wght@400;600&display=swap');
html,body,[class*="css"]{{font-family:'Outfit',sans-serif!important;-webkit-tap-highlight-color:transparent}}
.stApp{{background:linear-gradient(160deg,{t['bg1']} 0%,{t['bg2']} 50%,{t['bg1']} 100%)!important;min-height:100vh}}
.main .block-container{{padding:1rem 1rem 4rem 1rem!important;max-width:480px!important;margin:0 auto!important;background:transparent!important}}
#MainMenu,footer,header{{visibility:hidden}}[data-testid="stToolbar"]{{display:none}}
.oro-header{{text-align:center;padding:1.5rem 0 0.5rem}}
.oro-logo{{font-size:3rem;line-height:1;filter:drop-shadow(0 0 20px rgba(251,191,36,0.5));animation:pulse-glow 3s ease-in-out infinite}}
@keyframes pulse-glow{{0%,100%{{filter:drop-shadow(0 0 15px rgba(251,191,36,0.4))}}50%{{filter:drop-shadow(0 0 30px rgba(251,191,36,0.8))}}}}
.oro-title{{font-size:1.9rem;font-weight:800;background:{t['titulo_grad']}!important;-webkit-background-clip:text!important;-webkit-text-fill-color:transparent!important;background-clip:text!important;letter-spacing:-0.02em;margin:.2rem 0 .1rem}}
.oro-subtitle{{color:#4b6080;font-size:.82rem;font-weight:400;letter-spacing:.03em}}
[data-testid="stFileUploader"]{{background:transparent!important;border:none!important}}
[data-testid="stFileUploader"]>div{{background:linear-gradient(135deg,{t['card']},{t['card2']})!important;border:2px dashed {t['borde']}!important;border-radius:20px!important;padding:1.5rem!important}}
[data-testid="stFileUploader"] label{{color:{t['acento2']}!important;font-weight:600!important;font-size:1rem!important}}
.summary-card{{background:{t['card2']}!important;border:1px solid {t['borde2']}!important;border-left:4px solid {t['acento1']}!important;border-radius:18px;padding:1.2rem;margin:.8rem 0;color:{t['texto2']}!important;line-height:1.7;font-size:.9rem}}
.stButton>button{{background:{t['card']}!important;color:{t['texto3']}!important;border:1px solid {t['borde']}!important;border-radius:14px!important;font-weight:600!important;font-size:.82rem!important;min-height:3.2rem!important;width:100%!important;transition:all .15s!important;font-family:'Outfit',sans-serif!important}}
.stButton>button:hover{{border-color:{t['acento1']}!important;color:{t['acento2']}!important}}
[data-testid="stSidebar"]{{background:{t['bg1']}!important}}
</style>"""

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
    "resumen_error": False,
    "tab_activa": "resumen",
    "tema": "verde",
    "preview_cambio": None,
    "edicion_counter": 0,
    "texto_corregido": "",
}.items():
    if key not in st.session_state:
        st.session_state[key] = val

st.markdown(_get_all_css(st.session_state.get("tema","verde")), unsafe_allow_html=True)

# ==========================================
# CONEXIÓN GEMINI
# ==========================================
try:
    LLAVE_GEMINI = st.secrets["LLAVE_GEMINI"]
    genai.configure(api_key=LLAVE_GEMINI)
except Exception as e:
    st.error(f"🔑 Error configurando la IA: {e}")
    st.stop()

MODELOS_FALLBACK = [
    "gemini-3.1-flash-lite-preview",
    "gemini-3.1-flash-preview",
    "gemini-3.1-pro-preview",
]

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
    st.caption("Oro Asistente v2")

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
# UTILIDADES JSON & FUNCIONES IA
# ==========================================
def extraer_json_seguro(texto, es_lista=False):t = texto.replace("```json", "").replace("
```", "").strip()
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

# ==========================================
# EXPORTADORES
# ==========================================
def exportar_excel(texto, resumen_data=None, archivo_bytes=None, archivo_tipo=None, cambios=None):
    cambios = cambios or []
    
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
            header_cell(ws_res, fila, 1, "HALLAZGO DESTACADO", bg="1E40AF", size=11)
            fila += 1
            ws_res.merge_cells(f"A{fila}:D{fila+1}")
            cell = ws_res.cell(row=fila, column=1, value=resumen_data["hallazgo_destacado"])
            cell.fill = PatternFill("solid", fgColor="F0FDF4")
            cell.font = Font(italic=True, size=11)
            cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
            cell.border = thin
            ws_res.row_dimensions[fila].height = 40
            fila += 2

    buf = BytesIO()
    wb.save(buf)
    return buf.getvalue()

# ==========================================
# INTERFAZ Y LÓGICA PRINCIPAL
# ==========================================

archivo = st.file_uploader("Sube tu documento", type=["txt", "pdf"])

if archivo:
    if st.session_state.nombre_archivo != archivo.name:
        st.session_state.nombre_archivo = archivo.name
        # Lectura básica del texto para que funcione el ejemplo
        if archivo.name.endswith(".txt"):
            st.session_state.texto_extraido = archivo.getvalue().decode("utf-8", errors="ignore")
        elif archivo.name.endswith(".pdf"):
            try:
                pdf_reader = PyPDF2.PdfReader(archivo)
                st.session_state.texto_extraido = "\n".join([page.extract_text() for page in pdf_reader.pages])
            except Exception as e:
                st.session_state.texto_extraido = "Error extrayendo PDF."
        st.session_state.resumen_data = None 

    # BOTÓN CORREGIDO (UN SOLO CLIC)
    if st.button("🚀 Analizar Documento"):
        if st.session_state.texto_extraido:
            with st.spinner("Analizando con IA..."):
                resultado = solicitar_resumen_estructurado(st.session_state.texto_extraido)
                if resultado:
                    st.session_state.resumen_data = resultado
                    st.rerun() # Esto recarga y muestra de inmediato
                else:
                    st.error("Hubo un problema al generar el análisis.")

# Mostrar los resultados del análisis si existen
if st.session_state.resumen_data:
    res = st.session_state.resumen_data
    st.markdown(f"### {res.get('emoji_categoria', '📋')} {res.get('titulo', 'Resumen')}")
    st.markdown(f"<div class='summary-card'>{res.get('resumen_ejecutivo', '')}</div>", unsafe_allow_html=True)
    
    st.write("**Métricas Clave:**")
    st.json(res.get('metricas', {}))
    
    st.write("**Puntos Clave:**")
    for punto in res.get('puntos_clave', []):
        st.write(f"- {punto}")
