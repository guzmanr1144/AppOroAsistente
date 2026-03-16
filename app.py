import streamlit as st
import google.generativeai as genai

st.set_page_config(page_title="Escáner Oro", page_icon="🔍")

LLAVE_GEMINI = "AIzaSyADVQhbwbz6SZR-pT1rfpbf-tqJnFxRg-o"
genai.configure(api_key=LLAVE_GEMINI)

st.title("🔍 Escáner de Llave API")
st.write("Vamos a ver exactamente qué modelos te permite usar Google...")

try:
    modelos_permitidos = []
    for m in genai.list_models():
        if 'generateContent' in m.supported_generation_methods:
            modelos_permitidos.append(m.name)
            
    if modelos_permitidos:
        st.success("✅ ¡Google te permite usar estos modelos!")
        for mod in modelos_permitidos:
            st.code(mod)
    else:
        st.error("❌ Tu llave no tiene permisos para generar texto. Necesitas crear una llave nueva.")
except Exception as e:
    st.error(f"⚠️ Error al conectar con Google: {str(e)}")
