import streamlit as st
import google.generativeai as genai
import pickle
import os

st.set_page_config(page_title="Asistente MESOB", page_icon="🎓", layout="wide")
st.title("🎓 Asistente MESOB")
st.markdown("Pregunta sobre la documentación del programa MESOB")

# Sidebar
with st.sidebar:
    st.markdown("### 📞 Más Información")
    st.markdown("""
    **Web oficial:**
    [mesob.uam.es](https://www.uam.es/educacion/estudios/practicas-externas/informacion-practicas-posgrado/mesob)

    **Email consultas:**
    delegacion.educacion.practicasmesob@uam.es
    """)
    st.divider()
    st.caption("Asistente MESOB - UAM")

# Configura Gemini
api_key = st.secrets.get("GEMINI_API_KEY", os.environ.get("GEMINI_API_KEY", ""))
if not api_key:
    st.error("Falta la API key de Gemini. Configúrala en Streamlit Cloud > Secrets.")
    st.stop()

genai.configure(api_key=api_key)
model = genai.GenerativeModel("gemini-1.5-flash")

# Carga documentos procesados
@st.cache_resource
def load_documents():
    db_file = "./mesob_documents.pkl"
    if os.path.exists(db_file):
        with open(db_file, "rb") as f:
            return pickle.load(f)
    return []

documents = load_documents()

# Busca contexto relevante
def get_context(query, docs, max_chars=4000):
    if not docs:
        return ""
    query_words = set(query.lower().split())
    scored = []
    for doc in docs:
        score = sum(1 for w in query_words if w in doc.lower())
        scored.append((score, doc))
    scored.sort(reverse=True)
    context = ""
    for _, doc in scored[:3]:
        context += doc[:1500] + "\n\n"
    return context[:max_chars]

# Historial de chat
if "messages" not in st.session_state:
    st.session_state.messages = []

# Sugerencias iniciales
SUGGESTIONS = {
    "📅 ¿Cuál es el calendario de prácticas?": "¿Cuál es el calendario de prácticas externas?",
    "📋 ¿Cómo se asignan los centros?": "¿Cómo se asignan los centros de prácticas?",
    "📝 ¿Qué es el TFM?": "¿Qué información hay sobre el TFM?",
    "❓ ¿Cómo presento alegaciones?": "¿Cómo puedo presentar alegaciones?",
}

if not st.session_state.messages:
    selected = st.pills("Preguntas frecuentes:", list(SUGGESTIONS.keys()), label_visibility="collapsed")
    if selected:
        st.session_state.messages.append({"role": "user", "content": SUGGESTIONS[selected]})
        st.rerun()

# Muestra historial
for msg in st.session_state.messages:
    with st.chat_message(msg["role"]):
        st.markdown(msg["content"])

# Input usuario
if prompt := st.chat_input("¿Pregunta sobre MESOB?"):
    st.session_state.messages.append({"role": "user", "content": prompt})
    with st.chat_message("user"):
        st.markdown(prompt)

    with st.chat_message("assistant"):
        with st.spinner("Consultando documentación..."):
            context = get_context(prompt, documents)

            full_prompt = f"""Eres un asistente experto en el programa MESOB (Máster en Educación Secundaria Obligatoria y Bachillerato) de la UAM.
Responde en español basándote en la documentación proporcionada.
Si no encuentras la respuesta, indica que no está en la documentación y recomienda:
- Web: https://www.uam.es/educacion/estudios/practicas-externas/informacion-practicas-posgrado/mesob
- Email: delegacion.educacion.practicasmesob@uam.es

Documentación disponible:
{context}

Pregunta: {prompt}
"""
            response = model.generate_content(full_prompt)
            st.markdown(response.text)
            st.session_state.messages.append({"role": "assistant", "content": response.text})

# Botón limpiar
if st.session_state.messages:
    if st.button("🗑️ Limpiar chat"):
        st.session_state.messages = []
        st.rerun()
