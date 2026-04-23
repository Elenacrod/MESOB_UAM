import streamlit as st
import ollama
import pickle
import os

st.set_page_config(page_title="MESOB Assistant", layout="wide")
st.title("🎓 Asistente MESOB")
st.markdown("Pregunta sobre la documentación del programa MESOB")

# Sidebar con info de contacto
with st.sidebar:
    st.markdown("### 📞 Más Información")
    st.markdown("""
    **Web oficial:**
    https://www.uam.es/educacion/estudios/practicas-externas/informacion-practicas-posgrado/mesob

    **Email de consultas:**
    delegacion.educacion.practicasmesob@uam.es
    """)
    st.divider()
    st.markdown("💡 *Asistente MESOB - Powered by AI*")

# Cargar documentos en caché
@st.cache_resource
def load_documents():
    db_file = "./mesob_documents.pkl"
    if os.path.exists(db_file):
        with open(db_file, "rb") as f:
            return pickle.load(f)
    return []

documents = load_documents()

# Inicializa sesión de chat
if "messages" not in st.session_state:
    st.session_state.messages = []

# Muestra historial de chat
for msg in st.session_state.messages:
    with st.chat_message(msg["role"]):
        st.markdown(msg["content"])

# Input del usuario
if prompt := st.chat_input("¿Pregunta sobre MESOB?"):
    st.session_state.messages.append({"role": "user", "content": prompt})

    with st.chat_message("user"):
        st.markdown(prompt)

    with st.spinner("Procesando..."):
        try:
            # Busca contenido relevante en documentos
            context = ""
            if documents:
                # Busca manualmente por palabras clave
                for doc in documents[:3]:
                    if any(word.lower() in doc.lower() for word in prompt.split()):
                        context += f"\n{doc[:500]}...\n"

            if not context:
                context = "No se encontró documentación específica sobre este tema."

            # Genera respuesta usando Ollama
            system_prompt = """Eres un asistente experto en el programa MESOB (Máster en Educación Secundaria Obligatoria y Bachillerato).
Responde preguntas basándote en la documentación proporcionada.
Si no encuentras la respuesta en la documentación, indícalo e incluye:
- Web: https://www.uam.es/educacion/estudios/practicas-externas/informacion-practicas-posgrado/mesob
- Email: delegacion.educacion.practicasmesob@uam.es
Responde en español y sé conciso."""

            full_prompt = f"""Documentación:
{context}

Pregunta: {prompt}

Respuesta:"""

            # Genera respuesta
            with st.chat_message("assistant"):
                response_placeholder = st.empty()
                full_response = ""

                with ollama.stream(
                    model="mistral",
                    prompt=full_prompt,
                    system=system_prompt
                ) as stream:
                    for chunk in stream:
                        full_response += chunk
                        response_placeholder.markdown(full_response)

            st.session_state.messages.append({"role": "assistant", "content": full_response})

        except Exception as e:
            st.error(f"Error: {str(e)}")
            st.info("💡 Asegúrate que Ollama está corriendo (`ollama serve`)")

# Botón para limpiar chat
if st.button("🗑️ Limpiar chat"):
    st.session_state.messages = []
    st.rerun()
