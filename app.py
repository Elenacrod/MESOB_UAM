import streamlit as st
import chromadb
import ollama
from pathlib import Path

DB_PATH = "./mesob_db"

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
    st.markdown("💡 *Asistente local sin costos - Alimentado por IA*")

# Inicializa base de datos vectorial
@st.cache_resource
def load_db():
    try:
        client = chromadb.PersistentClient(path=DB_PATH)
        collection = client.get_or_create_collection(name="mesob_docs")
        return collection
    except Exception as e:
        st.error(f"Error cargando base de datos: {e}")
        return None

# Inicializa sesión de chat
if "messages" not in st.session_state:
    st.session_state.messages = []

collection = load_db()

# Muestra historial de chat
for msg in st.session_state.messages:
    with st.chat_message(msg["role"]):
        st.markdown(msg["content"])

# Input del usuario
if prompt := st.chat_input("¿Pregunta sobre MESOB?"):
    st.session_state.messages.append({"role": "user", "content": prompt})

    with st.chat_message("user"):
        st.markdown(prompt)

    # Busca documentos relevantes
    with st.spinner("Buscando en documentación..."):
        try:
            # Genera embedding de la pregunta
            question_embedding = ollama.embeddings(
                model="mistral",
                prompt=prompt
            )["embedding"]

            # Busca documentos similares
            results = collection.query(
                query_embeddings=[question_embedding],
                n_results=3
            )

            # Prepara contexto
            context = ""
            if results["documents"] and results["documents"][0]:
                for doc, source in zip(results["documents"][0], results["metadatas"][0]):
                    context += f"\n[{source['source']}]\n{doc}\n"

            # Genera respuesta usando Ollama
            system_prompt = """Eres un asistente experto en el programa MESOB (Máster en Educación Secundaria Obligatoria y Bachillerato).
Responde preguntas basándote en la documentación proporcionada.
Si no encuentras la respuesta en la documentación, indícalo claramente e incluye estos recursos:

📌 **Información oficial:**
- Web: https://www.uam.es/educacion/estudios/practicas-externas/informacion-practicas-posgrado/mesob
- Email: delegacion.educacion.practicasmesob@uam.es

Para dudas adicionales, recomienda contactar al email de la delegación.
Responde en español y sé conciso."""

            full_prompt = f"""Contexto de la documentación:
{context}

Pregunta: {prompt}

Respuesta:"""

            # Genera respuesta en streaming
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
            st.error(f"Error al generar respuesta: {e}")

# Botón para limpiar chat
col1, col2 = st.columns(2)
with col1:
    if st.button("🗑️ Limpiar chat"):
        st.session_state.messages = []
        st.rerun()

with col2:
    st.info("💡 Asistente local - sin costos de API")
