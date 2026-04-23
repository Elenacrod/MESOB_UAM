import os
import pdfplumber
import chromadb
from pathlib import Path

PDF_FOLDER = r"C:\Users\MC.5055521\Desktop\Claude acceso\Documentacion_MESOB"
DB_PATH = "./mesob_db"

def extract_pdf_text():
    """Extrae texto de todos los PDFs"""
    documents = []

    for pdf_file in Path(PDF_FOLDER).glob("*.pdf"):
        print(f"Procesando: {pdf_file.name}")
        try:
            with pdfplumber.open(pdf_file) as pdf:
                text = ""
                for page in pdf.pages:
                    text += page.extract_text() or ""

                if text.strip():
                    documents.append({
                        "id": pdf_file.stem,
                        "source": pdf_file.name,
                        "content": text
                    })
        except Exception as e:
            print(f"Error en {pdf_file.name}: {e}")

    return documents

def create_embeddings():
    """Crea embeddings y almacena en ChromaDB"""
    import ollama

    documents = extract_pdf_text()

    if not documents:
        print("No se encontraron documentos")
        return

    # Inicializa ChromaDB
    client = chromadb.PersistentClient(path=DB_PATH)
    collection = client.get_or_create_collection(name="mesob_docs")

    print(f"\nCreando embeddings para {len(documents)} documentos...")

    for doc in documents:
        # Divide el contenido en chunks para mejor búsqueda
        content = doc["content"]
        chunk_size = 1000
        chunks = [content[i:i+chunk_size] for i in range(0, len(content), chunk_size)]

        for idx, chunk in enumerate(chunks):
            if chunk.strip():
                chunk_id = f"{doc['id']}_chunk_{idx}"

                # Genera embedding usando Ollama
                try:
                    embedding_response = ollama.embeddings(
                        model="mistral",
                        prompt=chunk
                    )
                    embedding = embedding_response["embedding"]

                    collection.add(
                        ids=[chunk_id],
                        embeddings=[embedding],
                        documents=[chunk],
                        metadatas=[{"source": doc["source"], "chunk": idx}]
                    )
                    print(f"✓ {chunk_id}")
                except Exception as e:
                    print(f"Error en embedding {chunk_id}: {e}")

    print(f"\n✅ Base de datos creada en: {DB_PATH}")

if __name__ == "__main__":
    create_embeddings()
