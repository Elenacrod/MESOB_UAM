import pdfplumber
import pickle
from pathlib import Path

PDF_FOLDER = r"C:\Users\MC.5055521\Desktop\Claude acceso\Documentacion_MESOB"
DB_FILE = "./mesob_documents.pkl"

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
                    documents.append(text)
                    print(f"✓ {pdf_file.name}")
        except Exception as e:
            print(f"Error en {pdf_file.name}: {e}")

    return documents

def save_documents():
    """Extrae y guarda documentos en pickle"""
    documents = extract_pdf_text()

    if not documents:
        print("No se encontraron documentos")
        return

    # Guarda en pickle
    with open(DB_FILE, "wb") as f:
        pickle.dump(documents, f)

    print(f"\n✅ {len(documents)} documentos guardados en: {DB_FILE}")

if __name__ == "__main__":
    save_documents()
