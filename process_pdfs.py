import pdfplumber
from pathlib import Path

PDF_FOLDER = r"C:\Users\MC.5055521\Desktop\Claude acceso\Documentacion_MESOB"
OUTPUT_FOLDER = "./docs"

def convert_pdfs_to_txt():
    output_path = Path(OUTPUT_FOLDER)
    output_path.mkdir(exist_ok=True)

    for pdf_file in Path(PDF_FOLDER).glob("*.pdf"):
        print(f"Procesando: {pdf_file.name}")
        try:
            with pdfplumber.open(pdf_file) as pdf:
                text = ""
                for page in pdf.pages:
                    text += page.extract_text() or ""

            if text.strip():
                txt_file = output_path / (pdf_file.stem + ".txt")
                with open(txt_file, "w", encoding="utf-8") as f:
                    f.write(text)
                print(f"✓ Guardado: {txt_file.name}")
        except Exception as e:
            print(f"Error en {pdf_file.name}: {e}")

    print(f"\n✅ Archivos guardados en: {OUTPUT_FOLDER}")

if __name__ == "__main__":
    convert_pdfs_to_txt()
