import fitz  # PyMuPDF
import os

pdf_path = r"C:\Users\MC.5055521\Desktop\Claude code\CRIM_Intervals_Musicologia.pdf"
out_dir   = r"C:\Users\MC.5055521\Desktop\Claude code"

doc = fitz.open(pdf_path)
print(f"Pages: {len(doc)}")
for i, page in enumerate(doc):
    mat = fitz.Matrix(150/72, 150/72)  # 150 DPI
    pix = page.get_pixmap(matrix=mat)
    out_path = os.path.join(out_dir, f"slide-{i+1:02d}.jpg")
    pix.save(out_path)
    print(f"Saved {out_path}")
doc.close()
