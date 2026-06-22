import pypdfium2 as pdfium
import easyocr
import numpy as np
from PIL import Image
import re

def debug_pdf(pdf_path):
    print(f"Analyzing: {pdf_path}")
    reader = easyocr.Reader(['en'], gpu=False)
    doc = pdfium.PdfDocument(pdf_path)
    
    for i in range(len(doc)):
        print(f"\n--- Page {i} ---")
        page = doc[i]
        for rot in [0, 90, 180, 270]:
            bitmap = page.render(scale=2, rotation=rot)
            img = bitmap.to_pil()
            results = reader.readtext(np.array(img), detail=0)
            text = " ".join(results)
            print(f"ROT {rot}: {text}")
            
            # Check if our target exists here
            if "7369421124" in text.replace(" ", ""):
                print(f"!!! FOUND TRACKING at ROT {rot} !!!")
            if "4000003628" in text.replace(" ", ""):
                print(f"!!! FOUND REF at ROT {rot} !!!")
    doc.close()

if __name__ == "__main__":
    debug_pdf("pdf/4000003628.pdf")
