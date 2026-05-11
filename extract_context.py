from pypdf import PdfReader

pdf_path = r"d:\Estadistica Inferencial\Unidad 1\1.4 Tamaño de muestra para estimar la media y proporcion para muestreo estratificado.pdf"
reader = PdfReader(pdf_path)

for i in range(15, 23):
    print(f"--- PAGE {i+1} ---")
    try:
        print(reader.pages[i].extract_text()[:400]) # only the first 400 chars to avoid large dump if not needed
        text = reader.pages[i].extract_text()
        if 'costo' in text.lower() or '$' in text.lower():
            print(">>> HAS COST INFO:")
            print(text)
    except:
        pass
