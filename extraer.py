from docx import Document

doc = Document("CERTIFICADO DE PRÁCTICAS_Alvaro Martinez.docx.docx")

for p in doc.paragraphs:
    t = p.text.strip()
    if "Que" in t and "DNI" in t:
        print("TEXTO:")
        print(t)
        print("RUNS:", len(p.runs))
        print("-" * 50)
        break
