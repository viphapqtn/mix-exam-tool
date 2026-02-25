from docx import Document
doc = Document("De goc.docx")
lines = []
for para in doc.paragraphs:
    if para.text.strip():
        lines.append(para.text)
for x in lines[:50]:
    print(x)
