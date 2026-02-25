from docx import Document

try:
    doc = Document("Made 302.docx")
    with open("made_302_full.txt", "w", encoding="utf-8") as f:
        for p in doc.paragraphs:
            f.write(p.text + "\n")
except Exception as e:
    print("Error:", e)
