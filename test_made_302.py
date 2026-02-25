from docx import Document

doc = Document("Made 302.docx")
lines = [p.text for p in doc.paragraphs if p.text.strip()]
with open("test_made_302.txt", "w", encoding="utf-8") as f:
    for i, l in enumerate(lines[:30]):
        f.write(f"{i}: {l}\n")
