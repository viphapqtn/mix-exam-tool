from docx import Document

doc = Document("Made 302.docx")
for s in doc.styles:
    if s.name in ['Normal_35', 'Normal_1_0']:
        size = s.font.size.pt if s.font.size else "None"
        print(f"Style {s.name} size: {size}")
