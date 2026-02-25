from docx import Document

doc = Document("Made 302.docx")
for s in doc.styles:
    if s.name in ['Normal_35', 'Normal_1_0']:
        print(f"Style: {s.name}")
        base = s.base_style
        while base:
            print(f"  Base style: {base.name} - size: {base.font.size.pt if base.font.size else 'None'}")
            base = base.base_style
