from docx import Document

doc = Document("Made 302.docx")
with open("test_p_styles.txt", "w", encoding="utf-8") as f:
    for i, p in enumerate(doc.paragraphs):
        if "Silicon" in p.text or "Câu 10" in p.text or "Câu 9" in p.text:
            f.write(f"P index {i}: '{p.text[:30]}' style: {p.style.name} \n")
            for r in p.runs:
                r_size = r.font.size.pt if r.font.size else "None"
                f.write(f"  Run '{r.text}' - size {r_size}\n")
