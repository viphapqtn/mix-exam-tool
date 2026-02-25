from docx import Document

doc = Document("Made 302.docx")
with open("test_sizes_output.txt", "w", encoding="utf-8") as f:
    for i, p in enumerate(doc.paragraphs[:50]):
        sizes = []
        for r in p.runs:
            if r.font.size is not None:
                sizes.append(str(r.font.size.pt))
            else:
                sizes.append("None")
        if p.text.strip():
            f.write(f"P {i} sizes: {', '.join(sizes)} - text: {p.text[:30]}\n")
