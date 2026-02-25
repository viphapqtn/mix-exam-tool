from docx import Document

doc = Document("Made 302.docx")
lines = []
for p in doc.paragraphs:
    runs_info = []
    for r in p.runs:
        props = []
        if r.bold: props.append("B")
        if r.underline: props.append("U")
        if r.italic: props.append("I")
        if r.font.color and r.font.color.rgb: props.append(f"C({r.font.color.rgb})")
        if getattr(r.font, 'highlight_color', None): props.append(f"H({r.font.highlight_color})")
        
        runs_info.append(f"[{','.join(props)}]{r.text}")
    lines.append("".join(runs_info))

with open("test_made_302_runs.txt", "w", encoding="utf-8") as f:
    for i, l in enumerate(lines[:50]):
        if l.strip():
            f.write(f"{i}: {l}\n")
