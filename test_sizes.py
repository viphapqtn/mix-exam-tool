import codecs
import sys
from docx import Document

sys.stdout = codecs.getwriter("utf-8")(sys.stdout.detach())

doc = Document("Made 302.docx")
for i, p in enumerate(doc.paragraphs[:50]):
    sizes = []
    for r in p.runs:
        if r.font.size is not None:
            sizes.append(str(r.font.size.pt))
        else:
            sizes.append("None")
    if p.text.strip():
        print(f"P {i} sizes: {', '.join(sizes)} - text: {p.text[:30]}")
