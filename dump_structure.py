from docx import Document

doc = Document("De goc.docx")
table = doc.tables[0]
rows = table.rows

with open("structure.txt", "w", encoding="utf-8") as f:
    for i, row in enumerate(rows):
        col0 = row.cells[0].text.strip()
        f.write(f"{i}: {col0}\n")
