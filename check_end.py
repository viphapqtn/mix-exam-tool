import copy
from docx import Document

doc = Document("De goc.docx")
table = doc.tables[0]
rows = table.rows

with open("check_end.txt", "w", encoding="utf-8") as f:
    for i in [77, 83, 85, 87]:
        row = rows[i]
        f.write(f"Row {i} {row.cells[0].text.strip()}: {row.cells[1].text.strip()}\n")
