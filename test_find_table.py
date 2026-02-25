from docx import Document

doc = Document("test_out_linear.docx")

with open("test_find_table_out.txt", "w", encoding="utf-8") as f:
    for idx, t in enumerate(doc.tables):
        for r_idx, row in enumerate(t.rows):
            cells = [c.text.strip() for c in row.cells]
            f.write(f"Table {idx} Row {r_idx}: {cells}\n")
