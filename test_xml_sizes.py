from docx import Document
from lxml import etree

doc = Document("Made 302.docx")
lines_of_interest = []
for p in doc.paragraphs:
    if "Silicon" in p.text or "Câu 10" in p.text or "Câu 9" in p.text:
        lines_of_interest.append(p)

with open("test_xml_sizes_out_native.txt", "w", encoding="utf-8") as f:
    for p in lines_of_interest:
        f.write(f"Para: {p.text[:40]}\n")
        p_pr = p._p.xpath('./w:pPr')
        if p_pr:
            for ppr in p_pr:
                szs = ppr.xpath('.//w:sz/@w:val')
                if szs:
                    f.write(f"  pPr sz: {szs}\n")
        for r in p._p.xpath('./w:r'):
            szs = r.xpath('.//w:sz/@w:val')
            if szs:
                f.write(f"  rPr sz: {szs} - text: {''.join(r.xpath('.//w:t/text()'))}\n")
