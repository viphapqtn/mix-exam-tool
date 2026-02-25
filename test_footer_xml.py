from docx import Document
from lxml import etree

doc = Document("Made 302.docx")
with open("test_footer_out.xml", "w", encoding="utf-8") as f:
    for section in doc.sections:
        for p in section.footer.paragraphs:
            f.write("FOOTER PARA:\n")
            f.write(etree.tostring(p._p, pretty_print=True, encoding="unicode"))
            f.write("\n")
