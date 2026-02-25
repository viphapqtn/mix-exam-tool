from docx import Document
from lxml import etree
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

def normalize_p(p_node, font_name='Times New Roman', font_size_val='26'):
    # Clear paragraph style to standard Normal
    p_style = p_node.xpath('./w:pPr/w:pStyle')
    if p_style:
        p_style[0].set(qn('w:val'), 'Normal')
        
    for r in p_node.xpath('.//w:r'):
        rPrs = r.xpath('./w:rPr')
        if not rPrs:
            rPr = OxmlElement('w:rPr')
            r.insert(0, rPr)
            rPrs = [rPr]
        rPr = rPrs[0]
        
        # Fonts
        rFonts = rPr.xpath('./w:rFonts')
        if not rFonts:
            rFont = OxmlElement('w:rFonts')
            rPr.append(rFont)
            rFonts = [rFont]
        for rf in rFonts:
            rf.set(qn('w:ascii'), font_name)
            rf.set(qn('w:hAnsi'), font_name)
            rf.set(qn('w:cs'), font_name)
            
        # Size
        sz_nodes = rPr.xpath('./w:sz')
        if not sz_nodes:
            sz_node = OxmlElement('w:sz')
            rPr.append(sz_node)
            sz_nodes = [sz_node]
        for s in sz_nodes:
            s.set(qn('w:val'), font_size_val)
            
        szCs_nodes = rPr.xpath('./w:szCs')
        if not szCs_nodes:
            szCs_node = OxmlElement('w:szCs')
            rPr.append(szCs_node)
            szCs_nodes = [szCs_node]
        for s in szCs_nodes:
            s.set(qn('w:val'), font_size_val)

doc = Document("Made 302.docx")
lines_of_interest = []
for p in doc.paragraphs:
    if "Silicon" in p.text or "Câu 10" in p.text:
        lines_of_interest.append(p)

for p in lines_of_interest:
    normalize_p(p._p)

doc.save("test_out_normalized.docx")
print("Done")
