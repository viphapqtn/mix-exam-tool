import re
from docx import Document

def parse_text_docx(filename):
    doc = Document(filename)
    questions = []
    current_q = None
    current_part = 0
    
    for i, p in enumerate(doc.paragraphs):
        text = p.text.strip()
        if not text: continue
        
        part_match = re.search(r'PHẦN ([I]+)', text)
        if part_match:
            r = part_match.group(1)
            if r == 'I': current_part = 1
            if r == 'II': current_part = 2
            if r == 'III': current_part = 3
            continue
            
        câu_match = re.match(r'^Câu\s+\d+[:.]', text, re.IGNORECASE)
        if câu_match:
            if current_q: questions.append(current_q)
            current_q = {'part': current_part if current_part else 1, 'head': [], 'opts': []}
            current_q['head'].append(p)
            continue
            
        if current_q:
            # Check options
            opt_match = re.match(r'^([A-D]|[a-d])[.)]\s+', text)
            if opt_match:
                # Need to handle A. B. on same line?
                # Sometimes A and B are in the SAME paragraph text.
                # Oh... if they are in the exact same paragraph, then we have a problem swapping them.
                # We need to split the paragraph by runs or just let it be if it's too complex.
                current_q['opts'].append(p)
            elif "Đáp án" in text:
                current_q['ans'] = p.text
            else:
                if len(current_q['opts']) == 0:
                    current_q['head'].append(p)
                else:
                    current_q['opts'].append(p)

    if current_q: questions.append(current_q)
    return questions

questions = parse_text_docx("Made 302.docx")
for q in questions[:3]:
    print("Part:", q['part'], "Head len:", len(q['head']), "Opts len:", len(q['opts']))
    for o in q['opts']:
        print("Opt:", o.text)
