import re
from docx import Document

def parse_text_docx(filename):
    doc = Document(filename)
    questions = []
    current_q = None
    current_part = 0
    
    for p in doc.paragraphs:
        text = p.text.strip()
        if not text:
            continue
            
        if "PHẦN I." in text or "PHẦN I:" in text:
            current_part = 1
            continue
        elif "PHẦN II." in text or "PHẦN II:" in text:
            current_part = 2
            continue
        elif "PHẦN III." in text or "PHẦN III:" in text:
            current_part = 3
            continue
            
        # Match Câu 1:, Câu 2., vv
        câu_match = re.match(r'^Câu\s+\d+[:.]\s*(.*)', text, re.IGNORECASE)
        if câu_match:
            if current_q:
                questions.append(current_q)
            current_q = {
                'part': current_part if current_part else 1,
                'head_paragraphs': [p],
                'opts_paragraphs': [],
                'ans': None
            }
            continue
            
        if current_q:
            # Check if this paragraph contains options
            # Options A, B, C, D might be in one paragraph or multiple
            # If it starts with A., B., C., D. or a), b), c), d)
            opt_match = re.match(r'^([A-D]|[a-d])[.)]\s+', text)
            if opt_match:
                current_q['opts_paragraphs'].append(p)
            else:
                if len(current_q['opts_paragraphs']) == 0:
                    current_q['head_paragraphs'].append(p)
                else:
                    # Could be multiple lines of an option, or something else.
                    # Simplified: append to the last option
                    pass

    if current_q:
        questions.append(current_q)
        
    return questions

questions = parse_text_docx("Made 302.docx")
print("Parsed questions:", len(questions))
part1 = [q for q in questions if q['part'] == 1]
part2 = [q for q in questions if q['part'] == 2]
part3 = [q for q in questions if q['part'] == 3]
print(f"Part 1: {len(part1)}, Part 2: {len(part2)}, Part 3: {len(part3)}")
for q in part2:
    print("Part 2 Q has options:", len(q['opts_paragraphs']))
