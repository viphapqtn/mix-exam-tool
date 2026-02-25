import re
import copy
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

def has_answer_mark(paragraph):
    for run in paragraph.runs:
        if getattr(run.font, 'highlight_color', None):
            return True
        if run.underline:
            return True
    return False

def parse_freeform_docx(filename):
    doc = Document(filename)
    questions = []
    current_q = None
    current_part = 1
    
    for p in doc.paragraphs:
        text = p.text.strip()
        if not text:
            continue
            
        part_match = re.search(r'PHẦN ([I]+)', text)
        if part_match:
            r = part_match.group(1)
            if r == 'I': current_part = 1
            elif r == 'II': current_part = 2
            elif r == 'III': current_part = 3
            continue
            
        câu_match = re.match(r'^Câu\s+\d+[:.]', text, re.IGNORECASE)
        if câu_match:
            if current_q and (len(current_q['opts']) > 0 or current_q['part'] == 3):
                questions.append(current_q)
            current_q = {
                'part': current_part,
                'head': [p],
                'opts': [],
                'ans': None
            }
            continue
            
        if current_q:
            # Maybe it's part 3 which doesn't have A B C D options but might have "Đáp án: "
            if current_q['part'] == 3:
                # We need to extract the answer for part 3. Maybe they highlight the answer or just write it below.
                if text.upper().startswith("ĐÁP ÁN:"):
                    current_q['ans'] = text.split(":", 1)[1].strip()
                elif text.upper().startswith("ĐÁP ÁN"):
                    current_q['ans'] = text.replace("Đáp án", "").strip()
                else:
                    current_q['head'].append(p)
                continue

            opt_match = re.match(r'^([A-D]|[a-d])[.)]\s+(.*)', text)
            if opt_match:
                opt_char = opt_match.group(1).upper()
                is_correct = has_answer_mark(p)
                current_q['opts'].append({
                    'char': opt_char,
                    'paragraph': p,
                    'is_correct': is_correct
                })
                if is_correct:
                    if current_q['part'] == 1:
                        current_q['ans'] = opt_char
                    elif current_q['part'] == 2:
                        if not current_q['ans']:
                            current_q['ans'] = ""
                        # True/False format -> they highlight the true options? Let's say yes. True = highlighted
                        # Actually wait, Part 2 True/False: the user might highlight the letter a), b), c), d) if True, or write SDSD. 
                        # Let's support writing "Đ" or "S" or highlighting?
            else:
                if len(current_q['opts']) == 0:
                    current_q['head'].append(p)
                else:
                    # Append it!
                    pass

    if current_q and (len(current_q['opts']) > 0 or current_q['part'] == 3):
        questions.append(current_q)
        
    return questions

questions = parse_freeform_docx("Made 302.docx")
with open("test_out.txt", "w", encoding="utf-8") as f:
    f.write(f"Parsed queries: {len(questions)}\n")
    for q in questions:
        f.write(f"Part: {q['part']}, Head len: {len(q['head'])}, Opts: {len(q['opts'])}, Ans: {q['ans']}\n")
