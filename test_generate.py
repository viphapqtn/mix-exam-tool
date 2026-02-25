import os
from app import parse_docx, generate_exam

questions = parse_docx("De goc.docx")
print("Parsed questions:", len(questions))

exam_bytes, exam_ans = generate_exam(questions, "101")
with open("test_output_101.docx", "wb") as f:
    f.write(exam_bytes)
    
print("Exam ans:", exam_ans[:5])
print("Done writing test file.")
