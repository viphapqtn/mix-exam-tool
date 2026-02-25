from app import generate_exam_linear, parse_freeform_docx

questions = parse_freeform_docx("Made 302.docx")
print("Parsed questions:", len(questions))

exam_bytes, ans = generate_exam_linear(questions, "101", "Made 302.docx")
with open("test_out_linear.docx", "wb") as f:
    f.write(exam_bytes)

print("Done generating test linear exam")
