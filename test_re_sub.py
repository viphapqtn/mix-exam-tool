import re
text = "Câu 10: "
print(re.sub(r'^\s*Câu\s+\d+[:.]\s*', '', text, flags=re.IGNORECASE))
