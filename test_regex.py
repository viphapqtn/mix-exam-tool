import re
text = "-------------- HẾT ---------------"
print("Standard ASCII dash:", bool(re.search(r'[-_*]+\s*HẾT\s*[-_*]+', text, re.IGNORECASE)))
text2 = "———— HẾT ————"
print("Em-dash:", bool(re.search(r'[-_*]+\s*HẾT\s*[-_*]+', text2, re.IGNORECASE)))
