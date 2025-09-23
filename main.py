import re
from pymupdf import pymupdf
from pathlib import Path


BASE_DIR = Path(__file__).parent
PDF_PATH = BASE_DIR / "files" / "BRIAN School Leader Monthly Report 2025-26.pdf"

doc =  pymupdf.open(PDF_PATH)

# Pull text from first page of file
page = doc[0]
text = page.get_text()

# Find numbers from document
numbers = re.findall(r"\d+", text)
numbers = [int(n) for n in numbers]
refined_numbers = numbers[37:]

# print(refined_numbers)

# Index label objects 
index_labels = [
    {"label": "Sub", "index_value": 1},
    {"label": "IS K-8", "index_value": 15},
    {"label": "IS 9-12", "index_value": 2},
    {"label": "SWD K-8", "index_value": 4},
    {"label": "SWD 9-12", "index_value": 5},
    {"label": "OSS K-8", "index_value": 9},
    {"label": "OSS 9-12", "index_value": 10},
    {"label": "EX K-8", "index_value": 11},
    {"label": "EX 9-12", "index_value": 12},
    {"label": "ER", "index_value": 13},
    {"label": "MDM", "index_value": 14},
]

# Print target fields labels and values
for i in index_labels:
    label = i["label"]
    index = i["index_value"]
    value = refined_numbers[index]
    print(f"{label}: {value}")