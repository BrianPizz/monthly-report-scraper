import re
from pymupdf import pymupdf
from pathlib import Path


BASE_DIR = Path(__file__).parent
PDF_PATH = BASE_DIR / "files" / "BRIAN School Leader Monthly Report 2025-26.pdf"

doc =  pymupdf.open(PDF_PATH)

page = doc[0]

text = page.get_text()

# Find numbers from document
numbers = re.findall(r"\d+", text)
numbers = [int(n) for n in numbers]

print(numbers)