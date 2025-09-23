import re
from pymupdf import pymupdf
from pathlib import Path


BASE_DIR = Path(__file__).parent
PDF_PATH = BASE_DIR / "files" / "BRIAN School Leader Monthly Report 2025-26.pdf"

doc =  pymupdf.open(PDF_PATH)

# Pull text from first page of file
page = doc[0]
text = page.get_text()
page_dict = page.get_text("dict")
blocks = page.get_text("blocks")

# Find numbers from document
numbers = re.findall(r"\d+", text)
numbers = [int(n) for n in numbers]
refined_numbers = numbers[37:]

# print(refined_numbers)

# Define School object class
class School:
     def __init__(self, name, is_both, data_list):
          self.name = name
          self.is_both = is_both
          self.data_list = data_list

# Locate school name 
"""""
for block in blocks:
    x0, y0, x1, y1, text_content, block_no, block_type = block
    if block_no == 17: # School name is in block 17
        print(f"Block {block_no}:")
        print(f"  Text: {text_content}")
        print("-" * 20)
"""

# Extract first line of school name block
first_line_text = ""
school_name = ""
if page_dict and "blocks" in page_dict and len(page_dict["blocks"]) > 0:
    target_block = page_dict["blocks"][18] # School name is in block 17

    if "lines" in target_block and len(target_block["lines"]) > 0:
        first_line = target_block["lines"][0] # Get the first line of that block

        # Concatenate spans to get the full line text
        for span in first_line["spans"]:
                first_line_text += span["text"]
    school_name = first_line_text

# Initialize is_both referring to ES and HS
is_both = False
es_and_hs_schools = [
    "Arts and College Preparatory Academy",
    "Columbus Arts and Tech Academy",
    "Columbus Preparatory Academy",
    "Great River Connections Academy",
    "Northeast Ohio College Preparatory School",
    "Ohio Connections Academy",
    "Ohio Virtual Academy",
    "Wildwood Environmental Academy",
    # "Brian's Sample School"
]

# Set is_both to True if school name is found in list
if school_name in es_and_hs_schools:
     is_both = True


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

# Create target field list with labels and ref indexes
target_field_list = []
for i in index_labels:
    label = i["label"]
    index = i["index_value"]
    value = refined_numbers[index]

    # Create object with label and target value
    field_object = {"label": label, "value": value}
    # Add object to list 
    target_field_list.append(field_object) 


school = School(school_name, is_both, target_field_list)
print(school.name)
print(f"ES and HS ? : {school.is_both}")
for obj in school.data_list:
     label = obj["label"]
     value = obj["value"]
     print(f"{label}: {value}")
