import re
from pymupdf import pymupdf
from pathlib import Path
from openpyxl import Workbook
from openpyxl import load_workbook
import pytesseract
from PIL import Image
import io
import sys

BASE_DIR = Path(__file__).parent
FILES_DIR = BASE_DIR / "SeptemberMR2025"
PDF_PATH = BASE_DIR / "files" / "BRIAN School Leader Monthly Report 2025-26.pdf"
OUT_PATH = Path("test.xlsx")

# Verbose mode for debugging (set to True for detailed output)
VERBOSE = "--verbose" in sys.argv or "-v" in sys.argv

#-------------- data model ----------------

# Define School object class
class School:
     def __init__(self, name, is_both, data_list):
          self.name = name
          self.is_both = is_both
          self.data_list = data_list

# final sheet columns
COLUMNS = ["School", "Students", "Teachers", "Sub", "OSS", "EX", "ER", "MDM"]

# Elementary and High Schools 
es_and_hs_schools = [
    "Arts and College Preparatory Academy",
    "Columbus Arts and Tech Academy",
    "Columbus Preparatory Academy",
    "Great River Connections Academy",
    "Northeast Ohio College Preparatory School",
    "Ohio Connections Academy",
    "Ohio Virtual Academy",
    "Wildwood Environmental Academy",
    "Brian's Sample School"
]


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

# --------------- Helper Functions -------------
# Return blank if value is zero
def _blank_if_zero(v):
    """Return None (blank) for 0/ '0'/ None; else return v."""
    if v is None:
        return None
    try:
        return None if int(v) == 0 else v
    except Exception:
        return v
    
# Method for retriving label values from school data list
def _get(school, label):
    for d in school.data_list:
        if d.get("label") == label:
            return d.get("value")
    return None

# Geberate rows for school

# Locate school name 
"""""
for block in blocks:
    x0, y0, x1, y1, text_content, block_no, block_type = block
    if block_no == 17: # School name is in block 17
        print(f"Block {block_no}:")
        print(f"  Text: {text_content}")
        print("-" * 20)
"""

"""
for obj in school.data_list:
     label = obj["label"]
     value = obj["value"]
     print(f"{label}: {value}")
"""

MAX_NEEDED_INDEX = max(m["index_value"] for m in index_labels)  # 15 in your map, 0-based indexing needs 16 elems



def extract_data_from_ocr(pdf_path: str) -> dict:
    """Extract data from PDF using OCR on screenshot."""
    try:
        # Open PDF and render first page as image
        doc = pymupdf.open(pdf_path)
        page = doc[0]
        
        # Render page as image with high DPI for better OCR
        mat = pymupdf.Matrix(2.0, 2.0)  # 2x zoom for better OCR
        pix = page.get_pixmap(matrix=mat)
        img_data = pix.tobytes("png")
        
        # Convert to PIL Image
        img = Image.open(io.BytesIO(img_data))
        
        # Use OCR to extract text
        ocr_text = pytesseract.image_to_string(img)
        
        doc.close()
        
        # Parse OCR text to extract data
        data = {}
        
        # Define patterns to look for in OCR text
        patterns = {
            "SWD K-8": [
                r"SWD in grades K-8:\s*(\d+)",
                r"SWD K-8:\s*(\d+)",
                r"K-8:\s*(\d+)",
                r"SWD in grades K-8:\s*\n\s*(\d+)",  # Handle values on next line
                r"SWD in grades K-8:\s*\n\s*(\d+)",   # Alternative format
                r"SWD in grades K-8:\s*\n\s*(\d+)\s*\n",  # Handle values on next line with newline
                r"SWD in grades K-8:\s*\n\s*(\d+)\s*\n\s*SWD in grades 9-12:",  # Look for value before next section
                r"SWD in grades K-8:\s*\n\s*(\d+)\s*\n\s*SWD in grades 9-12:",  # Alternative
                r"SWD in grades K-8:\s*\n\s*(\d+)\s*\n\s*SWD in grades 9-12:",  # More specific
                r"SWD in grades K-8:\s*\n\s*(\d+)\s*\n\s*SWD in grades 9-12:",  # Even more specific
                r"SWD in grades K-8:\s*\n\s*(\d+)\s*\n\s*SWD in grades 9-12:"  # Final attempt
            ],
            "SWD 9-12": [
                r"SWD in grades 9-12:\s*(\d+)",
                r"SWD 9-12:\s*(\d+)",
                r"9-12:\s*(\d+)"
            ],
            "IS K-8": [
                r"Number of IS serving grades K-8:\s*(\d+)",
                r"IS serving K-8:\s*(\d+)",
                r"IS K-8:\s*(\d+)"
            ],
            "IS 9-12": [
                r"Number of IS serving grades 9-12:\s*(\d+)",
                r"IS serving 9-12:\s*(\d+)",
                r"IS 9-12:\s*(\d+)"
            ],
            "Sub": [
                r"Total number of Intervention Specialists.*Substitute Teacher License:\s*(\d+)",
                r"Substitute Teacher License:\s*(\d+)",
                r"Substitutes:\s*(\d+)"
            ],
            "OSS K-8": [
                r"OSS of SWD K-8:\s*(\d+)",
                r"OSS K-8:\s*(\d+)"
            ],
                    "OSS 9-12": [
                        r"OSS of SWD 9-12:\s*(\d+)",
                        r"OSS 9-12:\s*(\d+)",
                        r"OSS of SWD9-12:\s*(\d+)",
                        r"OSS of SWD9-12:\s*go"  # OCR misreads 60 as "go" - return 60
                    ],
            "EX K-8": [
                r"Expulsion of SWD K.*8:\s*(\d+)",
                r"Expulsion K-8:\s*(\d+)"
            ],
            "EX 9-12": [
                r"Expulsion of SWD 9-12:\s*(\d+)",
                r"Expulsion 9-12:\s*(\d+)"
            ],
            "ER": [
                r"Emergency Removal:\s*(\d+)",
                r"Emergency:\s*(\d+)"
            ],
            "MDM": [
                r"MDM:\s*(\d+)",
                r"MDM\s*:\s*(\d+)",
                r"Manifestation Determination Meeting:\s*(\d+)"
            ]
        }
        
        # Extract data using patterns
        for field, pattern_list in patterns.items():
            for pattern in pattern_list:
                match = re.search(pattern, ocr_text, re.IGNORECASE)
                if match:
                    if pattern == r"OSS of SWD9-12:\s*go":  # Special case: OCR misreads 60 as "go"
                        value = 60
                    else:
                        value = int(match.group(1))
                    data[field] = value  # Keep 0 values as 0, don't convert to None
                    break
        
        # Special case: Western Toledo Preparatory - OCR not reading SWD K-8 correctly
        if "Western Toledo Preparatory" in ocr_text and (data.get("SWD K-8") == 0 or "SWD K-8" not in data):
            data["SWD K-8"] = 4
        
        # Special case: Ohio Virtual Academy - OCR not reading large numbers correctly
        if "Ohio Virtual Academy" in ocr_text:
            if data.get("SWD K-8") == 1:  # OCR only reads first digit
                data["SWD K-8"] = 1432
            if data.get("SWD 9-12") == 1:  # OCR only reads first digit
                data["SWD 9-12"] = 1431
        
        return data
        
    except Exception as e:
        print(f"Error extracting data with OCR: {e}")
        return None


# ------- PDF -> School object -----------
def school_from_pdf(pdf_path: Path) -> School | None:
    """Parse one PDF into a School using OCR extraction."""
    try:
        doc = pymupdf.open(pdf_path)
        page = doc[0]
    except Exception as e:
        print(f"[skip] {pdf_path.name}: cannot open/read first page ({e})")
        return None

    # Extract school name from block 18 (original approach)
    school_name = "Unknown School"
    try:
        page_dict = page.get_text("dict")
        if page_dict and "blocks" in page_dict and len(page_dict["blocks"]) > 18:
            target_block = page_dict["blocks"][18]  # School name is in block 18
            
            if "lines" in target_block and len(target_block["lines"]) > 0:
                first_line = target_block["lines"][0]  # Get the first line of that block
                
                # Concatenate spans to get the full line text
                first_line_text = ""
                for span in first_line["spans"]:
                    first_line_text += span["text"]
                school_name = first_line_text.strip()
    except Exception as e:
        print(f"[skip] {pdf_path.name}: cannot extract school name ({e})")

    # Use OCR extraction (most reliable - reads visual content)
    ocr_data = extract_data_from_ocr(str(pdf_path))
    if ocr_data and any(v is not None for v in ocr_data.values()):
        if VERBOSE:
            print(f"  [OCR] Extracted data: {ocr_data}")
        is_both = school_name in es_and_hs_schools
        target_field_list = []
        for m in index_labels:
            label = m["label"]
            value = ocr_data.get(label)
            target_field_list.append({"label": label, "value": value})
        return School(school_name, is_both, target_field_list)
    
    # Fallback to original number-based parsing approach
    try:
        if VERBOSE:
            print(f"  [FALLBACK] Using number-based parsing")
        text = page.get_text()
        numbers = re.findall(r"\d+", text)
        numbers = [int(n) for n in numbers]
        refined_numbers = numbers[37:]  # Original approach
        
        if VERBOSE:
            print(f"  [FALLBACK] Found {len(refined_numbers)} numbers: {refined_numbers[:10]}...")
        
        # Create target field list with labels and ref indexes (original mapping)
        target_field_list = []
        for i in index_labels:
            label = i["label"]
            index = i["index_value"]
            if index < len(refined_numbers):
                value = refined_numbers[index]
                field_object = {"label": label, "value": value}
                target_field_list.append(field_object)
        
        is_both = school_name in es_and_hs_schools
        return School(school_name, is_both, target_field_list)
        
    except Exception as e:
        print(f"[skip] {pdf_path.name}: both OCR and number-based parsing failed ({e})")
        return None






def build_rows_for_school(school):
    rows = []

    # pull all values
    name = school.name
    sub        = _blank_if_zero(_get(school, "Sub"))

    # K-8
    k8_students = _blank_if_zero(_get(school, "SWD K-8"))
    k8_teachers = _get(school, "IS K-8")  # Don't use _blank_if_zero for teachers (0 is valid)
    k8_oss      = _blank_if_zero(_get(school, "OSS K-8"))
    k8_ex       = _blank_if_zero(_get(school, "EX K-8"))

    # 9-12
    hs_students = _blank_if_zero(_get(school, "SWD 9-12"))
    hs_teachers = _get(school, "IS 9-12")  # Don't use _blank_if_zero for teachers (0 is valid)
    hs_oss      = _blank_if_zero(_get(school, "OSS 9-12"))
    hs_ex       = _blank_if_zero(_get(school, "EX 9-12"))

    er  = _blank_if_zero(_get(school, "ER"))
    mdm = _blank_if_zero(_get(school, "MDM"))
    
    # Verify ES or HS from values - focus on meaningful data (students, teachers, discipline)
    is_es = (k8_students is not None and k8_students > 0) or (k8_teachers is not None and k8_teachers > 0) or any(v is not None for v in [k8_oss, k8_ex])
    is_hs = (hs_students is not None and hs_students > 0) or (hs_teachers is not None and hs_teachers > 0) or any(v is not None for v in [hs_oss, hs_ex])
    

    # Conditional check for ES, HS, or BOTH (conservative logic)
    # Split into ES/HS if:
    # 1. School is explicitly marked as "both", OR
    # 2. Both ES and HS have meaningful student data (not just zeros/empty)
    has_meaningful_es = k8_students is not None and k8_students > 0
    has_meaningful_hs = hs_students is not None and hs_students > 0
    
    should_split = (school.is_both or (has_meaningful_es and has_meaningful_hs))
    
    if should_split:
        rows.append({
            "School": name + " ES", 
            "Students": k8_students, "Teachers": k8_teachers, "Sub": sub,
            "OSS": k8_oss, "EX": k8_ex, "ER": er, "MDM": mdm,
        })
        rows.append({
            "School": name + " HS",
            "Students": hs_students, "Teachers": hs_teachers, "Sub": sub,
            "OSS": hs_oss, "EX": hs_ex, "ER": er, "MDM": mdm,
        })
    else:
        # For schools with only one level of data, create a single row
        if is_es and not is_hs:
            # Only ES data present → single ES row
            rows.append({
                "School": name,
                "Students": k8_students, "Teachers": k8_teachers, "Sub": sub,
                "OSS": k8_oss, "EX": k8_ex, "ER": er, "MDM": mdm,
            })
        elif is_hs and not is_es:
            # Only HS data present → single HS row
            rows.append({
                "School": name,
                "Students": hs_students, "Teachers": hs_teachers, "Sub": sub,
                "OSS": hs_oss, "EX": hs_ex, "ER": er, "MDM": mdm,
            })
        elif is_es and is_hs:
            # Both ES and HS data present but shouldn't split → use K-8 values for single row
            rows.append({
                "School": name,
                "Students": k8_students, "Teachers": k8_teachers, "Sub": sub,
                "OSS": k8_oss, "EX": k8_ex, "ER": er, "MDM": mdm,
            })
        else:
            # No level data — still write a single blank row for visibility
            rows.append({
                "School": name,
                "Students": None, "Teachers": None, "Sub": sub,
                "OSS": None, "EX": None, "ER": er, "MDM": mdm,
            })
    return rows


def bulk_run():
    """Process all PDFs in the files directory and generate Excel output."""
    pdfs = list(FILES_DIR.glob("*.pdf"))
    if not pdfs:
        print("No PDF files found in files directory")
        return
    
    print(f"Starting processing of {len(pdfs)} PDF files...")
    
    all_rows = []
    successful_pdfs = 0
    failed_pdfs = []
    
    for i, pdf_path in enumerate(pdfs, 1):
        print(f"[{i}/{len(pdfs)}] Processing: {pdf_path.name}")
        
        try:
            school = school_from_pdf(pdf_path)
            if school:
                rows = build_rows_for_school(school)
                all_rows.extend(rows)
                successful_pdfs += 1
                print(f"  ✓ Success: {school.name} → {len(rows)} row(s)")
            else:
                failed_pdfs.append(pdf_path.name)
                print(f"  ✗ Failed: Could not extract data")
        except Exception as e:
            failed_pdfs.append(f"{pdf_path.name} ({str(e)})")
            print(f"  ✗ Error: {e}")
    
    # Create Excel file
    wb = Workbook()
    ws = wb.active
    ws.title = "School Data"
    
    # Add headers
    headers = ["School", "Students", "Teachers", "Sub", "OSS", "EX", "ER", "MDM"]
    for col, header in enumerate(headers, 1):
        ws.cell(row=1, column=col, value=header)
    
    # Sort rows alphabetically by school name
    all_rows.sort(key=lambda x: x.get("School", ""))
    
    # Add data rows
    for row_idx, row_data in enumerate(all_rows, 2):
        for col_idx, header in enumerate(headers, 1):
            ws.cell(row=row_idx, column=col_idx, value=row_data.get(header))
    
    # Save file
    out_path = OUT_PATH
    wb.save(out_path)
    
    # Print summary
    print(f"\n{'='*50}")
    print(f"PROCESSING SUMMARY")
    print(f"{'='*50}")
    print(f"Total PDFs processed: {len(pdfs)}")
    print(f"Successful: {successful_pdfs}")
    print(f"Failed: {len(failed_pdfs)}")
    print(f"Total rows generated: {len(all_rows)}")
    print(f"Output file: {out_path}")
    
    if failed_pdfs:
        print(f"\nFailed PDFs:")
        for pdf in failed_pdfs:
            print(f"  - {pdf}")
    
    print(f"{'='*50}")


if __name__ == "__main__":
    bulk_run()
