import openpyxl
from PyPDF2 import PdfReader, PdfWriter
from PyPDF2.generic import NameObject, ArrayObject, DictionaryObject, NumberObject, FloatObject, TextStringObject, ArrayObject

# Hardcoded file names
EXCEL_FILE = "02-Exodus.xlsx"
PDF_INPUT = "02-Exodus.pdf"
PDF_OUTPUT = "02-Exodus-annotated.pdf"

# Color mapping
COLOR_MAP = {
    "RED": [1, 0, 0],
    "YELLOW": [1, 1, 0],
    "PURPLE": [0.5, 0, 0.5],
    "ORANGE": [1, 0.65, 0],
}

# Collect counts
color_counts = {color: 0 for color in COLOR_MAP.keys()}

# Load Excel workbook
wb = openpyxl.load_workbook(EXCEL_FILE)
sheet = wb.active

# Load PDF
reader = PdfReader(PDF_INPUT)
writer = PdfWriter()

# Helper: create annotation
def create_annotation(x, y, text, color):
    annotation = DictionaryObject()
    annotation.update({
        NameObject("/Type"): NameObject("/Annot"),
        NameObject("/Subtype"): NameObject("/Text"),
        NameObject("/Contents"): TextStringObject(text),
        NameObject("/Rect"): ArrayObject([FloatObject(x), FloatObject(y), FloatObject(x + 20), FloatObject(y + 20)]),
        NameObject("/C"): ArrayObject([FloatObject(c) for c in color]),
        NameObject("/T"): TextStringObject("Margin Checker"),
        NameObject("/F"): NumberObject(4),
        NameObject("/Name"): NameObject("Comment")
    })
    return annotation

# Iterate Excel cells
for row in sheet.iter_rows(min_row=2, values_only=True):
    if not row:
        continue
    page_num, x, y, issue, color_name = row[:5]
    if page_num is None or x is None or y is None or not color_name:
        continue

    color_name = str(color_name).strip().upper()
    if color_name not in COLOR_MAP:
        continue

    page = reader.pages[int(page_num) - 1]
    annotation = create_annotation(float(x), float(y), str(issue), COLOR_MAP[color_name])

    if "/Annots" not in page:
        last_page[NameObject("/Annots")] = ArrayObject()
    page["/Annots"].append(annotation)

    color_counts[color_name] += 1

# Add summary note on last page
summary_text = "Summary of Issues:\n" + "\n".join([f"{c}: {n}" for c, n in color_counts.items()])
summary_annotation = create_annotation(50, 50, summary_text, [0, 0, 1])
last_page = reader.pages[-1]
if "/Annots" not in last_page:
    last_page[NameObject("/Annots")] = ArrayObject()
last_page["/Annots"].append(summary_annotation)

# Write out PDF
for page in reader.pages:
    writer.add_page(page)
with open(PDF_OUTPUT, "wb") as f_out:
    writer.write(f_out)

# Print summary to console
print("\nAnnotation Summary:")
for c, n in color_counts.items():
    print(f"{c}: {n}")
print(f"\nAnnotated PDF saved as: {PDF_OUTPUT}")
