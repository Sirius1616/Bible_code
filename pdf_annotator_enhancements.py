# PDFWriter_extended.py

import openpyxl
import fitz  # PyMuPDF
import logging
import sys
import os
import glob
from datetime import datetime

# Configure logging
def setup_logging():
    if not os.path.exists('logs'):
        os.makedirs('logs')
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_filename = f'logs/pdf_margin_annotator_extended_{timestamp}.log'
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler(log_filename),
            logging.StreamHandler(sys.stdout)
        ]
    )
    logging.info("=== PDF Margin Annotator Extended Started ===")
    logging.info(f"Log file: {log_filename}")
    return log_filename

# === CONFIG ===
reference_file = "margin_baseline_reference.txt"

# === COLOR DETECTION ===
def is_red(cell):
    try:
        return (
            cell.fill.fill_type == 'solid' and
            cell.fill.start_color.rgb and
            cell.fill.start_color.rgb.upper().endswith("C7CE")
        )
    except:
        return False

def is_yellow(cell):
    try:
        return (
            cell.fill.fill_type == 'solid' and
            cell.fill.start_color.rgb and
            cell.fill.start_color.rgb.upper().endswith("EB9C")
        )
    except:
        return False

def is_purple(cell):
    try:
        return (
            cell.fill.fill_type == 'solid' and
            cell.fill.start_color.rgb and
            cell.fill.start_color.rgb.upper().endswith("00800080")
        )
    except:
        return False

def is_orange(cell):
    try:
        return (
            cell.fill.fill_type == 'solid' and
            cell.fill.start_color.rgb and
            cell.fill.start_color.rgb.upper().endswith("FFA500")
        )
    except:
        return False

# === LOAD REFERENCE VALUES ===
def load_reference_values(ref_file_path):
    logging.info(f"Loading reference values from: {ref_file_path}")
    ref_values = {}
    try:
        with open(ref_file_path, 'r') as f:
            for line in f:
                line = line.strip()
                if ':' in line:
                    key, value = line.split(':', 1)
                    try:
                        ref_values[key.strip()] = float(value.strip())
                    except ValueError:
                        continue
        logging.info(f"Loaded {len(ref_values)} reference values")
        return ref_values
    except FileNotFoundError:
        logging.error(f"Reference file '{ref_file_path}' not found")
        sys.exit(1)
    except Exception as e:
        logging.error(f"Error reading reference file: {e}")
        sys.exit(1)

# === FIND FILE PAIRS ===
def find_pdf_excel_pairs():
    pdf_files = glob.glob("*.pdf")
    pairs = []
    for pdf_file in pdf_files:
        base_name = os.path.splitext(pdf_file)[0]
        excel_file = f"{base_name}.xlsx"
        if os.path.exists(excel_file):
            output_pdf = f"{base_name}_annotated.pdf"
            pairs.append({'pdf': pdf_file, 'excel': excel_file, 'output': output_pdf, 'base_name': base_name})
            logging.info(f"Found pair: {pdf_file} + {excel_file}")
        else:
            logging.warning(f"PDF found but no matching Excel file: {pdf_file}")
    return pairs

# === POSITION HELPERS ===
def is_bottom_measurement(column_name):
    bottom_measurements = {
        "Bottom Scripture Baseline Left (in)",
        "Bottom Scripture Baseline Right (in)",
        "Bottom Scripture Baseline Column 1 (in)",
        "Bottom Scripture Baseline Column 2 (in)",
        "Footnote Baseline (in)",
        "Book Intro Baseline (in)",
        "Study Note Baseline (in)",
        "Article Baseline (in)",
        "Box Baseline (in)"
    }
    return column_name in bottom_measurements

def is_side_measurement(column_name):
    side_measurements = {
        "Column 1 Left Edge (in)",
        "Column 1 Right Edge (in)",
        "Column 2 Left Edge (in)", 
        "Column 2 Right Edge (in)",
        "Column 1 Max Width (in)",
        "Column 2 Max Width (in)",
        "Column Gap Width (in)"
    }
    return column_name in side_measurements

def create_comment_text(column_name, actual_value, reference_value, color_type):
    if is_bottom_measurement(column_name):
        from_position = "from the bottom"
    elif is_side_measurement(column_name):
        from_position = "from the side"
    else:
        from_position = "from the top"
    if reference_value is not None:
        comment = f"{color_type} {column_name}. Text is {actual_value} inches {from_position}. Normally text is {reference_value}"
    else:
        comment = f"{color_type} {column_name}. Text is {actual_value} inches {from_position}. No reference available"
    return comment

# === GET REFERENCE VALUE ===
def get_reference_value(column_name, page_side, ref_values):
    column_mappings = {
        "Top Scripture Baseline Left (in)": "Top",
        "Top Scripture Baseline Right (in)": "Top",
        "Bottom Scripture Baseline Left (in)": "Bottom",
        "Bottom Scripture Baseline Right (in)": "Bottom",
        "Footnote Baseline (in)": "Footnote",
        "Book Intro Baseline (in)": "Book Intro",
        "Study Note Baseline (in)": "Study Note",
        "Article Baseline (in)": "Article",
        "Running Head Baseline (in)": "Running Head",
        "Page Number Baseline (in)": "Page Number",
        "Column 1 Max Width (in)": "Column 1 Max Width (in)",
        "Column 2 Max Width (in)": "Column 2 Max Width (in)",
        "Column Gap Width (in)": "Column Gap Width (in)",
        "Box Baseline (in)": "Box Baseline"
    }
    page_side_mappings = {
        "Column 1 Left Edge (in)": f"{page_side} Pages - Column 1 Left Edge (in)",
        "Column 1 Right Edge (in)": f"{page_side} Pages - Column 1 Right Edge (in)", 
        "Column 2 Left Edge (in)": f"{page_side} Pages - Column 2 Left Edge (in)",
        "Column 2 Right Edge (in)": f"{page_side} Pages - Column 2 Right Edge (in)"
    }
    if column_name in page_side_mappings:
        return ref_values.get(page_side_mappings[column_name])
    return ref_values.get(column_mappings.get(column_name))

# === MAIN EXTENDED PROCESSING FUNCTION ===
def process_file_pair_extended(file_pair, ref_values):
    excel_file = file_pair['excel']
    pdf_file = file_pair['pdf']
    output_pdf = file_pair['output']
    base_name = file_pair['base_name']

    logging.info(f"=== Processing {base_name} ===")
    logging.info(f"Excel: {excel_file}, PDF: {pdf_file}, Output: {output_pdf}")

    center_x_inch = ref_values.get("Bible Text Area Center Point (in)", 3.144)
    inch_to_pts = 72

    try:
        wb = openpyxl.load_workbook(excel_file)
        ws = wb.active
    except Exception as e:
        logging.error(f"Error loading Excel file: {e}")
        return False

    headers = {col: ws.cell(row=1, column=col).value for col in range(1, ws.max_column + 1)}

    paged_comments = []

    for row_num, row in enumerate(ws.iter_rows(min_row=2), start=2):
        try:
            page_num = int(row[0].value)
            page_side = row[1].value
        except (ValueError, TypeError):
            continue

        for col_idx, cell in enumerate(row[2:], start=3):
            if cell.value is None or cell.value == "N/A":
                continue
            column_name = headers.get(col_idx, f"Column {col_idx}")
            color_type = None
            if is_red(cell):
                color_type = "RED"
            elif is_yellow(cell):
                color_type = "YELLOW"
            elif is_purple(cell):
                color_type = "PURPLE"
            elif is_orange(cell):
                color_type = "ORANGE"
            else:
                continue

            reference_value = get_reference_value(column_name, page_side, ref_values)
            comment_text = create_comment_text(column_name, cell.value, reference_value, color_type)
            y_inch = cell.value
            is_bottom = is_bottom_measurement(column_name)

            paged_comments.append({
                "page_num": page_num,
                "y_inch": y_inch,
                "comment": comment_text,
                "color": color_type.lower(),
                "is_bottom": is_bottom
            })

    try:
        doc = fitz.open(pdf_file)
    except Exception as e:
        logging.error(f"Error opening PDF: {e}")
        return False

    annotations_added = 0

    for entry in paged_comments:
        if entry["page_num"] < 1 or entry["page_num"] > len(doc):
            continue
        page = doc[entry["page_num"] - 1]
        page_height = page.rect.height
        x_pts = center_x_inch * inch_to_pts
        if "Column 1" in entry["comment"]:
            x_pts = (center_x_inch - 1.5) * inch_to_pts
        elif "Column 2" in entry["comment"]:
            x_pts = (center_x_inch + 1.5) * inch_to_pts

        if entry["is_bottom"]:
            y_pts_pdf = page_height - (entry["y_inch"] * inch_to_pts)
        else:
            y_pts_pdf = entry["y_inch"] * inch_to_pts

        annot = page.add_text_annot((x_pts, y_pts_pdf), entry["comment"])
        annot.set_info(title="Margin Check")
        # Set colors
        if entry["color"] == "red":
            annot.set_colors(stroke=[1,0,0], fill=[1,0.8,0.8])
        elif entry["color"] == "yellow":
            annot.set_colors(stroke=[1,1,0], fill=[1,1,0.8])
        elif entry["color"] == "purple":
            annot.set_colors(stroke=[0.5,0,0.5], fill=[0.8,0.6,0.8])
        elif entry["color"] == "orange":
            annot.set_colors(stroke=[1,0.5,0], fill=[1,0.8,0.6])
        annot.update()
        annotations_added += 1

    try:
        doc.save(output_pdf, incremental=False, garbage=4)
        doc.close()
    except Exception as e:
        logging.error(f"Error saving PDF: {e}")
        return False

    logging.info(f"=== {base_name} Complete: {annotations_added} annotations added ===")
    return True

# === MAIN ===
def main():
    setup_logging()
    ref_values = load_reference_values(reference_file)
    file_pairs = find_pdf_excel_pairs()
    if not file_pairs:
        logging.error("No PDF/Excel file pairs found!")
        sys.exit(1)

    for pair in file_pairs:
        process_file_pair_extended(pair, ref_values)

if __name__ == "__main__":
    main()
