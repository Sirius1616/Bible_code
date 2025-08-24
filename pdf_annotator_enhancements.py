import openpyxl
import fitz  # PyMuPDF
import logging
import sys
import os
import glob
from datetime import datetime
from openpyxl.utils import get_column_letter

# ==============================================
# Logging
# ==============================================

def setup_logging():
    """Set up detailed logging for troubleshooting."""
    if not os.path.exists('logs'):
        os.makedirs('logs')

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_filename = f'logs/pdf_margin_annotator_{timestamp}.log'

    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler(log_filename),
            logging.StreamHandler(sys.stdout)
        ]
    )

    logging.info(f"=== PDF Margin Annotator Started ===")
    logging.info(f"Log file: {log_filename}")
    return log_filename

# === CONFIG ===
reference_file = "margin_baseline_reference.txt"

# ==============================================
# Helpers: reference values
# ==============================================

def load_reference_values(ref_file_path):
    """Load reference values from the margin baseline reference file."""
    logging.info(f"Loading reference values from: {ref_file_path}")
    ref_values = {}

    try:
        with open(ref_file_path, 'r') as f:
            lines = f.readlines()

        for line_num, line in enumerate(lines, 1):
            line = line.strip()
            if ':' in line:
                key, value = line.split(':', 1)
                try:
                    parsed_value = float(value.strip())
                    ref_values[key.strip()] = parsed_value
                    logging.debug(f"Line {line_num}: '{key.strip()}' = {parsed_value}")
                except ValueError:
                    logging.warning(f"Line {line_num}: Could not parse value '{value.strip()}' as float")
                    continue

        logging.info(f"Successfully loaded {len(ref_values)} reference values")
        return ref_values

    except FileNotFoundError:
        logging.error(f"Reference file '{ref_file_path}' not found")
        sys.exit(1)
    except Exception as e:
        logging.error(f"Error reading reference file: {e}")
        sys.exit(1)


def get_reference_value(column_name, page_side, ref_values):
    """Get the appropriate reference value based on column name and page side."""
    logging.debug(f"Getting reference value for column '{column_name}' on {page_side} page")

    # Map Excel column names to reference keys in the txt file
    column_mappings = {
        # Top/Bottom baselines (accept either Left/Right or Column 1/Column 2 naming)
        "Top Scripture Baseline Left (in)": "Top",
        "Top Scripture Baseline Right (in)": "Top",
        "Top Scripture Baseline Column 1 (in)": "Top",
        "Top Scripture Baseline Column 2 (in)": "Top",

        "Bottom Scripture Baseline Left (in)": "Bottom",
        "Bottom Scripture Baseline Right (in)": "Bottom",
        "Bottom Scripture Baseline Column 1 (in)": "Bottom",
        "Bottom Scripture Baseline Column 2 (in)": "Bottom",

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
        ref_key = page_side_mappings[column_name]
        ref_value = ref_values.get(ref_key)
        if ref_value is not None:
            return ref_value
        logging.warning(f"Page-side dependent reference key '{ref_key}' not found in reference values")

    if column_name in column_mappings:
        ref_key = column_mappings[column_name]
        ref_value = ref_values.get(ref_key)
        if ref_value is not None:
            return ref_value
        logging.warning(f"Regular reference key '{ref_key}' not found in reference values")

    logging.warning(f"No mapping found for column '{column_name}'")
    return None


def is_bottom_measurement(column_name):
    """Check if a measurement is taken from the bottom of the page."""
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
    """Check if a measurement is taken from the side of the page."""
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
    """Create the comment text for the sticky note for RED/YELLOW."""
    if is_bottom_measurement(column_name):
        from_position = "from the bottom"
    elif is_side_measurement(column_name):
        from_position = "from the side"
    else:
        from_position = "from the top"

    comment = f"{color_type} {column_name}. Text is {actual_value} inches {from_position}. Normally text is {reference_value}"
    return comment

# ==============================================
# Color detection & normalization
# ==============================================

# Known color hex (last 6 digits) → canonical label
COLOR_HEX_MAP = {
    # Reds (common Excel fills: pure red; light red fill FFC7CE → endswith C7CE)
    "FF0000": "RED",
    # Yellows (pure yellow; light yellow fill FFEB9C → endswith EB9C)
    "FFFF00": "YELLOW",
    # Purple
    "800080": "PURPLE",
    # Oranges (pure + common Office orange variants)
    "FFA500": "ORANGE",
    "ED7D31": "ORANGE",
    "F4B083": "ORANGE",
}

PURPLE_TEXT = "The two columns on this page do not align and they are not at their typical location."
ORANGE_TEXT = "this column is not algined with the other column. The other column is in the correct position"


def get_cell_rgb_hex(cell):
    """Return the raw RGB hex code from an openpyxl cell fill, if present (e.g., '00FF0000')."""
    try:
        fill = cell.fill
        if not fill or fill.fill_type != 'solid':
            return None
        color = fill.start_color
        if not color:
            return None
        # Prefer .rgb when available
        if getattr(color, 'rgb', None):
            return color.rgb.upper()
        # Some files may use indexed/theme; we don't resolve those here
        if getattr(color, 'indexed', None) is not None:
            return f"INDEXED_{color.indexed}"
        if getattr(color, 'theme', None) is not None:
            return f"THEME_{color.theme}"
        return None
    except Exception:
        return None


def normalize_hex(raw_hex):
    """Normalize raw hex to the last 6 digits (strip leading alpha if present)."""
    if not raw_hex:
        return None
    # Only keep hex chars
    s = raw_hex.replace('#', '').upper()
    if s.startswith('INDEXED_') or s.startswith('THEME_'):
        return s  # return as-is so we can log it distinctly
    return s[-6:] if len(s) >= 6 else s


def detect_cell_color(cell, excel_row_idx, excel_col_idx, unknown_log_path=None):
    """Return 'RED' | 'YELLOW' | 'PURPLE' | 'ORANGE' | None; log unknowns."""
    raw = get_cell_rgb_hex(cell)
    norm = normalize_hex(raw)

    # Accept legacy partial matches used in earlier sheets
    if raw:
        up = raw.upper()
        if up.endswith('C7CE'):
            color = 'RED'
        elif up.endswith('EB9C'):
            color = 'YELLOW'
        else:
            color = COLOR_HEX_MAP.get(norm)
    else:
        color = None

    if not color and norm:
        addr = f"{get_column_letter(excel_col_idx)}{excel_row_idx}"
        msg = f"Unknown color '{raw}' (norm='{norm}') at {addr}"
        logging.warning(msg)
        if unknown_log_path:
            with open(unknown_log_path, 'a', encoding='utf-8') as f:
               f.write(msg + "\n")


    return color

# ==============================================
# File discovery
# ==============================================

def find_pdf_excel_pairs():
    """Find all PDF files and their corresponding Excel files in the current directory."""
    pdf_files = glob.glob("*.pdf")
    pairs = []

    for pdf_file in pdf_files:
        base_name = os.path.splitext(pdf_file)[0]
        excel_file = f"{base_name}.xlsx"

        if os.path.exists(excel_file):
            output_pdf = f"{base_name}_annotated.pdf"
            pairs.append({
                'pdf': pdf_file,
                'excel': excel_file,
                'output': output_pdf,
                'base_name': base_name
            })
            logging.info(f"Found pair: {pdf_file} + {excel_file} -> {output_pdf}")
        else:
            logging.warning(f"PDF found but no matching Excel file: {pdf_file} (looking for {excel_file})")

    return pairs

# ==============================================
# Core processing
# ==============================================

def get_x_tag_from_header(header: str) -> str:
    """Return 'col1', 'col2', or 'center' based on the header text."""
    h = (header or "").lower()
    if 'column 1' in h:
        return 'col1'
    if 'column 2' in h:
        return 'col2'
    return 'center'


def process_file_pair(file_pair, ref_values):
    """Process a single PDF/Excel file pair."""
    excel_file = file_pair['excel']
    pdf_file = file_pair['pdf']
    output_pdf = file_pair['output']
    base_name = file_pair['base_name']

    logging.info(f"=== Processing {base_name} ===")
    logging.info(f"Excel: {excel_file}")
    logging.info(f"PDF: {pdf_file}")
    logging.info(f"Output: {output_pdf}")

    center_x_inch = ref_values.get("Bible Text Area Center Point (in)", 3.144)
    page_height_inches = ref_values.get("Page Height (in)", 9.25)
    logging.info(f"Using center point: {center_x_inch} inches")
    logging.info(f"Using page height: {page_height_inches} inches")

    inch_to_pts = 72

    # === LOAD EXCEL ===
    logging.info(f"Loading Excel file: {excel_file}")
    try:
        wb = openpyxl.load_workbook(excel_file)
        ws = wb.active
        logging.info(f"Excel file loaded successfully")
    except Exception as e:
        logging.error(f"Error loading Excel file {excel_file}: {e}")
        return False

    # Get column headers
    headers = {}
    for col in range(1, ws.max_column + 1):
        cell_value = ws.cell(row=1, column=col).value
        if cell_value:
            headers[col] = str(cell_value)

    paged_comments = []
    annotations_found = 0

    # For comparison check
    expected_counts = {c: 0 for c in ['RED', 'YELLOW', 'PURPLE', 'ORANGE']}

    unknown_colors_log = os.path.join('logs', f"unknown_colors_{base_name}.log")

    # Process each row
    for row_idx, row in enumerate(ws.iter_rows(min_row=2), start=2):
        try:
            page_num = int(row[0].value)  # Column A = Page Number
            page_side = row[1].value      # Column B = Page Side (Left/Right)
        except (ValueError, TypeError):
            logging.debug(f"Row {row_idx}: Invalid page number or side, skipping")
            continue

        for col_idx, cell in enumerate(row[2:], start=3):  # Start from column C
            if cell.value is None or cell.value == "N/A":
                continue

            column_name = headers.get(col_idx, f"Column {col_idx}")

            # Detect color
            color_label = detect_cell_color(cell, row_idx, col_idx, unknown_log_path=unknown_colors_log)
            if not color_label:
                continue

            # Coerce numeric value
            try:
                actual_value = float(cell.value)
            except Exception:
                logging.warning(f"Row {row_idx} col {col_idx} ('{column_name}') has non-numeric value '{cell.value}', skipping")
                continue

            expected_counts[color_label] += 1

            # Build comment text
            if color_label in ("PURPLE", "ORANGE"):
                comment_text = PURPLE_TEXT if color_label == "PURPLE" else ORANGE_TEXT
            else:
                # RED / YELLOW retain detailed message
                reference_value = get_reference_value(column_name, page_side, ref_values)
                if reference_value is not None:
                    comment_text = create_comment_text(column_name, actual_value, reference_value, color_label)
                else:
                    # No reference available - create annotation anyway
                    if is_bottom_measurement(column_name):
                        from_position = "from the bottom"
                    elif is_side_measurement(column_name):
                        from_position = "from the side"
                    else:
                        from_position = "from the top"
                    comment_text = f"{color_label} {column_name}. Text is {actual_value} inches {from_position}. No reference available"

            # Determine vertical placement (top vs bottom rule)
            if is_bottom_measurement(column_name):
                y_inch = actual_value
                is_bottom_pos = True
            else:
                y_inch = actual_value
                is_bottom_pos = False

            # Determine horizontal placement tag (center/col1/col2)
            x_tag = get_x_tag_from_header(column_name)

            paged_comments.append({
                "page_num": page_num,
                "y_inch": y_inch,
                "comment": comment_text,
                "color": color_label.lower(),
                "is_bottom": is_bottom_pos,
                "x_tag": x_tag
            })

            annotations_found += 1
            logging.info(f"[{color_label}] Page {page_num}, '{column_name}', value={actual_value}, x_tag={x_tag}, bottom={is_bottom_pos}")

    logging.info(f"Total annotations prepared for {base_name}: {annotations_found}")

    # === OPEN PDF AND ADD ANNOTATIONS ===
    logging.info(f"Opening PDF file: {pdf_file}")
    try:
        doc = fitz.open(pdf_file)
        logging.info(f"PDF opened successfully. Pages: {len(doc)}")
    except Exception as e:
        logging.error(f"Error opening PDF {pdf_file}: {e}")
        return False

    annotations_added = 0
    added_counts = {c: 0 for c in ['RED', 'YELLOW', 'PURPLE', 'ORANGE']}

    # Process annotations
    for entry in paged_comments:
        if entry["page_num"] < 1 or entry["page_num"] > len(doc):
            logging.warning(f"Page {entry['page_num']} out of range, skipping")
            continue

        try:
            page = doc[entry["page_num"] - 1]
            page_height = page.rect.height

            # Calculate X position based on x_tag
            x_pts = center_x_inch * inch_to_pts  # default center
            if entry.get("x_tag") == 'col1':
                x_pts = (center_x_inch - 1.5) * inch_to_pts
            elif entry.get("x_tag") == 'col2':
                x_pts = (center_x_inch + 1.5) * inch_to_pts

            # Calculate Y position (handle bottom measurements)
            if entry.get("is_bottom", False):
                inches_from_bottom = entry["y_inch"]
                y_pts_pdf = page_height - (inches_from_bottom * inch_to_pts)
            else:
                y_pts_pdf = entry["y_inch"] * inch_to_pts

            # Clamp to page
            if y_pts_pdf < 0:
                y_pts_pdf = 0
            elif y_pts_pdf > page_height:
                y_pts_pdf = page_height

            annot = page.add_text_annot((x_pts, y_pts_pdf), entry["comment"])
            annot.set_info(title="Margin Check")

            # Set colors
            c = entry["color"]
            if c == "red":
                annot.set_colors(stroke=[1, 0, 0], fill=[1, 0.85, 0.85])
                added_counts['RED'] += 1
            elif c == "yellow":
                annot.set_colors(stroke=[1, 1, 0], fill=[1, 1, 0.85])
                added_counts['YELLOW'] += 1
            elif c == "purple":
                annot.set_colors(stroke=[0.5, 0, 0.5], fill=[0.93, 0.85, 0.96])
                added_counts['PURPLE'] += 1
            elif c == "orange":
                annot.set_colors(stroke=[1, 0.65, 0], fill=[1, 0.92, 0.85])
                added_counts['ORANGE'] += 1
            else:
                annot.set_colors(stroke=[0, 0, 0], fill=[0.9, 0.9, 0.9])

            annot.update()
            annotations_added += 1
            logging.info(f"Added {c} annotation to page {entry['page_num']} at ({x_pts/72:.2f} in, {y_pts_pdf/72:.2f} in)")

        except Exception as e:
            logging.error(f"Error adding annotation to page {entry['page_num']}: {e}")
            continue

    # === Summary note on first page ===
    try:
        page0 = doc[0]
        summary_msgs = []
        for color in ['RED', 'YELLOW', 'PURPLE', 'ORANGE']:
            exp_c = expected_counts[color]
            add_c = added_counts[color]
            if exp_c != add_c:
                missing = exp_c - add_c
                summary_msgs.append(f"{exp_c} {color} annotations expected, {add_c} written. {missing} missing.")

        if summary_msgs:
            # One note per color mismatch (top of page)
            y = 36.0
            for msg in summary_msgs:
                page0.add_text_annot((72, y), msg)
                y += 24.0
        else:
            ok_msg = ("All annotations successfully written into the PDF. "
                      f"RED: {added_counts['RED']}, YELLOW: {added_counts['YELLOW']}, "
                      f"PURPLE: {added_counts['PURPLE']}, ORANGE: {added_counts['ORANGE']}")
            page0.add_text_annot((72, 36), ok_msg)
    except Exception as e:
        logging.error(f"Failed to add summary note on first page: {e}")

    # Save the PDF
    logging.info(f"Saving annotated PDF to: {output_pdf}")
    try:
        doc.save(output_pdf, incremental=False, garbage=4)
        doc.close()
        logging.info(f"PDF saved successfully")
    except Exception as e:
        logging.error(f"Error saving PDF {output_pdf}: {e}")
        return False

    logging.info(f"=== {base_name} Complete: {annotations_added} annotations added ===")
    logging.info(f"Expected counts: {expected_counts}")
    logging.info(f"Added counts: {added_counts}")
    return True

# ==============================================
# Main
# ==============================================

def main():
    """Main function."""
    _ = setup_logging()

    # Load reference values (still needed for all files)
    ref_values = load_reference_values(reference_file)

    # Find all PDF/Excel pairs in the current directory
    file_pairs = find_pdf_excel_pairs()

    if not file_pairs:
        logging.error("No PDF/Excel file pairs found in current directory!")
        print("No PDF/Excel file pairs found!")
        print("Make sure you have matching PDF and XLSX files (same name, different extensions)")
        sys.exit(1)

    logging.info(f"Found {len(file_pairs)} file pair(s) to process")

    # Process each file pair
    total_processed = 0
    total_failed = 0

    for file_pair in file_pairs:
        try:
            success = process_file_pair(file_pair, ref_values)
            if success:
                total_processed += 1
                print(f"✓ Successfully processed: {file_pair['base_name']}")
            else:
                total_failed += 1
                print(f"✗ Failed to process: {file_pair['base_name']}")
        except Exception as e:
            logging.error(f"Unexpected error processing {file_pair['base_name']}: {e}")
            total_failed += 1
            print(f"✗ Failed to process: {file_pair['base_name']} - {e}")

    # Final summary
    logging.info(f"=== BATCH PROCESSING COMPLETE ===")
    logging.info(f"Successfully processed: {total_processed}")
    logging.info(f"Failed: {total_failed}")
    print(f"=== BATCH PROCESSING COMPLETE ===")
    print(f"Successfully processed: {total_processed} files")
    print(f"Failed: {total_failed} files")
    if total_processed > 0:
        print(f"Annotated PDFs saved with '_annotated' suffix")


if __name__ == "__main__":
    main()
import openpyxl
import fitz  # PyMuPDF
import logging
import sys
import os
import glob
from datetime import datetime
from openpyxl.utils import get_column_letter

# ==============================================
# Logging
# ==============================================

def setup_logging():
    """Set up detailed logging for troubleshooting."""
    if not os.path.exists('logs'):
        os.makedirs('logs')

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_filename = f'logs/pdf_margin_annotator_{timestamp}.log'

    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler(log_filename),
            logging.StreamHandler(sys.stdout)
        ]
    )

    logging.info(f"=== PDF Margin Annotator Started ===")
    logging.info(f"Log file: {log_filename}")
    return log_filename

# === CONFIG ===
reference_file = "margin_baseline_reference.txt"

# ==============================================
# Helpers: reference values
# ==============================================

def load_reference_values(ref_file_path):
    """Load reference values from the margin baseline reference file."""
    logging.info(f"Loading reference values from: {ref_file_path}")
    ref_values = {}

    try:
        with open(ref_file_path, 'r') as f:
            lines = f.readlines()

        for line_num, line in enumerate(lines, 1):
            line = line.strip()
            if ':' in line:
                key, value = line.split(':', 1)
                try:
                    parsed_value = float(value.strip())
                    ref_values[key.strip()] = parsed_value
                    logging.debug(f"Line {line_num}: '{key.strip()}' = {parsed_value}")
                except ValueError:
                    logging.warning(f"Line {line_num}: Could not parse value '{value.strip()}' as float")
                    continue

        logging.info(f"Successfully loaded {len(ref_values)} reference values")
        return ref_values

    except FileNotFoundError:
        logging.error(f"Reference file '{ref_file_path}' not found")
        sys.exit(1)
    except Exception as e:
        logging.error(f"Error reading reference file: {e}")
        sys.exit(1)


def get_reference_value(column_name, page_side, ref_values):
    """Get the appropriate reference value based on column name and page side."""
    logging.debug(f"Getting reference value for column '{column_name}' on {page_side} page")

    # Map Excel column names to reference keys in the txt file
    column_mappings = {
        # Top/Bottom baselines (accept either Left/Right or Column 1/Column 2 naming)
        "Top Scripture Baseline Left (in)": "Top",
        "Top Scripture Baseline Right (in)": "Top",
        "Top Scripture Baseline Column 1 (in)": "Top",
        "Top Scripture Baseline Column 2 (in)": "Top",

        "Bottom Scripture Baseline Left (in)": "Bottom",
        "Bottom Scripture Baseline Right (in)": "Bottom",
        "Bottom Scripture Baseline Column 1 (in)": "Bottom",
        "Bottom Scripture Baseline Column 2 (in)": "Bottom",

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
        ref_key = page_side_mappings[column_name]
        ref_value = ref_values.get(ref_key)
        if ref_value is not None:
            return ref_value
        logging.warning(f"Page-side dependent reference key '{ref_key}' not found in reference values")

    if column_name in column_mappings:
        ref_key = column_mappings[column_name]
        ref_value = ref_values.get(ref_key)
        if ref_value is not None:
            return ref_value
        logging.warning(f"Regular reference key '{ref_key}' not found in reference values")

    logging.warning(f"No mapping found for column '{column_name}'")
    return None


def is_bottom_measurement(column_name):
    """Check if a measurement is taken from the bottom of the page."""
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
    """Check if a measurement is taken from the side of the page."""
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
    """Create the comment text for the sticky note for RED/YELLOW."""
    if is_bottom_measurement(column_name):
        from_position = "from the bottom"
    elif is_side_measurement(column_name):
        from_position = "from the side"
    else:
        from_position = "from the top"

    comment = f"{color_type} {column_name}. Text is {actual_value} inches {from_position}. Normally text is {reference_value}"
    return comment

# ==============================================
# Color detection & normalization
# ==============================================

# Known color hex (last 6 digits) → canonical label
COLOR_HEX_MAP = {
    # Reds (common Excel fills: pure red; light red fill FFC7CE → endswith C7CE)
    "FF0000": "RED",
    # Yellows (pure yellow; light yellow fill FFEB9C → endswith EB9C)
    "FFFF00": "YELLOW",
    # Purple
    "800080": "PURPLE",
    # Oranges (pure + common Office orange variants)
    "FFA500": "ORANGE",
    "ED7D31": "ORANGE",
    "F4B083": "ORANGE",
}

PURPLE_TEXT = "The two columns on this page do not align and they are not at their typical location."
ORANGE_TEXT = "this column is not algined with the other column. The other column is in the correct position"


def get_cell_rgb_hex(cell):
    """Return the raw RGB hex code from an openpyxl cell fill, if present (e.g., '00FF0000')."""
    try:
        fill = cell.fill
        if not fill or fill.fill_type != 'solid':
            return None
        color = fill.start_color
        if not color:
            return None
        # Prefer .rgb when available
        if getattr(color, 'rgb', None):
            return color.rgb.upper()
        # Some files may use indexed/theme; we don't resolve those here
        if getattr(color, 'indexed', None) is not None:
            return f"INDEXED_{color.indexed}"
        if getattr(color, 'theme', None) is not None:
            return f"THEME_{color.theme}"
        return None
    except Exception:
        return None


def normalize_hex(raw_hex):
    """Normalize raw hex to the last 6 digits (strip leading alpha if present)."""
    if not raw_hex:
        return None
    # Only keep hex chars
    s = raw_hex.replace('#', '').upper()
    if s.startswith('INDEXED_') or s.startswith('THEME_'):
        return s  # return as-is so we can log it distinctly
    return s[-6:] if len(s) >= 6 else s


def detect_cell_color(cell, excel_row_idx, excel_col_idx, unknown_log_path=None):
    """Return 'RED' | 'YELLOW' | 'PURPLE' | 'ORANGE' | None; log unknowns."""
    raw = get_cell_rgb_hex(cell)
    norm = normalize_hex(raw)

    # Accept legacy partial matches used in earlier sheets
    if raw:
        up = raw.upper()
        if up.endswith('C7CE'):
            color = 'RED'
        elif up.endswith('EB9C'):
            color = 'YELLOW'
        else:
            color = COLOR_HEX_MAP.get(norm)
    else:
        color = None

    if not color and norm:
        addr = f"{get_column_letter(excel_col_idx)}{excel_row_idx}"
        msg = f"Unknown color '{raw}' (norm='{norm}') at {addr}"
        logging.warning(msg)
        if unknown_log_path:
            with open(unknown_log_path, 'a', encoding='utf-8') as f:
                f.write(msg + "")

    return color

# ==============================================
# File discovery
# ==============================================

def find_pdf_excel_pairs():
    """Find all PDF files and their corresponding Excel files in the current directory."""
    pdf_files = glob.glob("*.pdf")
    pairs = []

    for pdf_file in pdf_files:
        base_name = os.path.splitext(pdf_file)[0]
        excel_file = f"{base_name}.xlsx"

        if os.path.exists(excel_file):
            output_pdf = f"{base_name}_annotated.pdf"
            pairs.append({
                'pdf': pdf_file,
                'excel': excel_file,
                'output': output_pdf,
                'base_name': base_name
            })
            logging.info(f"Found pair: {pdf_file} + {excel_file} -> {output_pdf}")
        else:
            logging.warning(f"PDF found but no matching Excel file: {pdf_file} (looking for {excel_file})")

    return pairs

# ==============================================
# Core processing
# ==============================================

def get_x_tag_from_header(header: str) -> str:
    """Return 'col1', 'col2', or 'center' based on the header text."""
    h = (header or "").lower()
    if 'column 1' in h:
        return 'col1'
    if 'column 2' in h:
        return 'col2'
    return 'center'


def process_file_pair(file_pair, ref_values):
    """Process a single PDF/Excel file pair."""
    excel_file = file_pair['excel']
    pdf_file = file_pair['pdf']
    output_pdf = file_pair['output']
    base_name = file_pair['base_name']

    logging.info(f"=== Processing {base_name} ===")
    logging.info(f"Excel: {excel_file}")
    logging.info(f"PDF: {pdf_file}")
    logging.info(f"Output: {output_pdf}")

    center_x_inch = ref_values.get("Bible Text Area Center Point (in)", 3.144)
    page_height_inches = ref_values.get("Page Height (in)", 9.25)
    logging.info(f"Using center point: {center_x_inch} inches")
    logging.info(f"Using page height: {page_height_inches} inches")

    inch_to_pts = 72

    # === LOAD EXCEL ===
    logging.info(f"Loading Excel file: {excel_file}")
    try:
        wb = openpyxl.load_workbook(excel_file)
        ws = wb.active
        logging.info(f"Excel file loaded successfully")
    except Exception as e:
        logging.error(f"Error loading Excel file {excel_file}: {e}")
        return False

    # Get column headers
    headers = {}
    for col in range(1, ws.max_column + 1):
        cell_value = ws.cell(row=1, column=col).value
        if cell_value:
            headers[col] = str(cell_value)

    paged_comments = []
    annotations_found = 0

    # For comparison check
    expected_counts = {c: 0 for c in ['RED', 'YELLOW', 'PURPLE', 'ORANGE']}

    unknown_colors_log = os.path.join('logs', f"unknown_colors_{base_name}.log")

    # Process each row
    for row_idx, row in enumerate(ws.iter_rows(min_row=2), start=2):
        try:
            page_num = int(row[0].value)  # Column A = Page Number
            page_side = row[1].value      # Column B = Page Side (Left/Right)
        except (ValueError, TypeError):
            logging.debug(f"Row {row_idx}: Invalid page number or side, skipping")
            continue

        for col_idx, cell in enumerate(row[2:], start=3):  # Start from column C
            if cell.value is None or cell.value == "N/A":
                continue

            column_name = headers.get(col_idx, f"Column {col_idx}")

            # Detect color
            color_label = detect_cell_color(cell, row_idx, col_idx, unknown_log_path=unknown_colors_log)
            if not color_label:
                continue

            # Coerce numeric value
            try:
                actual_value = float(cell.value)
            except Exception:
                logging.warning(f"Row {row_idx} col {col_idx} ('{column_name}') has non-numeric value '{cell.value}', skipping")
                continue

            expected_counts[color_label] += 1

            # Build comment text
            if color_label in ("PURPLE", "ORANGE"):
                comment_text = PURPLE_TEXT if color_label == "PURPLE" else ORANGE_TEXT
            else:
                # RED / YELLOW retain detailed message
                reference_value = get_reference_value(column_name, page_side, ref_values)
                if reference_value is not None:
                    comment_text = create_comment_text(column_name, actual_value, reference_value, color_label)
                else:
                    # No reference available - create annotation anyway
                    if is_bottom_measurement(column_name):
                        from_position = "from the bottom"
                    elif is_side_measurement(column_name):
                        from_position = "from the side"
                    else:
                        from_position = "from the top"
                    comment_text = f"{color_label} {column_name}. Text is {actual_value} inches {from_position}. No reference available"

            # Determine vertical placement (top vs bottom rule)
            if is_bottom_measurement(column_name):
                y_inch = actual_value
                is_bottom_pos = True
            else:
                y_inch = actual_value
                is_bottom_pos = False

            # Determine horizontal placement tag (center/col1/col2)
            x_tag = get_x_tag_from_header(column_name)

            paged_comments.append({
                "page_num": page_num,
                "y_inch": y_inch,
                "comment": comment_text,
                "color": color_label.lower(),
                "is_bottom": is_bottom_pos,
                "x_tag": x_tag
            })

            annotations_found += 1
            logging.info(f"[{color_label}] Page {page_num}, '{column_name}', value={actual_value}, x_tag={x_tag}, bottom={is_bottom_pos}")

    logging.info(f"Total annotations prepared for {base_name}: {annotations_found}")

    # === OPEN PDF AND ADD ANNOTATIONS ===
    logging.info(f"Opening PDF file: {pdf_file}")
    try:
        doc = fitz.open(pdf_file)
        logging.info(f"PDF opened successfully. Pages: {len(doc)}")
    except Exception as e:
        logging.error(f"Error opening PDF {pdf_file}: {e}")
        return False

    annotations_added = 0
    added_counts = {c: 0 for c in ['RED', 'YELLOW', 'PURPLE', 'ORANGE']}

    # Process annotations
    for entry in paged_comments:
        if entry["page_num"] < 1 or entry["page_num"] > len(doc):
            logging.warning(f"Page {entry['page_num']} out of range, skipping")
            continue

        try:
            page = doc[entry["page_num"] - 1]
            page_height = page.rect.height

            # Calculate X position based on x_tag
            x_pts = center_x_inch * inch_to_pts  # default center
            if entry.get("x_tag") == 'col1':
                x_pts = (center_x_inch - 1.5) * inch_to_pts
            elif entry.get("x_tag") == 'col2':
                x_pts = (center_x_inch + 1.5) * inch_to_pts

            # Calculate Y position (handle bottom measurements)
            if entry.get("is_bottom", False):
                inches_from_bottom = entry["y_inch"]
                y_pts_pdf = page_height - (inches_from_bottom * inch_to_pts)
            else:
                y_pts_pdf = entry["y_inch"] * inch_to_pts

            # Clamp to page
            if y_pts_pdf < 0:
                y_pts_pdf = 0
            elif y_pts_pdf > page_height:
                y_pts_pdf = page_height

            annot = page.add_text_annot((x_pts, y_pts_pdf), entry["comment"])
            annot.set_info(title="Margin Check")

            # Set colors
            c = entry["color"]
            if c == "red":
                annot.set_colors(stroke=[1, 0, 0], fill=[1, 0.85, 0.85])
                added_counts['RED'] += 1
            elif c == "yellow":
                annot.set_colors(stroke=[1, 1, 0], fill=[1, 1, 0.85])
                added_counts['YELLOW'] += 1
            elif c == "purple":
                annot.set_colors(stroke=[0.5, 0, 0.5], fill=[0.93, 0.85, 0.96])
                added_counts['PURPLE'] += 1
            elif c == "orange":
                annot.set_colors(stroke=[1, 0.65, 0], fill=[1, 0.92, 0.85])
                added_counts['ORANGE'] += 1
            else:
                annot.set_colors(stroke=[0, 0, 0], fill=[0.9, 0.9, 0.9])

            annot.update()
            annotations_added += 1
            logging.info(f"Added {c} annotation to page {entry['page_num']} at ({x_pts/72:.2f} in, {y_pts_pdf/72:.2f} in)")

        except Exception as e:
            logging.error(f"Error adding annotation to page {entry['page_num']}: {e}")
            continue

    # === Summary note on first page ===
    try:
        page0 = doc[0]
        summary_msgs = []
        for color in ['RED', 'YELLOW', 'PURPLE', 'ORANGE']:
            exp_c = expected_counts[color]
            add_c = added_counts[color]
            if exp_c != add_c:
                missing = exp_c - add_c
                summary_msgs.append(f"{exp_c} {color} annotations expected, {add_c} written. {missing} missing.")

        if summary_msgs:
            # One note per color mismatch (top of page)
            y = 36.0
            for msg in summary_msgs:
                page0.add_text_annot((72, y), msg)
                y += 24.0
        else:
            ok_msg = ("All annotations successfully written into the PDF. "
                      f"RED: {added_counts['RED']}, YELLOW: {added_counts['YELLOW']}, "
                      f"PURPLE: {added_counts['PURPLE']}, ORANGE: {added_counts['ORANGE']}")
            page0.add_text_annot((72, 36), ok_msg)
    except Exception as e:
        logging.error(f"Failed to add summary note on first page: {e}")

    # Save the PDF
    logging.info(f"Saving annotated PDF to: {output_pdf}")
    try:
        doc.save(output_pdf, incremental=False, garbage=4)
        doc.close()
        logging.info(f"PDF saved successfully")
    except Exception as e:
        logging.error(f"Error saving PDF {output_pdf}: {e}")
        return False

    logging.info(f"=== {base_name} Complete: {annotations_added} annotations added ===")
    logging.info(f"Expected counts: {expected_counts}")
    logging.info(f"Added counts: {added_counts}")
    return True

# ==============================================
# Main
# ==============================================

def main():
    """Main function."""
    _ = setup_logging()

    # Load reference values (still needed for all files)
    ref_values = load_reference_values(reference_file)

    # Find all PDF/Excel pairs in the current directory
    file_pairs = find_pdf_excel_pairs()

    if not file_pairs:
        logging.error("No PDF/Excel file pairs found in current directory!")
        print("No PDF/Excel file pairs found!")
        print("Make sure you have matching PDF and XLSX files (same name, different extensions)")
        sys.exit(1)

    logging.info(f"Found {len(file_pairs)} file pair(s) to process")

    # Process each file pair
    total_processed = 0
    total_failed = 0

    for file_pair in file_pairs:
        try:
            success = process_file_pair(file_pair, ref_values)
            if success:
                total_processed += 1
                print(f"✓ Successfully processed: {file_pair['base_name']}")
            else:
                total_failed += 1
                print(f"✗ Failed to process: {file_pair['base_name']}")
        except Exception as e:
            logging.error(f"Unexpected error processing {file_pair['base_name']}: {e}")
            total_failed += 1
            print(f"✗ Failed to process: {file_pair['base_name']} - {e}")

    # Final summary
    logging.info(f"=== BATCH PROCESSING COMPLETE ===")
    logging.info(f"Successfully processed: {total_processed}")
    logging.info(f"Failed: {total_failed}")
    print(f"=== BATCH PROCESSING COMPLETE ===")
    print(f"Successfully processed: {total_processed} files")
    print(f"Failed: {total_failed} files")
    if total_processed > 0:
        print(f"Annotated PDFs saved with '_annotated' suffix")


if __name__ == "__main__":
    main()
