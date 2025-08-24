import openpyxl
import fitz  # PyMuPDF
import logging
import sys
import os
import glob
from datetime import datetime

# Configure logging
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
    
    # Define the mapping from Excel column names to reference keys
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
    
    # Handle page-side dependent columns
    page_side_mappings = {
        "Column 1 Left Edge (in)": f"{page_side} Pages - Column 1 Left Edge (in)",
        "Column 1 Right Edge (in)": f"{page_side} Pages - Column 1 Right Edge (in)", 
        "Column 2 Left Edge (in)": f"{page_side} Pages - Column 2 Left Edge (in)",
        "Column 2 Right Edge (in)": f"{page_side} Pages - Column 2 Right Edge (in)"
    }
    
    # First try page-side dependent mapping
    if column_name in page_side_mappings:
        ref_key = page_side_mappings[column_name]
        logging.debug(f"Using page-side dependent mapping: '{column_name}' -> '{ref_key}'")
        ref_value = ref_values.get(ref_key)
        if ref_value is not None:
            logging.debug(f"Found reference value: {ref_value}")
            return ref_value
        else:
            logging.warning(f"Page-side dependent reference key '{ref_key}' not found in reference values")
    
    # Then try regular mapping
    if column_name in column_mappings:
        ref_key = column_mappings[column_name]
        logging.debug(f"Using regular mapping: '{column_name}' -> '{ref_key}'")
        ref_value = ref_values.get(ref_key)
        if ref_value is not None:
            logging.debug(f"Found reference value: {ref_value}")
            return ref_value
        else:
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
    """Create the comment text for the sticky note."""
    if is_bottom_measurement(column_name):
        from_position = "from the bottom"
    elif is_side_measurement(column_name):
        from_position = "from the side"
    else:
        from_position = "from the top"
    
    comment = f"{color_type} {column_name}. Text is {actual_value} inches {from_position}. Normally text is {reference_value}"
    
    logging.debug(f"Created comment: '{comment}'")
    return comment

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
            headers[col] = cell_value
            logging.debug(f"Column {col}: '{cell_value}'")
    
    paged_comments = []
    annotations_found = 0
    
    # Process each row
    for row_num, row in enumerate(ws.iter_rows(min_row=2), start=2):
        try:
            page_num = int(row[0].value)  # Column A = Page Number
            page_side = row[1].value      # Column B = Page Side (Left/Right)
        except (ValueError, TypeError):
            logging.debug(f"Row {row_num}: Invalid page number or side, skipping")
            continue
        
        logging.debug(f"Processing row {row_num}: Page {page_num} ({page_side})")
        
        # Check each column for colored cells
        for col_idx, cell in enumerate(row[2:], start=3):  # Start from column C
            if cell.value is None or cell.value == "N/A":
                continue
                
            column_name = headers.get(col_idx, f"Column {col_idx}")
            
            # Check if cell is colored
            if is_red(cell) or is_yellow(cell):
                color_type = "RED" if is_red(cell) else "YELLOW"
                actual_value = cell.value
                
                logging.info(f"Found {color_type} cell: Page {page_num}, Column '{column_name}', Value: {actual_value}")
                
                # Get reference value
                reference_value = get_reference_value(column_name, page_side, ref_values)
                
                # Create detailed comment text (with or without reference)
                if reference_value is not None:
                    comment_text = create_comment_text(column_name, actual_value, reference_value, color_type)
                else:
                    # No reference available - create annotation anyway
                    logging.warning(f"No reference value found for '{column_name}' on {page_side} pages, creating annotation without reference")
                    
                    if is_bottom_measurement(column_name):
                        from_position = "from the bottom"
                    elif is_side_measurement(column_name):
                        from_position = "from the side"
                    else:
                        from_position = "from the top"
                    
                    comment_text = f"{color_type} {column_name}. Text is {actual_value} inches {from_position}. No reference available"
                
                # Calculate Y position (handle bottom measurements)
                if is_bottom_measurement(column_name):
                    y_inch = actual_value
                    is_bottom_pos = True
                    logging.debug(f"Bottom measurement: {actual_value} inches from bottom")
                else:
                    y_inch = actual_value
                    is_bottom_pos = False
                    logging.debug(f"Top measurement: {actual_value} inches from top")
                
                paged_comments.append({
                    "page_num": page_num,
                    "y_inch": y_inch,
                    "comment": comment_text,
                    "color": color_type.lower(),
                    "is_bottom": is_bottom_pos
                })
                
                annotations_found += 1
                logging.info(f"Added annotation #{annotations_found}: {comment_text}")
    
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
    
    # Process annotations
    for entry in paged_comments:
        if entry["page_num"] < 1 or entry["page_num"] > len(doc):
            logging.warning(f"Page {entry['page_num']} out of range, skipping")
            continue
            
        try:
            page = doc[entry["page_num"] - 1]
            page_height = page.rect.height
            
            # Calculate X position based on column reference
            x_pts = center_x_inch * inch_to_pts  # Default center position
            
            # Check if comment refers to Column 1 or Column 2 and adjust X position
            if "Column 1" in entry["comment"]:
                x_pts = (center_x_inch - 1.5) * inch_to_pts
                logging.info(f"Column 1 detected: Moving comment left 1.5 inches to {x_pts/inch_to_pts:.3f} inches")
            elif "Column 2" in entry["comment"]:
                x_pts = (center_x_inch + 1.5) * inch_to_pts
                logging.info(f"Column 2 detected: Moving comment right 1.5 inches to {x_pts/inch_to_pts:.3f} inches")
            else:
                logging.info(f"No column reference detected: Using center position {x_pts/inch_to_pts:.3f} inches")
            
            if entry.get("is_bottom", False):
                inches_from_bottom = entry["y_inch"]
                y_pts_pdf = page_height - (inches_from_bottom * inch_to_pts)
                logging.info(f"BOTTOM: Page {entry['page_num']}, {inches_from_bottom} inches from bottom = {y_pts_pdf} pts from top (page height: {page_height} pts)")
            else:
                y_pts = entry["y_inch"] * inch_to_pts
                y_pts_pdf = y_pts
                logging.info(f"TOP: Page {entry['page_num']}, {entry['y_inch']} inches from top = {y_pts_pdf} pts from top")
            
            # Ensure the annotation stays within page bounds
            if y_pts_pdf < 0:
                logging.warning(f"Annotation Y coordinate {y_pts_pdf} is above page, clamping to 0")
                y_pts_pdf = 0
            elif y_pts_pdf > page_height:
                logging.warning(f"Annotation Y coordinate {y_pts_pdf} is below page, clamping to {page_height}")
                y_pts_pdf = page_height
            
            annot = page.add_text_annot((x_pts, y_pts_pdf), entry["comment"])
            annot.set_info(title="Margin Check")
            
            # Set colors based on annotation type
            if entry["color"] == "red":
                annot.set_colors(stroke=[1, 0, 0], fill=[1, 0.8, 0.8])
            else:  # yellow
                annot.set_colors(stroke=[1, 1, 0], fill=[1, 1, 0.8])
                
            annot.update()
            annotations_added += 1
            logging.info(f"Added annotation to page {entry['page_num']} at ({x_pts}, {y_pts_pdf})")
            
        except Exception as e:
            logging.error(f"Error adding annotation to page {entry['page_num']}: {e}")
            continue
    
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
    return True

def main():
    """Main function."""
    log_file = setup_logging()
    
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
    print(f"\n=== BATCH PROCESSING COMPLETE ===")
    print(f"Successfully processed: {total_processed} files")
    print(f"Failed: {total_failed} files")
    if total_processed > 0:
        print(f"Annotated PDFs saved with '_annotated' suffix")

if __name__ == "__main__":
    main()