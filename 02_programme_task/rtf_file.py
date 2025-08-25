#!/usr/bin/env python3
"""
RTF → Excel Bible Verse Exporter

Converts batch RTF Bible book files into Excel workbooks where each row is a verse.
Column A: Full book name + reference (e.g., Ruth 1:1)
Column B: Verse text verbatim (cleaned for human readability)
"""

import os
import re
import glob
import shutil
import pandas as pd

# Configuration
CONFIG = {
    # File handling
    "INPUT_DIR": "./rtf_in/",
    "OUTPUT_DIR": "./xlsx_out/",
    "SOURCE_DIR": "./file_to_process/",

    # Filename pattern: number.bookname.rtf
    "FILENAME_PATTERN": r"^\d+\.(.+)\.rtf$",

    # Output headers
    "REFERENCE_HEADER": "Reference",
    "TEXT_HEADER": "Text"
}

def extract_book_name(filename):
    """Extract book name from filename using configured pattern"""
    match = re.match(CONFIG["FILENAME_PATTERN"], filename)
    if match:
        return match.group(1)
    return os.path.splitext(filename)[0]

def simple_rtf_to_text(rtf_content):
    """Simple RTF to text converter that removes RTF control words"""
    # Remove RTF header
    text = re.sub(r'\\rtf1.*?\\fs\d+', '', rtf_content, flags=re.DOTALL)

    # Remove control words
    text = re.sub(r'\\[a-z]+\*?', '', text)

    # Remove special characters
    text = re.sub(r'[{}]', '', text)

    # Remove extra whitespace
    text = re.sub(r'\s+', ' ', text)

    return text.strip()

def read_rtf_simple(filepath):
    """Read RTF file and convert to text using simple method"""
    with open(filepath, 'r', encoding='utf-8', errors='ignore') as file:
        rtf_content = file.read()
    return simple_rtf_to_text(rtf_content)

def process_text_content(text, book_name):
    """Process text content and extract verses"""
    verses = []
    verse_pattern = r'(\d+):(\d+)\s+(.*?)(?=\d+:\d+|$)'

    matches = re.findall(verse_pattern, text)
    for match in matches:
        chapter, verse, verse_text = match
        verses.append({
            CONFIG["REFERENCE_HEADER"]: f"{book_name} {chapter}:{verse}",
            CONFIG["TEXT_HEADER"]: verse_text.strip()
        })

    return verses

def prepare_input_dir():
    """Copy all RTF files from file_to_process into rtf_in"""
    source_dir = CONFIG["SOURCE_DIR"]
    target_dir = CONFIG["INPUT_DIR"]
    os.makedirs(target_dir, exist_ok=True)

    rtf_files = glob.glob(os.path.join(source_dir, "*.rtf"))
    if not rtf_files:
        print(f"No RTF files found in {source_dir}")
        return

    for filepath in rtf_files:
        filename = os.path.basename(filepath)
        target_path = os.path.join(target_dir, filename)
        shutil.copy(filepath, target_path)  # use shutil.move if you want to move instead
        print(f"Copied {filename} → {target_dir}")

def make_human_readable(text):
    """Clean verse text for human readability"""
    if not isinstance(text, str):
        return text

    # Remove RTF unicode escapes like \u12345?
    text = re.sub(r"\\u-?\d+\??", "", text)

    # Remove hex-like escaped sequences \'xx
    text = re.sub(r"\\'[0-9a-fA-F]{2}", "", text)

    # Remove stray asterisks, backslashes, and junk
    text = re.sub(r"[\\*]+", "", text)

    # Remove letter+number garbage like a201650, b1851106100
    text = re.sub(r"[a-zA-Z]\d{3,}", "", text)

    # Remove standalone large numbers (3+ digits)
    text = re.sub(r"\b\d{3,}\b", "", text)

    # Remove dangling 1–2 digit numbers that follow a removal
    text = re.sub(r"\s+\d{1,2}(?=\s|$)", "", text)

    # Normalize spaces around brackets
    text = re.sub(r"\s+\[|\[\s+", "[", text)
    text = re.sub(r"\s+\]", "]", text)

    # Fix multiple spaces
    text = re.sub(r"\s{2,}", " ", text)

    return text.strip()



def clean_excel_file(filepath):
    """Load Excel file, clean Column B (Text), and save again"""
    df = pd.read_excel(filepath)

    if CONFIG["TEXT_HEADER"] in df.columns:
        df[CONFIG["TEXT_HEADER"]] = df[CONFIG["TEXT_HEADER"]].apply(make_human_readable)
        df.to_excel(filepath, index=False)
        print(f"Cleaned text in {os.path.basename(filepath)}")

def main():
    """Main function to process all RTF files"""
    os.makedirs(CONFIG["OUTPUT_DIR"], exist_ok=True)
    os.makedirs(CONFIG["INPUT_DIR"], exist_ok=True)

    # Step 1: Copy files from file_to_process → rtf_in
    prepare_input_dir()

    # Step 2: Process RTF files
    rtf_files = glob.glob(os.path.join(CONFIG["INPUT_DIR"], "*.rtf"))

    if not rtf_files:
        print(f"No RTF files found in {CONFIG['INPUT_DIR']}")
        return

    for filepath in rtf_files:
        filename = os.path.basename(filepath)
        book_name = extract_book_name(filename)

        print(f"Processing {filename} ({book_name})...")

        try:
            # Read and convert RTF to text
            text = read_rtf_simple(filepath)

            # Process the text to extract verses
            verses = process_text_content(text, book_name)

            if verses:
                # Create DataFrame and save to Excel
                df = pd.DataFrame(verses)
                output_filename = f"{book_name}.xlsx"
                output_path = os.path.join(CONFIG["OUTPUT_DIR"], output_filename)
                df.to_excel(output_path, index=False)
                print(f"Saved {len(verses)} verses to {output_filename}")

                # Step 3: Clean the Excel file’s text column
                clean_excel_file(output_path)

            else:
                print(f"No verses found in {filename}")

        except Exception as e:
            print(f"Error processing {filename}: {str(e)}")

if __name__ == "__main__":
    main()
