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
    "INPUT_DIR": "./rtf_in/",
    "OUTPUT_DIR": "./xlsx_out/",
    "SOURCE_DIR": "./file_to_process/",
    "FILENAME_PATTERN": r"^\d+\.(.+)\.rtf$",
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
    text = re.sub(r'\\rtf1.*?\\fs\d+', '', rtf_content, flags=re.DOTALL)
    text = re.sub(r'\\[a-zA-Z]+\*?', '', text)
    text = re.sub(r'[{}]', '', text)
    text = re.sub(r'\s+', ' ', text)
    return text.strip()

def read_rtf_simple(filepath):
    """Read RTF file and convert to text using simple method"""
    with open(filepath, 'r', encoding='utf-8', errors='ignore') as file:
        rtf_content = file.read()
    return simple_rtf_to_text(rtf_content)

def remove_rtf_artifacts(text: str) -> str:
    """Remove strange RTF numeric/escape artifacts from verse text"""
    text = re.sub(r"\\\*\d+\s*[a-z]?", " ", text)
    text = re.sub(r"\\'[0-9a-fA-F]{2,}", " ", text)
    text = re.sub(r"\b\d{2,}\b", " ", text)
    text = re.sub(r"[?]", "", text)
    text = re.sub(r'\s+', ' ', text)
    return text.strip()

def make_human_readable(text: str) -> str:
    """Clean RTF artifacts and make text human-readable"""
    text = remove_rtf_artifacts(text)
    text = re.sub(r"\bLord\s*16\b", "Lord", text)
    text = re.sub(r"\b(\w+)\s+s\b", r"\1's", text)
    text = re.sub(r"\bIam\b", "I am", text)
    text = text.replace("Huram- ?abi", "Huram-abi")
    text = re.sub(r'\s+', ' ', text)
    return text.strip()

def process_text_content(text, book_name):
    """Process text content and extract verses"""
    verses = []
    verse_pattern = r'(\d+):(\d+)\s+(.*?)(?=\d+:\d+|$)'
    matches = re.findall(verse_pattern, text)

    for match in matches:
        chapter, verse, verse_text = match

        # Remove any trailing '16' artifact from verse numbers
        verse = re.sub(r'16$', '', verse)

        # Clean verse text
        verse_text = make_human_readable(verse_text)

        if verse_text:
            verses.append({
                CONFIG["REFERENCE_HEADER"]: f"{book_name} {chapter}:{verse}",
                CONFIG["TEXT_HEADER"]: verse_text
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
        shutil.copy(filepath, target_path)
        print(f"Copied {filename} → {target_dir}")

def main():
    """Main function to process all RTF files"""
    os.makedirs(CONFIG["OUTPUT_DIR"], exist_ok=True)
    os.makedirs(CONFIG["INPUT_DIR"], exist_ok=True)

    prepare_input_dir()

    rtf_files = glob.glob(os.path.join(CONFIG["INPUT_DIR"], "*.rtf"))
    if not rtf_files:
        print(f"No RTF files found in {CONFIG['INPUT_DIR']}")
        return

    for filepath in rtf_files:
        filename = os.path.basename(filepath)
        book_name = extract_book_name(filename)
        print(f"Processing {filename} ({book_name})...")

        try:
            text = read_rtf_simple(filepath)
            verses = process_text_content(text, book_name)
            if verses:
                df = pd.DataFrame(verses)
                output_filename = os.path.splitext(filename)[0] + ".xlsx"
                output_path = os.path.join(CONFIG["OUTPUT_DIR"], output_filename)
                df.to_excel(output_path, index=False)
                print(f"Saved {len(verses)} verses to {output_filename}")
            else:
                print(f"No verses found in {filename}")
        except Exception as e:
            print(f"Error processing {filename}: {str(e)}")

if __name__ == "__main__":
    main()
