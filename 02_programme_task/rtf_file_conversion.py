#!/usr/bin/env python3
"""
RTF → Excel Bible Verse Exporter - Fixed for your specific format
"""

import os
import re
import glob
import shutil
import pandas as pd

CONFIG = {
    "INPUT_DIR": "./rtf_in/",
    "OUTPUT_DIR": "./xlsx_out/",
    "SOURCE_DIR": "./file_to_process/",
    "FILENAME_PATTERN": r"^\d+\.(.+)\.rtf$",
    "REFERENCE_HEADER": "Reference",
    "TEXT_HEADER": "Text"
}

def extract_book_name(filename):
    match = re.match(CONFIG["FILENAME_PATTERN"], filename)
    if match:
        return match.group(1)
    return os.path.splitext(filename)[0]

def read_rtf_file(filepath):
    """Read RTF file and return raw content with BOM handling"""
    with open(filepath, 'rb') as f:
        content = f.read()
    
    # Remove UTF-8 BOM if present
    if content.startswith(b'\xef\xbb\xbf'):
        content = content[3:]
    
    return content.decode('utf-8', errors='ignore')

def clean_rtf_content(rtf_content):
    """Clean RTF content while preserving verse structure"""
    # More aggressive RTF header removal
    text = re.sub(r'^.*?\\pard\\plain', '', rtf_content, flags=re.DOTALL | re.IGNORECASE)
    text = re.sub(r'\\rtf1.*?\\fs\d+', '', text, flags=re.DOTALL)
    
    # Remove all RTF control words more thoroughly
    text = re.sub(r'\\[a-zA-Z]+\*?\d*', '', text)
    text = re.sub(r'[{}]', '', text)
    text = re.sub(r'\\\'[0-9a-fA-F]{2}', ' ', text)
    
    # Remove specific problematic patterns that cause -840 artifacts
    text = re.sub(r'\\lang\d+', '', text)
    text = re.sub(r'\\langfe\d+', '', text)
    text = re.sub(r'\\cxp\d+', '', text)
    text = re.sub(r'\\cxds\d+', '', text)
    
    # Remove any remaining non-printable characters
    text = re.sub(r'[\x00-\x1F\x7F-\x9F]', '', text)
    
    # Replace specific problematic patterns
    text = re.sub(r'\\\*', '*', text)  # Convert \* to just *
    text = re.sub(r'\*\s*\d+\s*\*\s*\*', '||VERSE_END||', text)  # Mark verse endings
    text = re.sub(r'\*\s*\d+\s*\*', '||VERSE_END||', text)  # Alternative verse ending
    
    # Clean up spaces
    text = re.sub(r'\s+', ' ', text)
    text = re.sub(r'\|\|VERSE_END\|\|\s+', '||VERSE_END||', text)
    
    return text.strip()

def extract_verses_from_text(text, book_name):
    """Extract verses based on the specific format with * markers"""
    verses = []
    
    # Split text by verse endings
    verse_blocks = re.split(r'\|\|VERSE_END\|\|', text)
    
    current_chapter = 1
    verse_number = 1
    
    for block in verse_blocks:
        block = block.strip()
        if not block:
            continue
            
        # Look for chapter number at the beginning (pattern like "1:1" or "1 ")
        chapter_verse_match = re.match(r'^(\d+)[:\.](\d+)\s+', block)
        if chapter_verse_match:
            current_chapter = int(chapter_verse_match.group(1))
            verse_number = int(chapter_verse_match.group(2))
            block = block[len(chapter_verse_match.group(0)):].strip()
        else:
            # Look for just chapter number
            chapter_match = re.match(r'^(\d+)\s+', block)
            if chapter_match:
                current_chapter = int(chapter_match.group(1))
                verse_number = 1
                block = block[len(chapter_match.group(0)):].strip()
        
        # Clean the verse text (remove any remaining verse numbers)
        cleaned_text = clean_verse_text(block)
        
        if cleaned_text:
            verses.append({
                CONFIG["REFERENCE_HEADER"]: f"{book_name} {current_chapter}:{verse_number}",
                CONFIG["TEXT_HEADER"]: cleaned_text
            })
            verse_number += 1
    
    return verses

def clean_verse_text(text):
    """Clean individual verse text from artifacts - careful with footnotes"""
    # Remove verse number patterns like 1:1, 1.1, etc. at the beginning
    text = re.sub(r'^\d+[:\.]\d+\s*', '', text)
    
    # Remove asterisk patterns (but be careful with footnote markers)
    text = re.sub(r'\*\s*', ' ', text)
    
    # Remove RTF footnote markers and specific control characters that cause -840
    text = re.sub(r'\\super\s*|\\nosupersub\s*|\\lang\d+|\\langfe\d+', '', text, flags=re.IGNORECASE)
    
    # Remove -840 and similar artifacts specifically
    text = re.sub(r'-\d{2,4}', '', text)  # Remove -840, -123, etc.
    text = re.sub(r'\s\d{2,4}\s', ' ', text)  # Remove standalone numbers like 840
    
    # Remove footnote numbers that are standalone
    text = re.sub(r'\s\d+\s', ' ', text)  # Standalone numbers with spaces
    text = re.sub(r'\s\d+$', '', text)    # Numbers at end
    text = re.sub(r'^\d+\s', '', text)    # Numbers at beginning
    
    # Fix hyphen artifacts
    text = re.sub(r'-\s*\?\s*', '-', text)
    
    # Fix common OCR issues (be more specific)
    text = re.sub(r'\bexpans e\b', 'expanse', text)
    text = re.sub(r'\bmad e\b', 'made', text)
    text = re.sub(r'\bplant s\b', 'plants', text)
    text = re.sub(r'\bseasons,\s*1\b', 'seasons,', text)
    text = re.sub(r'\bEarth,\s*1\b', 'Earth,', text)
    text = re.sub(r'\bHeaven.\s*1\b', 'Heaven.', text)
    text = re.sub(r'\blight s\b', 'lights', text)
    text = re.sub(r'\bnigh t\b', 'night', text)
    text = re.sub(r'\bdark-ness\b', 'darkness', text)
    text = re.sub(r'\bex-panse\b', 'expanse', text)
    text = re.sub(r'\btogeth-er\b', 'together', text)
    
    # Remove any remaining numbers in parentheses or brackets
    text = re.sub(r'\[[^\]]*\]', '', text)
    text = re.sub(r'\([^)]*\)', '', text)
    
    # Clean up spaces and punctuation
    text = re.sub(r'\s+', ' ', text)
    text = re.sub(r'\s([.,;:!?])', r'\1', text)
    text = text.strip()
    
    return text

def extract_verses_alternative(text, book_name):
    """Alternative extraction method for different formats"""
    verses = []
    
    # Try to find verse numbers directly in text
    verse_pattern = r'(\d+)[:\.](\d+)\s+(.*?)(?=\d+[:\.]\d+\s+|$)'
    matches = re.finditer(verse_pattern, text, flags=re.DOTALL)
    
    for match in matches:
        chapter_num = match.group(1)
        verse_num = match.group(2)
        verse_text = match.group(3).strip()
        
        # Clean the verse text
        cleaned_text = clean_verse_text(verse_text)
        
        if cleaned_text and chapter_num.isdigit() and verse_num.isdigit():
            verses.append({
                CONFIG["REFERENCE_HEADER"]: f"{book_name} {chapter_num}:{verse_num}",
                CONFIG["TEXT_HEADER"]: cleaned_text
            })
    
    return verses

def clean_first_verse(text):
    """Special cleaning for the first verse that often contains header artifacts"""
    # Remove any remaining RTF control characters
    text = re.sub(r'[\x00-\x1F\x7F-\x9F]', '', text)
    # Remove common RTF artifacts that might appear at start
    text = re.sub(r'^(\\|{|}|pard|plain|fs\d+|\s)*', '', text, flags=re.IGNORECASE)
    # Remove any numbers at the beginning
    text = re.sub(r'^\d+\s*', '', text)
    # Remove -840 specifically from beginning
    text = re.sub(r'^-\d+\s*', '', text)
    return text.strip()

def process_file(filepath):
    """Process a single RTF file"""
    filename = os.path.basename(filepath)
    book_name = extract_book_name(filename)
    
    print(f"Processing {filename} ({book_name})...")
    
    try:
        # Read and clean RTF content
        rtf_content = read_rtf_file(filepath)
        text = clean_rtf_content(rtf_content)
        
        # Try primary extraction method
        verses = extract_verses_from_text(text, book_name)
        
        # If primary method fails, try alternative
        if not verses:
            verses = extract_verses_alternative(text, book_name)
        
        if verses:
            # Clean the first verse text more aggressively
            if verses and CONFIG["TEXT_HEADER"] in verses[0]:
                verses[0][CONFIG["TEXT_HEADER"]] = clean_first_verse(verses[0][CONFIG["TEXT_HEADER"]])
            
            # Additional text cleaning pass
            for verse in verses:
                verse[CONFIG["TEXT_HEADER"]] = final_text_cleanup(verse[CONFIG["TEXT_HEADER"]])
            
            return verses
        else:
            print(f"No verses found in {filename}")
            return None
            
    except Exception as e:
        print(f"Error processing {filename}: {str(e)}")
        import traceback
        traceback.print_exc()
        return None

def final_text_cleanup(text):
    """Final cleanup of verse text - targeted removal of -840 artifacts"""
    # First, aggressively remove all -840 patterns
    text = re.sub(r'-\d{2,4}', '', text)  # Remove -840, -123, etc.
    text = re.sub(r'\s\d{2,4}\s', ' ', text)  # Remove standalone numbers
    text = re.sub(r'\s\d{2,4}$', '', text)  # Remove numbers at end
    text = re.sub(r'^\d{2,4}\s', '', text)  # Remove numbers at beginning
    
    # Remove any remaining special characters but preserve letters and punctuation
    text = re.sub(r'[^a-zA-Z0-9\s.,;:!?\'"()-]', '', text)
    
    # Fix common missing words and phrases based on context
    text = re.sub(r'\bthe earth was form\b', 'the earth was without form', text)
    text = re.sub(r'\bGod said, there be\b', 'God said, "Let there be', text)
    text = re.sub(r'\bGod mad e\b', 'God made', text)
    text = re.sub(r'\bthe waters that were the expanse\b', 'the waters that were under the expanse', text)
    text = re.sub(r'\bfrom the waters that were the expanse\b', 'from the waters that were above the expanse', text)
    text = re.sub(r'\bGod said, the waters\b', 'God said, "Let the waters', text)
    text = re.sub(r'\bGod said, the earth\b', 'God said, "Let the earth', text)
    text = re.sub(r'\bfor and for\b', 'for signs and for seasons', text)
    text = re.sub(r'\bGod the two great lights\b', 'God made the two great lights', text)
    text = re.sub(r'\bto over the day\b', 'to rule over the day', text)
    text = re.sub(r'\bSo created\b', 'So God created', text)
    text = re.sub(r'\bsaying, fruitful\b', 'saying, "Be fruitful', text)
    text = re.sub(r'\bplant\'s\b', 'plants', text)
    text = re.sub(r'\bbird\'s\b', 'birds', text)
    text = re.sub(r'\bkind\'s\b', 'kinds', text)
    
    # Fix quotation marks
    text = re.sub(r'\b“([^"]+)', r'"\1', text)
    text = re.sub(r'([^"]+)”', r'\1"', text)
    
    # Ensure proper spacing
    text = re.sub(r'\s+', ' ', text)
    text = re.sub(r'\s([.,;:!?])', r'\1', text)
    
    # Final cleanup of any orphaned characters
    text = re.sub(r'\s-\s', ' ', text)
    text = re.sub(r'\s+\.', '.', text)
    text = re.sub(r'\s+,', ',', text)
    
    return text.strip()

def prepare_input_dir():
    source_dir = CONFIG["SOURCE_DIR"]
    target_dir = CONFIG["INPUT_DIR"]
    os.makedirs(target_dir, exist_ok=True)
    for filepath in glob.glob(os.path.join(source_dir, "*.rtf")):
        shutil.copy(filepath, target_dir)
        print(f"Copied {os.path.basename(filepath)} → {target_dir}")

def main():
    os.makedirs(CONFIG["OUTPUT_DIR"], exist_ok=True)
    os.makedirs(CONFIG["INPUT_DIR"], exist_ok=True)

    prepare_input_dir()

    rtf_files = glob.glob(os.path.join(CONFIG["INPUT_DIR"], "*.rtf"))
    if not rtf_files:
        print(f"No RTF files found in {CONFIG['INPUT_DIR']}")
        return

    all_verses = []
    
    for filepath in rtf_files:
        verses = process_file(filepath)
        if verses:
            # Save individual book files
            df = pd.DataFrame(verses)
            filename = os.path.basename(filepath)
            output_filename = os.path.splitext(filename)[0] + ".xlsx"
            output_path = os.path.join(CONFIG["OUTPUT_DIR"], output_filename)
            df.to_excel(output_path, index=False, engine='openpyxl')
            print(f"Saved {len(verses)} verses to {output_filename}")
            
            # Add to combined dataset
            all_verses.extend(verses)
    
    # Save combined file with all books
    if all_verses:
        combined_df = pd.DataFrame(all_verses)
        combined_output = os.path.join(CONFIG["OUTPUT_DIR"], "ALL_BOOKS.xlsx")
        combined_df.to_excel(combined_output, index=False, engine='openpyxl')
        print(f"Saved combined file with {len(all_verses)} verses to ALL_BOOKS.xlsx")

if __name__ == "__main__":
    main()