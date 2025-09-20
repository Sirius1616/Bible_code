#!/usr/bin/env python
# 30_subheadfinder_auto.py

import os
import re
from openpyxl import load_workbook
import csv

def clean_text_content(text):
    """Clean text by handling quotes and special characters"""
    if not text:
        return ""
    
    # Replace different types of quotes with standard ones
    text = text.replace('"', '').replace("'", "").replace("´", "").replace("`", "")
    # Remove other special characters that might cause issues
    text = re.sub(r'[^\w\s\-–—:;,.!?()]', '', text)
    # Normalize whitespace
    text = re.sub(r'\s+', ' ', text).strip()
    return text

def load_txt_subheads(txt_file_path):
    """Load subhead titles from TXT file and clean them"""
    with open(txt_file_path, 'r', encoding='utf-8') as file:
        subheads = [line.strip() for line in file if line.strip()]
        return [clean_text_content(subhead) for subhead in subheads]

def extract_xlsx_structure(xlsx_file_path):
    """
    Extract the complete structure from XLSX file with proper text cleaning
    """
    wb = load_workbook(xlsx_file_path)
    ws = wb.active
    
    subheads = []  # (cleaned_text, row_number, original_text)
    verses = []    # (reference, verse_text, row_number)
    chapters = []  # (chapter_number, row_number)
    
    current_book = "Genesis"
    current_chapter = "1"
    current_verse = None
    verse_text_parts = []
    row_number = 0
    
    for row in ws.iter_rows(values_only=True):
        row_number += 1
        if not row or row[1] is None:
            continue
            
        row_type = str(row[1])
        content = str(row[2]) if len(row) > 2 and row[2] is not None else ""
        
        # SUBHEAD DETECTION
        if "SUBHEAD" in row_type and content:
            cleaned_content = clean_text_content(content)
            if cleaned_content:
                subheads.append((cleaned_content, row_number, content))  # Store both cleaned and original
        
        # CHAPTER DETECTION  
        elif "CHAPTER NUMBERS" in row_type and content.strip():
            # Extract chapter number
            chapter_match = ''.join(c for c in content if c.isdigit())
            if chapter_match:
                current_chapter = chapter_match
                chapters.append((current_chapter, row_number))
        
        # VERSE NUMBER DETECTION
        elif "VERSE NUMBERS" in row_type and content.strip():
            # Save previous verse if exists
            if current_verse is not None and verse_text_parts:
                full_text = " ".join(verse_text_parts).strip()
                if full_text:
                    reference = f"{current_book} {current_chapter}:{current_verse}"
                    verses.append((reference, full_text, row_number))
            
            # Start new verse
            verse_match = ''.join(c for c in content if c.isdigit())
            if verse_match:
                current_verse = verse_match
            verse_text_parts = []
        
        # VERSE TEXT DETECTION
        elif "SCRIPTURE TEXT" in row_type and content and content not in ['""', '']:
            clean_content = content.replace('"', '').replace('""', '').strip()
            if clean_content:
                verse_text_parts.append(clean_content)
    
    # Add the last verse
    if current_verse is not None and verse_text_parts:
        full_text = " ".join(verse_text_parts).strip()
        if full_text:
            reference = f"{current_book} {current_chapter}:{current_verse}"
            verses.append((reference, full_text, row_number))
    
    return subheads, verses, chapters

def find_verse_after_subhead(subhead_row, verses):
    """Find the first verse that comes after a subhead"""
    for ref, text, verse_row in verses:
        if verse_row > subhead_row:
            return ref, text
    return None, None

def match_subheads_to_verses(subhead_phrases, subheads, verses):
    """Match subhead titles to the verses that follow them with fuzzy matching"""
    results = []
    not_found = []
    
    # Create mappings for both cleaned and original subhead text
    subhead_clean_map = {}  # cleaned text -> (row_number, original_text)
    subhead_original_map = {}  # original text -> row_number
    
    for cleaned_text, row_num, original_text in subheads:
        subhead_clean_map[cleaned_text] = (row_num, original_text)
        subhead_original_map[original_text] = row_num
    
    for subhead_phrase in subhead_phrases:
        cleaned_phrase = clean_text_content(subhead_phrase)
        found = False
        
        # Try exact match with cleaned text first
        if cleaned_phrase in subhead_clean_map:
            row_num, original_text = subhead_clean_map[cleaned_phrase]
            verse_ref, verse_text = find_verse_after_subhead(row_num, verses)
            
            if verse_ref:
                results.append({
                    'reference': verse_ref,
                    'subhead': subhead_phrase,  # Use original phrase from TXT
                    'original_xlsx_subhead': original_text  # For debugging
                })
                found = True
        
        # If not found, try fuzzy matching
        if not found:
            best_match = None
            best_score = 0
            
            for clean_text, (row_num, original_text) in subhead_clean_map.items():
                # Simple fuzzy matching: check if one contains the other
                if cleaned_phrase in clean_text or clean_text in cleaned_phrase:
                    similarity = len(set(cleaned_phrase.split()) & set(clean_text.split())) / max(len(cleaned_phrase.split()), len(clean_text.split()))
                    if similarity > 0.7:  # 70% similarity threshold
                        verse_ref, verse_text = find_verse_after_subhead(row_num, verses)
                        if verse_ref:
                            results.append({
                                'reference': verse_ref,
                                'subhead': subhead_phrase,
                                'original_xlsx_subhead': original_text
                            })
                            found = True
                            break
            
            # Try matching with original text (before cleaning)
            if not found and subhead_phrase in subhead_original_map:
                row_num = subhead_original_map[subhead_phrase]
                verse_ref, verse_text = find_verse_after_subhead(row_num, verses)
                if verse_ref:
                    results.append({
                        'reference': verse_ref,
                        'subhead': subhead_phrase,
                        'original_xlsx_subhead': subhead_phrase
                    })
                    found = True
        
        if not found:
            not_found.append(subhead_phrase)
    
    return results, not_found

def find_matching_files():
    """Find matching TXT and XLSX files in the current directory"""
    files = os.listdir('.')
    
    # Look for the specific files we know exist
    txt_files = [f for f in files if f.lower().endswith('.txt') and 'subhead' in f.lower()]
    xlsx_files = [f for f in files if f.lower().endswith('.xlsx') and 'genesis' in f.lower()]
    
    matches = []
    
    # If we find both, use them
    if txt_files and xlsx_files:
        matches.append((txt_files[0], xlsx_files[0]))
    
    return matches

def process_single_pair(txt_file, xlsx_file):
    """Process a single pair of files"""
    print(f"Processing: {txt_file} with {xlsx_file}")
    
    # Load data
    subhead_phrases = load_txt_subheads(txt_file)
    print(f"  Loaded {len(subhead_phrases)} subhead titles from TXT")
    
    subheads, verses, chapters = extract_xlsx_structure(xlsx_file)
    print(f"  Found {len(subheads)} subheads, {len(verses)} verses, and {len(chapters)} chapters in XLSX")
    
    # Show samples for verification
    print("  Sample subheads from TXT (after cleaning):")
    for i, phrase in enumerate(subhead_phrases[:3], 1):
        print(f"    {i}. '{phrase}'")
    
    print("  Sample subheads from XLSX (original -> cleaned):")
    for i, (cleaned_text, row_num, original_text) in enumerate(subheads[:3], 1):
        print(f"    {i}. '{original_text}' -> '{cleaned_text}'")
    
    # Match subheads to verses
    results, not_found = match_subheads_to_verses(subhead_phrases, subheads, verses)
    print(f"  Results: {len(results)} matched, {len(not_found)} not found")
    
    # Create output filename
    output_file = os.path.splitext(txt_file)[0] + '.csv'
    
    # Write output
    with open(output_file, 'w', newline='', encoding='utf-8') as f:
        writer = csv.writer(f, delimiter='\t')
        
        # Unmatched items first
        if not_found:
            writer.writerow(["ERROR: Could not find verses for these subheads:"])
            writer.writerow([])
            for phrase in not_found:
                writer.writerow([f"NOT_FOUND: {phrase}"])
            writer.writerow([])
            writer.writerow([])
        
        # Matches
        writer.writerow(["Reference", "Subhead"])
        for result in results:
            writer.writerow([result['reference'], result['subhead']])
    
    print(f"  Output saved to: {output_file}")
    
    # Show some examples
    if results:
        print("  First 3 matches:")
        for result in results[:3]:
            print(f"    {result['reference']} -> {result['subhead']}")
            if 'original_xlsx_subhead' in result and result['subhead'] != result['original_xlsx_subhead']:
                print(f"      (XLSX had: '{result['original_xlsx_subhead']}')")
    
    # Debug: show some not found items
    if not_found:
        print("  Sample not found items:")
        for phrase in not_found[:5]:
            print(f"    NOT FOUND: '{phrase}'")
    
    return len(results), len(not_found)

def main():
    print("=== Automated Bible Subhead to Verse Matcher ===")
    print("Now with improved quote handling and fuzzy matching")
    
    # Find matching file pairs
    file_pairs = find_matching_files()
    
    if not file_pairs:
        print("No matching files found.")
        print("Files in current directory:")
        for f in os.listdir('.'):
            print(f"  {f}")
        return
    
    print(f"Found matching file pair: {file_pairs[0]}")
    print("\n" + "="*60)
    
    total_matched = 0
    total_not_found = 0
    
    # Process each file pair
    for txt_file, xlsx_file in file_pairs:
        try:
            matched, not_found = process_single_pair(txt_file, xlsx_file)
            total_matched += matched
            total_not_found += not_found
            print()
        except Exception as e:
            print(f"Error processing {txt_file}: {e}")
            import traceback
            traceback.print_exc()
            print()
    
    print("="*60)
    print("Processing complete!")
    print(f"Total: {total_matched} matched, {total_not_found} not found")

if __name__ == "__main__":
    main()