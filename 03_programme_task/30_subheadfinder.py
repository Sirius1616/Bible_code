#!/usr/bin/env python
# 30_subheadfinder_auto.py

import os
import re
from openpyxl import load_workbook
import csv

def load_txt_subheads(txt_file_path):
    """Load subhead titles from TXT file"""
    with open(txt_file_path, 'r', encoding='utf-8') as file:
        return [line.strip() for line in file if line.strip()]

def extract_xlsx_structure(xlsx_file_path):
    """
    Extract the complete structure from XLSX file:
    - Subheads with their row numbers
    - Verses with their references and row numbers
    - Chapter markers
    """
    wb = load_workbook(xlsx_file_path)
    ws = wb.active
    
    subheads = []  # (subhead_text, row_number)
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
            subheads.append((content, row_number))
        
        # CHAPTER DETECTION  
        elif "CHAPTER NUMBERS" in row_type and content.strip():
            # Extract chapter number (handle cases like "1 " with special characters)
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
            
            # Start new verse - extract verse number (handle special characters)
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
    """Match subhead titles to the verses that follow them"""
    results = []
    not_found = []
    
    # Create a mapping of subhead text to row number
    subhead_map = {}
    for subhead_text, row_num in subheads:
        subhead_map[subhead_text] = row_num
    
    for subhead_phrase in subhead_phrases:
        if subhead_phrase in subhead_map:
            subhead_row = subhead_map[subhead_phrase]
            verse_ref, verse_text = find_verse_after_subhead(subhead_row, verses)
            
            if verse_ref:
                results.append({
                    'reference': verse_ref,
                    'subhead': subhead_phrase
                })
            else:
                not_found.append(subhead_phrase)
        else:
            not_found.append(subhead_phrase)
    
    return results, not_found

def find_matching_files():
    """Find matching TXT and XLSX files in the current directory"""
    files = os.listdir('.')
    
    # Look for files with numbers at the beginning
    txt_files = [f for f in files if f.lower().endswith('.txt') and f[0].isdigit()]
    xlsx_files = [f for f in files if f.lower().endswith('.xlsx') and any(c.isdigit() for c in f)]
    
    # If no matches found with numbers, try any TXT and XLSX files
    if not txt_files or not xlsx_files:
        txt_files = [f for f in files if f.lower().endswith('.txt')]
        xlsx_files = [f for f in files if f.lower().endswith('.xlsx')]
    
    matches = []
    
    # Simple matching: if there's exactly one TXT and one XLSX, use them
    if len(txt_files) == 1 and len(xlsx_files) == 1:
        matches.append((txt_files[0], xlsx_files[0]))
    else:
        # Try to match by number prefix
        for txt_file in txt_files:
            # Extract number from beginning of filename
            txt_number_match = re.match(r'^(\d+)', txt_file)
            if txt_number_match:
                txt_number = txt_number_match.group(1)
                
                for xlsx_file in xlsx_files:
                    xlsx_number_match = re.search(r'(\d+)', xlsx_file)
                    if xlsx_number_match and xlsx_number_match.group(1) == txt_number:
                        matches.append((txt_file, xlsx_file))
                        break
    
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
    print("  Sample subheads from TXT:")
    for i, phrase in enumerate(subhead_phrases[:3], 1):
        print(f"    {i}. '{phrase}'")
    
    # Match subheads to verses
    results, not_found = match_subheads_to_verses(subhead_phrases, subheads, verses)
    print(f"  Results: {len(results)} matched, {len(not_found)} not found")
    
    # Create output filename (same as TXT file but with .csv extension)
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
    
    return len(results), len(not_found)

def main():
    print("=== Automated Bible Subhead to Verse Matcher ===")
    print("Searching for matching files in current directory...")
    
    # Find matching file pairs
    file_pairs = find_matching_files()
    
    if not file_pairs:
        # If no automatic matches found, try to use the specific files we know exist
        files = os.listdir('.')
        txt_candidates = [f for f in files if f.lower().endswith('.txt') and 'subhead' in f.lower()]
        xlsx_candidates = [f for f in files if f.lower().endswith('.xlsx') and 'genesis' in f.lower()]
        
        if txt_candidates and xlsx_candidates:
            file_pairs = [(txt_candidates[0], xlsx_candidates[0])]
            print("Using detected files:", file_pairs[0])
    
    if not file_pairs:
        print("No matching file pairs found.")
        print("Files in current directory:")
        for f in os.listdir('.'):
            print(f"  {f}")
        return
    
    print(f"Found {len(file_pairs)} matching file pair(s):")
    for txt, xlsx in file_pairs:
        print(f"  {txt} -> {xlsx}")
    
    print("\n" + "="*50)
    
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
            print()
    
    print("="*50)
    print("Processing complete!")
    print(f"Total across all files: {total_matched} matched, {total_not_found} not found")

if __name__ == "__main__":
    main()