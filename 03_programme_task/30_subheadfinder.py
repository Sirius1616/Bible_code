#!/usr/bin/env python3
"""
Bible Verse Reference Finder

This script processes an Excel file containing typesetting data for Genesis
and matches scripture phrases to their exact verse references.
"""

import pandas as pd
import re
from typing import Dict, List, Tuple, Optional
import ast


def parse_bbox(bbox_str: str) -> Tuple[float, float, float, float]:
    """Parse bounding box string into coordinates (x0, y0, x1, y1)."""
    try:
        # Handle string representation of tuple
        if isinstance(bbox_str, str):
            # Remove any whitespace and evaluate the string as a Python tuple
            bbox = ast.literal_eval(bbox_str.strip())
        else:
            bbox = bbox_str
        
        return tuple(map(float, bbox))
    except (ValueError, SyntaxError, TypeError):
        # Return a default bbox if parsing fails
        return (0.0, 0.0, 0.0, 0.0)


def sort_by_reading_order(df: pd.DataFrame) -> pd.DataFrame:
    """Sort dataframe by reading order: top to bottom, then left to right."""
    # Parse bounding boxes and create sort columns
    df = df.copy()
    df['bbox_parsed'] = df['Span Position (bbox)'].apply(parse_bbox)
    df['y0'] = df['bbox_parsed'].apply(lambda x: x[1])  # top coordinate
    df['x0'] = df['bbox_parsed'].apply(lambda x: x[0])  # left coordinate
    
    # Sort by y0 (top) first, then x0 (left)
    df_sorted = df.sort_values(['y0', 'x0'], ascending=[False, True])  # y decreases as we go down
    
    return df_sorted


def reconstruct_verses_from_excel(excel_path: str) -> Dict[str, str]:
    """
    Reconstruct complete verses from Excel typesetting data.
    
    Returns:
        Dict mapping verse references (e.g., "1:2") to complete verse text
    """
    print(f"Reading Excel file: {excel_path}")
    
    # Read the Excel file
    try:
        df = pd.read_excel(excel_path)
        print(f"Loaded {len(df)} rows from Excel file")
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        return {}
    
    # Display column names for debugging
    print("Column names:")
    for i, col in enumerate(df.columns):
        print(f"  {i}: '{col}'")
    
    # Check if required columns exist
    required_columns = ['Text Category', 'Span Content', 'Span Position (bbox)']
    missing_columns = [col for col in required_columns if col not in df.columns]
    
    if missing_columns:
        print(f"Error: Missing required columns: {missing_columns}")
        return {}
    
    # Sort by reading order
    df_sorted = sort_by_reading_order(df)
    
    # Extract verses
    verses = {}
    current_chapter = 1  # Default to chapter 1
    current_verse_num = None
    current_verse_text = []
    
    print("\nProcessing rows to reconstruct verses...")
    
    for idx, row in df_sorted.iterrows():
        text_category = str(row['Text Category']).strip()
        span_content = str(row['Span Content']).strip()
        
        # Skip empty content
        if not span_content or span_content == 'nan':
            continue
            
        # Check for chapter numbers
        if text_category == "CHAPTER NUMBERS":
            try:
                current_chapter = int(span_content)
                print(f"Found chapter: {current_chapter}")
            except ValueError:
                pass
        
        # Check for verse numbers
        elif text_category == "VERSE NUMBERS":
            # Save previous verse if we have one
            if current_verse_num is not None and current_verse_text:
                verse_ref = f"{current_chapter}:{current_verse_num}"
                verse_text = ' '.join(current_verse_text).strip()
                verses[verse_ref] = verse_text
                print(f"Saved verse {verse_ref}: {verse_text[:50]}...")
            
            # Start new verse
            try:
                current_verse_num = int(span_content)
                current_verse_text = []
            except ValueError:
                print(f"Warning: Could not parse verse number: '{span_content}'")
        
        # Collect scripture text
        elif text_category == "SCRIPTURE TEXT FONTS":
            if current_verse_num is not None:
                current_verse_text.append(span_content)
    
    # Don't forget the last verse
    if current_verse_num is not None and current_verse_text:
        verse_ref = f"{current_chapter}:{current_verse_num}"
        verse_text = ' '.join(current_verse_text).strip()
        verses[verse_ref] = verse_text
        print(f"Saved final verse {verse_ref}: {verse_text[:50]}...")
    
    print(f"\nReconstructed {len(verses)} verses")
    return verses


def load_phrases_from_txt(txt_path: str) -> List[str]:
    """Load phrases from text file, one per line."""
    print(f"\nReading phrases from: {txt_path}")
    
    try:
        with open(txt_path, 'r', encoding='utf-8') as f:
            phrases = [line.strip() for line in f if line.strip()]
        print(f"Loaded {len(phrases)} phrases")
        return phrases
    except Exception as e:
        print(f"Error reading text file: {e}")
        return []


def find_phrase_matches(phrases: List[str], verses: Dict[str, str]) -> List[Tuple[str, str]]:
    """
    Find which verse each phrase comes from.
    
    Returns:
        List of (reference, phrase) tuples. Reference is "NOT_FOUND" if no match.
    """
    print(f"\nMatching {len(phrases)} phrases against {len(verses)} verses...")
    
    results = []
    found_count = 0
    
    for phrase in phrases:
        found_match = False
        
        # Search through all verses
        for verse_ref, verse_text in verses.items():
            if phrase in verse_text:
                results.append((f"Genesis {verse_ref}", phrase))
                found_match = True
                found_count += 1
                print(f"Found match: '{phrase}' in Genesis {verse_ref}")
                break
        
        if not found_match:
            results.append(("NOT_FOUND", phrase))
            print(f"No match found for: '{phrase}'")
    
    print(f"\nMatching complete: {found_count}/{len(phrases)} phrases matched")
    return results


def write_results_to_csv(results: List[Tuple[str, str]], output_path: str):
    """Write results to tab-delimited CSV file with NOT_FOUND entries first."""
    print(f"\nWriting results to: {output_path}")
    
    # Separate found and not found results
    not_found = [result for result in results if result[0] == "NOT_FOUND"]
    found = [result for result in results if result[0] != "NOT_FOUND"]
    
    # Sort found results by reference (Genesis chapter:verse)
    def extract_chapter_verse(ref: str) -> Tuple[int, int]:
        """Extract chapter and verse numbers for sorting."""
        try:
            # Extract "1:2" from "Genesis 1:2"
            parts = ref.replace("Genesis ", "").split(":")
            return (int(parts[0]), int(parts[1]))
        except:
            return (999, 999)  # Put any parsing errors at the end
    
    found.sort(key=lambda x: extract_chapter_verse(x[0]))
    
    # Combine results with NOT_FOUND first
    all_results = not_found + found
    
    try:
        with open(output_path, 'w', encoding='utf-8') as f:
            for reference, phrase in all_results:
                f.write(f"{reference}\t{phrase}\n")
        
        print(f"Successfully wrote {len(all_results)} results")
        print(f"  - {len(not_found)} NOT_FOUND entries")
        print(f"  - {len(found)} matched entries")
        
    except Exception as e:
        print(f"Error writing output file: {e}")


def main():
    """Main function to orchestrate the verse finding process."""
    print("=== Bible Verse Reference Finder ===\n")
    
    # File paths (modify these as needed)
    excel_path = "CSB_GIFT_01-Genesis_all_spans.xlsx"
    txt_path = "01-Genesis body - 479 (1).txt"
    output_path = "output.csv"
    
    # Step 1: Reconstruct verses from Excel file
    verses = reconstruct_verses_from_excel(excel_path)
    if not verses:
        print("Error: Could not reconstruct verses from Excel file")
        return
    
    # Step 2: Load phrases from text file
    phrases = load_phrases_from_txt(txt_path)
    if not phrases:
        print("Error: Could not load phrases from text file")
        return
    
    # Step 3: Find matches
    results = find_phrase_matches(phrases, verses)
    
    # Step 4: Write results
    write_results_to_csv(results, output_path)
    
    print(f"\n=== Process Complete ===")
    print(f"Output saved to: {output_path}")


if __name__ == "__main__":
    main