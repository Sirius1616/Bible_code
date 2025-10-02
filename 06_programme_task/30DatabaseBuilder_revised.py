#!/usr/bin/env python

import pandas as pd
import re
from rapidfuzz import fuzz
import glob
import os

# --- XLSX file path ---
xlsx_file = "CSB_GIFT_01-Genesis_all_spans.xlsx"

# --- Load XLSX file ---
df = pd.read_excel(xlsx_file, engine='openpyxl')

# --- Clean text function ---
def clean_text(text):
    if pd.isna(text):
        return ""
    text = re.sub(r'[\u200b\u00ad]', '', str(text))
    text = text.replace('"', '"').replace('"', '"')
    text = text.replace(''', "'").replace(''', "'")
    text = text.replace('–', '-').replace('—', '-')
    return text.strip()

# --- Reconstruct full verses with chapter ---
verses = []
current_chapter = None
current_verse = None
current_body = ""

for idx, row in df.iterrows():
    category = str(row.get('Text Category', '')).upper()
    content = clean_text(row.get('Span Content', ''))

    if "CHAPTER NUMBERS" in category:
        current_chapter = content
        continue

    if category in ["OTHER ELEMENTS", "BOOK TITLES", "SUBHEAD HEADINGS IN BIBLE TEXT"]:
        continue

    if "VERSE NUMBERS" in category:
        if current_verse is not None and current_body:
            verses.append((current_chapter, current_verse, current_body.strip()))
        current_verse = content
        current_body = ""
    elif "SCRIPTURE TEXT FONTS" in category:
        current_body += " " + content

# Append last verse
if current_verse is not None and current_body:
    verses.append((current_chapter, current_verse, current_body.strip()))

# --- Process all TXT files in current folder ---
txt_files = glob.glob("*.txt")

for txt_file in txt_files:
    output_file = txt_file.rsplit(".", 1)[0] + ".csv"

    # Skip if CSV already exists
    if os.path.exists(output_file):
        print(f"Skipping {txt_file} because {output_file} already exists.")
        continue

    # --- Load TXT file ---
    with open(txt_file, 'r', encoding='utf-8') as f:
        txt_lines = [line.strip() for line in f if line.strip()]

    # --- Match TXT fragments using fuzzy matching ---
    output_rows = []
    used_verses = set()  # Keep track of already matched verses

    for line in txt_lines:
        match_num = re.match(r'^(\d+)\s*(.*)', line)
        if match_num:
            txt_verse_num = match_num.group(1)
            fragment_text = match_num.group(2).strip()
        else:
            txt_verse_num = None
            fragment_text = line.strip()

        found = False
        for chap, verse, body in verses:
            if (chap, verse) in used_verses:
                continue
            if txt_verse_num is not None and str(verse) == txt_verse_num:
                score = fuzz.partial_ratio(fragment_text.lower(), body.lower())
                if score >= 80:  # Threshold for approximate match
                    output_rows.append([f"Genesis {chap}:{verse}", line])
                    used_verses.add((chap, verse))
                    found = True
                    break

        if not found:
            output_rows.append(["Not matched", line])

    # --- Reorder rows: "Not matched" first, then the rest in original order ---
    # Create a list to store reordered rows
    reordered_rows = []
    
    # First, add all "Not matched" rows
    for row in output_rows:
        if row[0] == "Not matched":
            reordered_rows.append(row)
    
    # Then, add all matched rows
    for row in output_rows:
        if row[0] != "Not matched":
            reordered_rows.append(row)

    # --- Save output ---
    output_df = pd.DataFrame(reordered_rows, columns=['Reference', 'Body']
    output_df.to_csv(output_file, index=False, encoding='utf-8')
    print(f"Processed {txt_file} → {output_file}")