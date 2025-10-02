#!/usr/bin/env python

import pandas as pd
import os

def normalize_text(text: str) -> str:
    """Normalize text for comparison (strip, standardize quotes)."""
    if not isinstance(text, str):
        return ""
    return (
        text.strip()
        .replace("’", "'")
        .replace("‘", "'")
        .replace("“", '"')
        .replace("”", '"')
    )

def match_txt_to_excel(txt_file: str, excel_file: str, output_file: str):
    # Load Excel file
    df = pd.read_excel(excel_file)
    df["Normalized"] = df["Verse Text"].apply(normalize_text)

    # Read TXT file
    with open(txt_file, "r", encoding="utf-8") as f:
        txt_lines = [normalize_text(line.strip()) for line in f if line.strip()]

    results = []
    for line in txt_lines:
        matched_row = df[df["Normalized"] == line]

        if not matched_row.empty:
            ref = matched_row.iloc[0]["Reference"]
            verse = matched_row.iloc[0]["Verse Text"]
            results.append({
                "TXT Line": line,
                "Match Status": "Matched",
                "Reference": ref,
                "Verse Text": verse
            })
        else:
            results.append({
                "TXT Line": line,
                "Match Status": "Not Matched",
                "Reference": None,
                "Verse Text": None
            })

    # Save results
    output_df = pd.DataFrame(results)
    output_df.to_excel(output_file, index=False)
    print(f"Results saved to {output_file}")


# Example usage:
match_txt_to_excel(
    txt_file="01-Genesis Body - 217.txt",
    excel_file="01-Genesis_filtered_verses.xlsx",
    output_file="matched_results.xlsx"
)
