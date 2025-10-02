#!/usr/bin/env python

import pandas as pd
import os

def normalize_text(text: str) -> str:
    """Normalize text for comparison (strip, unify quotes)."""
    if not isinstance(text, str):
        return ""
    return (
        text.strip()
        .replace("’", "'")
        .replace("‘", "'")
        .replace("“", '"')
        .replace("”", '"')
    )

def process_txt_files(txt_folder: str, excel_file: str, output_file: str):
    # Load Excel file
    df = pd.read_excel(excel_file)
    df["Normalized"] = df["Verse Text"].apply(normalize_text)

    results = []

    for txt_file in os.listdir(txt_folder):
        if txt_file.endswith(".txt"):
            with open(os.path.join(txt_folder, txt_file), "r", encoding="utf-8") as f:
                txt_lines = [normalize_text(line) for line in f if line.strip()]

            for line in txt_lines:
                # look for verse text that STARTS with this line
                matched_row = df[df["Normalized"].str.startswith(line)]

                if not matched_row.empty:
                    ref = matched_row.iloc[0]["Reference"]
                    verse = matched_row.iloc[0]["Verse Text"]
                    results.append({
                        "Source File": txt_file,
                        "TXT Line": line,
                        "Match Status": "Matched",
                        "Reference": ref,
                        "Verse Text": verse
                    })
                else:
                    results.append({
                        "Source File": txt_file,
                        "TXT Line": line,
                        "Match Status": "Not Matched",
                        "Reference": None,
                        "Verse Text": None
                    })

    # Save results
    output_df = pd.DataFrame(results)
    output_df.to_excel(output_file, index=False)
    print(f"Results saved to {output_file}")


# Example usage
process_txt_files(
    txt_folder=".",  
    excel_file="01-Genesis_filtered_verses.xlsx",
    output_file="all_results.xlsx"
)
