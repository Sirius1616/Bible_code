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

def process_txt_files(txt_folder: str, excel_file: str, output_folder: str):
    # Load Excel file
    df = pd.read_excel(excel_file)
    df["Normalized"] = df["Verse Text"].apply(normalize_text)

    # Make sure output folder exists
    os.makedirs(output_folder, exist_ok=True)

    for txt_file in os.listdir(txt_folder):
        if txt_file.endswith(".txt"):
            results = []
            txt_path = os.path.join(txt_folder, txt_file)

            with open(txt_path, "r", encoding="utf-8") as f:
                txt_lines = [normalize_text(line) for line in f if line.strip()]

            for line in txt_lines:
                # look for verse text that CONTAINS this line
                matched_row = df[df["Normalized"].str.contains(line, na=False, regex=False)]

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

            # Convert results into DataFrame
            output_df = pd.DataFrame(results)

            # Move Not Matched rows to the top
            output_df = pd.concat([
                output_df[output_df["Match Status"] == "Not Matched"],
                output_df[output_df["Match Status"] == "Matched"]
            ])

            # Save results per txt file
            output_name = os.path.splitext(txt_file)[0] + "_results.xlsx"
            output_path = os.path.join(output_folder, output_name)
            output_df.to_excel(output_path, index=False)

            print(f"Results for {txt_file} saved to {output_path}")


# Example usage
process_txt_files(
    txt_folder=".",  
    excel_file="01-Genesis_filtered_verses.xlsx",
    output_folder="results"
)
