#!/usr/bin/env python
"""
Produce one CSV per TXT file by matching each TXT line against
'01-Genesis_filtered_verses.xlsx' (Reference, Verse Text).
Matching rule: normalized Excel Verse Text startswith(normalized TXT line).
Skips TXT files that already have a corresponding CSV.
"""

import pandas as pd
import os
from typing import Optional

EXCEL_FILE = "01-Genesis_filtered_verses.xlsx"  # change if needed
FOLDER = "."  # folder containing TXT and Excel files


def normalize_text(text: Optional[str]) -> str:
    """Normalize text for matching: strip and unify common curly quotes/apostrophes.
    (Hyphens are left alone on purpose per client requirement.)
    """
    if text is None:
        return ""
    s = str(text).strip()
    s = s.replace("’", "'").replace("‘", "'")
    s = s.replace("“", '"').replace("”", '"')
    # collapse inner multiple spaces to single (optional but helpful)
    s = " ".join(s.split())
    return s


def load_reference(excel_path: str) -> pd.DataFrame:
    df = pd.read_excel(excel_path)
    # ensure required columns exist
    if "Reference" not in df.columns or "Verse Text" not in df.columns:
        raise ValueError("Excel must contain 'Reference' and 'Verse Text' columns.")
    # create a normalized column for startswith matching
    df["Normalized"] = df["Verse Text"].apply(normalize_text)
    return df[["Reference", "Verse Text", "Normalized"]]


def process_one_txt(txt_path: str, df_ref: pd.DataFrame) -> pd.DataFrame:
    """Return a DataFrame with columns ['Reference','Body','MatchedVerseText'] for this TXT file."""
    with open(txt_path, "r", encoding="utf-8") as fh:
        raw_lines = [line.rstrip("\n") for line in fh if line.strip()]

    results = []
    for raw in raw_lines:
        norm = normalize_text(raw)
        if norm == "":
            # preserve empty lines as Not matched (optional)
            results.append(("Not matched", raw, None))
            continue

        # find first verse whose normalized text starts with the normalized txt line
        matches = df_ref[df_ref["Normalized"].str.startswith(norm, na=False)]
        if not matches.empty:
            first = matches.iloc[0]
            results.append((first["Reference"], raw, first["Verse Text"]))
        else:
            results.append(("Not matched", raw, None))

    # Move "Not matched" rows to the top while preserving order otherwise
    not_matched = [r for r in results if r[0] == "Not matched"]
    matched = [r for r in results if r[0] != "Not matched"]
    ordered = not_matched + matched

    out_df = pd.DataFrame(ordered, columns=["Reference", "Body", "MatchedVerseText"])
    return out_df


def main():
    df_ref = load_reference(EXCEL_FILE)
    txt_files = [f for f in os.listdir(FOLDER) if f.lower().endswith(".txt")]

    if not txt_files:
        print("No .txt files found in folder.")
        return

    for txt in txt_files:
        csv_name = os.path.splitext(txt)[0] + ".csv"
        if os.path.exists(csv_name):
            print(f"Skipping {txt} → {csv_name} already exists.")
            continue

        print(f"Processing {txt} ...")
        out_df = process_one_txt(txt, df_ref)
        # save CSV with Body as original TXT line (and the reference)
        out_df.to_csv(csv_name, index=False, encoding="utf-8")
        print(f"Wrote {csv_name} ({len(out_df)} rows)")

    print("Done.")


if __name__ == "__main__":
    main()
