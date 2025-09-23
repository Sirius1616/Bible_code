#!/usr/bin/env python

import csv
import os
import pandas as pd

def clean_subhead_file(raw_file, cleaned_file):
    """Split merged Reference+Subhead into two columns"""
    with open(raw_file, "r", newline="", encoding="utf-8") as infile, \
         open(cleaned_file, "w", newline="", encoding="utf-8") as outfile:

        reader = csv.reader(infile)
        writer = csv.writer(outfile)

        writer.writerow(["Reference", "Subhead"])  # new header

        next(reader)  # skip the original header
        for row in reader:
            if not row:
                continue
            line = row[0].strip().strip('"')
            if "\t" in line:
                ref, sub = line.split("\t", 1)
            elif " " in line:
                parts = line.split(" ", 1)
                ref, sub = parts[0], parts[1] if len(parts) > 1 else ""
            else:
                ref, sub = line, ""
            writer.writerow([ref.strip(), sub.strip()])

    print(f"✅ Cleaned file saved as: {cleaned_file}")


def match_subheads_with_spans(cleaned_file, span_file, output_file):
    """Match cleaned subheads with span contents from Excel"""
    spans_df = pd.read_excel(span_file)
    spans = spans_df.to_dict(orient="records")

    with open(cleaned_file, "r", newline="", encoding="utf-8") as infile, \
         open(output_file, "w", newline="", encoding="utf-8") as outfile:

        reader = csv.DictReader(infile)
        fieldnames = ["Reference", "Subhead", "Match Status", "X-Coord", "Page", "Even/Odd"]
        writer = csv.DictWriter(outfile, fieldnames=fieldnames)
        writer.writeheader()

        for row in reader:
            subhead = row["Subhead"].strip()
            match = None

            for span in spans:
                span_text = str(span.get("Span Content", "")).strip()
                if span_text == subhead:
                    match = span
                    break

            if match:
                page = match.get("Page Number", "")
                writer.writerow({
                    "Reference": row["Reference"],
                    "Subhead": row["Subhead"],
                    "Match Status": "MATCH",
                    "X-Coord": match.get("Span Position (bbox)", ""),
                    "Page": page,
                    "Even/Odd": "EVEN" if str(page).isdigit() and int(page) % 2 == 0 else "ODD"
                })
            else:
                writer.writerow({
                    "Reference": row["Reference"],
                    "Subhead": row["Subhead"],
                    "Match Status": "COULD NOT MATCH",
                    "X-Coord": "",
                    "Page": "",
                    "Even/Odd": ""
                })

    print(f"✅ Matching done. Output saved as: {output_file}")


if __name__ == "__main__":
    folder = "."  # current folder
    files = os.listdir(folder)

    # find all subhead CSVs
    subhead_files = [f for f in files if f.endswith(".csv") and "subhead" in f and not f.startswith("matched")]

    for subhead_file in subhead_files:
        prefix = subhead_file.split("_subhead")[0]  # e.g., "01-Genesis"
        span_file_candidates = [f for f in files if f.startswith(prefix) and f.endswith("_all_spans.xlsx")]

        if not span_file_candidates:
            print(f"❌ No span file found for {subhead_file}, skipping.")
            continue

        span_file = span_file_candidates[0]
        cleaned_file = f"{prefix}_subhead_clean.csv"
        output_file = f"matched_{subhead_file}"

        # Run cleaning and matching
        clean_subhead_file(subhead_file, cleaned_file)
        match_subheads_with_spans(cleaned_file, span_file, output_file)





    
  