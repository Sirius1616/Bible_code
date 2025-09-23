#!/usr/bin/env python

import csv
import os
import pandas as pd
import re


def clean_subhead_file(raw_file, cleaned_file):
    """Split merged Reference+Subhead into two columns"""
    if os.path.exists(cleaned_file):
        print(f"⚠️ Cleaned file already exists, skipping: {cleaned_file}")
        return

    with open(raw_file, "r", newline="", encoding="utf-8") as infile, \
         open(cleaned_file, "w", newline="", encoding="utf-8") as outfile:

        reader = csv.reader(infile)
        writer = csv.writer(outfile)

        writer.writerow(["Reference", "Subhead"])  # new header

        try:
            next(reader)  # skip the original header
        except StopIteration:
            print(f"⚠️ {raw_file} is empty or already cleaned, skipping.")
            return

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


def normalize_apostrophes(text: str) -> str:
    """Normalize straight and curly apostrophes for reliable matching."""
    if not isinstance(text, str):
        return ""
    return text.replace("’", "'").replace("‘", "'").strip()


def load_standards(standards_file):
    """Read Standards.txt and return list of acceptable coords + variance"""
    standards = []
    variance = 0.0
    with open(standards_file, "r", encoding="utf-8") as f:
        lines = [line.strip() for line in f if line.strip()]

    in_coords = False
    for line in lines:
        if line.startswith("Acceptable X Coordiantes"):
            in_coords = True
            continue
        if line.startswith("Acceptable Variance allowed"):
            in_coords = False
            continue
        if in_coords:
            try:
                standards.append(float(line))
            except ValueError:
                pass
        if line.startswith("Acceptable Variance allowed"):
            # extract number from "Acceptable Variance allowed +/- 0.3"
            match = re.search(r"([0-9]*\.?[0-9]+)", line)
            if match:
                variance = float(match.group(1))

    print(f"📏 Loaded Standards: {standards} with variance ±{variance}")
    return standards, variance


def check_within_standards(x_coord, standards, variance):
    """Check if x_coord is within variance of any standard value"""
    try:
        val = float(x_coord)
    except (ValueError, TypeError):
        return "ERROR"

    for std in standards:
        if std - variance <= val <= std + variance:
            return "YES"
    return "ERROR"


def match_subheads_with_spans(cleaned_file, span_file, output_file, standards, variance):
    """Match cleaned subheads with span contents from Excel"""
    spans_df = pd.read_excel(span_file)
    spans = spans_df.to_dict(orient="records")

    with open(cleaned_file, "r", newline="", encoding="utf-8") as infile, \
         open(output_file, "w", newline="", encoding="utf-8") as outfile:

        reader = csv.DictReader(infile)
        fieldnames = ["Reference", "Subhead", "Match Status", "X-Coord", "Page", "Even/Odd", "Location"]
        writer = csv.DictWriter(outfile, fieldnames=fieldnames)
        writer.writeheader()

        for row in reader:
            subhead = normalize_apostrophes(row["Subhead"])
            match = None

            for span in spans:
                span_text = normalize_apostrophes(str(span.get("Span Content", "")))
                if span_text == subhead:
                    match = span
                    break

            if match:
                page_raw = match.get("Page Number", "")
                try:
                    page_num = int(float(page_raw))  # handles 1, "1", 1.0, "1.0"
                    even_odd = "EVEN" if page_num % 2 == 0 else "ODD"
                except (ValueError, TypeError):
                    page_num = ""
                    even_odd = ""

                x_coord_raw = match.get("Span Position (bbox)", "")
                # extract first value from tuple string
                first_val = ""
                if isinstance(x_coord_raw, str) and x_coord_raw.startswith("("):
                    try:
                        first_val = float(x_coord_raw.strip("()").split(",")[0])
                    except Exception:
                        first_val = ""
                elif isinstance(x_coord_raw, (tuple, list)):
                    first_val = x_coord_raw[0] if x_coord_raw else ""
                else:
                    first_val = x_coord_raw

                location = check_within_standards(first_val, standards, variance)

                writer.writerow({
                    "Reference": row["Reference"],
                    "Subhead": row["Subhead"],  # keep original for clarity
                    "Match Status": "MATCH",
                    "X-Coord": x_coord_raw,
                    "Page": page_num,
                    "Even/Odd": even_odd,
                    "Location": location
                })
            else:
                writer.writerow({
                    "Reference": row["Reference"],
                    "Subhead": row["Subhead"],
                    "Match Status": "COULD NOT MATCH",
                    "X-Coord": "",
                    "Page": "",
                    "Even/Odd": "",
                    "Location": "ERROR"
                })

    print(f"✅ Matching done. Output saved as: {output_file}")


if __name__ == "__main__":
    folder = "."  # current folder
    files = os.listdir(folder)

    # load standards
    standards_file = os.path.join(folder, "Standards.txt")
    if not os.path.exists(standards_file):
        print("❌ Standards.txt not found in folder. Exiting.")
        exit(1)
    standards, variance = load_standards(standards_file)

    # find all subhead CSVs (skip already cleaned or matched files)
    subhead_files = [
        f for f in files
        if f.endswith(".csv")
        and "subhead" in f
        and not f.startswith("matched")
        and not f.endswith("_clean.csv")
    ]

    if not subhead_files:
        print("⚠️ No new raw subhead files found. Nothing to process.")
    else:
        for subhead_file in subhead_files:
            prefix = subhead_file[:3]  # e.g., "01-Genesis"
            span_file_candidates = [f for f in files if f.startswith(prefix) and "_all_spans" in f]

            if not span_file_candidates:
                print(f"❌ No span file found for {subhead_file}, skipping.")
                continue

            span_file = span_file_candidates[0]
            cleaned_file = f"{prefix}_subhead_clean.csv"
            output_file = f"matched_{subhead_file}"

            # Run cleaning and matching
            clean_subhead_file(subhead_file, cleaned_file)
            if os.path.exists(cleaned_file):
                match_subheads_with_spans(cleaned_file, span_file, output_file, standards, variance)
