#!/usr/bin/env python

import csv
import os
import pandas as pd
import ast
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
    """Load acceptable coordinates and variance from Standards.txt"""
    with open(standards_file, "r") as f:
        lines = f.readlines()

    acceptable_coords = []
    variance = 0.0
    reading_coords = False

    for line in lines:
        line = line.strip()
        if not line:
            continue

        if "Acceptable X" in line:
            reading_coords = True
            continue

        if "Acceptable Variance" in line:
            reading_coords = False
            continue

        # Collect coordinates
        if reading_coords:
            try:
                acceptable_coords.append(float(line))
            except ValueError:
                pass

        # Collect variance (line just below variance header)
        if not reading_coords and re.match(r"^[\d\.\-]+$", line):
            try:
                variance = float(line)
            except ValueError:
                pass

    return acceptable_coords, variance


def check_location(x_coord, acceptable_coords, variance):
    """Return YES if within acceptable ± variance, else ERROR"""
    if not x_coord:
        return "ERROR"
    try:
        tup = ast.literal_eval(str(x_coord))
        first_val = float(tup[0]) if isinstance(tup, (tuple, list)) else None
    except Exception:
        return "ERROR"

    if first_val is None:
        return "ERROR"

    for c in acceptable_coords:
        if abs(first_val - c) <= variance:
            return "YES"
    return "ERROR"


def match_subheads_with_spans(cleaned_file, span_file, output_file, standards_file):
    """Match cleaned subheads with span contents from Excel, then check standards"""
    spans_df = pd.read_excel(span_file)
    spans = spans_df.to_dict(orient="records")

    # Load standards once
    acceptable_coords, variance = load_standards(standards_file)

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
                    page_num = int(float(page_raw))
                    even_odd = "EVEN" if page_num % 2 == 0 else "ODD"
                except (ValueError, TypeError):
                    page_num = ""
                    even_odd = ""

                x_coord = match.get("Span Position (bbox)", "")
                location_status = check_location(x_coord, acceptable_coords, variance)

                writer.writerow({
                    "Reference": row["Reference"],
                    "Subhead": row["Subhead"],
                    "Match Status": "MATCH",
                    "X-Coord": x_coord,
                    "Page": page_num,
                    "Even/Odd": even_odd,
                    "Location": location_status
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

    print(f"✅ Matching + standards check done. Output saved as: {output_file}")


if __name__ == "__main__":
    folder = "."  # current folder
    files = os.listdir(folder)

    standards_file = "Standards.txt"
    if not os.path.exists(standards_file):
        print("❌ Standards.txt not found. Exiting.")
        exit(1)

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
            prefix = subhead_file[:3]
            span_file_candidates = [f for f in files if f.startswith(prefix) and "_all_spans" in f]

            if not span_file_candidates:
                print(f"❌ No span file found for {subhead_file}, skipping.")
                continue

            span_file = span_file_candidates[0]
            cleaned_file = f"{prefix}_subhead_clean.csv"
            output_file = f"matched_{subhead_file}"

            clean_subhead_file(subhead_file, cleaned_file)
            if os.path.exists(cleaned_file):
                match_subheads_with_spans(cleaned_file, span_file, output_file, standards_file)
