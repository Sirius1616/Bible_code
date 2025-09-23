#!/usr/bin/env python
import pandas as pd
import ast
import re

def load_standards(standards_file):
    """Load acceptable coordinates and variance from the standards file."""
    with open(standards_file, "r") as f:
        lines = f.readlines()

    acceptable_coords = []
    variance = 0.0
    reading_coords = False

    for line in lines:
        line = line.strip()
        if not line:
            continue

        # Start reading acceptable X coords
        if "Acceptable X" in line:
            reading_coords = True
            continue

        # Stop reading coords when variance section starts
        if "Acceptable Variance" in line:
            reading_coords = False
            continue

        # Collect coordinates
        if reading_coords:
            try:
                acceptable_coords.append(float(line))
            except ValueError:
                pass

        # Collect variance
        if line.startswith(".") or re.match(r"^\d+(\.\d+)?$", line):
            if "Variance" not in line and not reading_coords:
                try:
                    variance = float(line)
                except ValueError:
                    pass

    return acceptable_coords, variance


def process_excel(file_name, standards_file):
    # Read Excel file
    df = pd.read_excel(file_name)

    # Extract standards
    acceptable_coords, variance = load_standards(standards_file)

    # Extract first value from tuple string
    def get_first_value(x):
        try:
            tup = ast.literal_eval(x)
            return float(tup[0]) if isinstance(tup, (list, tuple)) else None
        except Exception:
            return None

    df["First_X"] = df["X-Coord"].apply(get_first_value)

    # Check against acceptable coords ± variance
    def check_value(val):
        if val is None:
            return "Error"
        for c in acceptable_coords:
            if abs(val - c) <= variance:
                return "Yes"
        return "Error"

    df["Check_Status"] = df["First_X"].apply(check_value)

    # Save results to a new file so original stays safe
    output_file = file_name.replace(".xlsx", "_checked.xlsx")
    df.to_excel(output_file, index=False)
    print(f"✅ File updated successfully: {output_file}")


if __name__ == "__main__":
    process_excel("test_file.xlsx", "Standards.txt")
