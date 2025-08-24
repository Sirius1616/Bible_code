import openpyxl

def get_rgb_from_cell(cell):
    """Extract RGB from Excel cell fill color."""
    fill = cell.fill
    if fill and fill.fgColor and fill.fgColor.type == "rgb":
        rgb = fill.fgColor.rgb
        if rgb:
            return rgb[-6:]  # strip alpha if ARGB
    return None

# Load workbook
wb = openpyxl.load_workbook("02-Exodus.xlsx")
ws = wb.active

color_counts = {"RED": 0, "YELLOW": 0, "PURPLE": 0, "ORANGE": 0}

print("Debugging Excel Cell Colors:\n")

for row in ws.iter_rows():
    for cell in row:
        rgb = get_rgb_from_cell(cell)
        if rgb:
            print(f"Cell {cell.coordinate}: Value={cell.value}, RGB={rgb}")

            # Check against known colors
            if rgb.upper() == "FF0000":
                color_counts["RED"] += 1
            elif rgb.upper() == "FFFF00":
                color_counts["YELLOW"] += 1
            elif rgb.upper() in ("800080", "00800080"):  # purple variants
                color_counts["PURPLE"] += 1
            elif rgb.upper() == "FFA500":  # orange
                color_counts["ORANGE"] += 1

print("\nAnnotation Summary (detected colors):")
for k, v in color_counts.items():
    print(f"{k}: {v}")
