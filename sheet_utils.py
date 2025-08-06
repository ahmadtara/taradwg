# sheet_utils.py
from datetime import datetime
import re
from math import dist

def get_latest_row_data(sheet):
    values = sheet.get_all_values()
    if len(values) < 2:
        return sheet.row_values(1), [""] * len(sheet.row_values(1))

    for row in reversed(values[1:]):
        if any(cell.strip() for cell in row):
            return sheet.row_values(1), row
    return sheet.row_values(1), [""] * len(sheet.row_values(1))

def parse_date_format(example_date):
    if "/" in example_date:
        return "%d/%m/%Y" if example_date.index("/") == 2 else "%Y/%m/%d"
    if "-" in example_date:
        return "%Y-%m-%d"
    return "%d-%m-%Y"

def extract_distance(path):
    total = 0.0
    for i in range(1, len(path)):
        p1 = path[i-1]
        p2 = path[i]
        total += dist(p1, p2)
    return round(total, 2)

def extract_fo_count(text):
    match = re.search(r"FO\s+(\d+)", text.upper())
    if match:
        return int(match.group(1))
    return None

def extract_tray_count(text):
    match = re.search(r"\/(\d+)T", text.upper())
    if match:
        return int(match.group(1))
    return None

def extract_length(text):
    match = re.search(r"AE[-\s]?(\d+)\s?M", text.upper())
    if match:
        return int(match.group(1))
    return None
