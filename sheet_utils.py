# sheet_utils.py

import gspread
from oauth2client.service_account import ServiceAccountCredentials
from datetime import datetime
import streamlit as st

def authenticate_google(secret_dict):
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    credentials = ServiceAccountCredentials.from_json_keyfile_dict(secret_dict, scope)
    client = gspread.authorize(credentials)
    return client

def get_latest_row_data(sheet):
    values = sheet.get_all_values()
    for i in range(len(values) - 1, 0, -1):
        if any(values[i]):
            return values[i]
    return ["" for _ in sheet.row_values(1)]

def parse_date_format(prev_value, today):
    try:
        if "/" in prev_value:
            return today.strftime("%d/%m/%Y")
        elif "-" in prev_value:
            return today.strftime("%Y-%m-%d")
        else:
            return today.strftime("%d-%m-%Y")
    except:
        return today.strftime("%d-%m-%Y")

def extract_distance(name_text):
    import re
    match = re.search(r'(AE[-\s]?)?(\d+)\s?M', name_text.upper())
    if match:
        return int(match.group(2))
    return None

def extract_cores(name_text):
    import re
    match = re.search(r'FO\s*(\d+)/', name_text.upper())
    if match:
        core_count = int(match.group(1))
        if core_count == 24:
            return 2, 24
        elif core_count == 48:
            return 4, 48
        elif core_count == 72:
            return 6, 72
        elif core_count == 96:
            return 8, 96
        elif core_count == 12:
            return 1, 12
    return None, None
