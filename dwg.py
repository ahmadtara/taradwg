import streamlit as st
import zipfile
import xml.etree.ElementTree as ET
from io import BytesIO
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import tempfile
from datetime import datetime
from math import dist
import re

SPREADSHEET_ID_3 = "1EnteHGDnRhwthlCO9B12zvHUuv3wtq5L2AKlV11qAOU"
SHEET_NAME_3 = "FDT Pekanbaru"

SPREADSHEET_ID_4 = "1D_OMm46yr-e80s3sCyvbSSsf8wrUCwpwiYsVBKPgszw"
SHEET_NAME_4 = "Cable Pekanbaru"

SPREADSHEET_ID_5 = "1paa8sT3nTZh_xxwHeKV8pwVIWacq7lC8U9A8BlX6LUw"
SHEET_NAME_5 = "Sheet1"

def authenticate_google():
    creds_dict = st.secrets["gcp_service_account"]
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    credentials = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
    client = gspread.authorize(credentials)
    return client

def extract_data_from_kmz(kmz_file):
    def parse_folder(folder_element, ns, folder_path=""):
        folder_name_el = folder_element.find("kml:name", ns)
        folder_name = folder_name_el.text.strip() if folder_name_el is not None else "Unknown"
        full_folder_path = f"{folder_path}/{folder_name}" if folder_path else folder_name

        placemarks = []
        for placemark in folder_element.findall("kml:Placemark", ns):
            name = placemark.find("kml:name", ns)
            name = name.text.strip() if name is not None else ""

            coords_tag = placemark.find(".//kml:coordinates", ns)
            coords = coords_tag.text.strip() if coords_tag is not None else ""
            coords = coords.split(",") if coords else ["", "", ""]

            description_tag = placemark.find("kml:description", ns)
            description = description_tag.text.strip() if description_tag is not None else ""

            placemarks.append({
                'name': name,
                'lon': coords[0],
                'lat': coords[1],
                'alt': coords[2] if len(coords) > 2 else "",
                'description': description,
                'folder': full_folder_path.split("/")[-1],
                'full_path': full_folder_path
            })

        for subfolder in folder_element.findall("kml:Folder", ns):
            placemarks.extend(parse_folder(subfolder, ns, full_folder_path))

        return placemarks

    folders = {}
    poles_7_4 = []
    new_pole_found = False

    with zipfile.ZipFile(kmz_file, 'r') as z:
        kml_filename = [f for f in z.namelist() if f.endswith('.kml')][0]
        with z.open(kml_filename) as kml_file:
            tree = ET.parse(kml_file)
            root = tree.getroot()

            ns = {'kml': 'http://www.opengis.net/kml/2.2'}

            for folder in root.findall(".//kml:Folder", ns):
                placemarks = parse_folder(folder, ns)
                for pm in placemarks:
                    folder_key = pm['folder']
                    folders.setdefault(folder_key, []).append(pm)

                    if 'NEW POLE' in folder_key.upper():
                        new_pole_found = True
                    if 'NEW POLE 7-4' in folder_key.upper():
                        poles_7_4.append(pm)

    if not new_pole_found:
        st.warning("⚠️ Tidak ditemukan titik tiang di folder yang diawali dengan 'NEW POLE' dalam KMZ.")
    else:
        st.success("✅ Titik tiang dari folder 'NEW POLE' berhasil ditemukan dalam KMZ.")

    return folders, poles_7_4
