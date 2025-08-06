import zipfile
import xml.etree.ElementTree as ET
from io import BytesIO
from datetime import datetime
import streamlit as st
import gspread
from oauth2client.service_account import ServiceAccountCredentials


def authenticate_google(secret_dict):
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    credentials = ServiceAccountCredentials.from_json_keyfile_dict(secret_dict, scope)
    client = gspread.authorize(credentials)
    return client

def extract_points_from_kmz(kmz_path):
    fdt_points, poles_74 = [], []

    def recurse_folder(folder, ns, path=""):
        items = []
        name_el = folder.find("kml:name", ns)
        folder_name = name_el.text.upper() if name_el is not None else "UNKNOWN"
        new_path = f"{path}/{folder_name}" if path else folder_name
        for sub in folder.findall("kml:Folder", ns):
            items += recurse_folder(sub, ns, new_path)
        for pm in folder.findall("kml:Placemark", ns):
            name_el = pm.find("kml:name", ns)
            desc_el = pm.find("kml:description", ns)
            coord_el = pm.find(".//kml:coordinates", ns)
            if name_el is not None and coord_el is not None:
                lon, lat = map(float, coord_el.text.strip().split(",")[:2])
                items.append({
                    "name": name_el.text.strip(),
                    "lat": lat,
                    "lon": lon,
                    "path": new_path,
                    "description": desc_el.text.strip() if desc_el is not None else ""
                })
        return items

    with zipfile.ZipFile(kmz_path, 'r') as zf:
        kml_file = next((f for f in zf.namelist() if f.lower().endswith(".kml")), None)
        if not kml_file:
            return [], []

        root = ET.parse(zf.open(kml_file)).getroot()
        ns = {"kml": "http://www.opengis.net/kml/2.2"}
        all_points = []
        for folder in root.findall(".//kml:Folder", ns):
            all_points += recurse_folder(folder, ns)

    for pt in all_points:
        base = pt["path"].split("/")[0].upper()
        if base == "FDT":
            fdt_points.append(pt)
        elif base == "NEW POLE 7-4":
            poles_74.append(pt)

    return fdt_points, poles_74

def extract_paths_from_kmz(kmz_path, folder_match="DISTRIBUTION CABLE"):
    paths = []

    def recurse_folder(folder, ns, path=""):
        items = []
        name_el = folder.find("kml:name", ns)
        folder_name = name_el.text.upper() if name_el is not None else "UNKNOWN"
        new_path = f"{path}/{folder_name}" if path else folder_name
        for sub in folder.findall("kml:Folder", ns):
            items += recurse_folder(sub, ns, new_path)
        for pm in folder.findall("kml:Placemark", ns):
            name_el = pm.find("kml:name", ns)
            coord_els = pm.findall(".//kml:coordinates", ns)
            if name_el is not None and coord_els:
                coords = []
                for coord_el in coord_els:
                    text = coord_el.text.strip()
                    for line in text.split():
                        parts = line.strip().split(',')
                        if len(parts) >= 2:
                            lon, lat = map(float, parts[:2])
                            coords.append((lat, lon))
                if len(coords) >= 2:
                    items.append({
                        "name": name_el.text.strip(),
                        "coords": coords,
                        "path": new_path
                    })
        return items

    with zipfile.ZipFile(kmz_path, 'r') as zf:
        kml_file = next((f for f in zf.namelist() if f.lower().endswith(".kml")), None)
        if not kml_file:
            return []

        root = ET.parse(zf.open(kml_file)).getroot()
        ns = {"kml": "http://www.opengis.net/kml/2.2"}
        all_paths = []
        for folder in root.findall(".//kml:Folder", ns):
            all_paths += recurse_folder(folder, ns)

    filtered = [p for p in all_paths if folder_match.upper() in p["path"]]
    return filtered
