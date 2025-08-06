# fdt_cable_writer.py

import zipfile
import xml.etree.ElementTree as ET
from math import radians, cos, sin, sqrt, atan2
from sheet_utils import get_latest_row_data, parse_date_format, extract_distance

def authenticate_google():
    import streamlit as st
    import gspread
    from oauth2client.service_account import ServiceAccountCredentials

    creds_dict = st.secrets["gcp_service_account"]
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    credentials = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
    client = gspread.authorize(credentials)
    return client

def haversine_distance(lat1, lon1, lat2, lon2):
    R = 6371000
    phi1, phi2 = radians(lat1), radians(lat2)
    delta_phi = radians(lat2 - lat1)
    delta_lambda = radians(lon2 - lon1)
    a = sin(delta_phi / 2)**2 + cos(phi1) * cos(phi2) * sin(delta_lambda / 2)**2
    c = 2 * atan2(sqrt(a), sqrt(1 - a))
    return R * c

def extract_points_from_kmz(kmz_path):
    points = []
    poles_74 = []
    ns = {"kml": "http://www.opengis.net/kml/2.2"}

    with zipfile.ZipFile(kmz_path, 'r') as zf:
        kml_file = next((f for f in zf.namelist() if f.lower().endswith(".kml")), None)
        if not kml_file:
            return [], []

        tree = ET.parse(zf.open(kml_file))
        root = tree.getroot()

        for folder in root.findall(".//kml:Folder", ns):
            folder_name_el = folder.find("kml:name", ns)
            folder_name = folder_name_el.text.strip().upper() if folder_name_el is not None else ""
            for pm in folder.findall("kml:Placemark", ns):
                name_el = pm.find("kml:name", ns)
                desc_el = pm.find("kml:description", ns)
                coord_el = pm.find(".//kml:coordinates", ns)

                if name_el is None or coord_el is None:
                    continue

                name = name_el.text.strip()
                description = desc_el.text.strip() if desc_el is not None and desc_el.text else ""
                coords = coord_el.text.strip().split(",")
                lon, lat = float(coords[0]), float(coords[1])

                item = {
                    "name": name,
                    "lat": lat,
                    "lon": lon,
                    "description": description
                }

                if folder_name == "FDT":
                    points.append(item)
                elif folder_name == "NEW POLE 7-4":
                    poles_74.append(item)

    return points, poles_74

def extract_paths_from_kmz(kmz_path, folder_match="DISTRIBUTION CABLE"):
    path_items = []
    ns = {"kml": "http://www.opengis.net/kml/2.2"}

    with zipfile.ZipFile(kmz_path, 'r') as zf:
        kml_file = next((f for f in zf.namelist() if f.lower().endswith(".kml")), None)
        if not kml_file:
            return []

        tree = ET.parse(zf.open(kml_file))
        root = tree.getroot()

        for folder in root.findall(".//kml:Folder", ns):
            folder_name_el = folder.find("kml:name", ns)
            folder_name = folder_name_el.text.strip().upper() if folder_name_el is not None else ""
            if folder_match not in folder_name:
                continue

            for pm in folder.findall("kml:Placemark", ns):
                name_el = pm.find("kml:name", ns)
                coord_el = pm.find(".//kml:coordinates", ns)
                if name_el is None or coord_el is None:
                    continue

                name = name_el.text.strip()
                coord_pairs = coord_el.text.strip().split()
                coords = [tuple(map(float, p.split(",")[:2])) for p in coord_pairs if ',' in p]

                total_length = 0
                for i in range(len(coords)-1):
                    lat1, lon1 = coords[i][1], coords[i][0]
                    lat2, lon2 = coords[i+1][1], coords[i+1][0]
                    total_length += haversine_distance(lat1, lon1, lat2, lon2)

                path_items.append({
                    "name": name,
                    "length_m": round(total_length),
                    "coords": coords,
                    "folder": folder_name
                })

    return path_items
