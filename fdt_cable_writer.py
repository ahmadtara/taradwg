import zipfile
import xml.etree.ElementTree as ET
from datetime import datetime
import math
import os
import re

def extract_fdt_and_poles(kmz_path):
    fdt_points, poles_74 = [], []

    def parse_description(desc):
        return desc.text.strip() if desc is not None else ""

    def recurse_folder(folder, ns, path=""):
        items = []
        folder_name = folder.find("kml:name", ns)
        current_path = f"{path}/{folder_name.text.strip()}" if folder_name is not None else path

        for sub in folder.findall("kml:Folder", ns):
            items += recurse_folder(sub, ns, current_path)

        for pm in folder.findall("kml:Placemark", ns):
            name = pm.find("kml:name", ns)
            desc = pm.find("kml:description", ns)
            coord = pm.find(".//kml:coordinates", ns)
            if name is not None and coord is not None and "," in coord.text:
                lon, lat = coord.text.strip().split(",")[:2]
                items.append({
                    "name": name.text.strip(),
                    "desc": parse_description(desc),
                    "lat": float(lat),
                    "lon": float(lon),
                    "path": current_path.upper()
                })
        return items

    with zipfile.ZipFile(kmz_path, 'r') as zf:
        kml_file = next((f for f in zf.namelist() if f.lower().endswith(".kml")), None)
        if not kml_file:
            raise ValueError("No .kml file found inside the .kmz archive.")

        root = ET.parse(zf.open(kml_file)).getroot()
        ns = {"kml": "http://www.opengis.net/kml/2.2"}

        all_pm = []
        for folder in root.findall(".//kml:Folder", ns):
            all_pm += recurse_folder(folder, ns)

    for item in all_pm:
        if "/FDT" in item["path"]:
            fdt_points.append(item)
        elif "/NEW POLE 7-4" in item["path"]:
            poles_74.append(item)

    return fdt_points, poles_74

def extract_distribution_paths(kmz_path, folder_name="DISTRIBUTION CABLE"):
    paths = []

    def recurse_folder(folder, ns):
        items = []
        for sub in folder.findall("kml:Folder", ns):
            items += recurse_folder(sub, ns)
        for pm in folder.findall("kml:Placemark", ns):
            name = pm.find("kml:name", ns)
            coord = pm.find(".//kml:coordinates", ns)
            if name is not None and coord is not None:
                coord_list = coord.text.strip().split()
                coords = [tuple(map(float, c.split(",")[:2])) for c in coord_list]
                items.append({"name": name.text.strip(), "coords": coords})
        return items

    with zipfile.ZipFile(kmz_path, 'r') as zf:
        kml_file = next((f for f in zf.namelist() if f.lower().endswith(".kml")), None)
        if not kml_file:
            raise ValueError("No .kml file found inside the .kmz archive.")

        root = ET.parse(zf.open(kml_file)).getroot()
        ns = {"kml": "http://www.opengis.net/kml/2.2"}

        for folder in root.findall(".//kml:Folder", ns):
            name = folder.find("kml:name", ns)
            if name is not None and folder_name in name.text.upper():
                return recurse_folder(folder, ns)

    return []

def extract_subfeeder_paths(kmz_path, folder_name="CABLE"):
    return extract_distribution_paths(kmz_path, folder_name)

def find_nearest_pole(fdt_point, poles):
    min_dist = float('inf')
    nearest_name = ""
    for pole in poles:
        d = math.dist([fdt_point['lat'], fdt_point['lon']], [pole['lat'], pole['lon']])
        if d < min_dist:
            min_dist = d
            nearest_name = pole['name']
    return nearest_name

def parse_fo_capacity(text):
    match = re.search(r"FO\s*(\d+)/(\d+)T", text.upper())
    if match:
        return int(match.group(2)), int(match.group(1))
    return None, None

def parse_distance(text):
    match = re.search(r"AE\s*-?\s*(\d+)\s*M", text.upper())
    if match:
        return int(match.group(1))
    return None

def measure_path_length(coords):
    total = 0
    for i in range(len(coords)-1):
        total += math.dist(coords[i], coords[i+1]) * 111139  # convert degrees to meters approx
    return round(total, 2)
