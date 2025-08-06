import zipfile
import xml.etree.ElementTree as ET
from shapely.geometry import LineString


def extract_kml_from_kmz(kmz_file):
    with zipfile.ZipFile(kmz_file, 'r') as zf:
        kml_filename = next(f for f in zf.namelist() if f.endswith('.kml'))
        return zf.read(kml_filename)


def parse_fdt_placemarks(kml_root):
    ns = {'kml': 'http://www.opengis.net/kml/2.2'}
    fdt_points = []

    for folder in kml_root.findall(".//kml:Folder", ns):
        name_el = folder.find("kml:name", ns)
        if name_el is not None and "FDT" in name_el.text.upper():
            for pm in folder.findall("kml:Placemark", ns):
                name = pm.find("kml:name", ns)
                desc = pm.find("kml:description", ns)
                coord = pm.find(".//kml:coordinates", ns)
                if name is not None and coord is not None:
                    lon, lat = map(float, coord.text.strip().split(",")[:2])
                    fdt_points.append({
                        "name": name.text.strip(),
                        "description": desc.text.strip() if desc is not None else "",
                        "lat": lat,
                        "lon": lon
                    })
    return fdt_points


def parse_poles_by_folder(kml_root, folder_name):
    ns = {'kml': 'http://www.opengis.net/kml/2.2'}
    poles = []

    for folder in kml_root.findall(".//kml:Folder", ns):
        name_el = folder.find("kml:name", ns)
        if name_el is not None and name_el.text.upper().strip() == folder_name.upper():
            for pm in folder.findall("kml:Placemark", ns):
                name = pm.find("kml:name", ns)
                coord = pm.find(".//kml:coordinates", ns)
                if name is not None and coord is not None:
                    lon, lat = map(float, coord.text.strip().split(",")[:2])
                    poles.append({
                        "name": name.text.strip(),
                        "lat": lat,
                        "lon": lon
                    })
    return poles


def parse_cables_by_folder(kml_root, folder_name):
    ns = {'kml': 'http://www.opengis.net/kml/2.2'}
    cables = []

    for folder in kml_root.findall(".//kml:Folder", ns):
        name_el = folder.find("kml:name", ns)
        if name_el is not None and name_el.text.upper().strip() == folder_name.upper():
            for pm in folder.findall("kml:Placemark", ns):
                name = pm.find("kml:name", ns)
                coord = pm.find(".//kml:coordinates", ns)
                if name is not None and coord is not None:
                    coords = [list(map(float, p.split(",")[:2])) for p in coord.text.strip().split() if "," in p]
                    try:
                        length = LineString(coords).length * 111139  # degree to meter approx
                    except Exception:
                        length = 0
                    cables.append({
                        "name": name.text.strip(),
                        "length": round(length, 2)
                    })
    return cables
