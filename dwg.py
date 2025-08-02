import streamlit as st
import zipfile
import os
from xml.etree import ElementTree as ET
import ezdxf
from pyproj import Transformer

st.set_page_config(page_title="KMZ → DXF Converter with Matchprop", layout="wide")

transformer = Transformer.from_crs("EPSG:4326", "EPSG:32760", always_xy=True)

target_folders = {
    'FDT', 'FAT', 'HP COVER', 'NEW POLE 7-3', 'NEW POLE 7-4', 'EXISTING POLE EMR 7-4', 'EXISTING POLE EMR 7-3'
}

def extract_kmz(kmz_path, extract_dir):
    with zipfile.ZipFile(kmz_path, 'r') as kmz_file:
        kmz_file.extractall(extract_dir)
    return os.path.join(extract_dir, "doc.kml")

def parse_kml(kml_path):
    ns = {'kml': 'http://www.opengis.net/kml/2.2'}
    with open(kml_path, 'rb') as f:
        tree = ET.parse(f)
    root = tree.getroot()
    folders = root.findall('.//kml:Folder', ns)
    points = []
    for folder in folders:
        folder_name_tag = folder.find('kml:name', ns)
        if folder_name_tag is None:
            continue
        folder_name = folder_name_tag.text.strip().upper()
        if folder_name not in target_folders:
            continue
        placemarks = folder.findall('.//kml:Placemark', ns)
        for pm in placemarks:
            name = pm.find('kml:name', ns)
            coord = pm.find('.//kml:coordinates', ns)
            if name is not None and coord is not None:
                name_text = name.text.strip()
                lon, lat, *_ = coord.text.strip().split(',')
                points.append({'name': name_text, 'latitude': float(lat), 'longitude': float(lon), 'folder': folder_name})
    return points

def latlon_to_xy(lat, lon):
    x, y = transformer.transform(lon, lat)
    return x, y

def apply_offset(points_xy):
    xs = [x for x, y in points_xy]
    ys = [y for x, y in points_xy]
    cx, cy = sum(xs)/len(xs), sum(ys)/len(ys)
    return [(x - cx, y - cy) for x, y in points_xy], (cx, cy)

def classify_points(points):
    classified = {
        "FDT": [], "FAT": [], "HP_COVER": [], "NEW_POLE": [], "EXISTING_POLE": [], "POLE": []
    }
    for p in points:
        folder = p['folder']
        if "FDT" in folder:
            classified["FDT"].append(p)
        elif "FAT" in folder:
            classified["FAT"].append(p)
        elif "HP COVER" in folder:
            classified["HP_COVER"].append(p)
        elif "NEW POLE" in folder:
            classified["NEW_POLE"].append(p)
        elif "EXISTING" in folder or "EMR" in folder:
            classified["EXISTING_POLE"].append(p)
        else:
            classified["POLE"].append(p)
    return classified

def draw_to_dxf(classified, template_path):
    template_doc = ezdxf.readfile(template_path)
    template_msp = template_doc.modelspace()

    doc = ezdxf.new(dxfversion="R2010")
    msp = doc.modelspace()

    # Copy all block definitions from template to output
    for block_name in template_doc.blocks.block_names():
        if block_name in doc.blocks:
            continue
        block_def = template_doc.blocks[block_name]
        new_block = doc.blocks.new(name=block_name)
        for entity in block_def:
            new_block.add_entity(entity.copy())

    # Find matchprop examples
    matchprop_hp = matchprop_pole = matchprop_sr = None
    for e in template_msp.query('TEXT'):
        txt = e.dxf.text.upper()
        if 'NN-' in txt:
            matchprop_hp = e.dxf
        elif 'MR.SRMRW16' in txt:
            matchprop_pole = e.dxf
        elif 'SRMRW16.067.B01' in txt:
            matchprop_sr = e.dxf

    all_points_xy = [latlon_to_xy(p['latitude'], p['longitude']) for cat in classified.values() for p in cat]
    if not all_points_xy:
        st.error("❌ Tidak ada titik ditemukan di KMZ!")
        return None

    shifted_points, (cx, cy) = apply_offset(all_points_xy)
    idx = 0
    for category in classified.values():
        for p in category:
            p['xy'] = shifted_points[idx]
            idx += 1

    for layer_name, data in classified.items():
        if layer_name not in doc.layers:
            doc.layers.add(name=layer_name)
        for obj in data:
            x, y = obj['xy']
            if layer_name != "HP_COVER":
                msp.add_circle((x, y), radius=2, dxfattribs={"layer": layer_name})

            matchprop = {
                "HP_COVER": matchprop_hp,
                "NEW_POLE": matchprop_pole,
                "EXISTING_POLE": matchprop_pole,
                "FAT": matchprop_sr,
                "FDT": matchprop_sr
            }.get(layer_name, None)

            attribs = {
                "height": matchprop.height if matchprop else 1.5,
                "layer": layer_name,
                "insert": (x + 2, y)
            }
            if matchprop:
                attribs["color"] = matchprop.color

            msp.add_text(obj["name"], dxfattribs=attribs)

            # Optional: insert block if exists
            if layer_name in ["NEW_POLE", "EXISTING_POLE"] and "NW" in doc.blocks:
                msp.add_blockref("NW", (x, y), dxfattribs={"layer": layer_name})

    return doc
