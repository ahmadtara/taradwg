import streamlit as st 
import zipfile
import os
import math
from xml.etree import ElementTree as ET
import ezdxf
from pyproj import Transformer

st.set_page_config(page_title="KMZ → DWG Converter", layout="wide")

transformer = Transformer.from_crs("EPSG:4326", "EPSG:32760", always_xy=True)

# Folder target yang akan diklasifikasikan
target_folders = {
    'FDT', 'FAT', 'HP COVER', 'NEW POLE 7-3', 'NEW POLE 7-4', 'EXISTING POLE EMR 7-4'
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

def parse_boundaries(kml_path):
    ns = {'kml': 'http://www.opengis.net/kml/2.2'}
    with open(kml_path, 'rb') as f:
        tree = ET.parse(f)
    root = tree.getroot()
    boundaries = []
    for folder in root.findall(".//kml:Folder", ns):
        name_tag = folder.find('kml:name', ns)
        if name_tag is not None and name_tag.text.strip().upper() == "BOUNDARY":
            placemarks = folder.findall(".//kml:Placemark", ns)
            for pm in placemarks:
                coords = pm.find('.//kml:coordinates', ns)
                if coords is not None:
                    coord_list = []
                    for pair in coords.text.strip().split():
                        lon, lat, *_ = map(float, pair.split(','))
                        coord_list.append(latlon_to_xy(lat, lon))
                    boundaries.append(coord_list)
    return boundaries

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
        name = p['name'].upper()
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

def draw_to_dxf(classified, boundaries):
    doc = ezdxf.new(dxfversion="R2010")
    msp = doc.modelspace()

    all_points_xy = []
    for category in classified.values():
        for p in category:
            all_points_xy.append(latlon_to_xy(p['latitude'], p['longitude']))

    if len(all_points_xy) == 0:
        st.error("❌ Tidak ada titik ditemukan di KMZ!")
        return None

    shifted_points, (cx, cy) = apply_offset(all_points_xy)

    idx = 0
    for category_name, category in classified.items():
        for i in range(len(category)):
            category[i]['xy'] = shifted_points[idx]
            idx += 1

    for layer_name, data in classified.items():
        if layer_name not in doc.layers:
            doc.layers.add(name=layer_name)
        for obj in data:
            x, y = obj['xy']
            msp.add_circle((x, y), radius=2, dxfattribs={"layer": layer_name})
            msp.add_text(obj["name"], dxfattribs={
                "height": 1.5,
                "layer": layer_name,
                "insert": (x + 2, y)  # ← posisi teks langsung di sini
            })

    for polygon in boundaries:
        shifted_polygon, _ = apply_offset(polygon)
        msp.add_lwpolyline(shifted_polygon, close=True, dxfattribs={"layer": "BOUNDARY"})

    return doc

# Streamlit UI
st.title("🏗️ KMZ → DXF Converter (No Template)")
st.write("Konversi file KMZ menjadi file DXF langsung tanpa upload template.")

uploaded_kmz = st.file_uploader("📂 Upload File KMZ", type=["kmz"])

if uploaded_kmz:
    extract_dir = "temp_kmz"
    os.makedirs(extract_dir, exist_ok=True)
    output_dxf = "converted_output.dxf"

    with st.spinner("🔍 Memproses data..."):
        kml_path = extract_kmz(uploaded_kmz, extract_dir)
        points = parse_kml(kml_path)
        classified = classify_points(points)
        boundaries = parse_boundaries(kml_path)

        updated_doc = draw_to_dxf(classified, boundaries)
        if updated_doc:
            updated_doc.saveas(output_dxf)

    if os.path.exists(output_dxf):
        st.success("✅ Konversi berhasil! DXF sudah dibuat.")
        with open(output_dxf, "rb") as f:
            st.download_button("⬇️ Download DXF", f, file_name="output_from_kmz.dxf")

        st.markdown("### 📊 Ringkasan Objek")
        for layer_name, objs in classified.items():
            st.write(f"- **{layer_name}**: {len(objs)} titik")
