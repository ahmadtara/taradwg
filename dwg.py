import streamlit as st
import zipfile
import os
import math
from xml.etree import ElementTree as ET
import ezdxf
from pyproj import Transformer
from io import BytesIO

# Konfigurasi halaman
st.set_page_config(page_title="KMZ ke DWG Converter", layout="wide")

# Transformer WGS84 -> UTM (zona 60S)
transformer = Transformer.from_crs("EPSG:4326", "EPSG:32760", always_xy=True)

# ---------------- FUNCTIONS ----------------
def extract_kmz(kmz_file, extract_dir):
    with zipfile.ZipFile(kmz_file, 'r') as kmz:
        kmz.extractall(extract_dir)
    return os.path.join(extract_dir, "doc.kml")

def parse_kml(kml_path):
    ns = {'kml': 'http://www.opengis.net/kml/2.2'}
    tree = ET.parse(kml_path)
    root = tree.getroot()
    placemarks = root.findall('.//kml:Placemark', ns)
    points = []
    for pm in placemarks:
        name = pm.find('kml:name', ns)
        coord = pm.find('.//kml:coordinates', ns)
        if name is not None and coord is not None:
            lon, lat, *_ = coord.text.strip().split(',')
            points.append({
                'name': name.text.strip(),
                'latitude': float(lat),
                'longitude': float(lon)
            })
    return points

def classify_points(points):
    classified = {
        "FDT": [], "FAT": [], "POLE": [], "HP_COVER": [],
        "NEW_POLE": [], "EXISTING_POLE": []
    }
    for p in points:
        name = p['name'].upper()
        if "FDT" in name:
            classified["FDT"].append(p)
        elif "FAT" in name:
            classified["FAT"].append(p)
        elif "HP" in name or "HOME" in name:
            classified["HP_COVER"].append(p)
        elif "NEW POLE" in name:
            classified["NEW_POLE"].append(p)
        elif "EXISTING" in name:
            classified["EXISTING_POLE"].append(p)
        elif "P" in name:
            classified["POLE"].append(p)
    return classified

def latlon_to_xy(lat, lon):
    return transformer.transform(lon, lat)

def merge_with_template(template_dwg):
    doc = ezdxf.readfile(template_dwg)
    return doc

def add_points_to_dwg(doc, classified):
    msp = doc.modelspace()
    for layer, points in classified.items():
        for p in points:
            x, y = latlon_to_xy(p['latitude'], p['longitude'])
            msp.add_text(p["name"], dxfattribs={"layer": layer}).set_pos((x, y), align='CENTER')
    return doc

# ---------------- STREAMLIT UI ----------------
st.title("📐 KMZ → DWG Converter")
st.markdown("Konversi file **KMZ** menjadi **DWG** sesuai template AutoCAD.")

col1, col2 = st.columns(2)

with col1:
    kmz_file = st.file_uploader("📂 Upload File KMZ", type=["kmz"])

with col2:
    template_dwg = st.file_uploader("📂 Upload Template DWG", type=["dwg"])

if kmz_file and template_dwg:
    if st.button("🚀 Convert ke DWG"):
        with st.spinner("Sedang memproses... Mohon tunggu"):
            extract_dir = "temp_kmz"
            os.makedirs(extract_dir, exist_ok=True)

            # Simpan KMZ sementara
            kmz_path = os.path.join(extract_dir, "input.kmz")
            with open(kmz_path, "wb") as f:
                f.write(kmz_file.read())

            # Simpan template DWG sementara
            template_path = "template.dwg"
            with open(template_path, "wb") as f:
                f.write(template_dwg.read())

            # Parsing KMZ
            kml_path = extract_kmz(kmz_path, extract_dir)
            points = parse_kml(kml_path)
            classified = classify_points(points)

            # Load template DWG
            doc = merge_with_template(template_path)

            # Tambah titik ke DWG
            doc = add_points_to_dwg(doc, classified)

            # Simpan ke memori untuk download
            output_buffer = BytesIO()
            doc.write(output_buffer)
            output_buffer.seek(0)

            st.success("✅ Konversi selesai!")
            st.download_button(
                "⬇️ Download DWG Hasil",
                data=output_buffer,
                file_name="hasil_konversi.dwg",
                mime="application/octet-stream"
            )
