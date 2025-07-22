import streamlit as st
import zipfile
import os
import math
import subprocess
import ezdxf
from pyproj import Transformer
from xml.etree import ElementTree as ET
from io import BytesIO

st.set_page_config(page_title="KMZ → DXF/DWG Converter", layout="wide")

# ------------------- STYLE -------------------
st.markdown("""
<style>
h1 { text-align: center; color: #4CAF50; }
.stButton>button {
    background-color: #4CAF50; color: white;
    font-size: 16px; padding: 10px 20px;
    border-radius: 8px;
}
</style>
""", unsafe_allow_html=True)

# ------------------- TITLE -------------------
st.title("📐 KMZ → DXF/DWG Converter")
st.markdown("Konversi file **KMZ** menjadi **DXF/DWG** sesuai template AutoCAD.")

# ------------------- UPLOAD SECTION -------------------
col1, col2 = st.columns(2)
with col1:
    kmz_file = st.file_uploader("📂 Upload File KMZ", type=["kmz"])
with col2:
    template_file = st.file_uploader("📂 Upload Template (DXF/DWG)", type=["dxf", "dwg"])

# ------------------- TRANSFORMER -------------------
transformer = Transformer.from_crs("EPSG:4326", "EPSG:32760", always_xy=True)

# ------------------- FUNGSI -------------------
def extract_kmz(kmz_path, extract_dir):
    with zipfile.ZipFile(kmz_path, 'r') as kmz_file:
        kmz_file.extractall(extract_dir)
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
    classified = {"FDT": [], "FAT": [], "POLE": [], "HP_COVER": [], "NEW_POLE": [], "EXISTING_POLE": []}
    for p in points:
        name = p['name'].upper()
        if "FDT" in name:
            classified["FDT"].append(p)
        elif "FAT" in name:
            classified["FAT"].append(p)
        elif "HP" in name or "HOME" in name or "COVER" in name:
            classified["HP_COVER"].append(p)
        elif "NEW POLE" in name:
            classified["NEW_POLE"].append(p)
        elif "EXISTING" in name or "EMR" in name:
            classified["EXISTING_POLE"].append(p)
        elif "P" in name:
            classified["POLE"].append(p)
    return classified

def latlon_to_xy(lat, lon):
    x, y = transformer.transform(lon, lat)
    return x, y

def convert_dwg_to_dxf(dwg_path, output_path):
    try:
        cmd = f"ODAFileConverter '{dwg_path}' '{output_path}' ACAD2010 DXF 1"
        subprocess.run(cmd, shell=True, check=True)
        return True
    except Exception:
        return False

def merge_with_template(template_path):
    return ezdxf.readfile(template_path)

def draw_to_template(doc, classified):
    msp = doc.modelspace()
    for hp in classified["HP_COVER"]:
        x, y = latlon_to_xy(hp["latitude"], hp["longitude"])
        msp.add_text(hp["name"], dxfattribs={"layer": "HP_COVER"}).set_pos((x, y), align='CENTER')
    return doc

# ------------------- PROSES KONVERSI -------------------
if st.button("🚀 Convert"):
    if kmz_file is None:
        st.error("❌ Harap upload file KMZ!")
    elif template_file is None:
        st.error("❌ Harap upload file template DXF/DWG!")
    else:
        with st.spinner("⏳ Sedang memproses..."):
            # Simpan file
            kmz_path = "uploaded.kmz"
            with open(kmz_path, "wb") as f:
                f.write(kmz_file.read())

            template_path = "template_input"
            with open(template_path, "wb") as f:
                f.write(template_file.read())

            # Jika DWG → konversi ke DXF
            if template_file.name.lower().endswith(".dwg"):
                st.warning("⚠ Template DWG terdeteksi. Mengonversi ke DXF...")
                if not convert_dwg_to_dxf(template_path, "./"):
                    st.error("❌ Gagal konversi DWG ke DXF! Pastikan ODA Converter terinstall.")
                    st.stop()
                template_path = "template_input.dxf"

            # Extract & parsing
            extract_dir = "temp_kmz"
            os.makedirs(extract_dir, exist_ok=True)
            kml_path = extract_kmz(kmz_path, extract_dir)
            points = parse_kml(kml_path)
            classified = classify_points(points)

            # Gabungkan dengan template
            doc = merge_with_template(template_path)
            doc = draw_to_template(doc, classified)

            # Simpan DXF
            dxf_buffer = BytesIO()
            doc.saveas(dxf_buffer)
            dxf_buffer.seek(0)

            # Convert DXF → DWG
            dxf_output = "hasil.dxf"
            with open(dxf_output, "wb") as f:
                f.write(dxf_buffer.getvalue())

            dwg_output = "hasil.dwg"
            convert_dwg_to_dxf(dxf_output, "./")  # jika ODA tersedia, akan buat DWG

            st.success("✅ Konversi berhasil!")
            st.download_button("⬇️ Download DXF", data=dxf_buffer, file_name="hasil.dxf", mime="application/dxf")

            if os.path.exists(dwg_output):
                with open(dwg_output, "rb") as f:
                    st.download_button("⬇️ Download DWG", data=f, file_name="hasil.dwg", mime="application/acad")
