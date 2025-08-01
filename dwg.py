import streamlit as st
from fastkml import kml
from pyproj import Transformer
import ezdxf
import tempfile
import os

st.title("Konversi KML ke DXF (UTM 60S)")

uploaded_file = st.file_uploader("Unggah file .kml", type=["kml"])

target_folders = {
    'FDT',
    'NEW POLE 7-3',
    'NEW POLE 7-4',
    'EXISTING POLE EMR 7-4',
    'FAT',
    'HP COVER'
}

transformer = Transformer.from_crs("EPSG:4326", "EPSG:32760", always_xy=True)

def parse_kml_from_temp_file(temp_path):
    k = kml.KML()
    with open(temp_path, 'rb') as f:
        k.from_string(f.read())  # read as bytes => aman dari error XML encoding
    points = []

    def extract(features):
        for f in features:
            if hasattr(f, 'features'):
                fname = f.name.strip() if f.name else ""
                if fname in target_folders:
                    for placemark in f.features():
                        if hasattr(placemark, 'geometry') and placemark.geometry.geom_type == 'Point':
                            lon, lat = placemark.geometry.x, placemark.geometry.y
                            easting, northing = transformer.transform(lon, lat)
                            points.append({
                                'folder': fname,
                                'placemark': placemark.name,
                                'easting': round(easting, 3),
                                'northing': round(northing, 3)
                            })
                extract(f.features())
    extract(k.features())
    return points

def generate_dxf(data):
    doc = ezdxf.new(dxfversion='R2010')
    msp = doc.modelspace()

    for item in data:
        x, y = item['easting'], item['northing']
        label = f"{item['placemark']}\n({item['folder']})"
        msp.add_point((x, y))
        msp.add_text(label, dxfattribs={'height': 2.5}).set_pos((x, y + 3), align='CENTER')

    temp_dxf = tempfile.NamedTemporaryFile(delete=False, suffix=".dxf")
    doc.saveas(temp_dxf.name)
    return temp_dxf.name

if uploaded_file:
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".kml") as tmp:
            tmp.write(uploaded_file.read())
            temp_path = tmp.name

        data = parse_kml_from_temp_file(temp_path)
        os.unlink(temp_path)  # hapus setelah dipakai

        if not data:
            st.warning("⚠️ Tidak ditemukan placemark dalam folder target.")
        else:
            output_path = generate_dxf(data)
            with open(output_path, "rb") as f:
                st.success("✅ File DXF berhasil dibuat!")
                st.download_button("📥 Download DXF", f, file_name="output.dxf")
            os.unlink(output_path)

    except Exception as e:
        st.error(f"❌ Terjadi error saat memproses file:\n\n{e}")
