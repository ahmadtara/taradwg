import streamlit as st
from fastkml import kml
from pyproj import Transformer
import ezdxf
import tempfile
import os

# Folder yang ingin diambil
target_folders = {
    'FDT',
    'NEW POLE 7-3',
    'NEW POLE 7-4',
    'EXISTING POLE EMR 7-4',
    'FAT',
    'HP COVER'
}

# Transformer WGS84 ke UTM zone 60S (EPSG:32760)
transformer = Transformer.from_crs("EPSG:4326", "EPSG:32760", always_xy=True)

def parse_kml(kml_file):
    raw_data = kml_file.read()  # jangan decode utf-8, langsung bytes
    k = kml.KML()
    k.from_string(raw_data)

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

def generate_dxf(data, output_path):
    doc = ezdxf.new(dxfversion='R2010')
    msp = doc.modelspace()

    for item in data:
        x, y = item['easting'], item['northing']
        label = f"{item['placemark']}\n({item['folder']})"
        msp.add_point((x, y))
        msp.add_text(label, dxfattribs={'height': 2.5}).set_pos((x, y + 3), align='CENTER')

    doc.saveas(output_path)

# Streamlit UI
st.title("Konversi KML ke DXF (UTM Zone 60S)")
st.markdown("Unggah file `.kml` untuk diubah menjadi file AutoCAD `.dxf`.")

uploaded_file = st.file_uploader("Pilih file KML", type=['kml'])

if uploaded_file:
    try:
        data = parse_kml(uploaded_file)

        if not data:
            st.warning("Tidak ada Placemark yang ditemukan dalam folder target.")
        else:
            with tempfile.NamedTemporaryFile(delete=False, suffix=".dxf") as tmpfile:
                generate_dxf(data, tmpfile.name)
                with open(tmpfile.name, "rb") as f:
                    st.success("✅ File DXF berhasil dibuat!")
                    st.download_button("⬇️ Download DXF", f, file_name="output.dxf")

    except Exception as e:
        st.error(f"❌ Terjadi error saat memproses file:\n\n{e}")
