# app.py
import streamlit as st
import tempfile
from datetime import datetime
from fdt_cable_writer import extract_points_from_kmz, extract_paths_from_kmz, authenticate_google, get_latest_row
from sheet_writer import fill_fdt_pekanbaru, fill_cable_pekanbaru, fill_subfeeder_pekanbaru

SPREADSHEET_ID_3 = "1EnteHGDnRhwthlCO9B12zvHUuv3wtq5L2AKlV11qAOU"
SPREADSHEET_ID_4 = "1D_OMm46yr-e80s3sCyvbSSsf8wrUCwpwiYsVBKPgszw"
SPREADSHEET_ID_5 = "1paa8sT3nTZh_xxwHeKV8pwVIWacq7lC8U9A8BlX6LUw"

SHEET_NAME_3 = "FDT Pekanbaru"
SHEET_NAME_4 = "Cable Pekanbaru"
SHEET_NAME_5 = "Sheet1"

st.set_page_config(page_title="Uploader FDT & Cable", layout="centered")
st.title("📡 Uploader FDT & Cable KMZ")

col1, col2, col3 = st.columns(3)
with col1:
    district_input = st.text_input("District (E)")
with col2:
    subdistrict_input = st.text_input("Subdistrict (F)")
with col3:
    vendor_input = st.text_input("Vendor Name (AB)")

uploaded_cluster = st.file_uploader("📤 Upload file .KMZ CLUSTER (berisi FDT & DISTRIBUTION CABLE)", type=["kmz"])
uploaded_subfeeder = st.file_uploader("📤 Upload file .KMZ SUBFEEDER (berisi CABLE)", type=["kmz"])

submit_clicked = st.button("🚀 Submit dan Kirim ke Spreadsheet")

if submit_clicked:
    if not district_input or not subdistrict_input or not vendor_input:
        st.warning("⚠️ Harap isi semua kolom input manual.")
    else:
        client = authenticate_google(st.secrets["gcp_service_account"])

        if uploaded_cluster:
            with tempfile.NamedTemporaryFile(delete=False, suffix=".kmz") as tmp:
                tmp.write(uploaded_cluster.read())
                kmz_path = tmp.name

            kmz_filename = uploaded_cluster.name
            with st.spinner("🔍 Membaca data dari KMZ CLUSTER..."):
                fdt_points, poles_74 = extract_points_from_kmz(kmz_path)
                cable_paths = extract_paths_from_kmz(kmz_path, folder_match="DISTRIBUTION CABLE")

            if fdt_points:
                sheet3 = client.open_by_key(SPREADSHEET_ID_3).worksheet(SHEET_NAME_3)
                count = fill_fdt_pekanbaru(sheet3, fdt_points, poles_74, district_input, subdistrict_input, vendor_input, kmz_filename)
                st.success(f"✅ {count} FDT berhasil dikirim ke Sheet FDT Pekanbaru")

            if cable_paths:
                sheet4 = client.open_by_key(SPREADSHEET_ID_4).worksheet(SHEET_NAME_4)
                count = fill_cable_pekanbaru(sheet4, cable_paths, vendor_input, kmz_filename)
                st.success(f"✅ {count} kabel DISTRIBUTION berhasil dikirim ke Sheet Cable Pekanbaru")

        if uploaded_subfeeder:
            with tempfile.NamedTemporaryFile(delete=False, suffix=".kmz") as tmp:
                tmp.write(uploaded_subfeeder.read())
                kmz_path = tmp.name

            kmz_filename = uploaded_subfeeder.name
            with st.spinner("🔍 Membaca data dari KMZ SUBFEEDER..."):
                sub_paths = extract_paths_from_kmz(kmz_path, folder_match="CABLE")

            if sub_paths:
                sheet5 = client.open_by_key(SPREADSHEET_ID_5).worksheet(SHEET_NAME_5)
                count = fill_subfeeder_pekanbaru(sheet5, sub_paths, vendor_input, kmz_filename)
                st.success(f"✅ {count} kabel SUBFEEDER berhasil dikirim ke Sheet1")
