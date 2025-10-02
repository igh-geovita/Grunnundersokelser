import streamlit as st
import tempfile
import os
import pandas as pd
from plot_pdf import (
    build_konus_sensitivity_series,
    export_sensitivity_pdf,
    export_curfc_pdf,
    export_cu_enaks_konus_pdf,
    build_enaks_deformation_series,
    export_enaks_deformation_pdf,
    read_series
)

# ✅ Always use repo logo
logo_path = os.path.join(os.path.dirname(__file__), "geovitalogo.png")

st.title("Geovita – Konus & Enaks Report Generator")

# Sidebar metadata
st.sidebar.header("Report Metadata")
rapport_nr = st.sidebar.text_input("Rapport Nr.", "SMS-20-A-11341")
dato       = st.sidebar.date_input("Dato")
tegn       = st.sidebar.text_input("Tegn", "IGH")
kontr      = st.sidebar.text_input("Kontr", "JOG")
godkj      = st.sidebar.text_input("Godkj", "AGR")

title_info = {
    "rapport_nr": rapport_nr,
    "dato": str(dato),
    "tegn": tegn,
    "kontr": kontr,
    "godkj": godkj,
}

# Upload files
terrain_file = st.file_uploader("Upload terrain_levels.xlsx", type=["xlsx"])
konus_files = st.file_uploader("Upload Konus Excel files", type=["xlsx","xlsm"], accept_multiple_files=True)
enaks_files = st.file_uploader("Upload Enaks Excel files", type=["xlsx","xlsm"], accept_multiple_files=True)

# Input ranges
sheet_name = "Sheet 001"
x_range_konus_undist  = 'L6:L30'
x_range_konus_remould = 'M6:M30'
y_range_depth         = 'F6:F30'
x_range_enaks_strength = 'G6:G30'
y_range_enaks_depth    = 'F6:F30'
x_range_enaks_deform   = 'H6:H30'

if st.button("Generate Reports"):
    if not terrain_file or not konus_files or not enaks_files:
        st.error("Please upload terrain, Konus, and Enaks files.")
    else:
        with tempfile.TemporaryDirectory() as tmpdir:
            # Save terrain
            terrain_path = os.path.join(tmpdir, "terrain.xlsx")
            with open(terrain_path, "wb") as f: f.write(terrain_file.getbuffer())

            # Save Konus files
            konus_dir = os.path.join(tmpdir, "konus"); os.makedirs(konus_dir, exist_ok=True)
            for uf in konus_files:
                with open(os.path.join(konus_dir, uf.name), "wb") as f: f.write(uf.getbuffer())

            # Save Enaks files
            enaks_dir = os.path.join(tmpdir, "enaks"); os.makedirs(enaks_dir, exist_ok=True)
            for uf in enaks_files:
                with open(os.path.join(enaks_dir, uf.name), "wb") as f: f.write(uf.getbuffer())

            # --- Run exports (call your functions from plot_pdf) ---

            # C2 – Sensitivity
            series_sens = build_konus_sensitivity_series(konus_dir, x_range_konus_undist, x_range_konus_remould, y_range_depth)
            out_c2_pdf = os.path.join(tmpdir, "C2_sensitivity.pdf")
            export_sensitivity_pdf(series_sens, outfile_pdf=out_c2_pdf, logo_path=logo_path, title_info={**title_info,"figur_nr":"C2"})
            with open(out_c2_pdf, "rb") as f:
                st.download_button("Download C2 – Sensitivity", f, file_name="C2_sensitivity.pdf")

            # C3 – Remoulded shear strength
            series_remould = []
            read_series(konus_dir, x_range_konus_remould, y_range_depth, series_remould)
            out_c3_pdf = os.path.join(tmpdir, "C3_curfc.pdf")
            export_curfc_pdf(series_remould, outfile_pdf=out_c3_pdf, logo_path=logo_path, title_info={**title_info,"figur_nr":"C3"})
            with open(out_c3_pdf, "rb") as f:
                st.download_button("Download C3 – Remoulded Strength", f, file_name="C3_curfc.pdf")

            # C4 – Konus (undisturbed) + Enaks
            series_konus = []
            read_series(konus_dir, x_range_konus_undist, y_range_depth, series_konus)
            series_enaks = []
            read_series(enaks_dir, x_range_enaks_strength, y_range_enaks_depth, series_enaks)
            out_c4_pdf = os.path.join(tmpdir, "C4_cu_enaks_konus.pdf")
            export_cu_enaks_konus_pdf(series_konus, series_enaks, outfile_pdf=out_c4_pdf, logo_path=logo_path, title_info={**title_info,"figur_nr":"C4"})
            with open(out_c4_pdf, "rb") as f:
                st.download_button("Download C4 – Konus + Enaks", f, file_name="C4_cu_enaks_konus.pdf")

            # C5 – Enaks deformation
            series_deform = build_enaks_deformation_series(enaks_dir, x_range_enaks_deform, y_range_enaks_depth)
            out_c5_pdf = os.path.join(tmpdir, "C5_enaks_deformation.pdf")
            export_enaks_deformation_pdf(series_deform, outfile_pdf=out_c5_pdf, logo_path=logo_path, title_info={**title_info,"figur_nr":"C5"})
            with open(out_c5_pdf, "rb") as f:
                st.download_button("Download C5 – Enaks Deformation", f, file_name="C5_enaks_deformation.pdf")
