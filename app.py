import streamlit as st
import tempfile
import os
import pandas as pd
from plot_pdf import (
    export_sensitivity_pdf,
    export_curfc_pdf,
    export_cu_enaks_konus_pdf,
    export_enaks_deformation_pdf,
)
from build_data import build_konus_series, build_enaks_series, export_combined_table

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
ranges = {
    "konus_undist": 'L6:L30',
    "konus_remould": 'M6:M30',
    "depth": 'F6:F30',
    "enaks_strength": 'G6:G30',
    "enaks_deform": 'H6:H30',
}

if st.button("Generate Reports"):
    if not terrain_file or not konus_files or not enaks_files:
        st.error("Please upload terrain, Konus, and Enaks files.")
    else:
        with tempfile.TemporaryDirectory() as tmpdir:
            # Save terrain
            terrain_path = os.path.join(tmpdir, "terrain.xlsx")
            with open(terrain_path, "wb") as f: f.write(terrain_file.getbuffer())
            terrain_df = pd.read_excel(terrain_path, usecols="A:B")
            terrain_df.columns = ["BH", "Z"]
            terrain_lookup = dict(zip(terrain_df["BH"], terrain_df["Z"]))

            # Save Konus files
            konus_dir = os.path.join(tmpdir, "konus"); os.makedirs(konus_dir, exist_ok=True)
            for uf in konus_files:
                with open(os.path.join(konus_dir, uf.name), "wb") as f: f.write(uf.getbuffer())

            # Save Enaks files
            enaks_dir = os.path.join(tmpdir, "enaks"); os.makedirs(enaks_dir, exist_ok=True)
            for uf in enaks_files:
                with open(os.path.join(enaks_dir, uf.name), "wb") as f: f.write(uf.getbuffer())

            # --- Build unified series ---
            konus_series = build_konus_series(konus_dir, sheet_name, ranges, terrain_lookup)
            enaks_series = build_enaks_series(enaks_dir, sheet_name, ranges, terrain_lookup)

            # --- Generate figures with preview + download ---

            # C2 – Sensitivity
            out_c2_pdf = os.path.join(tmpdir, "C2_sensitivity.pdf")
            out_c2_png = os.path.join(tmpdir, "C2_sensitivity.png")
            export_sensitivity_pdf(konus_series, outfile_pdf=out_c2_pdf, outfile_png=out_c2_png,
                                   logo_path=logo_path, title_info={**title_info,"figur_nr":"C2"})
            st.subheader("C2 – Sensitivity")
            st.image(out_c2_png, caption="Preview C2 – Sensitivity", use_column_width=True)
            with open(out_c2_pdf, "rb") as f:
                st.download_button("Download C2 – Sensitivity PDF", f, file_name="C2_sensitivity.pdf")

            # C3 – Remoulded shear strength
            out_c3_pdf = os.path.join(tmpdir, "C3_curfc.pdf")
            out_c3_png = os.path.join(tmpdir, "C3_curfc.png")
            export_curfc_pdf(konus_series, outfile_pdf=out_c3_pdf, outfile_png=out_c3_png,
                             logo_path=logo_path, title_info={**title_info,"figur_nr":"C3"})
            st.subheader("C3 – Remoulded Shear Strength")
            st.image(out_c3_png, caption="Preview C3 – Remoulded", use_column_width=True)
            with open(out_c3_pdf, "rb") as f:
                st.download_button("Download C3 – Remoulded Strength PDF", f, file_name="C3_curfc.pdf")

            # C4 – Konus (undisturbed) + Enaks
            out_c4_pdf = os.path.join(tmpdir, "C4_cu_enaks_konus.pdf")
            out_c4_png = os.path.join(tmpdir, "C4_cu_enaks_konus.png")
            export_cu_enaks_konus_pdf(konus_series, enaks_series, outfile_pdf=out_c4_pdf, outfile_png=out_c4_png,
                                      logo_path=logo_path, title_info={**title_info,"figur_nr":"C4"})
            st.subheader("C4 – Konus + Enaks")
            st.image(out_c4_png, caption="Preview C4 – Konus + Enaks", use_column_width=True)
            with open(out_c4_pdf, "rb") as f:
                st.download_button("Download C4 – Konus + Enaks PDF", f, file_name="C4_cu_enaks_konus.pdf")

            # C5 – Enaks deformation
            out_c5_pdf = os.path.join(tmpdir, "C5_enaks_deformation.pdf")
            out_c5_png = os.path.join(tmpdir, "C5_enaks_deformation.png")
            export_enaks_deformation_pdf(enaks_series, outfile_pdf=out_c5_pdf, outfile_png=out_c5_png,
                                         logo_path=logo_path, title_info={**title_info,"figur_nr":"C5"})
            st.subheader("C5 – Enaks Deformation")
            st.image(out_c5_png, caption="Preview C5 – Enaks Deformation", use_column_width=True)
            with open(out_c5_pdf, "rb") as f:
                st.download_button("Download C5 – Enaks Deformation PDF", f, file_name="C5_enaks_deformation.pdf")
