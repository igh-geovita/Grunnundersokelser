import os
import tempfile
import streamlit as st
import pandas as pd
from openpyxl import load_workbook
import matplotlib.pyplot as plt
import matplotlib.image as mpimg
from matplotlib import patches
import itertools

# Import your export function (copied from your script)
from yourmodule import export_curfc_pdf   # or paste export_curfc_pdf here


# -----------------------
# Streamlit UI
# -----------------------
st.title("Geovita Plot Generator")

# --- Inputs ---
data_files = st.file_uploader(
    "Upload your Excel data files", type=["xlsx", "xls", "xlsm"], accept_multiple_files=True
)

terrain_file = st.file_uploader("Upload Terrain Levels (Excel)", type=["xlsx"])
logo_file = st.file_uploader("Upload Logo (PNG)", type=["png"])

sheet_name = st.text_input("Sheet name", "Sheet 001")
x_range = st.text_input("X range (Excel notation)", "M6:M30")
y_range = st.text_input("Y range (Excel notation)", "F6:F30")

# Title block
st.subheader("Title block info")
rapport_nr = st.text_input("Rapport Nr.", "XX")
dato       = st.date_input("Dato")
tegn       = st.text_input("Tegn", "IGH")
kontr      = st.text_input("Kontr", "JOG")
godkj      = st.text_input("Godkj", "AGR")
figur_nr   = st.text_input("Figur Nr.", "C3")


# -----------------------
# Process when ready
# -----------------------
if st.button("Generate Report"):
    if not data_files or not terrain_file:
        st.error("Please upload at least one data file and a terrain file")
    else:
        # Save uploads to temp folder
        with tempfile.TemporaryDirectory() as tmpdir:
            # Save terrain
            terrain_path = os.path.join(tmpdir, terrain_file.name)
            with open(terrain_path, "wb") as f:
                f.write(terrain_file.getbuffer())
            terrain_df = pd.read_excel(terrain_path, usecols="A:B")
            terrain_df.columns = ["BH", "Z"]
            terrain_lookup = dict(zip(terrain_df["BH"], terrain_df["Z"]))

            # Save logo (optional)
            logo_path = None
            if logo_file:
                logo_path = os.path.join(tmpdir, logo_file.name)
                with open(logo_path, "wb") as f:
                    f.write(logo_file.getbuffer())

            # Collect data series
            data_series = []
            for file in data_files:
                file_path = os.path.join(tmpdir, file.name)
                with open(file_path, "wb") as f:
                    f.write(file.getbuffer())

                workbook_name = os.path.splitext(file.name)[0]
                terrain_level = terrain_lookup.get(workbook_name)

                if terrain_level is None:
                    st.warning(f"No terrain level found for {workbook_name}, skipping")
                    continue

                try:
                    wb = load_workbook(file_path, data_only=True)
                    if sheet_name not in wb.sheetnames:
                        st.warning(f"Sheet {sheet_name} not in {file.name}, skipping")
                        continue
                    ws = wb[sheet_name]

                    x_vals = [cell[0].value for cell in ws[x_range]]
                    y_vals = [cell[0].value for cell in ws[y_range]]

                    pts = [(x, d) for x, d in zip(x_vals, y_vals) if x is not None and d is not None]
                    if not pts:
                        st.warning(f"No valid points in {file.name}, skipping")
                        continue

                    x_f, y_f = zip(*pts)
                    elev = [terrain_level - d for d in y_f]
                    data_series.append((workbook_name, x_f, y_f, elev, terrain_level))
                except Exception as e:
                    st.error(f"Error reading {file.name}: {e}")

            # Run export function
            output_pdf = os.path.join(tmpdir, "curfc_report.pdf")
            output_png = os.path.join(tmpdir, "curfc_report.png")

            export_curfc_pdf(
                data_series,
                outfile_pdf=output_pdf,
                outfile_png=output_png,
                logo_path=logo_path,
                title_info={
                    "rapport_nr": rapport_nr,
                    "dato": str(dato),
                    "tegn": tegn,
                    "kontr": kontr,
                    "godkj": godkj,
                    "figur_nr": figur_nr,
                },
                depth_ylim=(0, 35),
                margin_cm=1.0,
            )

            # Show preview
            st.image(output_png, caption="Preview", use_column_width=True)

            # Download buttons
            with open(output_pdf, "rb") as f:
                st.download_button("Download PDF", f, file_name="curfc_report.pdf")
            with open(output_png, "rb") as f:
                st.download_button("Download PNG", f, file_name="curfc_report.png")
