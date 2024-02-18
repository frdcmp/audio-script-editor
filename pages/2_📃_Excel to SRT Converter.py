import streamlit as st
import pandas as pd
from io import BytesIO
import re
import os
import zipfile
from pathlib import Path
import base64
from docx import Document
import shutil
import io


st.set_page_config(layout="wide")
st.title("Murf.ai Script Converter")


# Streamlit UI
st.title("Excel to SRT Converter")

# Switch between single import and batch conversion
conversion_mode = st.radio("Select Conversion Mode", ["Single Import", "Batch Conversion"])

# Common checkboxes for both modes
substitute_multiple_line_breaks = st.checkbox("Substitute double or triple line breaks with a single one", value=True)
insert_period_before_line_break = st.checkbox("Insert period before line break", value=True)
replace_line_breaks = st.checkbox("Replace line breaks with a space", value=True)

# Additional checkboxes
trim_cell = st.checkbox("Trim the cell to remove extra spaces at the beginning and end", value=True)
add_dot_at_end = st.checkbox("Add a dot at the end of each cell if it doesn't finish already with a dot", value=True)

# Define function to process Excel data
def process_excel_data(excel_data, trim_cell, add_dot_at_end):
    srt_content = ""
    for index, row in excel_data.iterrows():
        text = row["Text"]

        if trim_cell:
            text = text.strip()

        if add_dot_at_end and text and not text.endswith(('.', '!', '?')):
            text += '.'

        if substitute_multiple_line_breaks:
            text = '\n'.join(line.strip() for line in text.splitlines() if line.strip())

        if insert_period_before_line_break:
            lines = text.split('\n')
            for i in range(len(lines)-1, 0, -1):
                if lines[i] and not lines[i-1].endswith(('.', '!', '?', ',')) and not lines[i-1].isspace():
                    lines[i-1] += '.'
            text = '\n'.join(lines)

        if replace_line_breaks:
            text = text.replace("\n", " ").replace("\r", " ")

        srt_content += f"{index + 1}\n00:00:00,000 --> 00:00:00,000\n{text}\n\n"

    return srt_content

# Upload file/files based on the selected mode
if conversion_mode == "Single Import":
    uploaded_files = st.file_uploader("Upload an Excel file", type=["xls", "xlsx"], accept_multiple_files=False)

    if uploaded_files:
        excel_data = pd.read_excel(uploaded_files)
        st.subheader("Uploaded Excel Data")
        st.write(excel_data)

        uploaded_file_name = os.path.splitext(uploaded_files.name)[0]
        srt_content = process_excel_data(excel_data, trim_cell, add_dot_at_end)

        st.subheader("Converted SRT Data")
        st.text_area(label="SRT Data", value=srt_content, height=400)
        st.download_button(
            label=f"Download {uploaded_file_name}.srt",
            data=srt_content,
            key="srt_download",
            file_name=f"{uploaded_file_name}_converted_data.srt",
        )
else:
    uploaded_files = st.file_uploader("Upload Excel files", type=["xls", "xlsx"], accept_multiple_files=True)

    if uploaded_files:
        zip_buffer = BytesIO()
        with zipfile.ZipFile(zip_buffer, 'a', zipfile.ZIP_DEFLATED, False) as zip_file:
            for uploaded_file in uploaded_files:
                excel_data = pd.read_excel(uploaded_file)
                uploaded_file_name = os.path.splitext(uploaded_file.name)[0]
                srt_content = process_excel_data(excel_data, trim_cell, add_dot_at_end)
                zip_file.writestr(f"{uploaded_file_name}_converted_data.srt", srt_content)

        zip_buffer.seek(0)
        st.subheader("Download Converted SRT Files (ZIP)")
        st.markdown("Click the link below to download a ZIP file containing all the converted SRT files.")
        st.download_button(
            label="Download All SRT Files",
            data=zip_buffer.read(),
            key="all_srt_download",
            file_name="converted_srt_files.zip",
        )