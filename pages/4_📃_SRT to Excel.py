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
st.title("Script Converter")
st.markdown("""
Audio script conversion tool:
""")



# Define the target directory
target_directory = "./temp_files/SRT to Excel Converter"

# Create the target directory if it doesn't exist
if not os.path.exists(target_directory):
    os.makedirs(target_directory)

def srt_to_dataframe(srt_text, keep_text_only=False, export_single_files=False):
    lines = srt_text.split('\n')
    data = []
    subtitle_id, start_time, end_time, text = None, None, None, []

    if export_single_files:
        df = pd.DataFrame(columns=['Start Time', 'End Time', 'Text'])

    for line in lines:
        line = line.strip()
        if not line:
            if subtitle_id is not None and start_time is not None and end_time is not None:
                if keep_text_only:
                    data.append([" ".join(text)])
                else:
                    if export_single_files:
                        single_file_df = pd.DataFrame([[start_time, end_time, " ".join(text)]], columns=['Start Time', 'End Time', 'Text'])
                        file_name = f"{subtitle_id}.xlsx"
                        file_path = os.path.join(target_directory, file_name)
                        single_file_df.to_excel(file_path, index=False, engine='openpyxl')
                    else:
                        data.append([subtitle_id, start_time, end_time, " ".join(text)])
            subtitle_id, start_time, end_time, text = None, None, None, []
        elif not subtitle_id:
            subtitle_id = line
        elif not start_time:
            start_time, end_time = line.split(" --> ")
        else:
            text.append(line)

    if not export_single_files:
        if keep_text_only:
            df = pd.DataFrame(data, columns=['Text'])
        else:
            df = pd.DataFrame(data, columns=['ID', 'Start Time', 'End Time', 'Text'])

    return df

st.title("SRT to Excel Converter")
st.markdown("""
This tool converts multiple SRT files into Excel files or multiple XLSX files (one for each ID in the SRT) within a ZIP archive.
1. Import SRT files
2. Convert and Download as ZIP

If the "Export as single files" checkbox is selected, it will export each SRT as a separate XLSX file.

Excel files will be formatted as:
- **Column 1**: ID (if not exporting as single files)
- **Column 2**: Start Time
- **Column 3**: End Time
- **Column 4**: Text

Use the checkbox to extract text only.
The program will also save the Excel file with SRT file names as "srt_file_names.xlsx" (if not exporting as single files).
""")

st.write("---")

keep_text_only = st.checkbox("Keep Text Only 2")
export_single_files = st.checkbox("Export as single files")
uploaded_srt_files = st.file_uploader("Upload SRT files", type=["srt"], accept_multiple_files=True)

if st.button("Convert and Download"):
    with st.spinner("Converting files..."):
        if export_single_files:
            for selected_file in uploaded_srt_files:
                srt_text = selected_file.read().decode('utf-8')
                srt_to_dataframe(srt_text, keep_text_only=keep_text_only, export_single_files=True)

        else:
            zip_buffer = BytesIO()
            with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zipf:
                srt_file_names = []

                for selected_file in uploaded_srt_files:
                    srt_text = selected_file.read().decode('utf-8')
                    df = srt_to_dataframe(srt_text, keep_text_only=keep_text_only)

                    original_file_name = os.path.splitext(selected_file.name)[0]
                    srt_file_names.append(original_file_name)

                    excel_buffer = BytesIO()
                    df.to_excel(excel_buffer, index=False, engine='openpyxl')
                    excel_buffer.seek(0)

                    excel_file_name = original_file_name + ".xlsx"
                    excel_file_path = os.path.join(target_directory, excel_file_name)
                    zipf.writestr(excel_file_name, excel_buffer.read())

                srt_names_df = pd.DataFrame({"File Names": srt_file_names})

                excel_buffer_srt_names = BytesIO()
                srt_names_df.to_excel(excel_buffer_srt_names, index=False, header=True, engine='openpyxl')
                excel_buffer_srt_names.seek(0)

                srt_names_file_path = os.path.join(target_directory, "srt_file_names.xlsx")
                zipf.writestr("srt_file_names.xlsx", excel_buffer_srt_names.read())

            zip_buffer.seek(0)

            # Provide a download button for the ZIP file
            st.download_button(
                label="Download ZIP",
                data=zip_buffer,
                key='zip',
                file_name='srt_data.zip',
            )