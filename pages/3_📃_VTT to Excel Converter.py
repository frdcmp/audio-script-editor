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


# Function to parse VTT and convert it to a DataFrame
def vtt_to_dataframe(vtt_text, keep_text_only=False):
    lines = vtt_text.split('\n')
    data = []
    start_time, end_time = None, None

    for line in lines:
        # Use regular expression to match timestamps and text
        timestamp_match = re.match(r'(\d{2}:\d{2}:\d{2}.\d{3}) --> (\d{2}:\d{2}:\d{2}.\d{3})', line)
        if timestamp_match:
            start_time, end_time = timestamp_match.groups()
        elif line.strip() and start_time is not None and end_time is not None:
            if keep_text_only:
                data.append([line])
            else:
                data.append([start_time, end_time, line])

    if keep_text_only:
        df = pd.DataFrame(data, columns=['Text'])
    else:
        df = pd.DataFrame(data, columns=['Start Time', 'End Time', 'Text'])

    return df

# Streamlit app

st.title("VTT to Excel Converter")
st.write("Input files format:")
st.image("./img/vtt.png")
st.markdown("""
This tool converts multiple VTT files into Excel files with the same name.         
1. Import VTT files
2. Convert and Download as ZIP
            
Every Excel file will be formatted as:
- **Column 1**: Start Time
- **Column 2**: End Time
- **Column 3**: Text

Use the checkbox to extract text only.      
The program will also Save the Excel file with VTT file names as "vtt_file_names.xlsx"
""")
st.write("---")

# Checkbox to decide whether to keep only the text or not
keep_text_only = st.checkbox("Keep Text Only2")

# Allow users to upload multiple VTT files
uploaded_vtt_files = st.file_uploader("Upload VTT files", type=["vtt"], accept_multiple_files=True)

if st.button("Convert and Download as ZIP"):
    with st.spinner("Converting and creating ZIP file..."):
        zip_buffer = BytesIO()
        with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zipf:
            vtt_file_names = []  # Store VTT file names

            for selected_file in uploaded_vtt_files:
                vtt_text = selected_file.read().decode('utf-8')
                df = vtt_to_dataframe(vtt_text, keep_text_only=keep_text_only)

                # Extract the original file name without extension
                original_file_name = os.path.splitext(selected_file.name)[0]
                vtt_file_names.append(original_file_name)

                # Create Excel file in memory
                excel_buffer = BytesIO()
                df.to_excel(excel_buffer, index=False, engine='openpyxl')
                excel_buffer.seek(0)

                # Save Excel file to the ZIP archive with the same name as the original VTT file
                excel_file_name = original_file_name + ".xlsx"
                zipf.writestr(excel_file_name, excel_buffer.read())

            # Create a DataFrame with the VTT file names
            vtt_names_df = pd.DataFrame({"File Names": vtt_file_names})

            # Save the VTT file names to an Excel file
            excel_buffer_vtt_names = BytesIO()
            vtt_names_df.to_excel(excel_buffer_vtt_names, index=False, header=True, engine='openpyxl')
            excel_buffer_vtt_names.seek(0)

            # Save the Excel file with VTT file names to the ZIP archive
            zipf.writestr("vtt_file_names.xlsx", excel_buffer_vtt_names.read())

        zip_buffer.seek(0)
        st.download_button(
            label="Download ZIP",
            data=zip_buffer,
            key='zip',
            file_name='vtt_data.zip',
        )

st.write("---")