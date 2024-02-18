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


# Streamlit UI
st.title("Excel to SRT Converter")
st.write("Input file format:")
st.image("./img/xlsx2.png")
st.markdown("""
This tool converts an Excel file into SubRip (SRT) format for subtitles. 
The Excel file should follow the following format:
- **Column 1**: File name
- **Column 2**: Text

Please make sure your Excel file adheres to this structure before proceeding.
""")

# Upload Excel file
uploaded_file = st.file_uploader("Upload an Excel file", type=["xls", "xlsx"])

if uploaded_file is not None:
    # Read the Excel file
    excel_data = pd.read_excel(uploaded_file)

    # Display the uploaded Excel data
    st.subheader("Uploaded Excel Data")
    st.write(excel_data)

    # Extract the "Text" column for SRT content
    srt_content = ""
    for index, row in excel_data.iterrows():
        text = row["Text"]
        srt_content += f"{index + 1}\n00:00:00,000 --> 00:00:00,000\n{text}\n\n"

    # Create SRT download link
    st.subheader("Converted SRT Data")
    st.text_area(label="SRT Data", value=srt_content, height=400)
    st.download_button(
        label="Download SRT",
        data=srt_content,
        key="srt_download",
        file_name="converted_data.srt",
    )


