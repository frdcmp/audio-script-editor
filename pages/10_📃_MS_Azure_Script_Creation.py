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
st.title("MS Azure Script Creation Tool")
st.markdown("""

""")




# Function to create .txt files with optional SSML indentation
def create_txt_files(df, file_name_col, text_col, output_dir, add_indent, voice, lang_code):
    text_files_dir = os.path.join(output_dir, "text_files")
    os.makedirs(text_files_dir, exist_ok=True)
    for index, row in df.iterrows():
        file_name = row[file_name_col]
        text = row[text_col]
        if add_indent:
            text = f'''<speak version="1.0" xmlns="http://www.w3.org/2001/10/synthesis" xml:lang="{lang_code}"><voice name="{voice}">{text}</voice></speak>'''
        file_path = os.path.join(text_files_dir, f"{file_name}.txt")
        with open(file_path, "w") as txt_file:
            txt_file.write(text)








# Streamlit app
st.title("Excel to Text Files")

lang_code = st.text_input("Language code", "")
voice = st.text_input("Voice Name", "")
# Upload Excel file
uploaded_file = st.file_uploader("Upload an Excel file", type=["xlsx", "xls"])
if uploaded_file is not None:
    df = pd.read_excel(uploaded_file)
    st.write("Preview of the Excel Data:")
    st.write(df)

    # Select columns for "File Name" and "Text"
    file_name_col = st.selectbox("Select the column for 'File Name'", df.columns)
    text_col = st.selectbox("Select the column for 'Text'", df.columns, index=1)  # Default to the second column

    # Checkbox to add SSML indent
    add_indent = st.checkbox("Add SSML Indent", value = True)

    # Create a folder for the text files and the zip file
    output_dir = "./temp_files"
    if os.path.exists(output_dir):
        shutil.rmtree(output_dir)  # Clean temp_files directory
    os.makedirs(output_dir)

    # Button to create text files and download them as a zip
    if st.button("Download as Text"):
        # Create text files
        create_txt_files(df, file_name_col, text_col, output_dir, add_indent, voice, lang_code)
        st.success("Text files created!")

        # Create a zip file
        text_files_dir = os.path.join(output_dir, "text_files")
        zip_file_name = os.path.join(output_dir, "text_files.zip")
        with st.spinner("Creating zip file..."):
            with zipfile.ZipFile(zip_file_name, "w", zipfile.ZIP_DEFLATED) as zipf:
                for root, _, files in os.walk(text_files_dir):
                    for file in files:
                        zipf.write(os.path.join(root, file), os.path.relpath(os.path.join(root, file), text_files_dir))

            st.success("Zip file created!")

        # Provide a download button for the zip file
        with open(zip_file_name, "rb") as f:
            zip_data = f.read()
        st.download_button(
            label="Download Zip File",
            data=zip_data,
            file_name="text_files.zip",
        )