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
st.title("Combine Excel Files App")
st.write("Input files format:")
st.image("./img/xlsx.png")
# Function to combine Excel files with optional period handling and trimming
def combine_excel_files(files, add_period, trim_sections):
    combined_df = pd.DataFrame(columns=["File Name", "Text"])
    skipped_files = []
    for file in files:
        try:
            # Read the Excel file
            df = pd.read_excel(file)
            
            # Check if 'Text' column exists in the DataFrame
            if 'Text' in df.columns:
                # Get the filename without extension
                file_name = Path(file.name).stem
                
                # Initialize combined_text with the first row's text
                combined_text = df.loc[0, "Text"]
                
                # Loop through the rows and concatenate text with optional period handling and trimming
                for i in range(1, len(df)):
                    next_text = df.loc[i, "Text"]
                    if trim_sections:
                        combined_text = combined_text.strip()
                        next_text = next_text.strip()
                    if add_period and next_text and next_text[0].isupper() and not combined_text.endswith("."):
                        combined_text += "."
                    combined_text += " " + next_text
                
                # Append to the combined dataframe
                combined_df = pd.concat([combined_df, pd.DataFrame({"File Name": [file_name], "Text": [combined_text]})], ignore_index=True)
            else:
                skipped_files.append(file.name)
        except Exception as e:
            skipped_files.append(file.name + f" (Error: {str(e)})")
    
    return combined_df, skipped_files



st.write("How to Prepare Excel Files")
# Introduction on how to prepare Excel files
st.markdown("""
1. Ensure your Excel files have three columns: 'Start Time', 'End Time', and 'Text'.
2. Place the text you want to combine in the 'Text' column of each file.
3. You can also choose to trim leading and trailing whitespace from sections in the 'Text' column.
4. Upload one or more prepared Excel files using the 'Upload Excel Files' section below.
5. Click the 'Combine and Download' button to combine the files.
    """)
st.write("")
uploaded_files = st.file_uploader("Upload one or more Excel files", accept_multiple_files=True, type=["xlsx"])

add_period_checkbox = st.checkbox("Add a period at the end of lines starting with a capital letter")
trim_sections_checkbox = st.checkbox("Trim leading and trailing whitespace from sections")

if uploaded_files:
    if st.button("Combine and Download"):
        combined_df, skipped_files = combine_excel_files(uploaded_files, add_period_checkbox, trim_sections_checkbox)
        
        if combined_df.empty:
            st.warning("No 'Text' column found in any imported file.")
        else:
            # Download the combined dataframe as an Excel file
            combined_file = Path("Combined_Excel_File.xlsx")
            
            # Change the column name from "Combined Text" to "Text"
            combined_df.rename(columns={"Combined Text": "Text"}, inplace=True)
            
            combined_df.to_excel(combined_file, index=False, engine="openpyxl")
            
            # Provide a link to download the combined Excel file
            with open(combined_file, "rb") as file:
                b64 = base64.b64encode(file.read()).decode()
                st.markdown(f"**[Download Combined Excel File](data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64})**", unsafe_allow_html=True)
        
        # Display skipped files
        if skipped_files:
            st.warning("The following files were skipped:")
            for skipped_file in skipped_files:
                st.write(skipped_file)

# Display the combined dataframe (if available)
if 'combined_df' in locals() and not combined_df.empty:
    st.write("Combined Data:")
    st.dataframe(combined_df)