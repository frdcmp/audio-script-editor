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


# Function to combine text rows into full sentences
def combine_sentences(dataframe):
    combined_data = []

    # Initialize variables to store combined sentences
    current_sentence = ""
    start_time = ""
    end_time = ""
    sentence_id = 0

    for index, row in dataframe.iterrows():
        text = row["Text"]

        if current_sentence == "":
            start_time = row["Start Time"]
            current_sentence = text
        else:
            current_sentence += " " + text
            end_time = row["End Time"]

        if text.endswith((".", "!", "?")):
            sentence_id += 1
            combined_data.append({"ID": f"{sentence_id}", "Start Time": start_time, "End Time": end_time, "Text": current_sentence})
            current_sentence = ""

    combined_df = pd.DataFrame(combined_data)

    return combined_df

# Function to create a download link for a DataFrame as an Excel file
def download_excel(df, filename):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df.to_excel(writer, sheet_name='Sheet1', index=False)
    b64 = base64.b64encode(output.getvalue()).decode()
    return f'<a href="data:application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;base64,{b64}" download="{filename}.xlsx">Download Excel File</a>'

# Streamlit app title
st.title("Excel Viewer App")

# Upload an Excel file
excel_file = st.file_uploader("Sentence concatenator: Upload an Excel file", type=["xlsx", "xls"])

if excel_file is not None:
    # Read the Excel file into a DataFrame
    df = pd.read_excel(excel_file)

    # Display the DataFrame
    st.write("Original DataFrame:")

    # Checkbox for trimming text
    trim_text = st.checkbox("Trim the lines in the text column")

    # Checkbox for fixing punctuation
    fix_punctuation = st.checkbox("Fix punctuation")

    # Checkbox to combine text rows into sentences
    combine_sentences_checkbox = st.checkbox("Combine text rows into sentences")

    # Column selection using a select box
    selected_column = st.selectbox("Select the text column:", df.columns)

    if trim_text:
        # Check if the selected column contains string data
        if df[selected_column].dtype == object:
            df[selected_column] = df[selected_column].str.strip()

    if fix_punctuation:
        try:
            if df[selected_column].dtype == object:
                for i in range(1, len(df)):
                    if df[selected_column].iloc[i][0].isupper():
                        previous_line = df[selected_column].iloc[i - 1]
                        current_line = df[selected_column].iloc[i]

                        # Define a list of specified punctuation
                        specified_punctuation = [',', ':', '::', "'", '"', '?', '!', "."]

                        # Add a full stop at the end of the previous line if it doesn't end with specified punctuation
                        if not any(previous_line.endswith(punct) for punct in specified_punctuation):
                            df[selected_column].iloc[i - 1] = previous_line + '.'
        except IndexError:
            st.warning("Please select a valid text column.")

    # Combine text rows into sentences if the checkbox is selected
    if combine_sentences_checkbox:
        combined_df = combine_sentences(df)
        st.write("DataFrame with Combined Sentences:")
        st.dataframe(combined_df)

        # Add a button to download the DataFrame as an Excel file
        if st.button("Download as Excel"):
            tmp_download_link = download_excel(combined_df, "output_data")
            st.markdown(tmp_download_link, unsafe_allow_html=True)

    # Display the original DataFrame or the one with selected options
    if not combine_sentences_checkbox:
        st.write("DataFrame with Text Trimming and Punctuation Fix:")
        st.dataframe(df)
        # Add a button to download the DataFrame as an Excel file
        if st.button("Download as Excel"):
            tmp_download_link = download_excel(combined_df, "output_data")
            st.markdown(tmp_download_link, unsafe_allow_html=True)


