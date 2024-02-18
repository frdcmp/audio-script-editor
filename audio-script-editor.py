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

with st.expander("VTT to Excel Converter", expanded=False):

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












with st.expander("Combine Excel Files App", expanded=False):

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

    st.write("---")








with st.expander("Excel to SRT Converter", expanded=False):



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




with st.expander("Extract Tab from Word into Excel", expanded=False):
    st.title(".docx to .xlsx")
    st.write("Input file format: [.docx]")
    st.image("./img/docx.png")
    st.markdown("""
    This tool extract a tab an Excel file into SubRip (SRT) format for subtitles. 
    The Excel file should follow the following format:
    - **Column 1**: File name
    - **Column 2**: Text
    
    Please make sure your Excel file adheres to this structure before proceeding.
    """)
    # Upload Word Document
    uploaded_file = st.file_uploader("Upload a Word document", type=["docx"])

    if uploaded_file is not None:
        # Read the Word document
        document = Document(uploaded_file)

        # Extract and Display Tables
        for i, table in enumerate(document.tables):
            st.header(f"Table {i + 1}")
            data = []
            keys = None

            for j, row in enumerate(table.rows):
                text = [cell.text for cell in row.cells]

                if j == 0:
                    keys = text
                else:
                    row_data = dict(zip(keys, text))
                    data.append(row_data)

            if data:
                df = pd.DataFrame(data)
                st.dataframe(df, use_container_width=True, hide_index=True)

                # Create a BytesIO buffer to store the Excel file
                excel_buffer = BytesIO()
                df.to_excel(excel_buffer, index=False, header=True)
                excel_buffer.seek(0)  # Move the cursor to the beginning of the buffer

                # Add a download button to download the DataFrame as an Excel file
                download_filename = os.path.splitext(uploaded_file.name)[0] + "_table.xlsx"
                st.download_button(
                    label=f"Download {download_filename}",
                    data=excel_buffer,
                    file_name=download_filename,
                    key=f"download_{i}",
                )





with st.expander("MS SSML text bacth generator", expanded=False):
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

    # Upload Excel file
    uploaded_file = st.file_uploader("Upload an Excel file", type=["xlsx", "xls"])
    if uploaded_file is not None:
        df = pd.read_excel(uploaded_file)
        st.write("Preview of the Excel Data:")
        st.write(df)

        lang_code = st.text_input("Language code", "")
        voice = st.text_input("Voice Name", "")
        # Select columns for "File Name" and "Text"
        file_name_col = st.selectbox("Select the column for 'File Name'", df.columns)
        text_col = st.selectbox("Select the column for 'Text'", df.columns, index=1)  # Default to the second column

        # Checkbox to add SSML indent
        add_indent = st.checkbox("Add SSML Indent")

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
















# Define the target directory
target_directory = "./temp_files/SRT to Excel Converter"

# Create the target directory if it doesn't exist
if not os.path.exists(target_directory):
    os.makedirs(target_directory)

with st.expander("SRT to Excel Converter", expanded=False):
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

    st.write("---")







with st.expander("Sentence Concatenator", expanded=False):
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


