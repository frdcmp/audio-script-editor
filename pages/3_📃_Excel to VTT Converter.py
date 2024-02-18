import streamlit as st
import pandas as pd
from io import BytesIO

def convert_to_vtt(data, timecode_in_col, timecode_out_col, text_col):
    vtt = "WEBVTT\n\n"
    for index, row in data.iterrows():
        time_in = row[timecode_in_col]
        time_out = row[timecode_out_col]
        text = row[text_col]
        vtt += f"{time_in} --> {time_out}\n{text}\n\n"
    return vtt


st.title("Excel to VTT Converter")

uploaded_file = st.file_uploader("Upload your Excel file", type=["xls", "xlsx"])

if uploaded_file is not None:
    file_contents = pd.read_excel(uploaded_file)
    
    columns = file_contents.columns.tolist()
    
    timecode_in_col = st.selectbox("Select column for Timecode IN", options=columns)
    timecode_out_col = st.selectbox("Select column for Timecode OUT", options=columns)
    text_col = st.selectbox("Select column for Text", options=columns)
    
    if st.button("Convert to VTT"):
        vtt_data = convert_to_vtt(file_contents, timecode_in_col, timecode_out_col, text_col)
        st.text_area("VTT Output", vtt_data, height=500)
        
        # Download button for VTT file
        download_button_str = f"Download {uploaded_file.name.split('.')[0]}.vtt"
        vtt_bytes = vtt_data.encode()
        st.download_button(label=download_button_str, data=BytesIO(vtt_bytes), file_name=f"{uploaded_file.name.split('.')[0]}.vtt")
