import streamlit as st
import pandas as pd
from docx import Document
import nltk

# Download nltk resources
nltk.download('punkt')

# Function to split text into sentences
def split_into_sentences(text):
    sentences = nltk.sent_tokenize(text)
    return sentences

st.set_page_config(layout="wide")
st.title("Word Document to Excel Converter")

st.title(".docx to .xlsx")
st.image("./img/docx.png")
st.markdown("""
This tool extract a tab an Excel file into SubRip (SRT) format for subtitles. 
The Excel file should follow the following format:
- **Column 1**: File name
- **Column 2**: Text

Please make sure your Excel file adheres to this structure before proceeding.
""")

st.write("Input file format: [.docx]")

uploaded_file = st.file_uploader("Upload a Word document", type=["docx"])

if uploaded_file is not None:
    document = Document(uploaded_file)
    tables = document.tables
    
    if tables:
        for i, table in enumerate(tables):
            st.header(f"Table {i + 1}")
            data = []
            keys = None

            for j, row in enumerate(table.rows):
                text = [cell.text.strip() for cell in row.cells]

                if j == 0:
                    keys = text
                else:
                    row_data = dict(zip(keys, text))
                    # Split text into sentences
                    row_data['Text'] = split_into_sentences(row_data['Text'])
                    data.append(row_data)

            if data:
                # Create DataFrame
                df = pd.DataFrame(data)
                # Explode the 'Text' column to split sentences into multiple rows
                df = df.explode('Text').reset_index(drop=True)
                
                # Selectbox for choosing the column to use as File name
                file_name_column = st.selectbox("Select column for File name", df.columns, index=0)
                
                # Selectbox for choosing the column to use as Text
                text_column = st.selectbox("Select column for Text", df.columns, index=1)
                
                st.dataframe(df, use_container_width=True, hide_index=True)

                excel_buffer = pd.ExcelWriter(f"table_{i + 1}.xlsx", engine='xlsxwriter')
                df.to_excel(excel_buffer, index=False, header=True)
                excel_buffer.close()

                st.download_button(
                    label=f"Download Table {i + 1} Excel",
                    data=open(f"table_{i + 1}.xlsx", 'rb'),
                    file_name=f"table_{i + 1}.xlsx",
                    key=f"download_{i}",
                )
