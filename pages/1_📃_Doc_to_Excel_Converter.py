import streamlit as st
import pandas as pd
from docx import Document
import os

st.set_page_config(layout="wide")
st.title("Word Table Extractor and Converter")

st.title("Word Document Table Extractor")
st.image("./img/docx.png")
st.markdown("""
This tool extracts tables from a Word document (.docx) and converts them into Excel (.xlsx) format.
The Word document should contain tables with the following structure:
- **Each table**: Represents a separate Excel file.
- **Column 1**: File name
- **Column 2**: Text
""")

st.write("Input file format: [.docx]")

uploaded_file = st.file_uploader("Upload a Word document", type=["docx"])

if uploaded_file is not None:
    document = Document(uploaded_file)
    tables = document.tables
    
    if tables:
        
        input_filename = uploaded_file.name.split('.')[0]  # Extracting filename without extension
        for i, table in enumerate(tables):
            st.header(f"{input_filename}")
            data = []
            keys = None

            for j, row in enumerate(table.rows):
                text = [cell.text.strip() for cell in row.cells]

                if j == 0:
                    keys = text
                else:
                    row_data = dict(zip(keys, text))
                    data.append(row_data)

            if data:
                df = pd.DataFrame(data)
                with st.expander("Show Dataframe"):
                    st.dataframe(df, use_container_width=True, hide_index=True)

                excel_filename = f"{input_filename}.xlsx"  # Generate output Excel filename
                excel_path = os.path.join("output", excel_filename)
                os.makedirs("output", exist_ok=True)

                excel_writer = pd.ExcelWriter(excel_path, engine='xlsxwriter')
                df.to_excel(excel_writer, index=False, header=True)
                excel_writer.close()  # Correct method to close the ExcelWriter object

                st.success(f"Script {input_filename} extracted and converted successfully!")
                st.download_button(
                    label=f"Download {input_filename} Excel",
                    data=open(excel_path, 'rb').read(),
                    file_name=excel_filename,
                    key=f"download_{i}",
                )

                # Clean up temporary Excel file
                os.remove(excel_path)