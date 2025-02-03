import streamlit as st
import csv
import os
import openpyxl
import chardet
from pathlib import Path
import tempfile
import codecs

# Set page configuration and style
st.set_page_config(
    page_title="Self CSV",
    page_icon="📊",
    layout="centered"
)

# Custom CSS for modern UI
st.markdown("""
    <style>
        .stButton>button {
            width: 100%;
            background-color: #FF4B4B;
            color: white;
            border-radius: 5px;
            height: 3em;
            margin-top: 20px;
        }
        .stButton>button:hover {
            background-color: #FF6B6B;
            border-color: #FF6B6B;
        }
        .upload-text {
            font-size: 1.2em;
            font-weight: bold;
            margin-bottom: 1em;
        }
        .success-message {
            padding: 1em;
            border-radius: 5px;
            background-color: #28a745;
            color: white;
        }
        .progress-container {
            margin: 1em 0;
        }
    </style>
""", unsafe_allow_html=True)

# Rest of your functions remain the same...

def merge_csv_to_excel(csv_files, progress_bar, status_text):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    with tempfile.TemporaryDirectory() as temp_dir:
        total_files = len(csv_files)
        for index, csv_file in enumerate(csv_files):
            try:
                status_text.text(f"Processing: {csv_file.name}")
                temp_file_path = os.path.join(temp_dir, csv_file.name)
                with open(temp_file_path, 'wb') as f:
                    f.write(csv_file.getvalue())

                encoding = try_encodings(temp_file_path)
                delimiter = detect_delimiter(temp_file_path, encoding)

                sheet_name = Path(csv_file.name).stem[:31]
                ws = wb.create_sheet(title=sheet_name)
                ws.append([f"File Name: {csv_file.name}"])

                with codecs.open(temp_file_path, 'r', encoding=encoding) as f:
                    reader = csv.reader(f, delimiter=delimiter)
                    for row in reader:
                        if row and any(cell.strip() for cell in row):
                            ws.append(row)

                progress_bar.progress((index + 1) / total_files)

            except Exception as e:
                status_text.text(f"Error processing {csv_file.name}: {str(e)}")
                continue

        status_text.text("Finalizing...")
        output_path = os.path.join(temp_dir, 'merged_output.xlsx')
        wb.save(output_path)
        
        with open(output_path, 'rb') as f:
            return f.read()

def main():
    st.title("Self CSV")
    
    # Container for the header section
    with st.container():
        st.markdown("""
        Welcome to Self CSV by [Shaikat Ray](https://shaikatray.com/) x [Self Canonical](https://selfcanonical.com/). 
        This tool merges multiple CSV files into a single Excel file, with each CSV becoming a separate sheet.
        """)

    # Features in columns
    col1, col2 = st.columns(2)
    with col1:
        st.markdown("""
        #### Features
        - Multiple file processing
        - Auto-encoding detection
        - Preserves file names
        """)
    with col2:
        st.markdown("""
        #### Supported Formats
        - UTF-8, UTF-16
        - CSV, TSV
        - Excel output
        """)

    # File upload section
    st.markdown("<div class='upload-text'>Upload Your Files</div>", unsafe_allow_html=True)
    uploaded_files = st.file_uploader(
        "Choose CSV files", 
        type=['csv'], 
        accept_multiple_files=True,
        help="You can select multiple CSV files"
    )

    if uploaded_files:
        st.markdown(f"**Selected Files:** {len(uploaded_files)}")
        for file in uploaded_files:
            st.markdown(f"- {file.name}")

        if st.button("Merge Files"):
            with st.container():
                progress_bar = st.progress(0)
                status_text = st.empty()
                
                try:
                    with st.spinner("Processing your files..."):
                        excel_data = merge_csv_to_excel(uploaded_files, progress_bar, status_text)
                        
                    st.markdown("<div class='success-message'>✅ Files merged successfully!</div>", 
                              unsafe_allow_html=True)
                    
                    # Download button
                    st.download_button(
                        label="📥 Download Merged File",
                        data=excel_data,
                        file_name="merged_files.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                except Exception as e:
                    st.error(f"An error occurred: {str(e)}")
                finally:
                    progress_bar.empty()

if __name__ == "__main__":
    main()
