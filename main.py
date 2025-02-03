import streamlit as st
import csv
import os
import openpyxl
import chardet
from pathlib import Path
import tempfile
import codecs

def try_encodings(file_path):
    encodings = ['utf-8-sig', 'utf-16', 'utf-16le', 'utf-16be', 'utf-8', 'ascii', 'iso-8859-1', 'cp1252']
    
    try:
        with open(file_path, 'rb') as file:
            raw_data = file.read()
            detected = chardet.detect(raw_data)
            if detected['confidence'] > 0.7:
                encodings.insert(0, detected['encoding'])
    except:
        pass

    for encoding in encodings:
        try:
            with codecs.open(file_path, 'r', encoding=encoding) as f:
                for _ in range(3):
                    f.readline()
            return encoding
        except (UnicodeDecodeError, UnicodeError):
            continue
    
    return 'utf-8'

def detect_delimiter(file_path, encoding):
    common_delimiters = [',', '\t', ';', '|']
    try:
        with codecs.open(file_path, 'r', encoding=encoding) as file:
            first_line = file.readline()
            counts = {delimiter: first_line.count(delimiter) for delimiter in common_delimiters}
            max_count = max(counts.values())
            if max_count > 0:
                return max(counts.items(), key=lambda x: x[1])[0]
    except:
        pass
    return ','

def merge_csv_to_excel(csv_files, progress_bar):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    with tempfile.TemporaryDirectory() as temp_dir:
        total_files = len(csv_files)
        for index, csv_file in enumerate(csv_files):
            try:
                temp_file_path = os.path.join(temp_dir, csv_file.name)
                with open(temp_file_path, 'wb') as f:
                    f.write(csv_file.getvalue())

                encoding = try_encodings(temp_file_path)
                delimiter = detect_delimiter(temp_file_path, encoding)

                sheet_name = Path(csv_file.name).stem[:31]
                ws = wb.create_sheet(title=sheet_name)

                # Add the file name as the first row
                ws.append([f"File Name: {csv_file.name}"])

                with codecs.open(temp_file_path, 'r', encoding=encoding) as f:
                    reader = csv.reader(f, delimiter=delimiter)
                    for row in reader:
                        if row and any(cell.strip() for cell in row):
                            ws.append(row)

                progress_bar.progress((index + 1) / total_files)

            except:
                continue

        output_path = os.path.join(temp_dir, 'merged_output.xlsx')
        wb.save(output_path)
        
        with open(output_path, 'rb') as f:
            return f.read()

def main():
    st.title("CSV Merger")

    uploaded_files = st.file_uploader(
        "Choose CSV files", 
        type=['csv'], 
        accept_multiple_files=True
    )

    if uploaded_files:
        if st.button("Merge Files"):
            with st.spinner("Merging files..."):
                progress_bar = st.progress(0)
                excel_data = merge_csv_to_excel(uploaded_files, progress_bar)
                progress_bar.empty()
                
                st.success("Files merged successfully!")
                st.download_button(
                    label="Download Merged File",
                    data=excel_data,
                    file_name="merged_output.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

if __name__ == "__main__":
    main()
