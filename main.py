import streamlit as st
import csv
import os
import openpyxl
import chardet
from pathlib import Path
import tempfile
import codecs

def try_encodings(file_path):
    """Try multiple encodings and return the first one that works."""
    encodings = [
        'utf-8-sig',
        'utf-16',
        'utf-16le',
        'utf-16be',
        'utf-8',
        'ascii',
        'iso-8859-1',
        'cp1252'
    ]
    
    # First try chardet
    try:
        with open(file_path, 'rb') as file:
            raw_data = file.read()
            detected = chardet.detect(raw_data)
            if detected['confidence'] > 0.7:  # Only use if confidence is high
                encodings.insert(0, detected['encoding'])
    except:
        pass

    # Try each encoding
    for encoding in encodings:
        try:
            with codecs.open(file_path, 'r', encoding=encoding) as f:
                # Try to read a few lines to confirm encoding works
                for _ in range(3):
                    f.readline()
            return encoding
        except (UnicodeDecodeError, UnicodeError):
            continue
    
    raise ValueError(f"Could not determine encoding for {file_path}")

def detect_delimiter(file_path, encoding):
    """Detect the delimiter of a CSV file."""
    common_delimiters = [',', '\t', ';', '|']
    
    try:
        with codecs.open(file_path, 'r', encoding=encoding) as file:
            first_line = file.readline()
            # Count occurrences of each delimiter
            counts = {delimiter: first_line.count(delimiter) for delimiter in common_delimiters}
            # Return the delimiter with maximum occurrences
            max_count = max(counts.values())
            if max_count > 0:
                return max(counts.items(), key=lambda x: x[1])[0]
    except Exception as e:
        st.warning(f"Delimiter detection failed: {str(e)}. Defaulting to comma.")
    
    return ','

def read_csv_safely(file_path, encoding, delimiter):
    """Read CSV with error handling and verification."""
    rows = []
    try:
        with codecs.open(file_path, 'r', encoding=encoding) as f:
            reader = csv.reader(f, delimiter=delimiter)
            for row in reader:
                # Verify row is not empty and contains actual data
                if row and any(cell.strip() for cell in row):
                    rows.append(row)
        return rows
    except Exception as e:
        raise ValueError(f"Error reading CSV: {str(e)}")

def merge_csv_to_excel(csv_files):
    # Create a new Excel workbook
    wb = openpyxl.Workbook()
    wb.remove(wb.active)  # Remove the default sheet

    # Create a temporary directory to store the Excel file
    with tempfile.TemporaryDirectory() as temp_dir:
        # Iterate through the list of CSV files and write each one to a separate sheet
        for csv_file in csv_files:
            try:
                # Save uploaded file temporarily
                temp_file_path = os.path.join(temp_dir, csv_file.name)
                with open(temp_file_path, 'wb') as f:
                    f.write(csv_file.getvalue())

                # Detect encoding
                encoding = try_encodings(temp_file_path)
                st.info(f"Detected encoding for {csv_file.name}: {encoding}")

                # Detect delimiter
                delimiter = detect_delimiter(temp_file_path, encoding)
                st.info(f"Detected delimiter for {csv_file.name}: '{delimiter}'")

                # Create a new sheet for each CSV file
                sheet_name = Path(csv_file.name).stem[:31]  # Get file name without extension
                ws = wb.create_sheet(title=sheet_name)

                # Add the file name to the first row of the sheet
                ws.append([f"File Name: {csv_file.name}"])
                ws.append([f"Encoding: {encoding}"])
                ws.append([f"Delimiter: {delimiter}"])
                ws.append([])  # Empty row for separation

                # Read and write the CSV data
                rows = read_csv_safely(temp_file_path, encoding, delimiter)
                for row in rows:
                    ws.append(row)

                st.success(f"Successfully processed {csv_file.name}")

            except Exception as e:
                st.error(f"Failed to process {csv_file.name}. Error: {str(e)}")
                continue

        # Save the Excel workbook
        output_path = os.path.join(temp_dir, 'merged_output.xlsx')
        wb.save(output_path)
        
        # Read the file for downloading
        with open(output_path, 'rb') as f:
            return f.read()

def main():
    st.title("Advanced CSV Merger App")
    st.write("Upload your CSV files with any encoding (UTF-8, UTF-16, etc.) and merge them into a single Excel file.")

    # File uploader
    uploaded_files = st.file_uploader(
        "Choose CSV files", 
        type=['csv'], 
        accept_multiple_files=True,
        help="You can select multiple CSV files with different encodings"
    )

    if uploaded_files:
        st.write(f"Selected {len(uploaded_files)} files:")
        for file in uploaded_files:
            st.write(f"- {file.name}")

        if st.button("Merge Files"):
            with st.spinner("Merging files... This might take a moment."):
                try:
                    excel_data = merge_csv_to_excel(uploaded_files)
                    
                    # Offer the merged file for download
                    st.download_button(
                        label="Download Merged Excel File",
                        data=excel_data,
                        file_name="merged_output.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                    
                    st.success("Files merged successfully! Click the download button above to get your merged file.")
                except Exception as e:
                    st.error(f"An error occurred while merging: {str(e)}")

    # Add information about supported formats
    with st.expander("Supported Formats"):
        st.write("""
        This app supports CSV files with various encodings including:
        - UTF-8
        - UTF-16 (both LE and BE)
        - ASCII
        - ISO-8859-1
        - Windows-1252
        
        And various delimiters including:
        - Comma (,)
        - Tab (\\t)
        - Semicolon (;)
        - Pipe (|)
        
        Each file will be automatically analyzed for its encoding and delimiter.
        """)

if __name__ == "__main__":
    main()
