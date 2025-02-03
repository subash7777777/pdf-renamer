import streamlit as st
import pandas as pd
import zipfile
import os
import shutil
from io import BytesIO

def process_files(pdf_zip_file, excel_file):
    # Create temporary directories
    temp_folder = "temp_pdfs"
    renamed_folder = "renamed_pdfs"
    os.makedirs(temp_folder, exist_ok=True)
    os.makedirs(renamed_folder, exist_ok=True)

    try:
        # Read Excel file
        data = pd.read_excel(excel_file)
        
        # Verify Excel structure
        if 'Account number' not in data.columns:
            st.error("Error: The Excel file must contain a column named 'Account number'.")
            return None

        # Extract PDFs from zip file
        with zipfile.ZipFile(pdf_zip_file, 'r') as zip_ref:
            zip_ref.extractall(temp_folder)

        # Get list of PDF files
        pdf_files = sorted(
            [f for f in os.listdir(temp_folder) if f.endswith('.pdf')],
            key=lambda x: int(os.path.splitext(x)[0])
        )

        # Verify file count matches
        if len(pdf_files) != len(data):
            st.error("Error: The number of PDFs does not match the number of rows in the Excel file.")
            return None

        # Rename PDFs
        for i, pdf_file in enumerate(pdf_files):
            account_number = str(data.iloc[i]['Account number'])
            old_path = os.path.join(temp_folder, pdf_file)
            new_path = os.path.join(renamed_folder, f"{account_number}.pdf")
            os.rename(old_path, new_path)

        # Create zip file in memory
        zip_buffer = BytesIO()
        with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for root, _, files in os.walk(renamed_folder):
                for file in files:
                    file_path = os.path.join(root, file)
                    arcname = os.path.basename(file_path)
                    zipf.write(file_path, arcname=arcname)

        return zip_buffer

    finally:
        # Clean up temporary folders
        if os.path.exists(temp_folder):
            shutil.rmtree(temp_folder)
        if os.path.exists(renamed_folder):
            shutil.rmtree(renamed_folder)

def main():
    st.title("PDF Renamer Tool")
    st.write("""
    This tool helps you rename PDF files based on account numbers from an Excel file.
    
    Instructions:
    1. Upload a ZIP file containing numbered PDF files (e.g., 1.pdf, 2.pdf, etc.)
    2. Upload an Excel file with an 'Account number' column
    3. Click 'Process Files' to rename and download the PDFs
    """)

    # File uploaders
    pdf_zip_file = st.file_uploader("Upload ZIP file containing PDFs", type=['zip'])
    excel_file = st.file_uploader("Upload Excel file with account numbers", type=['xlsx', 'xls'])

    if st.button("Process Files"):
        if pdf_zip_file is None or excel_file is None:
            st.error("Please upload both files before processing.")
            return

        with st.spinner("Processing files..."):
            zip_buffer = process_files(pdf_zip_file, excel_file)
            
            if zip_buffer:
                # Offer the processed file for download
                st.success("Processing complete! Click below to download your renamed PDFs.")
                st.download_button(
                    label="Download Renamed PDFs",
                    data=zip_buffer.getvalue(),
                    file_name="renamed_pdfs.zip",
                    mime="application/zip"
                )

if __name__ == "__main__":
    main()
