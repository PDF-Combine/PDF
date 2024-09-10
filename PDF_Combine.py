import io
import re
import shutil
import tempfile
import uuid
from pathlib import Path
from subprocess import PIPE, run
from PyPDF2 import PdfMerger
from openpyxl import load_workbook
from fpdf import FPDF
from PIL import Image
import streamlit as st
from docx import Document

# Page configuration
st.set_page_config(
    page_title="Nicola's PDF Puzzle",
    page_icon="📄",
    layout="centered",
    initial_sidebar_state="auto"
)

# Main title and subheader
st.title("📄 Nicola's PDF Puzzle")
st.subheader("From chaos to order—one PDF at a time! 🚀")

def convert_doc_to_pdf_native(doc_file: Path, output_dir: Path = Path("."), timeout: int = 60):
    """Converts a doc file to pdf using LibreOffice."""
    exception = None
    output = None
    try:
        process = run(['soffice', '--headless', '--convert-to',
                        'pdf:writer_pdf_Export', '--outdir', output_dir.resolve(), doc_file.resolve()],
                       stdout=PIPE, stderr=PIPE,
                       timeout=timeout, check=True)
        stdout = process.stdout.decode("utf-8")
        re_filename = re.search(r'-> (.*?) using filter', stdout)
        output = Path(re_filename[1]).resolve()
    except Exception as e:
        exception = e
    return output, exception

def convert_excel_to_pdf(excel_file):
    """Convert Excel file to PDF."""
    try:
        wb = load_workbook(excel_file)
        sheet = wb.active
        pdf = FPDF()
        pdf.add_page()
        for row in sheet.iter_rows(values_only=True):
            row_text = ' '.join([str(cell) for cell in row if cell is not None])
            pdf.set_font("Arial", size=12)
            pdf.multi_cell(0, 10, row_text)
        
        pdf_output = io.BytesIO()
        pdf.output(pdf_output, 'S')
        pdf_output.seek(0)
        return pdf_output if pdf_output.getbuffer().nbytes > 0 else None
    except Exception as e:
        st.error(f"⚠️ An error occurred while converting Excel to PDF: {str(e)}")
        return None

def convert_image_to_pdf(image_file):
    """Convert image file to PDF."""
    try:
        img = Image.open(image_file)
        pdf_output = io.BytesIO()
        img.convert('RGB').save(pdf_output, format='PDF')
        pdf_output.seek(0)
        return pdf_output if pdf_output.getbuffer().nbytes > 0 else None
    except Exception as e:
        st.error(f"⚠️ An error occurred while converting image to PDF: {str(e)}")
        return None

# Instructions
st.write("""
**Upload up to 15 Word, Excel, Image, or PDF documents below.**
We'll help you combine them into one single, neat PDF file that you can download.
""")

# File upload
uploaded_files = st.file_uploader(
    "Choose files",
    type=["pdf", "docx", "xlsx", "png", "jpg", "jpeg"],
    accept_multiple_files=True
)

# Display a limit on the number of files
if uploaded_files and len(uploaded_files) > 15:
    st.error("🚨 You can upload a maximum of 15 files. Please try again.")
elif uploaded_files:
    st.write("All files uploaded successfully! Press the **Create PDF** button below to merge them.")
    
    # Create a temporary directory for file conversion
    temp_dir = Path(tempfile.mkdtemp())
    
    # Button to start the merging process
    if st.button("🎉 Create PDF!"):
        merger = PdfMerger()
        for file in uploaded_files:
            file_type = file.name.split('.')[-1].lower()
            pdf_file = None
            
            # Convert and merge based on file type
            if file_type == 'pdf':
                pdf_file = file
            elif file_type == 'docx':
                doc_file_path = Path(tempfile.mktemp())  # Temporary file for DOCX
                with open(doc_file_path, 'wb') as f:
                    f.write(file.getbuffer())
                pdf_file_path, exception = convert_doc_to_pdf_native(doc_file_path, temp_dir)
                if exception:
                    st.error(f"⚠️ An error occurred while converting {file.name} to PDF: {str(exception)}")
                if pdf_file_path:
                    pdf_file = pdf_file_path
            elif file_type == 'xlsx':
                pdf_file = convert_excel_to_pdf(file)
            elif file_type in ['png', 'jpg', 'jpeg']:
                pdf_file = convert_image_to_pdf(file)
            
            if pdf_file:
                try:
                    if isinstance(pdf_file, Path):  # If it's a file path, open it
                        pdf_file = open(pdf_file, 'rb')
                    merger.append(pdf_file)
                except Exception as e:
                    st.error(f"Error merging {file.name}: {str(e)}")

        # Cleanup temp directory
        shutil.rmtree(temp_dir)
        
        # Output merged PDF
        merged_pdf = io.BytesIO()
        merger.write(merged_pdf)
        merged_pdf.seek(0)
        merger.close()

        st.success("🎉 PDF created successfully! Download your merged PDF below.")
        st.download_button(
            label="📥 Download Merged PDF",
            data=merged_pdf,
            file_name="merged_document.pdf",
            mime="application/pdf"
        )
