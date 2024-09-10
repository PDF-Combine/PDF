
import streamlit as st
from PyPDF2 import PdfMerger
from docx import Document
from openpyxl import load_workbook
from fpdf import FPDF
from PIL import Image
import io

# Helper functions (same as before)
def convert_docx_to_pdf(docx_file):
    doc = Document(docx_file)
    pdf = FPDF()
    pdf.add_page()
    for para in doc.paragraphs:
        pdf.set_font("Arial", size=12)
        pdf.multi_cell(0, 10, para.text)
    
    pdf_output = io.BytesIO()  # Create a BytesIO stream
    pdf.output(pdf_output, 'S')  # Pass 'S' to output the PDF as a string, not a file
    pdf_output.seek(0)  # Move to the start of the BytesIO stream
    return pdf_output if pdf_output.getbuffer().nbytes > 0 else None  # Ensure it's not empty

def convert_excel_to_pdf(excel_file):
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
    return pdf_output if pdf_output.getbuffer().nbytes > 0 else None  # Ensure it's not empty

def convert_image_to_pdf(image_file):
    img = Image.open(image_file)
    pdf_output = io.BytesIO()
    img.convert('RGB').save(pdf_output, format='PDF')
    pdf_output.seek(0)
    return pdf_output if pdf_output.getbuffer().nbytes > 0 else None  # Ensure it's not empty

# Page configuration
st.set_page_config(
    page_title="The Ultimate PDF Mixer",
    page_icon="📄",
    layout="centered",
    initial_sidebar_state="auto"
)

# Main title and subheader
st.title("📄 Nicola's PDF Puzzle")
st.subheader("From chaos to order—one PDF at a time! 🚀")


# Instructions
st.write("""
**Upload up to 15 Word, Excel, Image, or PDF documents below.**
We'll help you combine them into one single, neat PDF file that you can download.
""")

# File upload
uploaded_files = st.file_uploader("Choose files", type=["pdf", "docx", "xlsx", "png", "jpg", "jpeg"], accept_multiple_files=True)

# Display a limit on the number of files
if uploaded_files and len(uploaded_files) > 15:
    st.error("🚨 You can upload a maximum of 15 files. Please try again.")
elif uploaded_files:
    st.write("All files uploaded successfully! Press the **Create PDF** button below to merge them.")
    
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
                pdf_file = convert_docx_to_pdf(file)
            elif file_type == 'xlsx':
                pdf_file = convert_excel_to_pdf(file)
            elif file_type in ['png', 'jpg', 'jpeg']:
                pdf_file = convert_image_to_pdf(file)
            
            if pdf_file:
                try:
                    merger.append(pdf_file)  # Append only if pdf_file is not None
                except Exception as e:
                    st.error(f"Error merging {file.name}: {str(e)}")

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
