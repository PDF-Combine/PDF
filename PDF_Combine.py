import streamlit as st
from PyPDF2 import PdfMerger
from openpyxl import load_workbook
from fpdf import FPDF
from PIL import Image, ImageDraw, ImageFont
import io
import tempfile
from datetime import datetime

# Page configuration - should be at the top
st.set_page_config(
    page_title="Nicola's PDF Puzzle",
    page_icon="📄",
    layout="centered",
    initial_sidebar_state="auto"
)

# Main title and subheader
st.title("📄 Nicola's PDF Puzzle")
st.subheader("From chaos to order—one PDF at a time! 🚀")

# Helper functions
def convert_docx_to_image(docx_file):
    try:
        from docx import Document
        document = Document(docx_file)

        images = []
        width, height = 800, 1000  # Adjust dimensions as needed
        
        for para in document.paragraphs:
            image = Image.new('RGB', (width, height), 'white')
            draw = ImageDraw.Draw(image)
            font = ImageFont.load_default()
            
            y = 10
            draw.text((10, y), para.text, font=font, fill='black')
            y += 20
            
            # Save image to a BytesIO object
            img_io = io.BytesIO()
            image.save(img_io, format='JPEG')
            img_io.seek(0)
            images.append(img_io)
        
        return images
    
    except Exception as e:
        st.error(f"⚠️ An error occurred while converting DOCX to image: {str(e)}")
        return None

def convert_images_to_pdf(image_list):
    try:
        pdf_output = io.BytesIO()
        images = [Image.open(img) for img in image_list]
        
        # Save the first image and append the rest
        images[0].save(pdf_output, save_all=True, append_images=images[1:], resolution=100.0, quality=95, optimize=True)
        pdf_output.seek(0)
        return pdf_output
    
    except Exception as e:
        st.error(f"⚠️ An error occurred while converting images to PDF: {str(e)}")
        return None

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
    return pdf_output if pdf_output.getbuffer().nbytes > 0 else None

def convert_image_to_pdf(image_file):
    img = Image.open(image_file)
    pdf_output = io.BytesIO()
    img.convert('RGB').save(pdf_output, format='PDF')
    pdf_output.seek(0)
    return pdf_output if pdf_output.getbuffer().nbytes > 0 else None

# Instructions
st.write("""
**Upload up to 15 Images or PDF documents below.**
We'll help you combine them into one single, neat PDF file that you can download.
""")

# File upload
uploaded_files = st.file_uploader("Choose files", type=["pdf", "docx", "xlsx", "png", "jpg", "jpeg"], accept_multiple_files=True)

# Display a limit on the number of files
if uploaded_files and len(uploaded_files) > 15:
    st.error("🚨 You can upload a maximum of 15 files. Please try again.")
elif uploaded_files:
    st.write("All files uploaded successfully! Press the **Create PDF** button below to merge them.")
    
    # Display files with numbering
    st.write("### Files Uploaded:")
    for idx, file in enumerate(uploaded_files, start=1):
        st.write(f"**Attachment {idx}:** {file.name}")

    # Button to start the merging process
    if st.button("🎉 Create PDF!"):
        merger = PdfMerger()
        for file in uploaded_files:
            file_type = file.name.split('.')[-1].lower()
            pdf_file = None
            
            if file_type == 'pdf':
                pdf_file = file
            elif file_type == 'docx':
                images = convert_docx_to_image(file)
                if images:
                    pdf_file = convert_images_to_pdf(images)
            elif file_type == 'xlsx':
                pdf_file = convert_excel_to_pdf(file)
            elif file_type in ['png', 'jpg', 'jpeg']:
                pdf_file = convert_image_to_pdf(file)
            
            if pdf_file:
                try:
                    merger.append(pdf_file)
                except Exception as e:
                    st.error(f"Error merging {file.name}: {str(e)}")

        merged_pdf = io.BytesIO()
        merger.write(merged_pdf)
        merged_pdf.seek(0)
        merger.close()

        # Generate timestamp for file name
        timestamp = datetime.now().strftime("%Y%m%d%H%M")
        file_name = f"merged_document_{timestamp}.pdf"

        st.success("🎉 PDF created successfully! Download your merged PDF below.")
        st.download_button(
            label="📥 Download Merged PDF",
            data=merged_pdf,
            file_name=file_name,
            mime="application/pdf"
        )
