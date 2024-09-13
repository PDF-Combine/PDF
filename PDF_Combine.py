# import streamlit as st
# from PyPDF2 import PdfMerger
# from openpyxl import load_workbook
# from fpdf import FPDF
# from PIL import Image
# from io import BytesIO
# from docx import Document
# from reportlab.lib.pagesizes import letter
# from reportlab.pdfgen import canvas
# import io
# from datetime import datetime
# import tempfile

# # Page configuration - should be at the top
# st.set_page_config(
#     page_title="Nicola's PDF Puzzle",
#     page_icon="📄",
#     layout="centered",
#     initial_sidebar_state="auto"
# )

# # Main title and subheader
# st.title("📄 Nicola's PDF Puzzle")
# st.subheader("From chaos to order—one PDF at a time! 🚀")

# # Helper functions
# def convert_word_to_pdf(docx_file):
#     try:
#         # Handle if the input file is BytesIO
#         if isinstance(docx_file, BytesIO):
#             # Save the BytesIO content to a temporary file
#             with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as temp_docx:
#                 temp_docx.write(docx_file.getbuffer())
#                 temp_docx_path = temp_docx.name

#             # Read the temporary DOCX file
#             doc = Document(temp_docx_path)

#             # Check if the document is empty
#             if not doc.paragraphs:
#                 raise ValueError("The DOCX file is empty.")

#             # Create a PDF in memory
#             pdf_output = BytesIO()
#             c = canvas.Canvas(pdf_output, pagesize=letter)
#             width, height = letter

#             # Write each paragraph of the DOCX file to the PDF
#             text = c.beginText(40, height - 40)
#             text.setFont("Helvetica", 12)

#             for para in doc.paragraphs:
#                 text.textLine(para.text)

#             c.drawText(text)
#             c.showPage()
#             c.save()

#             pdf_output.seek(0)
#             return pdf_output

#         else:
#             raise ValueError("The input file is not a valid DOCX file.")
#     except Exception as e:
#         st.error(f"⚠️ An error occurred while converting DOCX to PDF: {str(e)}")
#         return None

# def convert_excel_to_pdf(excel_file):
#     try:
#         # Handle if the input file is BytesIO
#         if isinstance(excel_file, BytesIO):
#             # Load the workbook from the BytesIO object
#             wb = load_workbook(excel_file)
#             sheet = wb.active

#             # Create a PDF in memory
#             pdf_output = BytesIO()
#             c = canvas.Canvas(pdf_output, pagesize=letter)
#             width, height = letter

#             # Write each row of the Excel file to the PDF
#             text = c.beginText(40, height - 40)
#             text.setFont("Helvetica", 12)

#             for row in sheet.iter_rows(values_only=True):
#                 row_text = ' '.join([str(cell) for cell in row if cell is not None])
#                 text.textLine(row_text)

#             c.drawText(text)
#             c.showPage()
#             c.save()

#             pdf_output.seek(0)
#             return pdf_output

#         else:
#             raise ValueError("The input file is not a valid Excel file.")
#     except Exception as e:
#         st.error(f"⚠️ An error occurred while converting Excel to PDF: {str(e)}")
#         return None

# def convert_image_to_pdf(image_file):
#     try:
#         img = Image.open(image_file)
#         pdf_output = io.BytesIO()
#         img.convert('RGB').save(pdf_output, format='PDF')
#         pdf_output.seek(0)
#         return pdf_output if pdf_output.getbuffer().nbytes > 0 else None
#     except Exception as e:
#         st.error(f"⚠️ An error occurred while converting image to PDF: {str(e)}")
#         return None

# # Instructions
# st.write("""
# **Upload up to 15 Images, Word, Excel, or PDF documents below.**
# We'll help you combine them into one single, neat PDF file that you can download.
# """)

# # File upload
# uploaded_files = st.file_uploader("Choose files", type=["pdf", "png", "jpg", "jpeg", "docx", "xlsx"], accept_multiple_files=True)

# if uploaded_files:
#     # Display a limit on the number of files
#     if len(uploaded_files) > 15:
#         st.error("🚨 You can upload a maximum of 15 files. Please try again.")
#     else:
#         st.write("All files uploaded successfully! Press the **Create PDF** button below to merge them.")

#         # Display files with dropdowns for reordering
#         file_names = [file.name for file in uploaded_files]
#         reordered_files = []

#         st.write("### Files Uploaded:")
#         for idx, file_name in enumerate(file_names):
#             selected_position = st.selectbox(
#                 label=f"Position for Attachment {idx + 1}: {file_name}",
#                 options=list(range(1, len(file_names) + 1)),
#                 index=idx,
#                 key=f"select_{idx}"
#             )
#             reordered_files.append((selected_position, uploaded_files[idx]))

#         # Sort files based on selected positions
#         reordered_files.sort(key=lambda x: x[0])
#         sorted_files = [file[1] for file in reordered_files]

#         # Button to start the merging process
#         if st.button("🎉 Create PDF!"):
#             merger = PdfMerger()
#             for file in sorted_files:
#                 file_type = file.name.split('.')[-1].lower()
#                 pdf_file = None
                
#                 if file_type == 'pdf':
#                     pdf_file = file
#                 elif file_type == 'docx':
#                     pdf_file = convert_word_to_pdf(file)
#                 elif file_type == 'xlsx':
#                     pdf_file = convert_excel_to_pdf(file)
#                 elif file_type in ['png', 'jpg', 'jpeg']:
#                     pdf_file = convert_image_to_pdf(file)
                
#                 if pdf_file:
#                     try:
#                         merger.append(pdf_file)
#                     except Exception as e:
#                         st.error(f"Error merging {file.name}: {str(e)}")
#                 else:
#                     st.error(f"Conversion failed for {file.name}. Skipping this file.")

#             merged_pdf = io.BytesIO()
#             merger.write(merged_pdf)
#             merged_pdf.seek(0)
#             merger.close()

#             # Generate timestamp for file name
#             timestamp = datetime.now().strftime("%Y%m%d%H%M")
#             file_name = f"merged_document_{timestamp}.pdf"

#             st.success("🎉 PDF created successfully! Download your merged PDF below.")
#             st.download_button(
#                 label="📥 Download Merged PDF",
#                 data=merged_pdf,
#                 file_name=file_name,
#                 mime="application/pdf"
#             )


import streamlit as st
from PyPDF2 import PdfMerger, PdfWriter
from openpyxl import load_workbook
from fpdf import FPDF
from PIL import Image
from io import BytesIO
from docx import Document
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
import io
from datetime import datetime
import tempfile
import pypandoc
from pdf2docx import Converter

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

# Helper function to convert DOCX to PDF
def convert_word_to_pdf(docx_file):
    try:
        if isinstance(docx_file, BytesIO):
            with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as temp_docx:
                temp_docx.write(docx_file.getbuffer())
                temp_docx_path = temp_docx.name

            doc = Document(temp_docx_path)
            if not doc.paragraphs:
                raise ValueError("The DOCX file is empty.")

            pdf_output = BytesIO()
            c = canvas.Canvas(pdf_output, pagesize=letter)
            width, height = letter

            text = c.beginText(40, height - 40)
            text.setFont("Helvetica", 12)

            for para in doc.paragraphs:
                text.textLine(para.text)

            c.drawText(text)
            c.showPage()
            c.save()

            pdf_output.seek(0)
            return pdf_output
        else:
            raise ValueError("The input file is not a valid DOCX file.")
    except Exception as e:
        st.error(f"⚠️ Error converting DOCX to PDF: {str(e)}")
        return None

# Helper function to convert Excel to PDF
def convert_excel_to_pdf(excel_file):
    try:
        if isinstance(excel_file, BytesIO):
            wb = load_workbook(excel_file)
            sheet = wb.active

            pdf_output = BytesIO()
            c = canvas.Canvas(pdf_output, pagesize=letter)
            width, height = letter

            text = c.beginText(40, height - 40)
            text.setFont("Helvetica", 12)

            for row in sheet.iter_rows(values_only=True):
                row_text = ' '.join([str(cell) for cell in row if cell is not None])
                text.textLine(row_text)

            c.drawText(text)
            c.showPage()
            c.save()

            pdf_output.seek(0)
            return pdf_output

        else:
            raise ValueError("The input file is not a valid Excel file.")
    except Exception as e:
        st.error(f"⚠️ Error converting Excel to PDF: {str(e)}")
        return None

# Helper function to convert image to PDF
def convert_image_to_pdf(image_file):
    try:
        img = Image.open(image_file)
        pdf_output = io.BytesIO()
        img.convert('RGB').save(pdf_output, format='PDF')
        pdf_output.seek(0)
        return pdf_output if pdf_output.getbuffer().nbytes > 0 else None
    except Exception as e:
        st.error(f"⚠️ Error converting image to PDF: {str(e)}")
        return None

# Convert PDF to DOCX using pdf2docx
def convert_pdf_to_docx(pdf_data):
    try:
        with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as temp_pdf:
            temp_pdf.write(pdf_data)
            temp_pdf_path = temp_pdf.name

        docx_output = io.BytesIO()
        converter = Converter(temp_pdf_path)
        converter.convert(docx_output)  # Save output to BytesIO
        converter.close()
        docx_output.seek(0)
        return docx_output
    except Exception as e:
        st.error(f"⚠️ Error converting PDF to DOCX: {str(e)}")
        return None

# Helper function to display PDF using pdf.js
def show_pdf(pdf_data):
    base64_pdf = base64.b64encode(pdf_data).decode('utf-8')
    pdf_display = f'<iframe src="data:application/pdf;base64,{base64_pdf}" width="700" height="1000" type="application/pdf"></iframe>'
    st.markdown(pdf_display, unsafe_allow_html=True)

# Instructions
st.write("""
**Upload up to 15 Images, Word, Excel, or PDF documents below.**
We'll help you combine them into one single, neat PDF file that you can download.
""")

# File upload
uploaded_files = st.file_uploader("Choose files", type=["pdf", "png", "jpg", "jpeg", "docx", "xlsx"], accept_multiple_files=True)

if uploaded_files:
    if len(uploaded_files) > 15:
        st.error("🚨 You can upload a maximum of 15 files. Please try again.")
    else:
        st.write("All files uploaded successfully! Press the **Create PDF** button below to merge them.")

        file_names = [file.name for file in uploaded_files]
        reordered_files = []

        st.write("### Files Uploaded:")
        for idx, file_name in enumerate(file_names):
            selected_position = st.selectbox(
                label=f"Position for Attachment {idx + 1}: {file_name}",
                options=list(range(1, len(file_names) + 1)),
                index=idx,
                key=f"select_{idx}"
            )
            reordered_files.append((selected_position, uploaded_files[idx]))

        reordered_files.sort(key=lambda x: x[0])
        sorted_files = [file[1] for file in reordered_files]

        if st.button("🎉 Create PDF!") or "merged_pdf" in st.session_state:
            if "merged_pdf" not in st.session_state:
                merger = PdfMerger()
                for file in sorted_files:
                    file_type = file.name.split('.')[-1].lower()
                    pdf_file = None

                    if file_type == 'pdf':
                        pdf_file = file
                    elif file_type == 'docx':
                        pdf_file = convert_word_to_pdf(file)
                    elif file_type == 'xlsx':
                        pdf_file = convert_excel_to_pdf(file)
                    elif file_type in ['png', 'jpg', 'jpeg']:
                        pdf_file = convert_image_to_pdf(file)

                    if pdf_file:
                        try:
                            merger.append(pdf_file)
                        except Exception as e:
                            st.error(f"Error merging {file.name}: {str(e)}")
                    else:
                        st.error(f"Conversion failed for {file.name}. Skipping this file.")

                merged_pdf = io.BytesIO()
                merger.write(merged_pdf)
                merged_pdf.seek(0)
                merger.close()

                st.session_state.merged_pdf = merged_pdf

            add_password = st.radio("Would you like to add a password?", ("No", "Yes"))

            if add_password == "Yes":
                password = st.text_input("Enter password for the PDF", type="password")

                if password and "protected_pdf" not in st.session_state:
                    writer = PdfWriter()
                    reader = PdfMerger(open=st.session_state.merged_pdf)
                    for page in reader.pages:
                        writer.add_page(page)
                    writer.encrypt(user_pwd=password)
                    protected_pdf = io.BytesIO()
                    writer.write(protected_pdf)
                    protected_pdf.seek(0)
                    st.session_state.merged_pdf = protected_pdf
                    st.success("🔒 Password has been added to your PDF.")

            st.write("### Preview of Merged PDF:")
            show_pdf(st.session_state.merged_pdf.getvalue())

            timestamp = datetime.now().strftime("%Y%m%d%H%M")
            file_name = f"merged_document_{timestamp}.pdf"

            st.success("🎉 PDF created successfully! Download your merged PDF below.")
            st.download_button(
                label="📥 Download Merged PDF",
                data=st.session_state.merged_pdf,
                file_name=file_name,
                mime="application/pdf"
            )

            # Convert merged PDF to DOCX and provide download option
            docx_data = convert_pdf_to_docx(st.session_state.merged_pdf.getvalue())
            st.download_button(
                label="📥 Download as DOCX",
                data=docx_data,
                file_name=f"merged_document_{timestamp}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
r
