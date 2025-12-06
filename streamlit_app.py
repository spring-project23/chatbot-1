import streamlit as st
from googletrans import Translator
import fitz  # PyMuPDF
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from io import BytesIO
from PIL import Image
import pytesseract
import base64

# ============================
# EMBEDDED LOGO (BASE64 STRING)
# Replace this with your own Base64 if needed
# ============================
RAJYUG_LOGO_BASE64 = """
<PUT_YOUR_BASE64_STRING_HERE>
"""

# Function to extract text and formatting from PDF using PyMuPDF (fitz)
def extract_text_from_pdf(pdf_document):
    elements = []
    for page_num in range(len(pdf_document)):
        page = pdf_document.load_page(page_num)
        page_width = page.rect.width

        text_blocks = page.get_text("dict")["blocks"]
        for block in text_blocks:
            if block["type"] == 0:
                for line in block["lines"]:
                    line_text = "".join([span["text"] for span in line["spans"]])
                    font_size = line["spans"][0]["size"] if line["spans"] else None
                    font_name = line["spans"][0]["font"] if line["spans"] else None
                    is_bold = "Bold" in (line["spans"][0]["font"] if line["spans"] else "")
                    alignment = WD_PARAGRAPH_ALIGNMENT.LEFT

                    x0 = line["bbox"][0]
                    if x0 < page_width * 0.3:
                        alignment = WD_PARAGRAPH_ALIGNMENT.LEFT
                    elif x0 > page_width * 0.7:
                        alignment = WD_PARAGRAPH_ALIGNMENT.RIGHT
                    else:
                        alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

                    elements.append(('text', line_text, font_name, font_size, is_bold, alignment))

    return elements

# OCR extract
def extract_text_from_image(image):
    return pytesseract.image_to_string(image)

# Create DOCX
def create_docx_with_formatting(elements):
    doc = Document()

    for element_type, *data in elements:
        if element_type == 'text':
            text, font_name, font_size, is_bold, alignment = data
            paragraph = doc.add_paragraph()
            run = paragraph.add_run(text.strip())

            if font_size:
                run.font.size = Pt(font_size)
            if font_name:
                run.font.name = font_name

            if is_bold:
                run.bold = True

            paragraph.alignment = alignment

    doc_io = BytesIO()
    doc.save(doc_io)
    doc_io.seek(0)
    return doc_io

# Translate text
def translate_text(text, target_language):
    translator = Translator()
    try:
        return translator.translate(text, dest=target_language).text
    except Exception as e:
        return f"An error occurred: {e}"

# =====================================
# STREAMLIT APP
# =====================================
def main():

    # Display embedded logo
    st.markdown(
        f"""
        <div style='text-align: center; padding-bottom: 30px;'>
            <img src="data:image/jpeg;base64,{RAJYUG_LOGO_BASE64}" width="250">
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.sidebar.title("Image & PDF Translator")
    st.sidebar.write("Upload a PDF or Image → Translate → Download DOCX")

    uploaded_file = st.sidebar.file_uploader("Upload", type=["pdf", "jpg", "jpeg", "png"])

    if uploaded_file:
        if uploaded_file.type == "application/pdf":
            st.write("Extracting text from PDF...")
            pdf_document = fitz.open(stream=uploaded_file.read(), filetype="pdf")
            all_elements = extract_text_from_pdf(pdf_document)

        else:
            st.write("Extracting text from image...")
            image = Image.open(uploaded_file)
            text = extract_text_from_image(image)
            all_elements = [('text', text, None, None, False, WD_PARAGRAPH_ALIGNMENT.LEFT)]

        target_language = st.sidebar.selectbox(
            "Select target language", ["Marathi", "Hindi"]
        )
        language_code = "mr" if target_language == "Marathi" else "hi"

        translated_elements = []
        for element_type, *data in all_elements:
            if element_type == "text":
                text = data[0]
                translated_text = translate_text(text, language_code)
                translated_elements.append(('text', translated_text, *data[1:]))

        st.write(f"**Translated Content → {target_language}**")
        st.write("Converting to DOCX...")

        docx_file = create_docx_with_formatting(translated_elements)

        st.download_button(
            label=f"Download DOCX ({target_language})",
            data=docx_file,
            file_name=f"translated_{target_language.lower()}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )

    else:
        st.error("Please upload a file.")

    # Footer
    st.markdown(
        """
        <style>
        .footer {
            position: fixed;
            bottom: 0;
            left: 170px;
            width: 100%;
            background-color: #000;
            text-align: center;
            padding: 10px;
            color: #fff;
        }
        </style>
        <div class="footer">
        <p>Developed by @Rajyug IT Solutions Pvt. Ltd</p>
        </div>
        """,
        unsafe_allow_html=True,
    )


if __name__ == "__main__":
    main()
