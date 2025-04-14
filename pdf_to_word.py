# import fitz  
# from docx import Document
# from docx.shared import Inches
# import os

# def pdf_to_text_images_word(pdf_path, word_path):
#     doc = Document()
#     pdf = fitz.open(pdf_path)
#     image_folder = "pdf_images"  
    
#     if not os.path.exists(image_folder):
#         os.makedirs(image_folder)

#     for page_num, page in enumerate(pdf):
        
#         full_text = " ".join(page.get_text("text").split())  
        
#         if full_text:
#             para = doc.add_paragraph(full_text)  
#             para.alignment = 3  

#         img_list = page.get_images(full=True)
#         for img_index, img in enumerate(img_list):
#             xref = img[0]  
#             base_image = pdf.extract_image(xref)
#             image_bytes = base_image["image"]

#             image_path = f"{image_folder}/image_{page_num}_{img_index}.png"
#             with open(image_path, "wb") as img_file:
#                 img_file.write(image_bytes)

#             doc.add_picture(image_path, width=Inches(5))  
#             doc.add_paragraph(f"[Image from page {page_num + 1}]") 

#     doc.save(word_path)
#     print(f"Conversion completed! Word document saved at: {word_path}")

# pdf_to_text_images_word(
#     "C:/Users/Srijan/Desktop/word translator/sample_word_document.pdf",
#     "C:/Users/Srijan/Desktop/converted_sample_word_document.docx"
# )
# from pdf2docx import Converter
# import os

# def convert_pdf_exactly(pdf_path, word_path):
    
#     if not os.path.exists(pdf_path):
#         print(f"❌ Error: PDF file not found at {pdf_path}")
#         return

#     print(f"🔄 Converting:\n{pdf_path} \n→\n{word_path}")

    
#     try:
#         cv = Converter(pdf_path)
#         cv.convert(word_path, start=0, end=None)  
#         cv.close()
#         print(f"✅ Conversion complete! Word file saved at:\n{word_path}")
#     except Exception as e:
#         print(f"❌ An error occurred during conversion: {e}")


# pdf_file = ""
# word_output = ""

# convert_pdf_exactly(pdf_file, word_output)


# import os
# import fitz  # PyMuPDF
# import pdfplumber
# from pdf2docx import Converter
# from docx import Document
# from docx.shared import Inches

# def has_tables(pdf_path):
#     with pdfplumber.open(pdf_path) as pdf:
#         for page in pdf.pages:
#             tables = page.extract_tables()
#             if any(tables):
#                 return True
#     return False

# def remove_highlights(docx_path):
#     from docx import Document
#     doc = Document(docx_path)
#     for para in doc.paragraphs:
#         for run in para.runs:
#             run.font.highlight_color = None
#     doc.save(docx_path)

# def convert_with_pdf2docx(pdf_path, word_path):
#     print(f"📄 Converting with pdf2docx: {pdf_path}")
#     cv = Converter(pdf_path)
#     cv.convert(word_path, start=0, end=None)
#     cv.close()
#     print(f"✅ Layout/table conversion done: {word_path}")

# def insert_images(pdf_path, word_path):
#     print("🖼️ Adding images using PyMuPDF...")
#     doc = Document(word_path)
#     pdf = fitz.open(pdf_path)
#     image_folder = "pdf_images"
#     os.makedirs(image_folder, exist_ok=True)

#     for page_num, page in enumerate(pdf):
#         for img_index, img in enumerate(page.get_images(full=True)):
#             xref = img[0]
#             base_image = pdf.extract_image(xref)
#             image_bytes = base_image["image"]
#             ext = base_image["ext"]
#             image_path = f"{image_folder}/image_{page_num}_{img_index}.{ext}"
#             with open(image_path, "wb") as f:
#                 f.write(image_bytes)

#             doc.add_picture(image_path, width=Inches(5))
#             doc.add_paragraph(f"[Image from page {page_num + 1}]")

#     doc.save(word_path)
#     print("✅ Images inserted.")

# def convert_with_fitz_text_images(pdf_path, word_path):
#     print("✍️ Converting text + images with PyMuPDF")
#     doc = Document()
#     pdf = fitz.open(pdf_path)
#     image_folder = "pdf_images"
#     os.makedirs(image_folder, exist_ok=True)

#     for page_num, page in enumerate(pdf):
#         # Use layout-aware text block extraction
#         blocks = page.get_text("blocks")
#         blocks.sort(key=lambda b: (b[1], b[0]))  # Sort top-to-bottom, then left-to-right

#         for block in blocks:
#             text = block[4].strip()
#             if text:
#                 para = doc.add_paragraph(text)
#                 para.alignment = 3  # Justify

#         # Extract images
#         for img_index, img in enumerate(page.get_images(full=True)):
#             xref = img[0]
#             base_image = pdf.extract_image(xref)
#             image_bytes = base_image["image"]
#             ext = base_image["ext"]
#             image_path = f"{image_folder}/image_{page_num}_{img_index}.{ext}"
#             with open(image_path, "wb") as f:
#                 f.write(image_bytes)

#             doc.add_picture(image_path, width=Inches(5))
#             doc.add_paragraph(f"[Image from page {page_num + 1}]")

#     doc.save(word_path)
#     print(f"✅ Saved with clean paragraphs + images: {word_path}")

# def smart_pdf_to_word(pdf_path, word_path):
#     if has_tables(pdf_path):
#         convert_with_pdf2docx(pdf_path, word_path)
#     else:
#         convert_with_fitz_text_images(pdf_path, word_path)

#     insert_images(pdf_path, word_path)
#     remove_highlights(word_path)

# === Run it on both PDFs ===

# 1. File with text/images (like two-column article)
# smart_pdf_to_word(
#     "/Users/shayna/Downloads/Why Bharat Matters.pdf",
#     "/Users/shayna/Downloads/WHYBM_FINAL.docx"
# )
import aspose.pdf as apdf
from io import FileIO
from os import path
import logging # Add logging

logger = logging.getLogger(__name__) # Setup logger for this module

def convert_pdf_with_aspose(pdf_path, word_path):
    """Converts a PDF file to DOCX using Aspose.PDF."""
    logger.info(f"Attempting PDF to DOCX conversion using Aspose: {pdf_path} -> {word_path}")
    try:
        document = apdf.Document(pdf_path)
        save_options = apdf.DocSaveOptions()
        save_options.format = apdf.DocSaveOptions.DocFormat.DOC_X
        # Add any other specific Aspose options here if needed
        # e.g., save_options.recognize_bullets = True
        document.save(word_path, save_options)
        logger.info(f"Aspose conversion successful: {word_path}")
    except Exception as e:
        logger.error(f"Aspose PDF conversion failed for {pdf_path}. Error: {e}")
        raise # Re-raise the exception to be caught by the web server

# === Comment out or remove the old standalone code block ===
# path_infile = ""
# path_outfile = ""
# 
# document = apdf.Document(path_infile)
# save_options = apdf.DocSaveOptions()
# save_options.format = apdf.DocSaveOptions.DocFormat.DOC_X
# document.save(path_outfile, save_options)
# === End of commented/removed block ===

# Keep other functions like smart_pdf_to_word if they might be used elsewhere or for comparison
# ... (rest of the original file content, if any) ...