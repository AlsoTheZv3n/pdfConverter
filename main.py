import os

from docx2pdf import convert
from fpdf import FPDF
import fitz  # PyMuPDF


def convert_to_pdf(file_path):
    file_name, file_ext = os.path.splitext(file_path)
    if file_ext.lower() == '.pdf':
        print("Already a PDF file.")
    elif file_ext.lower() == '.docx':
        try:
            convert(file_path)
            print("File converted to PDF successfully!")
        except Exception as e:
            print(f"Error converting file to PDF: {e}")
    else:
        print("Unsupported file format for conversion.")


def convert_to_docx(file_path):
    file_name, file_ext = os.path.splitext(file_path)
    if file_ext.lower() == '.pdf':
        try:
            doc = fitz.open(file_path)
            text = ''
            for page_num in range(len(doc)):
                page = doc.load_page(page_num)
                text += page.get_text()

            docx_output_path = file_name + '.docx'
            with open(docx_output_path, 'w', encoding='utf-8') as docx_file:
                docx_file.write(text)
            print("PDF converted to DOCX successfully!")
        except Exception as e:
            print(f"Error converting PDF to DOCX: {e}")
    else:
        print("Unsupported file format for conversion.")


def main():
    input_file = input("Enter the file path to convert: ")
    if not os.path.exists(input_file):
        print("Input file does not exist.")
        return

    output_format = input("Enter the output format (pdf or docx): ")

    if output_format.lower() == 'pdf':
        convert_to_pdf(input_file)
    elif output_format.lower() == 'docx':
        convert_to_docx(input_file)
    else:
        print("Unsupported output format.")


if __name__ == "__main__":
    main()
