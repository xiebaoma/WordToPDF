import os
from converter.word_to_pdf import word_to_pdf
from converter.excel_to_pdf import excel_to_pdf
from converter.ppt_to_pdf import ppt_to_pdf
from converter.image_to_pdf import image_to_pdf

def convert_file(input_file, output_dir):
    file_extension = os.path.splitext(input_file)[1].lower()
    pdf_file = os.path.join(output_dir, os.path.basename(input_file).replace(file_extension, ".pdf"))

    if file_extension in [".doc", ".docx"]:
        return word_to_pdf(input_file, pdf_file)
    elif file_extension in [".xls", ".xlsx"]:
        return excel_to_pdf(input_file, pdf_file)
    elif file_extension in [".ppt", ".pptx"]:
        return ppt_to_pdf(input_file, pdf_file)
    elif file_extension in [".png", ".jpg", ".jpeg", ".bmp"]:
        return image_to_pdf(input_file, pdf_file)
    else:
        print(f"Unsupported file type: {file_extension}")
        return False
