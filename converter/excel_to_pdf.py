import win32com.client
import os

def excel_to_pdf(excel_file, pdf_file):
    """
    Converts an Excel file to a PDF file.

    :param excel_file: Path to the Excel file to be converted.
    :param pdf_file: Path to save the converted PDF file.
    :return: True if the conversion was successful, False otherwise.
    """
    excel = None
    workbook = None
    try:
        excel_file = os.path.abspath(excel_file)  # Ensure absolute path
        pdf_file = os.path.abspath(pdf_file)  # Ensure absolute path

        # Check if the Excel file exists
        if not os.path.exists(excel_file):
            print(f"Error: File not found -> {excel_file}")
            return False

        # Start Excel application
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False  # Hide Excel window

        # Open the Excel file
        workbook = excel.Workbooks.Open(excel_file)

        # Export to PDF (0 represents PDF format)
        workbook.ExportAsFixedFormat(0, pdf_file)
        print(f"Successfully converted: {excel_file} -> {pdf_file}")
        return True

    except Exception as e:
        print(f"Error converting Excel file {excel_file}: {e}")
        return False

    finally:
        # Ensure workbook and application are properly closed
        if workbook:
            workbook.Close(False)
        if excel:
            excel.Quit()
