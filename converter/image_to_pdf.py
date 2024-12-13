from PIL import Image
import os

def image_to_pdf(image_file, pdf_file):
    """
    Converts an image file to a PDF file.

    :param image_file: Path to the image file to be converted.
    :param pdf_file: Path to save the converted PDF file.
    :return: True if the conversion was successful, False otherwise.
    """
    try:
        image_file = os.path.abspath(image_file)  # Ensure absolute path
        pdf_file = os.path.abspath(pdf_file)  # Ensure absolute path

        # Check if the image file exists
        if not os.path.exists(image_file):
            print(f"Error: File not found -> {image_file}")
            return False

        # Open and convert the image
        image = Image.open(image_file)
        if image.mode != 'RGB':
            image = image.convert('RGB')  # Ensure it is in RGB mode

        # Save the image as a PDF
        image.save(pdf_file, "PDF")
        print(f"Successfully converted: {image_file} -> {pdf_file}")
        return True

    except Exception as e:
        print(f"Error converting Image file {image_file}: {e}")
        return False
