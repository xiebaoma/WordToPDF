import win32com.client
import os


def ppt_to_pdf(ppt_file, pdf_file):
    powerpoint = None
    presentation = None
    try:
        ppt_file = os.path.abspath(ppt_file)  # Ensure absolute path
        pdf_file = os.path.abspath(pdf_file)  # Ensure absolute path

        # Check if the PowerPoint file exists
        if not os.path.exists(ppt_file):
            print(f"Error: File not found -> {ppt_file}")
            return False

        powerpoint = win32com.client.Dispatch("PowerPoint.Application")
        powerpoint.Visible = 1  # Make PowerPoint visible (set to 0 to hide)

        # Open the PowerPoint file
        presentation = powerpoint.Presentations.Open(ppt_file, WithWindow=False)

        # Save as PDF (format 32 indicates PDF)
        presentation.SaveAs(pdf_file, 32)
        print(f"Successfully converted: {ppt_file} -> {pdf_file}")
        return True

    except Exception as e:
        print(f"Error converting PowerPoint file {ppt_file}: {e}")
        return False

    finally:
        # Ensure the presentation and PowerPoint application are closed
        if presentation:
            presentation.Close()
        if powerpoint:
            powerpoint.Quit()
