import win32com.client
import os

def word_to_pdf(word_file, pdf_file):
    word = None
    doc = None
    try:
        word_file = os.path.abspath(word_file)  # 确保是绝对路径
        pdf_file = os.path.abspath(pdf_file)  # 确保是绝对路径

        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False  # 隐藏 Word 窗口

        # 检查文件是否存在
        if not os.path.exists(word_file):
            print(f"Error: File not found -> {word_file}")
            return False

        doc = word.Documents.Open(word_file)
        wdFormatPDF = 17  # 指定 PDF 格式代码
        doc.SaveAs(pdf_file, FileFormat=wdFormatPDF)
        print(f"Successfully converted: {word_file} -> {pdf_file}")
        return True
    except Exception as e:
        print(f"Error converting {word_file}: {e}")
        return False
    finally:
        # 确保文档和应用程序被关闭
        if doc:
            doc.Close()
        if word:
            word.Quit()
