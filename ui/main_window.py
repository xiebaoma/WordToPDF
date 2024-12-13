import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinterdnd2 import TkinterDnD, DND_FILES
from converter.file_converter import convert_file


def start_app():
    root = TkinterDnD.Tk()  # 使用 TkinterDnD 提供的 Tk 类
    root.title("File to PDF Converter")
    root.geometry("500x400")
    root.configure(bg="#f0f0f0")  # 设置背景颜色

    # 标题标签
    title_label = tk.Label(root, text="File to PDF Converter", font=("Arial", 16, "bold"), bg="#f0f0f0")
    title_label.pack(pady=10)

    # 提示标签
    label = tk.Label(root, text="Drag and drop files here or click to browse", font=("Arial", 12), bg="#f0f0f0")
    label.pack(pady=5)

    # 文件列表框
    file_list_frame = tk.Frame(root, bg="#ffffff", bd=2, relief=tk.SUNKEN)
    file_list_frame.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)

    file_list = tk.Listbox(file_list_frame, font=("Arial", 10), selectmode=tk.SINGLE, bg="#ffffff")
    file_list.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

    scrollbar = tk.Scrollbar(file_list_frame, orient=tk.VERTICAL, command=file_list.yview)
    scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
    file_list.configure(yscrollcommand=scrollbar.set)

    default_output_dir = "converter/output"  # 默认输出目录
    os.makedirs(default_output_dir, exist_ok=True)

    def on_drop(event):
        files = root.tk.splitlist(event.data)  # 获取拖放的文件路径列表
        for file in files:
            if os.path.isfile(file):
                file_list.insert(tk.END, file)  # 显示到文件列表
        messagebox.showinfo("Files Added", "Files have been added to the list.")

    def open_file_dialog():
        filetypes = [
            ("Supported files", "*.docx;*.doc;*.xlsx;*.xls;*.pptx;*.ppt;*.png;*.jpg;*.jpeg;*.bmp"),
            ("All files", "*.*"),
        ]
        files = filedialog.askopenfilenames(filetypes=filetypes)
        for file in files:
            if os.path.isfile(file):
                file_list.insert(tk.END, file)  # 显示到文件列表

    def choose_output_directory():
        output_dir = filedialog.askdirectory(title="Select Output Directory")
        if output_dir:
            return output_dir
        return default_output_dir  # 如果用户没有选择路径，则使用默认路径

    def convert_files():
        if file_list.size() == 0:
            messagebox.showwarning("No Files", "Please add some files to convert.")
            return

        # 获取用户选择的输出目录
        output_dir = choose_output_directory()

        for idx in range(file_list.size()):
            file = file_list.get(idx)
            if os.path.isfile(file):
                success = convert_file(file, output_dir)
                if success:
                    print(f"Converted: {file}")

        # 转换完成后清空文件列表
        file_list.delete(0, tk.END)

        messagebox.showinfo("Success", f"All supported files have been converted and saved in {output_dir}!")

    # 按钮
    button_frame = tk.Frame(root, bg="#f0f0f0")
    button_frame.pack(pady=10)

    browse_button = ttk.Button(button_frame, text="Browse Files", command=open_file_dialog)
    browse_button.pack(side=tk.LEFT, padx=5)

    convert_button = ttk.Button(button_frame, text="Convert to PDF", command=convert_files)
    convert_button.pack(side=tk.LEFT, padx=5)

    # 输出路径标签
    output_label = tk.Label(root, text=f"Output Directory: {default_output_dir}", font=("Arial", 10), bg="#f0f0f0")
    output_label.pack(pady=5)

    # 绑定拖放事件
    root.drop_target_register(DND_FILES)
    root.dnd_bind('<<Drop>>', on_drop)

    root.mainloop()

# 如果需要运行，请调用 start_app()
