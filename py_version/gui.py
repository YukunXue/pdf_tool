import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
from tkinter import ttk
from ttkthemes import ThemedTk
import ttkbootstrap as ttk
import threading
from doc2pdf import batch_convert_word_to_pdf, batch_convert_pdf_to_word
from utils import get_version, update_progress
from pdf_op import insert_pages
import time

# 声明为全局变量
current_row = 0
current_col = 0
file_label_input_blank = None
file_label_before_blank = None
file_label_after_blank = None
file_label_output_folder_blank = None
file_label_word_folder_blank = None
file_label_pdf_folder_blank = None


# Global variables for progress
global_progress_var = None
global_progressbar = None

def initialize_globals():
    global global_progress_var, global_progressbar
    global_progress_var = ttk.DoubleVar()
    global_progressbar = ttk.Progressbar(bootstyle="success", mode="determinate", variable=global_progress_var)

def save_file():
    output_folder = filedialog.askdirectory()
    output_file = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")], initialdir=output_folder)
    file_label_output_folder_blank.delete(0, tk.END)
    file_label_output_folder_blank.insert(0, output_file)
    return output_file

def browse_folder(entry_widget):
    folder_path = filedialog.askdirectory()
    entry_widget.delete(0, tk.END)  # 清除之前的文本
    entry_widget.insert(0, folder_path)  # 插入选择的文件夹路径

def browse_file(entry_widget, filetypes):
    filename = filedialog.askopenfilename(initialdir="/", title="Select a file", filetypes=filetypes)
    entry_widget.delete(0, tk.END)  # Clear any previous text
    entry_widget.insert(0, filename)  # Insert the selected filename

def on_insert_pages_complete(progressbar):
    progressbar.stop()  # 停止进度条
    messagebox.showinfo("Success", "PDF pages inserted successfully!")

def run_batch_convert_pdf():
    input_folder = file_label_word_folder_blank.get()
    output_folder = file_label_pdf_folder_blank.get()
    threading.Thread(target=batch_convert_word_to_pdf, args=(input_folder, output_folder)).start()

def run_batch_convert_word():
    input_folder = file_label_word_folder_blank.get()
    output_folder = file_label_pdf_folder_blank.get()
    threading.Thread(target=batch_convert_pdf_to_word, args=(input_folder, output_folder)).start()

def run_insert_pages(start_page, root):    
    input_pdf = file_label_input_blank.get()
    insert_before_pdf = file_label_before_blank.get()
    insert_after_pdf = file_label_after_blank.get()
    output_pdf = file_label_output_folder_blank.get()
    print(str(insert_before_pdf))
    print(str(insert_after_pdf))
    print(str(output_pdf))
    print(str(input_pdf))

    if input_pdf and insert_before_pdf and insert_after_pdf and output_pdf:
        try:
            global_progress_var.set(0)  # Reset global progress bar value to 0
            global_progressbar.start(10)  # Start global progress bar animation
            threading.Thread(target=insert_pages, args=(input_pdf, insert_before_pdf, insert_after_pdf, output_pdf, start_page, global_progress_var, global_progressbar, root)).start()
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {str(e)}")
    else:
        messagebox.showerror("Error", "Please select all PDF files and output path.")

def run_app(current_way, start_page, root):
    if current_way == '间隔插入':
        run_insert_pages(start_page, root)
    elif current_way == '批量word转pdf':
        run_batch_convert_pdf()
    elif current_way == '批量pdf转word':
        run_batch_convert_word()
        

def create_gui(current_row):
    global file_label_input_blank, file_label_before_blank, file_label_after_blank, \
           file_label_output_folder_blank, file_label_word_folder_blank, file_label_pdf_folder_blank
    # Create the main window
    root = ThemedTk(theme="equilux")
    root.title("PDF Pages Insertion - Version " + get_version())

    root.geometry("600x650")  # 设置窗口宽度为400像素，高度为300像素
    # set_custom_icon()
    # Set the custom icon
    # icon_path = "E:/1.ico"
    # root.iconbitmap(icon_path)  # Replace 'path_to_your_icon.ico' with the path to your icon file

    way_sel_entry = tk.Label(root, text="操作方式:")
    way_sel_entry.grid(row=current_row, column=0)
    way_sel = ttk.Combobox(bootstyle="primary")
    way_sel.config(state="readonly")
    way_sel.grid(row=current_row, column=2)
    # 更新选项
    options = ['间隔插入', '批量word转pdf', '批量pdf转word']
    way_sel['values'] = options
    # 设置默认选项
    way_sel.set(options[0])
    # 创建并放置标签和按钮用于选择 PDF 文件和输出路径
    current_row +=5
    file_label_input = tk.Label(root, text="选择输入 PDF 文件:")
    file_label_input.grid(row=current_row, column=0, sticky=tk.E)
    file_label_input_blank = ttk.Entry(bootstyle="info")
    file_label_input_blank.grid(row=current_row, column=2, sticky=tk.E)

    input_button = tk.Button(root, text="浏览", command=lambda: browse_file(file_label_input_blank, [("PDF files", "*.pdf")]))
    input_button.grid(row=current_row, column=4)

    tk.Label(text="").grid(row=1)


    start_page_label = tk.Label(root, text="开始插入的页数:")
    start_page_label.grid(row=current_row, column=8, sticky=tk.E)

    start_page_entry = ttk.Entry(bootstyle="info")
    start_page_entry.grid(row=current_row, column=10, columnspan=2, sticky=tk.E)
    start_page_entry.insert(0, "18")

    current_row+=2

    file_label_before = tk.Label(root, text="选择插入前 PDF 文件:")
    file_label_before.grid(row=current_row, column=0, sticky=tk.E)
    file_label_before_blank = ttk.Entry(bootstyle="info")
    file_label_before_blank.grid(row=current_row, column=2, sticky=tk.E)

    before_button = tk.Button(root, text="浏览", command=lambda: browse_file(file_label_before_blank, [("PDF files", "*.pdf")]))
    before_button.grid(row=current_row, column=4)

    tk.Label(text="").grid(row=3)

    current_row +=2

    file_label_after = tk.Label(root, text="选择插入后 PDF 文件:")
    file_label_after.grid(row=current_row, column=0, sticky=tk.E)
    file_label_after_blank = ttk.Entry(bootstyle="info")
    file_label_after_blank.grid(row=current_row, column=2, sticky=tk.E)

    after_button = tk.Button(root, text="浏览", command=lambda: browse_file(file_label_after_blank, [("PDF files", "*.pdf")]))
    after_button.grid(row=current_row, column=4)

    tk.Label(text="").grid(row=current_row+1)

    current_row +=2

    file_label_output_folder = tk.Label(root, text="选择输出 PDF 文件夹:")
    file_label_output_folder.grid(row=current_row, column=0, sticky=tk.E)
    file_label_output_folder_blank = ttk.Entry(bootstyle="info")
    file_label_output_folder_blank.grid(row=current_row, column=2, sticky=tk.E)

    output_folder_button = tk.Button(root, text="浏览", command=lambda: browse_folder(file_label_output_folder_blank))
    output_folder_button.grid(row=current_row, column=4)


    # 创建并放置标签和按钮用于选择 Word 文件夹和输出 PDF 文件夹
    tk.Label(text="").grid(row=current_row + 2)
    current_row += 8
    tk.Label(text="批量转换：").grid(row=current_row)
    current_row += 4
    file_label_word_folder = tk.Label(root, text="选择输入 Word 文件夹:")
    file_label_word_folder.grid(row=current_row, column=0, sticky=tk.E)
    file_label_word_folder_blank = ttk.Entry(bootstyle="info")
    file_label_word_folder_blank.grid(row=current_row, column=2, sticky=tk.E)

    word_folder_button = tk.Button(root, text="浏览", command=lambda: browse_folder(file_label_word_folder_blank))
    word_folder_button.grid(row=current_row, column=4)

    tk.Label(text="").grid(row=current_row + 1)
    current_row += 2

    file_label_pdf_folder = tk.Label(root, text="选择输出 PDF 文件夹:")
    file_label_pdf_folder.grid(row=current_row, column=0, sticky=tk.E)
    file_label_pdf_folder_blank = ttk.Entry(bootstyle="info")
    file_label_pdf_folder_blank.grid(row=current_row, column=2, sticky=tk.E)

    pdf_folder_button = tk.Button(root, text="浏览", command=lambda: browse_folder(file_label_pdf_folder_blank))
    pdf_folder_button.grid(row=current_row, column=4)

    # tk.Label(text="").grid(row=current_row + 1)
    # current_row += 2

    # run_batch_convert_button = tk.Button(root, text="批量转换 Word 到 PDF", command=run_batch_convert)
    # run_batch_convert_button.grid(row=current_row, column=0, columnspan=2, sticky=tk.W + tk.E)

    tk.Label(text="").grid(row=current_row+1)
    current_row +=2
    run_button = tk.Button(root, text="运行", command=lambda:run_app(way_sel.get(), int(start_page_entry.get()), root))
    run_button.grid(row=current_row, column=0, columnspan=2, sticky=tk.W + tk.E)

    initialize_globals()
    global_progressbar.grid(row=current_row+1, column=0, columnspan=5, sticky=tk.W + tk.E)

    
    tk.Label(text="").grid(row=current_row+2)
    tk.Label(text="").grid(row=current_row+3)
    current_row +=4
    save_button = tk.Button(root, text="另存为", command=lambda:save_file)
    save_button.grid(row=current_row, column=0, columnspan=2, sticky=tk.W + tk.E)

    # 启动主循环
    root.mainloop()
    
# Entry point for the program
if __name__ == "__main__":
    create_gui(current_row)
