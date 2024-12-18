import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
from tkinter import ttk
from ttkthemes import ThemedTk
import ttkbootstrap as ttk

from PIL import Image, ImageTk
import PyPDF2
import threading
import time
import os
from docx2pdf import convert as doc2pdf_conv


def get_version():
    return "0.0.1"


def insert_pages(input_pdf, insert_before_pdf, insert_after_pdf, output_pdf, progress_var, start_page, root, callback):
    with open(input_pdf, 'rb') as file:
        pdf_reader = PyPDF2.PdfReader(file)
        pdf_writer = PyPDF2.PdfWriter()

        with open(insert_before_pdf, 'rb') as insert_before_file:
            insert_before_pdf_reader = PyPDF2.PdfReader(insert_before_file)

            with open(insert_after_pdf, 'rb') as insert_after_file:
                insert_after_pdf_reader = PyPDF2.PdfReader(insert_after_file)

                total_pages = len(pdf_reader.pages)
                progress_var.set(0)  # 重置进度条值为0

                for page_number in range(total_pages):
                    # 更新进度条
                    progress_var.set((page_number + 1) * 100 / total_pages)
                    root.update_idletasks()
                    root.after(1000)  # 添加小延迟，以观察进度条的变化
                    print(str(progress_var.get()) + "\n")
                    if page_number < start_page:
                        pdf_writer.add_page(pdf_reader.pages[page_number])
                    else:
                        insert_before_page = insert_before_pdf_reader.pages[0]
                        pdf_writer.add_page(insert_before_page)

                        pdf_writer.add_page(pdf_reader.pages[page_number])

                        insert_after_page = insert_after_pdf_reader.pages[0]
                        pdf_writer.add_page(insert_after_page)

                with open(output_pdf, 'wb') as output_file:
                    pdf_writer.write(output_file)
                
                # 调度回调函数，以确保在主线程中执行
                root.after(1, callback)

                
                    
def browse_file(entry_widget, filetypes):
    filename = filedialog.askopenfilename(initialdir="/", title="Select a file", filetypes=filetypes)
    entry_widget.delete(0, tk.END)  # Clear any previous text
    entry_widget.insert(0, filename)  # Insert the selected filename


def save_file():
    output_folder = filedialog.askdirectory()
    output_file = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")], initialdir=output_folder)
    file_label_output_folder_blank.delete(0, tk.END)
    file_label_output_folder_blank.insert(0, output_file)
    return output_file

def run_insert_pages():
    # input_pdf = file_label_input.cget("text")
    # insert_before_pdf = file_label_before.cget("text")
    # insert_after_pdf = file_label_after.cget("text")
    # output_pdf = file_label_output.cget("text")
    
    input_pdf = file_label_input_blank.get()
    insert_before_pdf = file_label_before_blank.get()
    insert_after_pdf = file_label_after_blank.get()
    output_pdf = file_label_output_folder_blank.get()
    start_page = int(start_page_entry.get())

    if input_pdf and insert_before_pdf and insert_after_pdf and output_pdf:
        try:
            progress_var.set(0)  # 重置进度条值为0
            progressbar.start(10)  # 启动进度条动画
            threading.Thread(target=insert_pages, args=(input_pdf, insert_before_pdf, insert_after_pdf, output_pdf, progress_var, start_page, on_insert_pages_complete)).start()
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {str(e)}")
    else:
        messagebox.showerror("Error", "Please select all PDF files and output path.")

def on_insert_pages_complete():
    progressbar.stop()  # 停止进度条
    messagebox.showinfo("Success", "PDF pages inserted successfully!")

def convert_to_pdf(input_path, output_path):
    doc2pdf_conv(input_path,output_path)
    print(f"Conversion completed: {output_path}")
    


def batch_convert_word_to_pdf(input_folder, output_folder):
    # Check if the output folder is provided
    print(f"Converting Word documents in folder: {input_folder}")
    if not output_folder:
        print("Error: Output folder not provided.")
        return

    # Ensure output folder exists
    os.makedirs(output_folder, exist_ok=True)

    # Iterate through all files in the input folder
    for filename in os.listdir(input_folder):
        if filename.endswith(".docx") and not filename.startswith("~$"):
            input_path = os.path.join(input_folder, filename)
            
            # Ensure the output folder for PDFs exists
            pdf_output_folder = os.path.join(output_folder, "pdf_output")
            os.makedirs(pdf_output_folder, exist_ok=True)

            file_name, _ = os.path.splitext(filename)
            output_path = os.path.join(pdf_output_folder, file_name + ".pdf")
            print(str(input_path))
            convert_to_pdf(input_path, output_path)
    messagebox.showinfo("Success", "PDF pages inserted successfully!")
    
def run_batch_convert():
    input_folder = file_label_word_folder_blank.get()
    output_folder = file_label_pdf_folder_blank.get()
    threading.Thread(target=batch_convert_word_to_pdf, args=(input_folder, output_folder)).start()
    
# def set_custom_icon():
#     icon_path = "E:/1.ico"
#     icon_image = Image.open(icon_path)
#     icon_photo = ImageTk.PhotoImage(icon_image)
#     root.tk.call('wm', 'iconphoto', root._w, icon_photo)

def browse_folder(entry_widget):
    folder_path = filedialog.askdirectory()
    entry_widget.delete(0, tk.END)  # 清除之前的文本
    entry_widget.insert(0, folder_path)  # 插入选择的文件夹路径




def run_app():
    current_way = way_sel.get()

    if current_way == '间隔插入':
        run_insert_pages()
    elif current_way == '批量word转pdf':
        run_batch_convert()
        
# Create the main window
root = ThemedTk(theme="equilux")
root.title("PDF Pages Insertion - Version " + get_version())

root.geometry("600x650")  # 设置窗口宽度为400像素，高度为300像素
# set_custom_icon()
# Set the custom icon
# icon_path = "E:/1.ico"
# root.iconbitmap(icon_path)  # Replace 'path_to_your_icon.ico' with the path to your icon file


# 创建并放置标签和按钮用于选择 PDF 文件和输出路径
current_row = 0
current_col = 0

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
run_button = tk.Button(root, text="运行", command=run_app)
run_button.grid(row=current_row, column=0, columnspan=2, sticky=tk.W + tk.E)

progress_var = ttk.DoubleVar()
progressbar = ttk.Progressbar(bootstyle="success", mode="determinate", variable=progress_var)
progressbar.grid(row=current_row+1, column=0, columnspan=5, sticky=tk.W + tk.E)

tk.Label(text="").grid(row=current_row+2)
tk.Label(text="").grid(row=current_row+3)
current_row +=4
save_button = tk.Button(root, text="另存为", command=save_file)
save_button.grid(row=current_row, column=0, columnspan=2, sticky=tk.W + tk.E)

# 启动主循环
root.mainloop()
