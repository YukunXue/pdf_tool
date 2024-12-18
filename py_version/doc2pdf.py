from docx2pdf import convert as doc2pdf_conv
from pdf2docx import Converter as pdf2doc_conv
import os
from tkinter import messagebox
from utils import update_progress

def convert_to_pdf(input_path, output_path):
    doc2pdf_conv(input_path,output_path)
    print(f"Conversion completed: {output_path}")

def batch_convert_word_to_pdf(input_folder, output_folder, progress_var, root):
    # Check if the output folder is provided
    print(f"Converting Word documents in folder: {input_folder}")
    if not output_folder:
        print("Error: Output folder not provided.")
        return
    # 获取所有 Word 文件数量
    word_files = [filename for filename in os.listdir(input_folder) if filename.endswith(".docx")]
    total_word_files = len(word_files)
    # Ensure output folder exists
    os.makedirs(output_folder, exist_ok=True)

    # 初始化进度条
    current_progress = 0
    # Iterate through all files in the input folder
    for filename in os.listdir(input_folder):
        if filename.endswith(".docx") and not filename.startswith("~$"):
            input_path = os.path.join(input_folder, filename)
            
            # Ensure the output folder for PDFs exists
            pdf_output_folder = os.path.join(output_folder, "pdf_output")
            os.makedirs(pdf_output_folder, exist_ok=True)

            file_name, _ = os.path.splitext(filename)
            output_path = os.path.join(pdf_output_folder, file_name + ".pdf")
            convert_to_pdf(input_path, output_path)
            # 更新进度条
            current_progress += 1
            progress = (current_progress / total_word_files) * 100
            update_progress(progress_var, progress, root, "Conversion in progress...")
    messagebox.showinfo("Success", "PDF Convert successfully!")
    
def convert_to_word(input_path, output_path):
    cv = pdf2doc_conv(input_path)
    cv.convert(output_path, start=0, end=None)
    cv.close()

def batch_convert_pdf_to_word(input_folder, output_folder, progress_var, root):
    # Check if the output folder is provided
    print(f"Converting Word documents in folder: {input_folder}")
    if not output_folder:
        print("Error: Output folder not provided.")
        return
    # 获取所有 Word 文件数量
    word_files = [filename for filename in os.listdir(input_folder) if filename.endswith(".docx")]
    total_word_files = len(word_files)
    # Ensure output folder exists
    os.makedirs(output_folder, exist_ok=True)

    # 初始化进度条
    current_progress = 0
    # Iterate through all files in the input folder
    for filename in os.listdir(input_folder):
        if filename.endswith(".pdf") and not filename.startswith("~$"):
            input_path = os.path.join(input_folder, filename)
            
            # Ensure the output folder for PDFs exists
            pdf_output_folder = os.path.join(output_folder, "pdf_output")
            os.makedirs(pdf_output_folder, exist_ok=True)

            file_name, _ = os.path.splitext(filename)
            output_path = os.path.join(pdf_output_folder, file_name + ".word")
            convert_to_word(input_path, output_path)
            # 更新进度条
            current_progress += 1
            progress = (current_progress / total_word_files) * 100
            update_progress(progress_var, progress, root, "Conversion in progress...")
    messagebox.showinfo("Success", "WORD Convert successfully!")