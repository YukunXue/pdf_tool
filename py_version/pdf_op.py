from utils import update_progress
from tkinter import messagebox
import threading
import PyPDF2

def insert_pages(input_pdf, insert_before_pdf, insert_after_pdf, output_pdf, progress_var, start_page, callback, root):
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
                    
                    update_progress(progress_var, progress_var.get(), root, "Insert in progress...")
                    
                with open(output_pdf, 'wb') as output_file:
                    pdf_writer.write(output_file)
                
                # 调度回调函数，以确保在主线程中执行
                root.after(1, callback)


                
