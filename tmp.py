import PyPDF2

def remove_last_page(input_pdf, output_pdf):
    with open(input_pdf, 'rb') as file:
        pdf_reader = PyPDF2.PdfReader(file)
        num_pages = len(pdf_reader.pages)

        if num_pages <= 1:
            print("The PDF has only one page or is empty. Nothing to remove.")
            return

        pdf_writer = PyPDF2.PdfWriter()

        # Add pages to the new PDF except the last one
        for page_num in range(num_pages - 1):
            page = pdf_reader.pages[page_num]
            pdf_writer.add_page(page)

        # Write the modified PDF to the output file
        with open(output_pdf, 'wb') as output_file:
            pdf_writer.write(output_file)

if __name__ == "__main__":
    input_pdf_path = "1.pdf"  # Replace with your input PDF file path
    output_pdf_path = "2.pdf"  # Replace with your output PDF file path

    remove_last_page(input_pdf_path, output_pdf_path)
    print(f"Last page removed. Result saved to {output_pdf_path}.")
