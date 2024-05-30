import os
from PyPDF2 import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter


def add_watermark(input_file, output_file):
    # Create a PDF writer object
    pdf_writer = PdfWriter()

    # Open the input PDF file
    with open(input_file, 'rb') as pdf_file:
        pdf_reader = PdfReader(pdf_file)
        num_pages = len(pdf_reader.pages)

        # Create a PDF canvas for drawing the watermark
        watermark_canvas = canvas.Canvas('watermark.pdf', pagesize=letter)
        watermark_canvas.setFont('Helvetica', 80)
        watermark_canvas.setFillGray(0.5)

        # Draw the watermark text diagonally across the page
        watermark_canvas.saveState()
        watermark_canvas.translate(340, 355)
        watermark_canvas.rotate(45)
        watermark_canvas.setFillAlpha(0.5)
        watermark_canvas.drawCentredString(0, 0, "PRELIMINARY")
        watermark_canvas.restoreState()
        watermark_canvas.save()

        # Add the watermark to each page of the input PDF file
        for page_num in range(num_pages):
            page = pdf_reader.pages[page_num]
            page.merge_page(PdfReader('watermark.pdf').pages[0])
            pdf_writer.add_page(page)

    # Write the output PDF file with watermark
    with open(output_file, 'wb') as output_pdf:
        pdf_writer.write(output_pdf)

    # Remove the temporary watermark file
    os.remove('watermark.pdf')


# Iterate over all PDF files in the folder and add watermark to each
def add_directory_watermark(input_folder_path, output_folder_path):
    """Adds a watermark to each file in a given directory, saves the watermarked file into a new directory,
    then deletes the original file."""
    for file_name in os.listdir(input_folder_path):
        if file_name.endswith('.pdf'):
            input_file_path = os.path.join(input_folder_path, file_name)
            output_file_path = os.path.join(output_folder_path, f'PRELIMINARY_{file_name}')
            add_watermark(input_file_path, output_file_path)
            # check if the line below works
            # os.remove(input_file_path)

    print('Watermark added to all PDF files in the folder.')
