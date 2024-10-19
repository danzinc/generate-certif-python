from reportlab.pdfgen import canvas
from PyPDF2 import PdfReader, PdfWriter
import pandas as pd
import copy
from decimal import Decimal

# import excel file
excel_path = "data/participant.xlsx"
list=pd.read_excel(excel_path)
# extract data dari excel ambil list namanya 
participant_names = list['Name'].tolist()

#  import template pdf nya 
pdf_path = "data/template.pdf" 
reader = PdfReader(pdf_path)
original_page = reader.pages[0] 
 
#  generate certificate looping berdasarkan nama dari template excel
for name in participant_names:
    # untuk mendapatkan lebar dan tinggi halaman dan template pdf
    packet = canvas.Canvas('pdf/output.pdf', pagesize=(original_page.mediabox.width, original_page.mediabox.height))
    page_width = float(reader.pages[0].mediabox.width)
    page_height = float(reader.pages[0].mediabox.height)

    # ini untuk menggambar certificate dan menyimpan nama ke dalam sertif
    packet.setFont("Helvetica", 24)
    text_width = packet.stringWidth(name, "Helvetica", 24)
    x_position = (page_width - text_width) / 2
    y_position = page_height / 2
    packet.drawString(x_position, y_position, name) 

    packet.save()
 
    temp_pdf = PdfReader('pdf/output.pdf')
    page = copy.copy(original_page)
    page.merge_page(temp_pdf.pages[0])
 
    output_path = f'pdf/Certificate_{name.replace(" ", "_")}.pdf'
    writer = PdfWriter()
    writer.add_page(page)
    with open(output_path, 'wb') as output_file:
        writer.write(output_file)


print("Certificates generated successfully!")