from reportlab.pdfgen import canvas
from PyPDF2 import PdfFileWriter, PdfFileReader
import os


from reportlab.platypus import SimpleDocTemplate, Image, Paragraph
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

pdfmetrics.registerFont(TTFont('SimSun', 'SimSun.ttf'))  #注册字体


# Program to extract number
# of rows using Python
import xlrd





"""
https://blog.51cto.com/walkerqt/1378142
Refer to an image if you want to add an image to a watermark.
Fill in text if you want to watermark with text.
Alternatively, following settings will skip this.
picture_path = None
text = None
"""


picture_path = ''
text = '领航'

# Folder in which PDF files will be watermarked. (Could be shared folder)
folder_path = './original_file_folder'
out_put_path = './out_put_folder'
student_name_path= './student_list'

# Give the location of the file
loc = student_name_path + '/student name.xlsx'

wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(0)
sheet.cell_value(0, 0)

# Extracting number of rows
print(sheet.nrows)
for i in range(sheet.nrows):
    # skip the first rrow
    if i == 0:
        continue

    # retrieve student information
    print(sheet.cell_value(i, 0))

    text = sheet.cell_value(i, 0)
    pdf_name = text + '.pdf'

    c = canvas.Canvas(pdf_name)

# if picture_path:
#     c.drawImage(picture_path, 15, 15)

    if text:
        c.setFontSize(22)
        c.setFont('SimSun', 36)
        # 指定填充颜色
        c.setFillColorRGB(0.6, 0, 0)
        # 设置透明度，1为不透明
        c.setFillAlpha(0.3)
        c.drawString(15, 15, text)

    c.save()

    watermark = PdfFileReader(open(pdf_name, "rb"))

    for file in os.listdir(folder_path):
        if file.endswith(".pdf"):

            output_file = PdfFileWriter()
            input_file = PdfFileReader(open(folder_path + '/' + file, "rb"))

            page_count = input_file.getNumPages()
            for page_number in range(page_count):
                input_page = input_file.getPage(page_number)
                input_page.mergePage(watermark.getPage(0))
                output_file.addPage(input_page)

            output_path = out_put_path + '/' + file.split('.pdf')[0] + '_' + pdf_name + '_watermarked' + '.pdf'
            with open(output_path, "wb") as outputStream:
                output_file.write(outputStream)







from reportlab.lib.units import cm


def create_watermark(content):
    # 默认大小为21cm*29.7cm
    c = canvas.Canvas("mark.pdf", pagesize=(30 * cm, 30 * cm))
    # 移动坐标原点(坐标系左下为(0,0))
    c.translate(10 * cm, 5 * cm)

    # 设置字体
    c.setFont('SimSun', 80)
    # 指定描边的颜色
    #c.setStrokeColorRGB(0, 1, 0)
    # 指定填充颜色
    # c.setFillColorRGB(0, 1, 0)
    # 画一个矩形
    # c.rect(cm, cm, 7 * cm, 17 * cm, fill=1)

    # 旋转45度，坐标系被旋转
    c.rotate(45)
    # 指定填充颜色
    c.setFillColorRGB(0.6, 0, 0)
    # 设置透明度，1为不透明
    c.setFillAlpha(0.3)
    # 画几个文本，注意坐标系旋转的影响
    c.drawString(3 * cm, 0 * cm, content)

    #c.setFillAlpha(0.6)
    # c.drawString(6 * cm, 3 * cm, content)
    # c.setFillAlpha(1)
    # c.drawString(9 * cm, 6 * cm, content)

    # 关闭并保存pdf文件
    c.save()


create_watermark('领航')