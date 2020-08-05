from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.units import cm
from PyPDF2 import PdfFileWriter, PdfFileReader
import xlrd
import os

# 注册中文字体
pdfmetrics.registerFont(TTFont('SimSun', 'SimSun.ttf'))


# 水印图片（领航）
picture_path = 'watermark.png'
# 未加水印的原始文件放在此folder中
folder_path = './original_file_folder'
# 文件输出路径
out_put_path = './out_put_folder'
# 学生信息文件路径
student_name_path= './student_list'
# 临时文件路径
temp_file_path = './temp_file_folder'
# 学生信息文件路径及文件名
loc = student_name_path + '/student name.xlsx'


# 读取Excel中的学生信息
wb = xlrd.open_workbook(loc)
sheet = wb.sheet_by_index(0)

# Extracting number of rows
for i in range(sheet.nrows):
    # skip the first row
    if i == 0:
        continue

    # retrieve student information
    student_info = sheet.cell_value(i, 0)
    #要生成的PDF文件名字 for example, 张三.pdf
    pdf_name = temp_file_path + '/' + student_info + '.pdf'

    c = canvas.Canvas(pdf_name)

    if student_info:
        c.setFontSize(22)
        c.setFont('SimSun', 36)
        # 指定填充颜色
        c.setFillColorRGB(0.6, 0, 0)
        # 设置透明度，1为不透明
        c.setFillAlpha(0.1)
        c.drawString(15, 15, student_info)

    if picture_path:
        c.translate(5 * cm, 2.5 * cm)
        c.rotate(45)
        c.drawImage(picture_path,  15, 15, 600, 120)

    # 生成临时的 pdf
    c.save()

    # 读取临时生成的 pdf
    watermark = PdfFileReader(open(pdf_name, "rb"))

    # 读取原始文件夹中的所有pdf
    for file in os.listdir(folder_path):
        if file.endswith(".pdf"):
            output_file = PdfFileWriter()
            input_file = PdfFileReader(open(folder_path + '/' + file, "rb"))

            page_count = input_file.getNumPages()
            for page_number in range(page_count):
                input_page = input_file.getPage(page_number)
                input_page.mergePage(watermark.getPage(0))
                output_file.addPage(input_page)

            # 要输出的文件路径及名称
            output_path = out_put_path + '/' + file.split('.pdf')[0] + '_' + student_info + '_watermarked' + '.pdf'
            with open(output_path, "wb") as outputStream:
                output_file.write(outputStream)



# 删除临时生成的pdf??

