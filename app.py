from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.units import cm
from PyPDF2 import PdfFileWriter, PdfFileReader
from pathlib import Path
import xlrd
import os
import shutil


# 注册中文字体
pdfmetrics.registerFont(TTFont('SimSun', 'SimSun.ttf'))


# 水印图片路径
picture_path = 'watermark.png'
# 未加水印的原始文件放在此folder中
folder_path = './original_file_folder'
# 文件输出路径。制作好的文件输出到这个文件夹中
out_put_path = './out_put_folder'
# 学生信息文件路径，
student_name_path= './student_list'
# 学生信息文件路径及文件名。手动建立一个excel表格，名字为：student name.xlsx.
# 里面存放水印文字信息
loc = student_name_path + '/student name.xlsx'
# 临时文件路径，
temp_file_path = './temp_file_folder'

#创建路径
Path(folder_path).mkdir(parents=True, exist_ok=True)
Path(out_put_path).mkdir(parents=True, exist_ok=True)
Path(student_name_path).mkdir(parents=True, exist_ok=True)
Path(temp_file_path).mkdir(parents=True, exist_ok=True)


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
            # 要输出的文件路径及名称
            output_file_name = file.split('.pdf')[0] + '_' + student_info + '_watermarked' + '.pdf'
            print('正在生成：' + output_file_name)
            print('请等待……')
            output_file = PdfFileWriter()
            input_file = PdfFileReader(open(folder_path + '/' + file, "rb"))

            page_count = input_file.getNumPages()
            for page_number in range(page_count):
                input_page = input_file.getPage(page_number)
                input_page.mergePage(watermark.getPage(0))
                output_file.addPage(input_page)

            # 要输出的文件路径及名称
            output_path = out_put_path + '/' + output_file_name
            with open(output_path, "wb") as outputStream:
                output_file.write(outputStream)
            print('生成：' + output_file_name)
            print('....................................................')

print('任务完毕')

# 删除临时生成的pdf
shutil.rmtree(temp_file_path)
