import openpyxl
import os
from filesfunction import opfiles
import re

def natural_sort_key(s):
    """自然排序的自定义键函数"""
    return int(re.findall(r'\d+', s)[0])

# 图片文件夹路径
image_folder, parent_folder = opfiles.OpFiles.select_folder()

# 创建一个新的Excel文件
workbook = openpyxl.Workbook()
sheet = workbook.active

# 获取文件夹中所有图片文件名，并按数字排序
image_files = [f for f in os.listdir(image_folder) if f.endswith('.jpg')]
#image_files.sort(key=lambda x: int(x.split('_')[1].split('.')[0]))
image_files.sort(key=natural_sort_key)

# 设置单元格宽度，以适应图片大小
sheet.column_dimensions['A'].width = 200  # 调整宽度值以适应你的图片

# 批量插入图片
row = 1
for img in image_files:
    img_path = os.path.join(image_folder, img)
    img = openpyxl.drawing.image.Image(img_path)
    img.anchor = 'A' + str(row)
    sheet.add_image(img)
    row += 1

# 保存Excel文件
excel_file_name = image_folder.split("/")[-1] + ".xlsx"
images_in_excel_path = os.path.normpath(os.path.join(parent_folder, excel_file_name))
    # 保存Excel文件
workbook.save(images_in_excel_path)