import openpyxl
import os

# 图片文件夹路径
image_folder = "C:\\Users\\admin\\Desktop\\all_font_img"  # 替换为你的图片文件夹路径

# 创建一个新的Excel文件
workbook = openpyxl.Workbook()
sheet = workbook.active

# 获取文件夹中所有图片文件名，并按数字排序
image_files = [f for f in os.listdir(image_folder) if f.endswith('.jpg')]
image_files.sort(key=lambda x: int(x.split('_')[1].split('.')[0]))

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
workbook.save("images_in_excel.xlsx")