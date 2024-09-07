import json
import openpyxl

# 读取JSON文件
with open("C:\\Users\\admin\\Desktop\\result(1).json", 'r', encoding='utf-8') as f:
    data = json.load(f)

# 创建一个新的Excel工作簿和工作表
workbook = openpyxl.Workbook()
sheet = workbook.active

# 将字符写入Excel的C列
row = 1
for char, coordinates in data.items():
    sheet.cell(row=row, column=5, value=coordinates)
    row += 1

# 保存Excel文件
workbook.save('output2.xlsx')