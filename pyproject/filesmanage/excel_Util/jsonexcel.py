#json对的值写入excel
import json
import openpyxl
import tkinter as tk
from tkinter import filedialog
import os

root = tk.Tk()
root.withdraw()
file_path = filedialog.askopenfilename(filetypes=[("JSON文件", "*.json")])

if not file_path:
    print("未选择文件")
else:
    parent_folder = os.path.dirname(file_path)

     # 读取JSON文件
    with open(file_path, 'r', encoding='utf-8') as f:
        data = json.load(f)

    # 创建一个新的Excel工作簿和工作表
    workbook = openpyxl.Workbook()
    sheet = workbook.active

    # 将字符写入Excel的C列
    row = 1
    for char, coordinates in data.items():
        sheet.cell(row=row, column=3, value=coordinates)
        row += 1

    # 获取JSON文件名，并生成对应的Excel文件名
    excel_file_name = file_path.split("/")[-1].split(".")[0] + ".xlsx"
    images_in_excel_path = os.path.normpath(os.path.join(parent_folder, excel_file_name))

    # 保存Excel文件
    workbook.save(images_in_excel_path)