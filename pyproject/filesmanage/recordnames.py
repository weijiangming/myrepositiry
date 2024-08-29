import os
from openpyxl import Workbook
from filesfunction import opfiles

def write_filenames_to_excelT(folder_path, excel_file_path):
    """
    将指定文件夹中的文件名写入到Excel文件中

    Args:
        folder_path: 目标文件夹路径
        excel_file_path: Excel文件保存路径
    """

    workbook = Workbook()
    sheet = workbook.active
    sheet.append(['文件名'])

    for root, dirs, files in os.walk(folder_path):
        for file in files:
            sheet.append([file])

    workbook.save(excel_file_path)


def write_filenames_to_excel(folder_path, excel_file_path, include_subfolders=False):
    """
    将指定文件夹及其子文件夹中的文件名写入到Excel文件中

    Args:
        folder_path: 目标文件夹路径
        excel_file_path: Excel文件保存路径
        include_subfolders: 是否包含子文件夹中的文件，默认为True
    """

    workbook = Workbook()
    sheet = workbook.active

    # 根据是否包含子文件夹选择遍历方式
    if include_subfolders:
        for root, _, files in os.walk(folder_path):
            for file in files:
                sheet.append([os.path.join(root, file)])
    else:
        for file in os.listdir(folder_path):
            if os.path.isfile(os.path.join(folder_path, file)):
                sheet.append([file])

    workbook.save(excel_file_path)


# 示例用法
folder_path, parent_folder = opfiles.OpFiles.select_folder()#你的文件夹路径
excel_file_path = opfiles.OpFiles.select_excel_file()#Excel文件名

write_filenames_to_excel(folder_path, excel_file_path)


# import tkinter as tk
# from tkinter import filedialog

# root = tk.Tk()
# root.title("选择文件")

# # 选择文件夹
# def select_folder():
#     global selected_folder
#     selected_folder = filedialog.askdirectory()
#     # 在这里处理选定的文件夹路径，比如将其显示在标签上
#     label_folder_path.config(text=selected_folder)

# # 选择Excel文件
# def select_excel_file():
#     # ... (同上)

# # 创建按钮
# button_folder = tk.Button(root, text="选择文件夹", command=select_folder)
# button_file = tk.Button(root, text="选择Excel文件", command=select_excel_file)
