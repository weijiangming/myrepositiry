#将指定文件夹及其子文件夹中的文件名写入到Excel文件中
import os
from openpyxl import Workbook
import sys
from pathlib import Path

# 获取当前文件的父目录
parent_dir = str(Path(__file__).resolve().parent.parent)
# 将父目录添加到sys.path
sys.path.append(parent_dir)

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


def write_filenames_to_excel(folder_path, excel_file_path):
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
    for root, dirnames, _ in os.walk(folder_path):
        if root == folder_path:
            for dirname in dirnames:
                sheet.append([os.path.join(root, dirname)])

    workbook.save(excel_file_path)

# 示例用法
folder_path, parent_folder = opfiles.OpFiles.select_folder()#你的文件夹路径
excel_file_path = opfiles.OpFiles.select_excel_file()#Excel文件名

write_filenames_to_excel(folder_path, excel_file_path)
