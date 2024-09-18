#删除"切片带格式""的值中，内容带有 src=\"data:image/png;base64 的<img>标签
import tkinter as tk
from tkinter import filedialog
#from openpyxl import load_workbook
import os
import sys
from pathlib import Path
import json
import re

# 获取当前文件的父目录
parent_dir = str(Path(__file__).resolve().parent.parent)
# 将父目录添加到sys.path
sys.path.append(parent_dir)

from filesfunction import opfiles

# 创建 Tkinter 主窗口（通常隐藏）
root = tk.Tk()
root.withdraw()

source_folder, parent_folder = opfiles.OpFiles.select_folder()

newfolder = source_folder.split("/")[-1] + "_newname"
newfolder_path = os.path.normpath(os.path.join(parent_folder, newfolder))

# 遍历源文件夹中的所有文件
for filename in os.listdir(source_folder):
    if filename.endswith('.json') or filename.endswith('.Json'):
        file_path = os.path.join(source_folder, filename)
  

        with open(file_path, 'r', encoding='utf-8') as json_file:
            data = json.load(json_file)

        try:
            #修改 "文档名称"
            for item in data:
                if "切片不带格式" in item:
                    value = item['切片带格式']
                    # 匹配 src="data:image/png;base64" 的 <img> 标签
                    if 'src="data:image/png;base64' in value:
                        pattern = r'<img.*?/>'
                        valueRes = re.sub(pattern, '', value)
                        item['切片带格式'] = valueRes

            # 将修改后的数据写入新的json文件
            with open(file_path, 'w', encoding='utf-8') as file:
                json.dump(data, file, ensure_ascii=False , indent=4)        
                        
        except json.JSONDecodeError:
            print(f'{filename} 不是有效的 JSON 文件')


