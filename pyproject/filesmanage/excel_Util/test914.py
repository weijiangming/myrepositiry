import os
import sys
from pathlib import Path

parent_dir = str(Path(__file__).resolve().parent.parent)# 获取当前文件的父目录
sys.path.append(parent_dir)# 将父目录添加到sys.path
from filesfunction import opfiles

source_folder, parent_folder = opfiles.OpFiles.select_folder()

# 遍历源文件夹中的所有文件
bmod = True

for jsonname in os.listdir(source_folder):
    bmod = False