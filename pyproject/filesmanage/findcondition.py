#一个文件夹下有一个以上文件
import os
import shutil
from filesfunction import opfiles

def samedoc_delete(folder_path, parent_floder):
    for root, dirs, files in os.walk(folder_path):
        #文件
        filenames = []
        for file in files:
            filenames.append(file)
        if len(filenames) >1:
            filenames
             
# 指定文件夹路径
selected_folder, parent_folder = opfiles.OpFiles.select_folder()
samedoc_delete(selected_folder, parent_folder)