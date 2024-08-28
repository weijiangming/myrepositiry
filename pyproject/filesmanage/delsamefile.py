#删除文件如一个文件夹下有文件：厂房扩建工程施工方案.pdf和厂房扩建工程施工方案_638586599358442232.pdf大小相同删掉第二个文件
import os
import shutil
from filesfunction import opfiles

def samedoc_delete(folder_path, parent_floder):
    for root, dirs, files in os.walk(folder_path):
        #文件
        filenames = []
        if len(files) > 1:
            dfs = 1
        for file in files:
            if file.endswith('.pdf'):
                file_Nosuf = opfiles.OpFiles.remove_suffix(file, 4)
            elif file.endswith('.PDF'):
                file_Nosuf = opfiles.OpFiles.remove_suffix(file, 4)
            else:
                continue
            filenames.append(file_Nosuf)
        filepathsdel = []  
        filesdel = []  
        for i in range(len(filenames) - 1):
            if filenames[i] in filesdel:
                continue
            for j in range(i + 1, len(filenames)):
                if filenames[j] in filesdel:
                    continue
                # 文件大小相同且文件名包含关系成立
                filepath_i = os.path.normpath(os.path.join(root, filenames[i] + ".pdf"))
                filepath_j = os.path.normpath(os.path.join(root, filenames[j] + ".pdf"))
                size_i = os.path.getsize(filepath_i)
                size_j = os.path.getsize(filepath_j)
                diff_value = abs(size_i - size_j)
                if diff_value < 2000000:
                    if filenames[i].find(filenames[j]) != -1:
                        filepathsdel.append(filepath_i)
                        filesdel.append(filenames[i])
                    elif  filenames[j].find(filenames[i]) != -1:
                        filepathsdel.append(filepath_j)
                        filesdel.append(filenames[j])
        for filepathT in filepathsdel:
            os.remove(filepathT)
        
                    
                  

            
            

# 指定文件夹路径
selected_folder, parent_folder = opfiles.OpFiles.select_folder()
samedoc_delete(selected_folder, parent_folder)