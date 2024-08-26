#文件夹下有同名的doc和docx移到外面的目录下（父文件夹名+_samedoc_x）
import os
import shutil
from filesfunction import opfiles

def samedoc_moveout(folder_path, parent_floder):
    des_path = os.path.normpath(os.path.join(parent_folder, folder_path+"_samedoc_x"))
    try:
        os.makedirs(des_path)
    except FileExistsError:
        pass
    for root, dirs, files in os.walk(folder_path):
        #文件
        filenames = []
        for file in files:
            if file.endswith('.docx'):
                file_Nosuf = opfiles.OpFiles.remove_suffix(file, 5)
            elif file.endswith('.doc'):
                file_Nosuf = opfiles.OpFiles.remove_suffix(file, 4)
            else:
                continue
            filenames.append(file_Nosuf)
        # 检查是否有重复项
        if len(filenames) != len(set(filenames)):
            print("列表中有重复的数据。")
            pathT = root
            uppath = os.path.dirname(pathT)
            while True:
                if uppath == folder_path:
                    break
                pathT = uppath
                uppath = os.path.dirname(uppath)

            try:
                shutil.move(pathT,des_path)
            except shutil.Error as e:
                print(uppath+"未成功移动!")  
                  

            
            

# 指定文件夹路径
selected_folder, parent_folder = opfiles.OpFiles.select_folder()
samedoc_moveout(selected_folder, parent_folder)
