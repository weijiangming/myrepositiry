#将第一个选择的目录里与第二个选择的目录里的文件夹名称相同时，将第一个目录里的该文件夹移动名为“第一个选择的目录+_same_as_sample的文件夹里
import os
import shutil
from filesfunction import opfiles


def samedoc_delete(folder_path, parent_floder, selected_folderexample):
    foldnames = []
    pathnames = []
    des_path = os.path.normpath(os.path.join(parent_folder, folder_path+"_same_as_sample"))
    try:
        os.makedirs(des_path)
    except FileExistsError:
        pass

    for root, dirs, files in os.walk(selected_folderexample):
        if root == selected_folderexample :
            for dire in dirs:
                path = os.path.normpath(os.path.join(root, dire))
                foldnames.append(dire)
                pathnames.append(path)
 
    pathnames_sel = []
    for root, dirs, files in os.walk(folder_path):
        if root == folder_path :
            for dire in dirs:
                if dire in foldnames:
                    path = os.path.normpath(os.path.join(root, dire))
                    pathnames_sel.append(path)

    for path_sel in pathnames_sel:
        try:
            shutil.move(path_sel,des_path)
        except shutil.Error as e:
            print(path+"未成功移动!")  


# 指定文件夹路径
selected_folder, parent_folder = opfiles.OpFiles.select_folder()
selected_folderexample, parent_folderexample = opfiles.OpFiles.select_folder()
samedoc_delete(selected_folder, parent_folder,selected_folderexample)
