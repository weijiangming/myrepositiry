#把已经解压过的移到压缩包同目录unzip2zip_position
#解压前把已经解压过的临时移出来movefiles_out
import os
import opfiles
import shutil

def movefiles_out(folder_path, parent_folder):
    des_path = os.path.normpath(os.path.join(parent_folder, folder_path+"_beforeUnzip"))
    try:
        os.makedirs(des_path)
    except FileExistsError:
        pass
    remarklist = []
    for root, dirs, files in os.walk(folder_path):
        if root == folder_path:#只处理指定文件夹下这一级
            for dir in dirs:
                dir_path = os.path.normpath(os.path.join(folder_path, dir))
                for root2, dirs2, files2 in os.walk(dir_path):
                    if root2 != dir_path:
                        continue
                    for file2 in files2:
                        if file2.endswith('.rar') or file2.endswith('.zip'):
                            filename2 = opfiles.OpFiles.remove_suffix(file2, 4)
                            for dir2 in dirs2:
                                if dir2 == filename2:
                                    try:
                                        shutil.move(dir_path,des_path)
                                    except shutil.Error as e:
                                        print(dir_path+"未成功移动!")       


#把已经解压过的移到压缩包同目录
def unzip2zip_position(folder_path, parent_folder):
    des_path = os.path.normpath(os.path.join(parent_folder, folder_path+"_beforeUnzip"))
    try:
        os.makedirs(des_path)
    except FileExistsError:
        pass
    remarklist = []

    dict_addname = {}
    for rootT, dirsT, filesT in os.walk(folder_path):
        if rootT == folder_path:#只处理指定文件夹下这一级
            for dirSrc in dirsT:
                dict_addname[dirSrc] = dirSrc
                                         
    for root, dirs, files in os.walk(folder_path):
        if root == folder_path:#只处理指定文件夹下这一级
            icount = 0
            for dir in dirs:
                icount += 1
                if icount % 1000 == 0:
                    print(f"icount 的值为：{icount}")

                dir_path = os.path.normpath(os.path.join(folder_path, dir))
                for root2, dirs2, files2 in os.walk(dir_path):
                    if root2 != dir_path:
                        continue
                    for file2 in files2:
                        if file2.endswith('.rar') or file2.endswith('.zip'):
                            filename2 = opfiles.OpFiles.remove_suffix(file2, 4)
                            bmoved = False

                            if filename2 in dict_addname:
                                dir_Srcpath = os.path.normpath(os.path.join(folder_path, filename2))
                                try:
                                    shutil.move(dir_Srcpath,dir_path)
                                except shutil.Error as e:
                                    print(dir_path+"未成功移动!")  


# 指定文件夹路径
selected_folder, parent_folder = opfiles.OpFiles.select_folder()

#unzip2zip_position(selected_folder, parent_folder)

movefiles_out(selected_folder, parent_folder)

