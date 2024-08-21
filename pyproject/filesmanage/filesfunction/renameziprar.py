#压缩文件改名:加文件大小
import os
import opfiles


def rename_files(folder_path, parent_folder):
    remarklist = []
    for root, dirs, files in os.walk(folder_path):
        if root == folder_path:#只处理指定文件夹下这一级
            for dir in dirs:
                dir_path = os.path.join(folder_path, dir)
                for root2, dirs2, files2 in os.walk(dir_path):
                    for file2 in files2:
                        if file2.endswith('.rar') or file2.endswith('.zip'):
                            #用断点查看非本级目录的情况
                            if root2 != dir_path:
                                 print(f"压缩包不在二级目录的目录名(root2): {root2}")
                                 #先记录到excel，必要时处理这类细节
                                 remarklist.append(root2)
                                 break
                            if "_addsize" in file2:
                                break
                            file_path = os.path.join(dir_path, file2)
                            # 使用 os.path.getsize() 获取文件大小 (压缩文件的大小)
                            size_bytes = os.path.getsize(file_path)
                            size_kb = size_bytes // 1024
                            new_name = f"{os.path.splitext(file_path)[0]}_addsize{size_kb}kb{os.path.splitext(file_path)[1]}"
                            new_path = os.path.join(root, new_name)
                            # 重命名文件
                            os.rename(file_path, new_path)
                            #print(f"已重命名: {file_path} -> {new_path}")
    
    sheet_name='未重命名成功压缩包不在二级目录'

    foldername = folder_path
    last_slash = folder_path.rfind('/')
    if last_slash != -1:
        foldername = folder_path[last_slash + 1:]
 
    exl_name=foldername + "_程序执行过程指定信息记录.xlsx"
    excel_file = os.path.join(parent_folder, exl_name)
    opfiles.OpFiles.write_1d_list_to_excel(remarklist, excel_file, sheet_name)


# 指定文件夹路径
selected_folder, parent_folder = opfiles.OpFiles.select_folder()
#folder_path = r"C:\Users\jimee\Desktop\新文件夹"  # 替换为你的文件夹路径
rename_files(selected_folder, parent_folder)

