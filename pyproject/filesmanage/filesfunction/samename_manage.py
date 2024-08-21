import os
import opfiles
import shutil

#处理有与pdf同名的doc和docx,注意二级三级目录，目前只支持第三级 不支持第二级
def samenameop(folder_path, parent_folder):
    folder_samedoc = os.path.join(parent_folder,f"{folder_path}_samedoc")
    folder_samedocx = os.path.join(parent_folder,f"{folder_path}_samedocx")
    if not os.path.exists(folder_samedoc):
        os.makedirs(folder_samedoc)
    if not os.path.exists(folder_samedocx):
        os.makedirs(folder_samedocx)

    for root, dirs, _ in os.walk(folder_path):
        if root == folder_path:#只处理指定文件夹下这一级
            for dir in dirs:
                if dir == "北京某地铁站深基坑施工组织设计96P_pdf_855.47KB":
                     dir
                dir_path2 = os.path.join(folder_path, dir)
                for root2, dirs2, files2 in os.walk(dir_path2):
                    if not root2 == dir_path2:
                        continue
                    #二级
                    pdffilename = ""
                    bBreak = False
                    for file2 in files2:
                        if file2.lower().endswith('.pdf'):
                            pdffilename = opfiles.OpFiles.remove_suffix(file2, 4)
                            is_samedoc = False
                            is_samedocx = False                                  
                            filenamedoc = ""
                            filenamedocx = ""                  
                            for file2T in files2:
                                if file2 == file2T:
                                    continue
                                if file2T.lower().endswith('.doc'):
                                    filenamedoc = opfiles.OpFiles.remove_suffix(file2T, 4)
                                elif file2T.lower().endswith('.docx'):
                                    filenamedocx = opfiles.OpFiles.remove_suffix(file2T, 5)
                                
                                if not is_samedoc:
                                    if pdffilename == filenamedoc:
                                        is_samedoc = True
                                        break
                                if not is_samedocx:
                                    if pdffilename == filenamedocx:
                                        is_samedocx = True
                                
                            if is_samedoc:
                                shutil.move(dir_path2,folder_samedoc)
                                bBreak = True
                            elif is_samedocx:
                                shutil.move(dir_path2,folder_samedocx)
                                bBreak = True
                        if bBreak:
                            break
                    
                    #三级目录即压缩包文件夹
                    bBreak = False
                    for dir2 in dirs2:
                        dir_path3 = os.path.join(dir_path2, dir2)
                        #解压文件夹
                        for root3, _, files3 in os.walk(dir_path3):
                            if not root3 == dir_path3:
                                continue
                            pdffilename = ""
                            for file3 in files3:
                                if file3.lower().endswith('.pdf'):
                                    pdffilename = opfiles.OpFiles.remove_suffix(file3, 4)
                                    if "北京某地铁站深基坑施工组织设计" in file3:
                                        pdffilename
                                    is_samedoc = False
                                    is_samedocx = False                                  
                                    filenamedoc = ""
                                    filenamedocx = ""                  
                                    for file3T in files3:
                                        if file3 == file3T:
                                            continue
                                        if file3T.lower().endswith('.doc'):
                                            filenamedoc = opfiles.OpFiles.remove_suffix(file3T, 4)
                                        elif file3T.lower().endswith('.docx'):
                                            filenamedocx = opfiles.OpFiles.remove_suffix(file3T, 5)
                                        
                                        if not is_samedoc:
                                            if pdffilename == filenamedoc:
                                                is_samedoc = True
                                                break
                                        if not is_samedocx:
                                            if pdffilename == filenamedocx:
                                                is_samedocx = True
                                        
                                    if is_samedoc:
                                        shutil.move(dir_path2,folder_samedoc)
                                        bBreak = True
                                    elif is_samedocx:
                                        shutil.move(dir_path2,folder_samedocx)
                                        bBreak = True
                                if bBreak:
                                    break
                            if bBreak:
                                    break
                        if bBreak:
                                    break

# 指定文件夹路径
selected_folder, parent_folder = opfiles.OpFiles.select_folder()
#folder_path = r"C:\Users\jimee\Desktop\新文件夹"  # 替换为你的文件夹路径
samenameop(selected_folder, parent_folder)