import os
import tkinter as tk
from tkinter import filedialog
import win32com.client

# 选择文件夹
root = tk.Tk()
root.withdraw()
selected_folder = filedialog.askdirectory()


if selected_folder:
    # 获取上级目录
    parent_folder = os.path.dirname(selected_folder)

    # 启动 Word 应用程序
    word = win32com.client.Dispatch("Word.Application")

    try:
        # 遍历指定目录下的一级文件夹
        for root, dirs, files in os.walk(selected_folder):
            if root == selected_folder:  # 只处理指定文件夹下的这一层
                for dir in dirs:
                    dir_path = os.path.join(selected_folder, dir)
                    for root2, _, files2 in os.walk(dir_path):
                        if root2 != dir_path:
                            break

                        for file2 in files2:
                            # file2_full_path = os.path.join(root2, file2)
                            file2_full_path = os.path.normpath(os.path.join(root2, file2))
                            file_size = os.path.getsize(file2_full_path) / 1024 / 1024  # 文件大小（MB）
                            
                            if file2.endswith('.doc') or file2.endswith('.docx'):
                                try:
                                    # 打开指定的文档
                                    doc = word.Documents.Open(file2_full_path)
                                    # 获取文档的页数
                                    page_count = doc.ComputeStatistics(2)  # 2代表页数
                                    # 关闭文档
                                    doc.Close()
                                    # 退出Word应用程序


                                    print(f"文件: {file2_full_path}")
                                    print(f"大小: {file_size:.2f} MB")
                                    print(f"页数: {page_count} 页")
                                    print("=" * 40)

                                except Exception as e:
                                    print(f"处理文件 {file2_full_path} 时出错: {e}")
                                    if 'doc' in locals():
                                        doc.Close()
    finally:
        # 退出 Word 应用程序
        word.Quit()

else:
    print("未选择文件夹")




# import os
# import win32com.client
# import time
# import opfiles

# def get_page_count(doc_path):
#     # 启动Word应用程序
#     word = win32com.client.Dispatch("Word.Application")
#     # 打开指定的文档
#     doc = word.Documents.Open(doc_path)
    
#     # 在这里等待10秒
#     #time.sleep(10)
    
#     # 获取文档的页数
#     page_count = doc.ComputeStatistics(2)  # 2代表页数
#     # 关闭文档
#     doc.Close()
#     # 退出Word应用程序
#     word.Quit()
#     return page_count

# # doc_path = r'C:\Users\jimee\Desktop\ttyy\甘肃省屋面施工方案.docx'
# # print(f'The document has {get_page_count(doc_path)} pages.')

# selected_folder, parent_folder = opfiles.OpFiles.select_folder()

# for root, dirs, files in os.walk(selected_folder):
#         if root == selected_folder:#只处理指定文件夹下这一级
#             for dir in dirs:
#                 dir_path = os.path.join(selected_folder, dir)
#                 for root2, _, files2 in os.walk(dir_path):
#                     if root2 != dir_path:
#                          break
                
#                     for file2 in files2:
                        
#                         file_size = os.path.getsize(file_namefull) / 1024 /1024
#                         if file_size < 1.5:
#                              continue
#                         if file2.endswith('.docx'):
#                             file_namefull = os.path.join(root2, file2)
#                             ipage = get_page_count(file_namefull)
#                             print(file_namefull,ipage)