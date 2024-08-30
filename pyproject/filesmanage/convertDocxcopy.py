import win32com.client
import time
import os
from filesmanage import opfiles

def get_page_count(doc_path):
    # 启动Word应用程序
    word = win32com.client.Dispatch("Word.Application")
    # 打开指定的文档
    doc = word.Documents.Open(doc_path)
    
    # 在这里等待10秒
    time.sleep(10)
    
    # 获取文档的页数
    page_count = doc.ComputeStatistics(2)  # 2代表页数
    # 关闭文档
    doc.Close()
    # 退出Word应用程序
    word.Quit()
    return page_count

# doc_path = r'C:\Users\jimee\Desktop\ttyy\甘肃省屋面施工方案.docx'
# print(f'The document has {get_page_count(doc_path)} pages.')

selected_folder, parent_folder = opfiles.select_folder()

for root, dirs, files in os.walk(selected_folder):
        if root == selected_folder:#只处理指定文件夹下这一级
            for dir in dirs:
                dir
                dir_path = os.path.join(selected_folder, dir)
                for root2, _, files2 in os.walk(dir_path):
                    if root2 != dir_path:
                         break
                    bdoc = False
                    bdocx =False
                    bpdf = False
                    for file2 in files2:
                        file_namefull = os.path.join(dir_path, file2)
                        file_size = os.path.getsize(file_namefull) / 1024 /1024
                        if file_size < 1.5:
                             continue
                        if file2.endswith('.docx'):
                            ipage = get_page_count(file_namefull)