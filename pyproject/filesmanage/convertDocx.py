import os
#file_size = os.path.getsize("F:\文档\钢结构-工程测量-加固改造--智能建造test - 副本\(北京)工业厂房施工组织设计_doc_839.00KB\北京工业厂房施工组织设计.docx")
file_size = os.path.getsize(r'C:\Users\jimee\Desktop\ttyy\北京工业厂房施工组织设计.docx')
file_size_kb = file_size / 1024
print(f"文件大小：{file_size_kb:.2f} KB")

