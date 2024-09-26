#工标库json修改
import re
import os
import sys
from pathlib import Path
import json
import openpyxl
import random

parent_dir = str(Path(__file__).resolve().parent.parent)# 获取当前文件的父目录
sys.path.append(parent_dir)# 将父目录添加到sys.path
from filesfunction import opfiles

#1.“切片”开头同“条文编号”是数字没有点或附录+字母的，那么删除开头同“条文编号”的部分（字符个数），加"所属章节标题"的内容判断后面有空格，没有加空格。
#2.“切片”开头同“条文编号”的后面字符不是空格时 加个空格。

#"文档id": "242de7ff-3c34-4dc2-9c26-be646b96993a",
# "切片id": "3f694971-b5fc-4019-a57f-7ac47b127625",
# "文档名称": "《现浇改性石膏墙体应用技术规程》T/CECS 971—2021；T/CREA 008—2021 ",
# "所属章节标题": "1 总 则",
# "切片不带格式": "1.0.2 本规程适用于抗震设防烈度8度及以下地区，一般工业与民用建筑中采用现浇改性石膏墙体作为内隔墙的设计、施工及验收。",
# "切片带格式": "1.0.2 本规程适用于抗震设防烈度8度及以下地区，一般工业与民用建筑中采用现浇改性石膏墙体作为内隔墙的设计、施工及验收。",
# "条文编号": "1.0.2",
# "版本": "1",

# 定义文件夹路径
source_folder, parent_folder = opfiles.OpFiles.select_folder()

# 创建一个新的Excel工作簿和工作表
workbook = openpyxl.Workbook()
sheet = workbook.active

row = 0
icount = 0
icount2 = 0
# 遍历源文件夹中的所有文件
bmod = True
for jsonname in os.listdir(source_folder):
    bmod = False
    #定义数组
    jsonnames = []#。json文件的文件名
    filenames = []#"文档名称"对应的值
    sliceuuids = []#"切片id"
    slicetexts = []#"切片不带格式"
    slicetext_format_list = []#"切片带格式"
    articlecodes = []#"条文编号"
    if jsonname.endswith('.json') or jsonname.endswith('.Json'):
        file_path = os.path.join(source_folder, jsonname)
        #jsonnames.append(jsonname)
        # 打开并读取JSON文件
        with open(file_path, 'r', encoding='utf-8') as json_file:
            data = json.load(json_file)
            icount2 = icount2 + 1
        try:
            for entry in data:  
                if "文档名称" in entry and "条文编号" in entry and "切片不带格式" in entry and "切片带格式" in entry and "所属章节标题" in entry:
                    filename = entry["文档名称"]
                    slicetext = entry["切片不带格式"]
                    slicetext_format = entry["切片带格式"]
                    articlecode = entry["条文编号"]
                    chaptertitil = entry["所属章节标题"]
#1.“切片”开头同“条文编号”是数字没有点或附录+字母的，那么删除开头同“条文编号”的部分（字符个数），加"所属章节标题"的内容判断后面有空格，没有加空格。
#2.“切片”开头同“条文编号”的后面字符不是空格时 加个空格。
                    bAppendix = "附" in str(articlecode) and "录" in str(articlecode)
                    bNumPure = "." not in str(articlecode)
                    bNumT = type(articlecode) in (int, float, complex)
                    if bAppendix or bNumPure:
                        bmod = True
                        substring = slicetext[:len(articlecode)]
                        substringf = slicetext_format[:len(articlecode)]
                        if substring == substringf and substring == str(articlecode):
                            #保证后面有空格
                            newsubstring = slicetext[len(articlecode):]
                            newsubstringf = slicetext_format[len(articlecode):]
                            if newsubstring and newsubstring[0] == ' ':
                                slicetext = chaptertitil + newsubstring
                            else:
                                slicetext = chaptertitil + ' ' + newsubstring

                            if newsubstringf and newsubstringf[0] == ' ':
                                slicetext_format = chaptertitil + newsubstringf
                            else:
                                slicetext_format = chaptertitil + ' ' + newsubstringf
                    else:
                        substring = slicetext[:len(articlecode)]
                        substringf = slicetext_format[:len(articlecode)]
                        if substring == substringf and substring == str(articlecode):
                            #保证后面有空格
                            newsubstring = slicetext[len(articlecode):]
                            newsubstringf = slicetext_format[len(articlecode):]
                            if newsubstring[0] == ' ':
                                pass
                            else:
                                slicetext = substring + ' ' + newsubstring
                                bmod = True

                            if newsubstringf[0] == ' ':
                                pass
                            else:
                                slicetext_format = substringf + ' ' + newsubstringf
                                bmod = True

                    entry["切片不带格式"] = slicetext
                    entry["切片带格式"] = slicetext_format
                        
                else:
                    print(f'{jsonname} 五字段不全缺；条文编号：{str(articlecode)}')#
            # 将修改后的数据写入新的json文件
            # if bmod:
            #     with open(file_path, 'w', encoding='utf-8') as file:
            #         json.dump(data, file, ensure_ascii=False , indent=4) 
            # else:
            #     pass              

        except json.JSONDecodeError:
            print(f'{jsonname} 该文件单独查是什么问题')

pass

excel_file_name = source_folder.split("/")[-1] + "_提取文件名.xlsx"
excelfolder_path = os.path.normpath(os.path.join(parent_folder, excel_file_name))
workbook.save(excelfolder_path)
