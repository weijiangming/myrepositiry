#工标库删除条文说明的json文件；条纹编号重复切片（注：适用正文）
import os
import sys
from pathlib import Path
import json
import openpyxl

parent_dir = str(Path(__file__).resolve().parent.parent.parent)# 获取当前文件的父目录
sys.path.append(parent_dir)# 将父目录添加到sys.path
parent_dir = str(Path(__file__).resolve().parent.parent)# 获取当前文件的父目录
sys.path.append(parent_dir)# 将父目录添加到sys.path
from filesfunction import opfiles



def issametext(slicetextbase, slicetext, articlecode):

    #切片的头部剪掉与"条文编号"相同的部分
    newsubstringbase = slicetextbase
    newsubstring = slicetext

    indexfindT = -1
    indexfindT = slicetext.find(str(articlecode))
    if indexfindT == 0:
        newsubstring = slicetext[len(articlecode):]
    else:
        pass

    newsubstring = newsubstring.lstrip()

    indexfindT = -1
    indexfindT = slicetextbase.find(str(articlecode))
    if indexfindT == 0:
        newsubstringbase = slicetextbase[len(articlecode):]
    else:
        pass

    newsubstringbase = newsubstringbase.lstrip()

    return newsubstringbase == newsubstring



# 定义文件夹路径
source_folder, parent_folder = opfiles.OpFiles.select_folder()

# 创建一个新的Excel工作簿和工作表
workbook = openpyxl.Workbook()
sheet = workbook.active

row = 0
icount = 0
icount2 = 0 #测试用
icount3 = 0
icount4 = 0 
icount5 = 0
icount6 = 0
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

    #重复的
     #定义字典 一个重复组：重复组第一项的序号+重复组第一项以外i的其他项的条文编号的数组（Index2codelist）；字典一项对应一个重复组；字典由多个重复组组成。
    repeat_dict = {}
    repeat_dict_text = {}
    if jsonname.endswith('.json') or jsonname.endswith('.Json'):
        file_path = os.path.join(source_folder, jsonname)
        # 打开并读取JSON文件
        new_data = []
        with open(file_path, 'r', encoding='utf-8') as json_file:
            data = json.load(json_file)
            icount = icount + 1
            if icount % 100 == 0:
                print(icount)
        try:
            for entry in data:  
                if "文档名称" in entry and "条文编号" in entry and "切片不带格式" in entry and "切片带格式" in entry:
                    icount2 = icount2 + 1
                    filename = entry["文档名称"]
                    slicetext = entry["切片不带格式"]
                    slicetext_format = entry["切片带格式"]
                    articlecode = entry["条文编号"]
                    sliceuuid = entry["切片id"]

                    if articlecode in articlecodes:
                        #说明第二次出现，重复了
                        indexT = articlecodes.index(articlecode)
                        if indexT in repeat_dict:#
                            slicetextbase = slicetexts[indexT]
                            if issametext(slicetextbase, slicetext, articlecode):
                                #有记录 重复项已在字典里，这时加入
                                repeat_dict[indexT].append(sliceuuid)
                                repeat_dict_text[indexT].append(slicetext)
                        else:
                            #说明重复项字典里还没记录
                            slicetextbase = slicetexts[indexT]
                            if issametext(slicetextbase, slicetext, articlecode):
                                repeat_dict[indexT] = []
                                repeat_dict_text[indexT] = []
                                repeat_dict[indexT].append(sliceuuid)
                                repeat_dict_text[indexT].append(slicetext)
                            
                    articlecodes.append(articlecode) 
                    slicetexts.append(slicetext)          
                else:
                    print(f'{jsonname} 字段不全缺；条文编号：{str(articlecode)}')
        
        except json.JSONDecodeError:
            print(f'{jsonname} 该文件单独查是什么问题')


        #开始处理一个文件
        icount3 = icount3 + 1
        # ①记录所有重复项的条文编号字段，留下第一个，其余的重复项分片删除
        
        sync_index = -1#确保第一个的序号是零
        for entry in data:
            #确保和上一次for entry in data:规则相同，保证序号相同。
            if "文档名称" in entry and "条文编号" in entry and "切片不带格式" in entry and "切片带格式" in entry:
                sync_index = sync_index + 1
            #
            twbhT = entry["切片id"]
            # 遍历字典
            found = False
            for key, value_list in repeat_dict.items():
                if twbhT in value_list:
                    found = True
                    break 

            if found == False:#需要保留的都在new_data里
                new_data.append(entry)

        # 将修改后的数据写回文件
        with open(file_path, 'w', encoding='utf-8') as json_file:
            json.dump(new_data, json_file, ensure_ascii=False, indent=4)

