#工标库删除json条纹编号重复切片
import re
import os
import sys
from pathlib import Path
import json
import openpyxl
import random

parent_dir = str(Path(__file__).resolve().parent.parent.parent)# 获取当前文件的父目录
sys.path.append(parent_dir)# 将父目录添加到sys.path
parent_dir = str(Path(__file__).resolve().parent.parent)# 获取当前文件的父目录
sys.path.append(parent_dir)# 将父目录添加到sys.path
from filesfunction import opfiles

# 搜索出“切片不带格式”和“切片带格式”中除了行首的条文编号不同，其余部分的文字内容一样的分片，进行如下处理：

# ①记录所有重复项的条文编号字段，留下第一个，其余的重复项分片删除

# ②对于去重后的分片处理：删除“切片不带格式”和“切片带格式”行首的条文编号，然后将所有重复分片各自的条文编号拿出来填被留下分片的条文编号字段中。并用“、”隔开

# ③如果被留下分片“切片不带格式”和“切片带格式”的行首没有“x.x.x”或“x.x”样式的编号，那么就将重新填的条文编号放在行首，如果这串条文编号是连续的（例如“1.0.1、1.0.2、1.0.3、1.0.4”）那么久简写成“1.0.1~1.0.4”样式；如果有不连续的号，就单独拿出来放在最后，用“、”隔开（例如“1.0.1、1.0.2、1.0.3、1.0.4、1.0.8”简写成“1.0.1~1.0.4、1.0.8”；如果被留下分片“切片不带格式”和“切片带格式”的行首有“x.x.x”或“x.x”样式的编号，那么就不用重新填编号

# ④编号和正文之间还是要添加一个空格

def simplify_versions(versions):
    # 将输入的字符串按照逗号分隔，生成版本列表，并去除多余的空格
    version_list = [v.strip() for v in versions.split('、')]
    
    # 初始化简化版本列表
    simplified = []
    
    # 初始化一个列表，用于存储连续的版本号
    current_range = []
    
    # 遍历版本列表
    for i in range(len(version_list)):
        if not current_range:  # 如果current_range为空，添加当前版本
            current_range.append(version_list[i])
        else:
            # 检查当前版本与上一个版本是否为连续版本
            current_version_parts = version_list[i].split('.')
            prev_version_parts = current_range[-1].split('.')
            
            # 比较最后一部分是否相差1，且前面部分是否相同
            if (len(current_version_parts) == len(prev_version_parts) and
                current_version_parts[:-1] == prev_version_parts[:-1] and
                int(current_version_parts[-1]) == int(prev_version_parts[-1]) + 1):
                current_range.append(version_list[i])  # 如果连续，加入current_range
            else:
                # 如果不连续，结束当前range，并存储到simplified
                if len(current_range) > 2:  # 如果连续超过两个，才简写为范围
                    simplified.append(f"{current_range[0]}~{current_range[-1]}")
                else:
                    # 如果是两个连续版本或单独版本，用“、”隔开
                    simplified.extend(current_range)
                current_range = [version_list[i]]  # 开始一个新的range
    
    # 最后一次的range处理
    if current_range:
        if len(current_range) > 2:  # 如果最后一段有超过两个版本连续
            simplified.append(f"{current_range[0]}~{current_range[-1]}")
        else:
            simplified.extend(current_range)  # 两个连续版本或单独版本
    
    return '、'.join(simplified)


def get_string_before_last_dot(input_string):
    # 找到最后一个"."的位置
    last_dot_index = input_string.rfind('.')
    # 如果找到了"."，返回"."之前的部分；如果没有找到，返回原始字符串
    if last_dot_index != -1:
        return input_string[:last_dot_index]
    else:
        return input_string


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
                            #有记录 重复项已在字典里，这时加入
                            repeat_dict[indexT].append(sliceuuid)
                            sliceuuids.append(sliceuuid)
                        else:
                            #说明重复项字典里还没记录
                            repeat_dict[indexT] = []
                            repeat_dict[indexT].append(sliceuuid)
                            sliceuuids.append(sliceuuid)

                    articlecodes.append(articlecode)                
                else:
                    print(f'{jsonname} 字段不全缺；条文编号：{str(articlecode)}')

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


        except json.JSONDecodeError:
            print(f'{jsonname} 该文件单独查是什么问题')
pass


# excel_file_name = source_folder.split("/")[-1] + "_条文说明略写前后对比记录.xlsx"
# excelfolder_path = os.path.normpath(os.path.join(parent_folder, excel_file_name))
# workbook.save(excelfolder_path)

