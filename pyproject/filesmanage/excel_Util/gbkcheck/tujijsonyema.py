#图集son文件页码补如 34 36 补35
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


    tujiyes = []#图集页码列表

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
                if "tujiye" in entry and "tuji_name" in entry:
                    icount2 = icount2 + 1

                    tuji_name = entry["tuji_name"]
                    tujiye = entry["tujiye"]
                    tujiyes.append(tujiye)
      
                else:
                    print(f'{jsonname} 字段不全缺；条文编号：{str(tuji_name)}')
        
        except json.JSONDecodeError:
            print(f'{jsonname} 该文件单独查是什么问题')


        #开始处理一个文件
        icount3 = icount3 + 1
        # ①记录所有重复项的条文编号字段，留下第一个，其余的重复项分片删除
        
        sync_index = -1#确保第一个的序号是零
        for entry in data:
            #确保和上一次for entry in data:规则相同，保证序号相同。
            if "tujiye" in entry and "tuji_name" in entry:
                sync_index = sync_index + 1
            #
            tujiye = entry["tujiye"]
            lenye = len(tujiyes)
            if tujiye == "" and sync_index > 0 and sync_index < lenye -1:
                pass
                freye = tujiyes[sync_index - 1]
                nextye = tujiyes[sync_index + 1]

                if freye.isdigit() and nextye.isdigit():
                    # 将字符串转换为整数
                    num1 = int(freye)
                    num2 = int(nextye)
                    # 计算差值
                    difference = num2 - num1
                    if difference == 2:
                        entry["tujiye"] = num1 + 1 
                        icount6 = icount6 + 1

            new_data.append(entry)

        # 将修改后的数据写回文件
        with open(file_path, 'w', encoding='utf-8') as json_file:
            json.dump(new_data, json_file, ensure_ascii=False, indent=4)

print(f'共补了{icount6} 项')
