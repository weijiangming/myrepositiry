import os
import json
import shutil
from filesfunction import opfiles

# 定义文件夹路径
source_folder, parent_folder = opfiles.OpFiles.select_folder()

target_folder = os.path.join(parent_folder,f"{source_folder}_version1")

# 如果目标文件夹不存在，创建它
if not os.path.exists(target_folder):
    os.makedirs(target_folder)

# 遍历源文件夹中的所有文件
for filename in os.listdir(source_folder):
    if filename.endswith('.json'):
        file_path = os.path.join(source_folder, filename)
        
        # 打开并读取JSON文件
        with open(file_path, 'r', encoding='utf-8') as json_file:
            data = json.load(json_file)
        try:
            # 检查是否有 "版本": 1
            for entry in data:
                if "版本" in entry:
                    versionTTT = entry["版本"]
                    if versionTTT == 1:
                        # 移动文件到 "版本一" 文件夹
                        shutil.move(file_path, target_folder)
                        #print(f'{filename} 已移动到 {target_folder}')
                        # if os.path.exists(file_path):
                        #     os.remove(file_path)
                        #     print(f"{file_path} 已成功删除")
                        # else:
                        #     print(f"文件 {file_path} 不存在")
                         
                        break
                else:
                    print("entry does not have a '版本' key")
            

        except json.JSONDecodeError:
            print(f'{filename} 不是有效的 JSON 文件')