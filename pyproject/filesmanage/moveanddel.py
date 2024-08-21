import os
import openpyxl
import shutil
from filesop import opfiles

#解压文件拷回主目录
def is_rar_file(filename):
    return filename.lower().endswith(".rar")

def remove_rar_suffix(filename):
  """
  从文件名中去除 .rar 后缀

  Args:
    filename: 包含 .rar 后缀的文件名

  Returns:
    去除 .rar 后缀后的文件名
  """

  # 使用字符串切片，直接截取到 .rar 前面的部分
  return filename[:-4]



def get_folder2fullpaths(root_dir):
    """
    Args:
        root_dir: 指定根目录路径
    Returns:
        指定目录下所有文件夹的全路径名的列表
    """
    folder2fullpaths = {}
    unzipfilepath = (f"{root_dir}_解压")

    for uziproot, uzipdirs, _ in os.walk(unzipfilepath): #xxx_解压里的文件夹
        if not uziproot==unzipfilepath:
            continue
        
        for uzipdir in uzipdirs:
            is_moved = False

            #主目录
            for root, dirs, _ in os.walk(root_dir):
                if not root==root_dir:
                    continue
                #主目录下的文件夹
                for dir in dirs:
                    full_path = os.path.join(root, dir)
                    ##主目录下的文件夹里的压缩文件
                    for root2, _, filenames in os.walk(full_path):
                        if not root2==full_path:
                            continue
                        for filename in filenames:
                            full_pathname = os.path.join(root2, filename)
                            if is_rar_file(filename):#判断是.rar文件
                                #filename  删除full_pathname
                                #获取已经解压的
                                new_filename = remove_rar_suffix(filename)

                                if uzipdir == new_filename:
                                    uzipfull_path = os.path.join(uziproot, uzipdir)
                                    despath = os.path.join(full_path, uzipdir)
                                    if not os.path.exists(despath):
                                        shutil.move(uzipfull_path,root2)
                                        is_moved = True
                            if is_moved:
                                break
                        if is_moved:
                            break
                    if is_moved:
                        break
                if is_moved:
                    break

            #
            folder2fullpaths[dir] = full_path
    return folder2fullpaths

# 指定文件夹路径
selected_folder, parent_folder = opfiles.OpFiles.select_folder()
folder2fullpaths = get_folder2fullpaths(selected_folder)
#print(folder2fullpaths)