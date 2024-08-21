import os

def rename_and_move(folder_path):
    """
    批量重命名文件，并以文件名（不含扩展名）创建文件夹并移动文件。

    Args:
        folder_path: 需要处理的文件夹路径。
    """

    for file_name in os.listdir(folder_path):
        # 构造新文件名
        new_file_name = "#" + file_name
        # 构造新文件路径
        new_file_path = os.path.join(folder_path, new_file_name)  # 这里对 new_file_path 赋值
        # 获取文件名（不含扩展名）
        base_name, _ = os.path.splitext(file_name)
        base_name_new, _ = os.path.splitext(new_file_path)
        # 构造新文件夹路径
        new_folder_path = os.path.join(folder_path, base_name_new)

        # 重命名文件
        old_file_path = os.path.join(folder_path, file_name)
        os.rename(old_file_path, new_file_path)

        # 创建新文件夹
        os.makedirs(new_folder_path, exist_ok=True)

        # 移动文件到新文件夹
        os.rename(new_file_path, os.path.join(new_folder_path, new_file_name))

if __name__ == "__main__":
    # 替换为你的文件夹路径
    folder_path = r"G:\9DECSAzhmOOm"
    rename_and_move(folder_path)