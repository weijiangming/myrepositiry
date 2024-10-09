
import tkinter as tk
from tkinter import filedialog
import pandas as pd

def check_duplicates(excel_file, column_name):
    """
    检查Excel表格中指定列是否有重复字符，并标记

    Args:
        excel_file (str): Excel文件路径
        column_name (str): 要检查的列名

    Returns:
        pandas.DataFrame: 处理后的DataFrame
    """

    # 读取Excel文件
    df = pd.read_excel(excel_file)

    # 检查是否有重复字符
    df['has_duplicates'] = df[column_name].str.contains(r'(.+).*\1')

    # 标记重复项
    df['duplicate_flag'] = df['has_duplicates'].astype(int)

    # 保存结果
    df.to_excel(excel_file, index=False)

    return df

# 示例用法
excel_file = filename = filedialog.askopenfilename(initialdir="/", title="选择Excel文件", filetypes=(("Excel files", "*.xlsx"), ("all files", "*.*")))
column_to_check = 'column_name'  # 替换为实际的列名

result_df = check_duplicates(excel_file, column_to_check)
print(result_df)