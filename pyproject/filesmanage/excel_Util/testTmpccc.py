from tkinter import filedialog
import pandas as pd

# 加载Excel文件
# 打开文件选择对话框
file_path = filedialog.askopenfilename(title='请选择Excel文件', filetypes=[('Excel文件', '*.xlsx')])

df = pd.read_excel(file_path)  # 读取Excel文件到DataFrame

# 假设我们要统计第一列的不同项数量
column_name = df.columns[0]  # 第一列的列名
unique_count = df[column_name].nunique()  # 计算不同项的数量

print(f"Column '{column_name}' has {unique_count} unique items.")