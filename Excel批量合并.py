import os
import pandas as pd

# 环境需要使用xlsxwriter模块
# pip install xlsxwriter


# 指定文件夹路径
folder_path = input("\n请输入要处理的文件夹路径: \n")

# 获取文件夹下所有Excel文件的文件名
excel_files = [f for f in os.listdir(folder_path) if f.endswith(".xlsx")]

# 创建一个新的Excel写入器
writer = pd.ExcelWriter(folder_path+"/merged_excel.xlsx", engine="xlsxwriter")

# 遍历每个Excel文件并将其写入新的Excel文件中
for file in excel_files:
    file_path = os.path.join(folder_path, file)
    sheet_name = os.path.splitext(file)[0]  # 使用文件名作为sheet名称

    # 读取Excel文件
    df = pd.read_excel(file_path)

    # 将数据写入新的Excel文件中
    df.to_excel(writer, sheet_name=sheet_name, index=False)

# 保存并关闭写入器
writer.close()