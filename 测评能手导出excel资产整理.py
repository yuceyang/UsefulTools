import os
import re
import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill, Border, Side, Alignment, Font
from tkinter import Tk
from tkinter.filedialog import askdirectory
##
#   环境需要使用xlsxwriter模块
#   pip install xlsxwriter
#
# #

def start():
    # 创建一个Tkinter根窗口，并隐藏它
    root = Tk()
    root.withdraw()
    # 使用文件资源管理器获取文件夹路径
    folder_path = askdirectory(title="请选择要处理的文件夹")
    # 检查是否选择了文件夹路径
    if folder_path:
        print(f"你选择的文件夹路径是: {folder_path}")

        # 检查merged_excel.xlsx文件是否已经存在
        merged_file_path = os.path.join(folder_path, "merged_excel.xlsx")
        if os.path.exists(merged_file_path):
            print(f"文件【merged_excel.xlsx】已经存在，跳过文件处理步骤，请删除【merged_excel.xlsx】文件后重试!!!。")
        else:
            perform_replace_task(folder_path)
    else:
        print("你没有选择任何文件夹。")


def perform_replace_task(folder_path):
    # 输入文件夹路径
    #folder_path = input("\n请输入要处理的文件夹路径: \n")

    #删除多余文件
    removeFiles(folder_path+r"\02一安全区域边界一区域边界.xlsx")
    print("已删除多余文件："+folder_path+r"\02一安全区域边界一区域边界.xlsx")
    removeFiles(folder_path+r"\11一安全管理中心一安全管理中心.xlsx")
    print("已删除多余文件："+folder_path+r"\11一安全管理中心一安全管理中心.xlsx")
    removeFiles(folder_path+r"\15一全局对象.xlsx")
    print("已删除多余文件："+folder_path+r"\15一全局对象.xlsx")

# 输入要替换的字符串和对应的新字符串
    replace_rules = r"01一安全物理环境一==01,03一安全计算环境一==02,04一安全计算环境一==03,05一安全计算环境一==04,06一安全计算环境一==05,07一安全计算环境一==06,08一安全计算环境一==07,09一安全计算环境一==08,10一安全计算环境一==09,12一==10,13一==11,14一==12"

    # 将替换规则按逗号分隔为列表
    rule_list = replace_rules.split(",")

    # 创建一个空字典用于存储替换规则
    rule_dict = {}

    # 遍历替换规则列表
    for rule in rule_list:
        # 将每个规则按"=="分隔为旧字段和新字段
        fields = rule.split("==")

        # 检查每个规则是否包含旧字段和新字段
        if len(fields) != 2:
            print(f"替换规则 '{rule}' 格式不正确,请使用 '旧字段==新字段' 的格式!")
            return

        # 将旧字段和新字段添加到字典中
        old_str, new_str = fields
        rule_dict[old_str] = new_str

    # 获取文件夹下所有文件和目录
    all_entries = os.listdir(folder_path)

    # 初始化计数器
    total_files = len(all_entries)
    replaced_files = 0
    skipped_files = 0
    error_files = []

    print("")

    # 遍历文件夹下的所有文件和目录
    for entry in all_entries:
        # 初始化新文件名为原文件名
        new_filename = entry

        # 依次应用每个替换规则
        for old_str, new_str in rule_dict.items():
            # 检查文件名中是否包含要替换的字符串
            if re.search(re.escape(old_str), new_filename):
                # 替换文件名中的字符串
                new_filename = re.sub(re.escape(old_str), new_str, new_filename)

        # 如果文件名发生了变化,则重命名文件
        if new_filename != entry:
            try:
                # 拼接完整的源文件路径和目标文件路径
                src = os.path.join(folder_path, entry)
                dst = os.path.join(folder_path, new_filename)

                # 重命名文件
                os.rename(src, dst)
                print(f"【已替换】文件 '{entry}' 已成功重命名为 '{new_filename}'")
                replaced_files += 1
            except Exception as e:
                print(f"[ERROR] 重命名文件 '{entry}' 时发生错误: {e}")
                error_files.append(entry)
        else:
            print(f"-->[已跳过] 文件 '{entry}' 不存在替换条件!")
            skipped_files += 1

    # 输出统计信息
    print(f"\n统计信息:")
    print(f"文件夹下共有 {total_files} 个文件")
    print(f"已替换 {replaced_files} 个文件")
    print(f"已跳过 {skipped_files} 个文件")
    print(f"错误文件数: {len(error_files)}")
    if error_files:
        print("错误文件列表:")
        for file in error_files:
            print("    "+file)
    print("------------------------------------------------------------------------------")
    print("------------------------------------------------------------------------------")

    #合并excel
    _excel_hb(folder_path)
    #操作excel，删除行、列、表格美化等
    _excel_operate(folder_path+"/merged_excel.xlsx")

def removeFiles(file_path): # 指定要删除的Excel文件绝对路径
    # 检查文件是否存在
    if os.path.exists(file_path):
        try:
            # 删除文件
            os.remove(file_path)
            print(f"文件 '{file_path}' 已成功删除。")
        except PermissionError:
            print(f"没有权限删除文件 '{file_path}'。")
        except OSError as e:
            print(f"删除文件 '{file_path}' 时出错: {e}")
    else:
        print(f"文件 '{file_path}' 不存在。")

def _excel_hb(folder_path):
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


def _excel_operate(folder_path):
    # 打开Excel文件
    workbook = openpyxl.load_workbook(folder_path)
    # 定义每个sheet需要删除的行和列,并设置行高、列宽
    #例如：
    # '06服务器一存储设备': {'rows': [1],'cols': [8,11,12,13],'row_heights': {1: 14.4, 2: 9, 3: 16},'col_widths': {'A': 20, 'B': 15, 'C': 30}}
    sheet_deletions = {
        '01机房': {'rows': [1],'cols': [5,6],'row_heights': {None: 14.4},'col_widths': {'A': 9, 'B': 19, 'C': 60 ,'D':18}},
        '02网络设备': {'rows': [1],'cols': [7,10,11,12],'row_heights': {None: 14.4},'col_widths': {'A': 9, 'B': 36, 'C': 18 ,'D':35 ,'E':35 ,'F':35 ,'G':9 ,'H':18 }},
        '03安全设备': {'rows': [1],'cols': [7,10,11,12],'row_heights': {None: 14.4},'col_widths':  {'A': 9, 'B': 36, 'C': 18 ,'D':35 ,'E':35 ,'F':35 ,'G':9 ,'H':18 }},
        '04业务应用软件一平台': {'rows': [1],'cols': [7,8,9],'row_heights': {None: 14.4},'col_widths':  {'A': 9, 'B': 36, 'C': 18 ,'D':35 ,'E':35 ,'F':18 }},
        '05系统管理平台一全局扩展': {'rows': [1],'cols': [6,8,9,10],'row_heights': {None: 14.4},'col_widths':  {'A': 9, 'B': 36, 'C': 18 ,'D':35 ,'E':35 ,'F':18 }},
        '06服务器一存储设备': {'rows': [1],'cols': [8,11,12,13],'row_heights': {None: 14.4},'col_widths':  {'A': 9, 'B': 36, 'C': 18 ,'D':35 ,'E':35 ,'F':35,'G':22 ,'H':9,'I':18}},
        '07终端一感知设备一现场设备': {'rows': [1],'cols': [9,10,11],'row_heights': {None: 14.4},'col_widths': {'A': 9, 'B': 36, 'C': 18 ,'D':35 ,'E':35 ,'F':35,'G':22 ,'H':9}},
        '08数据库管理系统': {'rows': [1],'cols': [11,12,13],'row_heights': {None: 14.4},'col_widths': {'A': 9, 'B': 36, 'C': 18 ,'D':35 ,'E':35 ,'F':35,'G':35 ,'H':35,'I':18,'J':18}},
        '09关键数据类别': {'rows': [1],'cols': [6,7,8,9,10,11,13,14],'row_heights': {None: 14.4},'col_widths': {'A': 9, 'B': 18, 'C': 35 ,'D':35 ,'E':35 ,'F':18}},
        '10密码产品': {'rows': [1],'cols': [8],'row_heights': {None: 14.4},'col_widths': {'A': 9, 'B': 36, 'C': 36 ,'D':35 ,'E':35 ,'F':35,'G':9}},
        '11安全相关人员': {'rows': [1],'cols': [6],'row_heights': {None: 14.4},'col_widths': {'A': 9, 'B': 18, 'C': 62 ,'D':23 ,'E':55 }},
        '12安全管理文档': {'rows': [1],'cols': [4],'row_heights': {None: 14.4},'col_widths':{'A': 9, 'B': 50, 'C': 162 }}
    }

    # 遍历每个sheet并进行删除和美化操作
    for sheet_name, deletions in sheet_deletions.items():

        sheet = workbook[sheet_name]
        # 删除指定行
        for row_idx in sorted(deletions['rows'], reverse=True):
            sheet.delete_rows(row_idx)
        # 删除指定列
        for col_idx in sorted(deletions['cols'], reverse=True):
            sheet.delete_cols(col_idx)
        # 设置全表白色填充
        for row in sheet.iter_rows():
            for cell in row:
                cell.fill = PatternFill(start_color='FFFFFF', end_color='FFFFFF', fill_type='solid')
        # 设置指定列宽
        for col_letter, width in deletions['col_widths'].items():
            sheet.column_dimensions[col_letter].width = width
        # 设置默认行高
        default_height = deletions['row_heights'].get(None)
        if default_height:
            sheet.default_row_height = default_height
        # 设置特定行高
        for row_idx, height in deletions['row_heights'].items():
            if row_idx is not None:
                sheet.row_dimensions[row_idx].height = height
        # 设置标题加粗蓝底
        header_fill = PatternFill(start_color='0066CC', end_color='0066CC', fill_type='solid')
        header_font = Font(bold=True, color='FFFFFF')
        for cell in sheet[1]:
            cell.fill = header_fill
            cell.font = header_font

        # 设置单元格内外边框
        border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
        for row in sheet.iter_rows():
            for cell in row:
                cell.border = border

        # 设置单元格字体居中
        for row in sheet.iter_rows():
            for cell in row:
                cell.alignment = Alignment(horizontal='center', vertical='center')

    # 保存修改后的Excel文件
    workbook.save(folder_path)

# 主程序循环
while True:
    start()

    # 提示用户选择下一步操作
    choice = input("\n按 1 重新进行下一次任务,按其他任意键输入则退出程序: ")

    if choice != "1":
        break

print("程序已退出。")