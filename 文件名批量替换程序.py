import os
import re

def perform_replace_task():
    # 输入文件夹路径
    folder_path = input("\n请输入要处理的文件夹路径: \n")
    # 确保输入的路径存在
    if not os.path.exists(folder_path):
        print(f"路径 {folder_path} 不存在!")
        return

    # 输入要替换的字符串和对应的新字符串
    replace_rules = input("\n请输入替换规则(格式: 旧字段1==新字段1,旧字段2==新字段2,...): \n")

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


# 主程序循环
while True:
    perform_replace_task()

    # 提示用户选择下一步操作
    choice = input("\n按 1 重新进行下一次替换任务,按其他任意键输入则退出程序: ")

    if choice != "1":
        break

print("程序已退出。")