import os
import re

def read_excel_files(directory):
    excel_files = []
    for filename in os.listdir(directory):
        if (filename.endswith(".xlsx") or filename.endswith(".xls")) and not filename.startswith("~$"):
            excel_files.append(filename)
    return excel_files

def extract_number_from_filename(filename):
    match = re.search(r'\d{4} - (\d{3})', filename)
    if match:
        return int(match.group(1))
    return None

def debug_matching_files(directory, output_directory, excel_files, condition):
    matched_files = []
    for filename in excel_files:
        if re.search(r'MOS \d{4}', filename) and "DO & INV" in filename:
            excel_path = os.path.join(directory, filename)
            file_number = extract_number_from_filename(filename)
            if file_number is not None and condition(file_number):
                matched_files.append(excel_path)
                print(f"匹配成功的文件: {excel_path} (文件编号: {file_number})")
    return matched_files

def convert_all_excels(directory, output_directory, excel_files):
    condition = lambda number: True
    matched_files = debug_matching_files(directory, output_directory, excel_files, condition)
    print(f"总匹配文件数: {len(matched_files)}")

def convert_range_excels(directory, output_directory, start_num, end_num, excel_files):
    print(f"转换范围: {str(start_num).zfill(3)} 到 {str(end_num).zfill(3)}")
    condition = lambda number: start_num <= number <= end_num
    matched_files = debug_matching_files(directory, output_directory, excel_files, condition)
    print(f"总匹配文件数: {len(matched_files)}")

def convert_keyword_excels(directory, output_directory, keyword, excel_files):
    print(f"转换包含关键词 '{keyword}' 的文件")
    condition = lambda filename: keyword in filename
    matched_files = debug_matching_files(directory, output_directory, excel_files, condition)
    print(f"总匹配文件数: {len(matched_files)}")

def main():
    directory = input("请输入Excel文件所在的目录路径: ")
    output_directory = input("请输入输出目录路径: ")
    if not os.path.isdir(directory):
        print(f"路径无效: {directory}")
        return
    if not os.path.isdir(output_directory):
        os.makedirs(output_directory)

    # 读取目录中的所有 Excel 文件
    excel_files = read_excel_files(directory)

    print("请选择功能选项：")
    print("1. 转换全部")
    print("2. 输入数字序号只转换范围内的文件")
    print("3. 输入关键词转换相关的文件")
    choice = input("请输入1、2或3: ")

    if choice == '1':
        convert_all_excels(directory, output_directory, excel_files)
    elif choice == '2':
        start_num = int(input("请输入起始数字序号: "))
        end_num = int(input("请输入结束数字序号: "))
        convert_range_excels(directory, output_directory, start_num, end_num, excel_files)
    elif choice == '3':
        keyword = input("请输入关键词: ")
        convert_keyword_excels(directory, output_directory, keyword, excel_files)
    else:
        print("无效的选项")

if __name__ == "__main__":
    main()
