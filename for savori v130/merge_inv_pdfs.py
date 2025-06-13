# programe name merge_inv_pdfs.py
import os
import re
from PyPDF2 import PdfMerger

def merge_INV_pdfs_in_range(directory, start_num, end_num):
    pdf_files = []
    pattern = re.compile(r'-\s*(\d{3})\s*-')
    for filename in os.listdir(directory):
        if filename.endswith(".pdf"):
            match = pattern.search(filename)
            if match:
                file_num = int(match.group(1))
                if start_num <= file_num <= end_num and "INV" in filename:
                    pdf_files.append(os.path.join(directory, filename))
    return pdf_files

def merge_INV_pdfs_by_keywords(directory, keywords):
    pdf_files = []
    for filename in os.listdir(directory):
        if filename.endswith(".pdf"):
            for keyword in keywords:
                if keyword in filename and "INV" in filename:
                    pdf_files.append(os.path.join(directory, filename))
                    break
    return pdf_files

def merge_all_inv_pdfs(directory):
    pdf_files = []
    for filename in os.listdir(directory):
        if filename.endswith(".pdf") and "INV" in filename:
            pdf_files.append(os.path.join(directory, filename))
    return pdf_files

def merge_INV_pdfs_by_numbers(directory, numbers):
    pdf_files = []
    pattern = re.compile(r'-\s*(\d{3})\s*-')
    for filename in os.listdir(directory):
        if filename.endswith(".pdf"):
            match = pattern.search(filename)
            if match:
                file_num = int(match.group(1))
                if file_num in numbers and "INV" in filename:
                    pdf_files.append(os.path.join(directory, filename))
    return pdf_files

def merge_pdfs(pdf_files, output_filename):
    try:
        merger = PdfMerger()
        for pdf in pdf_files:
            merger.append(pdf)
        
        output_dir = os.path.dirname(output_filename)
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)

        merger.write(output_filename)
        merger.close()
        print(f"合并后的PDF文件已保存为: {output_filename}")
    except Exception as e:
        print(f"合并PDF文件时出错: {e}")

def main():
    directory = input("请输入文件夹路径: ").strip()
    if not os.path.isdir(directory):
        print("无效的文件夹路径。")
        return

    choice = input("请选择功能选项：1. 按范围查找PDF 2. 按关键词查找PDF 3. 合并所有带有INV的文件 4. 按数字序列查找PDF\n请输入1、2、3或4: ").strip()

    if choice == '2':
        try:
            start_num = int(input("请输入起始数字序号: ").strip())
            end_num = int(input("请输入结束数字序号: ").strip())
            if start_num > end_num:
                print("起始序号不能大于结束序号。")
                return
            pdf_files = merge_INV_pdfs_in_range(directory, start_num, end_num)
        except ValueError:
            print("请输入有效的数字序号。")
            return
    elif choice == '3':
        keywords_input = input("请输入关键词，用逗号分隔（例如：stadium,park,arena）: ").strip()
        keywords = [keyword.strip() for keyword in keywords_input.split(',')]
        pdf_files = merge_INV_pdfs_by_keywords(directory, keywords)
    elif choice == '4':
        pdf_files = merge_all_inv_pdfs(directory)
    elif choice == '5':
        numbers_input = input("请输入数字序列，用逗号分隔（例如：004,015,100）: ").strip()
        numbers = [int(num.strip()) for num in numbers_input.split(',')]
        pdf_files = merge_INV_pdfs_by_numbers(directory, numbers)
    else:
        print("无效的选项")
        return

    if not pdf_files:
        print("没有找到符合条件的PDF文件。")
        return

    output_filename = input("请输入输出合并PDF文件的路径和文件名（例如：output/merged.pdf）: ").strip()
    if not output_filename.endswith(".pdf"):
        print("输出文件名必须以 .pdf 结尾。")
        return

    merge_pdfs(pdf_files, output_filename)

if __name__ == "__main__":
    main()
