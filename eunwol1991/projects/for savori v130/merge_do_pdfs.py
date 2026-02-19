# programe name merge_do_pdfs.py
import os
from PyPDF2 import PdfMerger

def merge_do_pdfs_in_range(directory, start_num, end_num):
    pdf_files = []
    directory = os.path.abspath(directory)
    
    for filename in os.listdir(directory):
        if filename.endswith(".pdf"):
            for num in range(start_num, end_num + 1):
                num_str = str(num).zfill(3)
                if num_str in filename and "DO" in filename:
                    pdf_files.append(os.path.join(directory, filename))
                    break
    return pdf_files

def merge_pdfs(pdf_files, output_filename):
    merger = PdfMerger()
    for pdf in pdf_files:
        merger.append(pdf)
    
    output_dir = os.path.dirname(output_filename)
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    merger.write(output_filename)
    merger.close()

def main():
    directory = input("请输入文件夹路径: ")
    start_num = int(input("请输入起始数字序号: "))
    end_num = int(input("请输入结束数字序号: "))
    output_filename = input("请输入输出合并PDF文件的路径和文件名（例如：output/merged.pdf）: ")

    if not output_filename.endswith(".pdf"):
        print("输出文件名必须以 .pdf 结尾。")
        return

    pdf_files = merge_do_pdfs_in_range(directory, start_num, end_num)
    if not pdf_files:
        print("没有找到符合条件的PDF文件。")
        return

    merge_pdfs(pdf_files, output_filename)
    print(f"合并后的PDF文件已保存为: {output_filename}")

if __name__ == "__main__":
    main()
