# programe name convert_do_pdf.py
import os
import re
import win32com.client as win32
from filelock import FileLock

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
def excel_to_pdf(excel_path, output_path):
    print(f"正在转换文件: {excel_path}")

    from win32com.client import makepy
    makepy.GenerateFromTypeLibSpec('Excel.Application')

    lock_path = excel_path + '.lock'
    lock = FileLock(lock_path)

    with lock:
        excel = win32.DispatchEx('Excel.Application')  # 使用 DispatchEx
        excel.Visible = False

        wb = excel.Workbooks.Open(excel_path)

        try:
            ws = wb.Worksheets('DO')
            pdf_filename = create_pdf_filename(excel_path)
            pdf_path = os.path.join(output_path, pdf_filename)
            ws.ExportAsFixedFormat(0, pdf_path)
            print(f"PDF saved as {pdf_path}")
        except Exception as e:
            print(f"Error: {e}")
        finally:
            wb.Close(SaveChanges=False)
            excel.Quit()

def create_pdf_filename(excel_path):
    excel_filename = os.path.basename(excel_path)
    base_name = re.sub(r"DO & INV", "DO", excel_filename).replace(".xlsx", "").replace(".xls", "")
    return base_name + ".pdf"

def convert_all_excels(directory, output_directory, excel_files):
    directory = os.path.abspath(directory)
    output_directory = os.path.abspath(output_directory)
    for filename in excel_files:
        if "DO & INV" in filename:
            excel_path = os.path.join(directory, filename)
            excel_to_pdf(excel_path, output_directory)

def convert_range_excels(directory, output_directory, start_num, end_num, excel_files):
    directory = os.path.abspath(directory)
    output_directory = os.path.abspath(output_directory)
    print(f"转换范围: {str(start_num).zfill(3)} 到 {str(end_num).zfill(3)}")
    for filename in excel_files:
        file_number = extract_number_from_filename(filename)
        if file_number is not None and start_num <= file_number <= end_num and "DO & INV" in filename:
            print(f"Found matching file: {filename}")
            excel_path = os.path.join(directory, filename)
            excel_to_pdf(excel_path, output_directory)

def main():
    directory = input("请输入Excel文件所在的目录路径: ")
    output_directory = input("请输入PDF文件输出目录路径: ")
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
    choice = input("请输入1或2: ")

    if choice == '1':
        convert_all_excels(directory, output_directory, excel_files)
    elif choice == '2':
        start_num = int(input("请输入起始数字序号: "))
        end_num = int(input("请输入结束数字序号: "))
        convert_range_excels(directory, output_directory, start_num, end_num, excel_files)
    else:
        print("无效的选项")

if __name__ == "__main__":
    main()
