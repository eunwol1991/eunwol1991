# programe name convert_inv_pdf.py
import os
import re
import win32com.client as win32
import time

def read_excel_files(directory):
    print(f"Reading Excel files from directory: {directory}")
    excel_files = []
    for filename in os.listdir(directory):
        if (filename.endswith(".xlsx") or filename.endswith(".xls")) and not filename.startswith("~$"):
            excel_files.append(filename)
    print(f"Found {len(excel_files)} Excel files.")
    return excel_files

def extract_number_from_filename(filename):
    match = re.search(r'\d{4} - (\d{3})', filename)
    if match:
        print(f"Extracted number {match.group(1)} from filename {filename}")
        return int(match.group(1))
    print(f"No number found in filename {filename}")
    return None
def excel_to_pdf(excel_path, output_path):
    for attempt in range(3):
        try:
            print(f"Attempting to open Excel file: {excel_path} (Attempt {attempt + 1})")
            excel = win32.DispatchEx('Excel.Application')  # 使用 DispatchEx
            excel.Visible = False

            wb = excel.Workbooks.Open(excel_path)

            try:
                ws = wb.Worksheets('Invoice')
                pdf_filename = create_pdf_filename(excel_path)
                pdf_path = os.path.join(output_path, pdf_filename)
                ws.ExportAsFixedFormat(0, pdf_path)
                print(f"PDF saved as {pdf_path}")
                break
            except Exception as e:
                print(f"Error saving PDF: {e}")
            finally:
                wb.Close(SaveChanges=False)
                excel.Quit()
        except Exception as e:
            print(f"Error opening file on attempt {attempt + 1}: {e}")
            time.sleep(2)
    else:
        print(f"Failed to open file after multiple attempts: {excel_path}")

def create_pdf_filename(excel_path):
    excel_filename = os.path.basename(excel_path)
    base_name = re.sub(r"DO & INV", "INV", excel_filename).replace(".xlsx", "").replace(".xls", "")
    pdf_filename = base_name + ".pdf"
    print(f"Created PDF filename: {pdf_filename}")
    return pdf_filename

def debug_matching_files(directory, output_directory, excel_files, condition, is_keyword=False):
    print(f"Starting to debug matching files in directory: {directory}")
    matched_files = []
    for filename in excel_files:
        print(f"Processing file: {filename}")
        if re.search(r' \d{4}', filename) and "DO & INV" in filename:
            excel_path = os.path.join(directory, filename)
            if is_keyword:
                if condition(filename):
                    matched_files.append(excel_path)
                    print(f"Matched file: {excel_path} (Filename: {filename})")
            else:
                file_number = extract_number_from_filename(filename)
                print(f"File number: {file_number}")
                if file_number is not None and condition(file_number):
                    matched_files.append(excel_path)
                    print(f"Matched file: {excel_path} (File number: {file_number})")
    print(f"Total matched files: {len(matched_files)}")
    return matched_files

def convert_all_excels(directory, output_directory, excel_files):
    condition = lambda number: True
    matched_files = debug_matching_files(directory, output_directory, excel_files, condition)
    for file in matched_files:
        excel_to_pdf(file, output_directory)

def convert_range_excels(directory, output_directory, start_num, end_num, excel_files):
    print(f"Converting files in range: {str(start_num).zfill(3)} to {str(end_num).zfill(3)}")
    condition = lambda number: start_num <= number <= end_num
    matched_files = debug_matching_files(directory, output_directory, excel_files, condition)
    for file in matched_files:
        excel_to_pdf(file, output_directory)

def convert_keyword_excels(directory, output_directory, keyword, excel_files):
    print(f"Converting files containing keyword '{keyword}'")
    condition = lambda filename: keyword in filename
    matched_files = debug_matching_files(directory, output_directory, excel_files, condition, is_keyword=True)
    for file in matched_files:
        excel_to_pdf(file, output_directory)

def main():
    directory = input("请输入Excel文件所在的目录路径: ")
    output_directory = input("请输入输出目录路径: ")
    if not os.path.isdir(directory):
        print(f"Invalid directory: {directory}")
        return
    if not os.path.isdir(output_directory):
        os.makedirs(output_directory)
        print(f"Output directory created: {output_directory}")

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




