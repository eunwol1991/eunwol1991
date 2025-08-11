import os
from tqdm import tqdm
import win32com.client as win32

# 目标文件夹路径
FOLDER_PATH = r"C:\Users\User\Dropbox\DO & INV\DO & INV 2025"

# 初始化 Excel
excel = win32.Dispatch("Excel.Application")
excel.Visible = False
excel.DisplayAlerts = False

skipped_files = []

def set_black_white(file_path):
    try:
        wb = excel.Workbooks.Open(Filename=file_path, UpdateLinks=0, ReadOnly=False)
        count = 0
        for sheet in wb.Sheets:
            try:
                sheet.PageSetup.BlackAndWhite = True
                count += 1
            except Exception as e:
                print(f"  [Sheet Skip] {sheet.Name} in {file_path} — {e}")
        wb.Save()
        wb.Close(SaveChanges=True)
        return count
    except Exception as e:
        skipped_files.append(file_path)
        return None

def main():
    all_files = []
    for root, _, files in os.walk(FOLDER_PATH):
        for file in files:
            if file.lower().endswith((".xlsx", ".xls", ".xlsm")):
                all_files.append(os.path.join(root, file))

    for file_path in tqdm(all_files, desc="处理Excel文件", unit="file"):
        count = set_black_white(file_path)
        if count is not None:
            print(f"[OK] {file_path} — 设置为黑白的表：{count}")
        else:
            print(f"[FAIL] {file_path}")

    if skipped_files:
        print("\n跳过的文件：")
        for f in skipped_files:
            print(f)

    excel.Quit()

if __name__ == "__main__":
    main()
