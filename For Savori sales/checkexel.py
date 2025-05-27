import os
import openpyxl

folder = r"C:\Users\User\Dropbox\DO & INV\DO & INV 2025"

for file in os.listdir(folder):
    if file.lower().endswith('.xlsx') and 'sales summary' in file.lower():
        file_path = os.path.join(folder, file)
        print(f"\n正在处理文件: {file}")

        try:
            wb = openpyxl.load_workbook(file_path)
            modified = False

            for sheet_name in wb.sheetnames:
                if 'sales' in sheet_name.lower():
                    ws = wb[sheet_name]
                    print(f"  - 处理Sheet: {sheet_name}")
                    row_count = 0
                    modify_count = 0

                    for row in ws.iter_rows(min_row=2, min_col=7, max_col=7):  # G列
                        cell = row[0]
                        row_count += 1
                        old_value = cell.value

                        if old_value and isinstance(old_value, str):
                            new_value = old_value.strip()
                            if new_value != old_value:
                                print(f"    * 第{cell.row}行: '{old_value}' -> '{new_value}'")
                                cell.value = new_value
                                modified = True
                                modify_count += 1

                    print(f"    → 共检查 {row_count} 行, 修改 {modify_count} 行。")

            if modified:
                wb.save(file_path)
                print(f"文件已保存: {file}")
            else:
                print("未检测到需要修改的内容。")

        except Exception as e:
            print(f"处理文件时出错: {file}\n  错误详情: {e}")
