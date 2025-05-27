import os
import openpyxl

folder = r"C:\Users\User\Dropbox\DO & INV"
print("程序启动，递归查找所有子目录下的文件……")

# os.walk递归遍历所有目录和文件
for dirpath, dirnames, filenames in os.walk(folder):
    for file in filenames:
        if file.lower().endswith('.xlsx') and 'sales summary' in file.lower():
            file_path = os.path.join(dirpath, file)
            print(f"\n==========\n打开文件: {file_path}")

            try:
                wb = openpyxl.load_workbook(file_path)
                modified = False
                sheet_found = False

                for sheet_name in wb.sheetnames:
                    if 'sales' in sheet_name.lower():
                        sheet_found = True
                        ws = wb[sheet_name]
                        print(f"  检查Sheet: {sheet_name}")
                        row_count = 0
                        modify_count = 0

                        for row in ws.iter_rows(min_row=2, min_col=7, max_col=7):  # G列
                            cell = row[0]
                            row_count += 1
                            old_value = cell.value

                            if old_value and isinstance(old_value, str):
                                new_value = old_value.strip()
                                if new_value != old_value:
                                    print(f"    - 第{cell.row}行 G列: 原='{old_value}' → 新='{new_value}'")
                                    cell.value = new_value
                                    modified = True
                                    modify_count += 1

                        print(f"    Sheet检查行数: {row_count}，被修改行数: {modify_count}")

                if not sheet_found:
                    print("  没有找到包含'sales'的Sheet。")
                if modified:
                    wb.save(file_path)
                    print(f"  结果: 已修改并保存文件。")
                else:
                    print("  结果: 没有需要修改的内容，无需保存。")

            except Exception as e:
                print(f"  错误: 处理文件时出错！\n    详情: {e}")
