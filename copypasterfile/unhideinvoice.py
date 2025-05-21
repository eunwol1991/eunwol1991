import os
import xlwings as xw

def activate_sheets(directory, debug=False):
    """
    批量处理 Excel 文件：
    1. 仅处理包含 "DO" 或 "Invoice" 工作表的 Excel 文件，否则跳过。
    2. 依次激活 "DO" → "Invoice" 工作表。
    3. 处理完成后保存并关闭文件。
    4. 仅处理文件名包含 "xx25" 的 Excel 文件 (.xlsx, .xlsm)。
    """

    # 只启动一次 Excel，提高效率
    app = xw.App(visible=True)  # 设置 True 可看到操作过程
    app.display_alerts = False
    app.screen_updating = False

    processed_count = 0  # 统计成功处理的文件数
    skipped_count = 0  # 统计跳过的文件数
    error_files = []  # 记录处理失败的文件

    for root, _, files in os.walk(directory):
        for file in files:
            # 只处理包含 "xx25" 的 Excel 文件
            if "xx25" in file.lower() and file.lower().endswith(('.xlsx', '.xlsm')):
                file_path = os.path.join(root, file)
                if debug:
                    print(f"🔄 正在处理: {file_path}")

                try:
                    wb = app.books.open(file_path)  # 打开 Excel 文件
                    sheet_names = [sheet.name.lower() for sheet in wb.sheets]

                    # 检查是否包含 "DO" 或 "Invoice" 工作表
                    if "do" not in sheet_names and "invoice" not in sheet_names:
                        wb.close()  # 关闭 Excel 文件
                        skipped_count += 1
                        if debug:
                            print(f"⚠️ 跳过: {file_path}（无 DO 或 Invoice 工作表）")
                        continue  # 跳过这个文件

                    # 依次选择 "DO" → "Invoice" 工作表
                    for sheet_name in ["do", "invoice"]:
                        if sheet_name in sheet_names:
                            wb.sheets[sheet_name].activate()
                            if debug:
                                print(f"✅ 已激活: {sheet_name} in {file}")

                    # 保存修改（如果不需要修改可以去掉 wb.save()）
                    wb.save()
                    processed_count += 1
                    if debug:
                        print(f"💾 已保存: {file_path}")

                    # 关闭当前 Excel 文件
                    wb.close()

                except Exception as e:
                    error_files.append(file_path)
                    if debug:
                        print(f"❌ 处理失败: {file_path}，错误: {e}")

    # 关闭 Excel
    app.quit()

    print(f"✅ 处理完成，共修改 {processed_count} 个文件，跳过 {skipped_count} 个文件。")
    if error_files:
        print("⚠️ 以下文件处理失败：")
        for error_file in error_files:
            print(f"  ❌ {error_file}")

# 设定目标路径
directory_path = r"C:\Users\User\Dropbox\DO & INV\DO & INV 2025"
activate_sheets(directory_path, debug=True)
