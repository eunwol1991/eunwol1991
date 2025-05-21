import os
import xlwings as xw

def process_excel_files(folder_path):
    """
    批量处理 Excel 文件：
    - 对 "do" 工作表，选中 A1:K11，并将焦点移到 K11。
    - 对 "invoice" 工作表，选中 A1:I11，并将焦点移到 I11。
    - 只启动一次 Excel，提高效率。
    """

    # 判断文件夹是否存在
    if not os.path.isdir(folder_path):
        print(f'路径 "{folder_path}" 不存在或不是文件夹。')
        return

    # 只启动一次 Excel 应用
    app = xw.App(visible=False)  # 设为 False，不显示 Excel 窗口，提高运行速度
    app.display_alerts = False
    app.screen_updating = False

    try:
        # 统计处理的文件数量
        processed_count = 0  

        # 遍历文件夹内的所有 Excel 文件
        for filename in os.listdir(folder_path):
            if filename.lower().endswith(('.xlsx', '.xlsm', '.xlsb', '.xls')):
                file_path = os.path.join(folder_path, filename)
                print(f'正在处理文件: {file_path}')

                try:
                    wb = app.books.open(file_path)  # 打开 Excel 文件
                    modified = False

                    # 设定目标工作表和对应的目标单元格
                    target_sheets = {
                        'do': ('A1:K11', 'K11'),
                        'invoice': ('A1:I11', 'I11')
                    }

                    # 遍历工作表
                    for sheet in wb.sheets:
                        sheet_name = sheet.name.lower()
                        if sheet_name in target_sheets:
                            range_to_select, focus_cell = target_sheets[sheet_name]

                            # 选中 A1 到指定列的 11 行区域
                            try:
                                sheet.api.Application.Goto(sheet.range(range_to_select).api, True)
                                sheet.range(focus_cell).select()
                                print(f'在 "{sheet.name}" 选中 {range_to_select} 并定位 {focus_cell}')
                                modified = True
                            except Exception as e:
                                print(f'⚠️ 处理 "{sheet.name}" 时出错: {e}')

                    # 只有修改了工作表，才保存文件
                    if modified:
                        wb.save()

                    # 关闭当前文件（避免 Excel 进程占用过多内存）
                    wb.close()
                    processed_count += 1

                except Exception as e:
                    print(f'❌ 处理文件 {file_path} 时出错: {e}')

        print(f'✅ 处理完成，共修改了 {processed_count} 个文件。')

    finally:
        app.quit()  # 处理完所有文件后，关闭 Excel 实例

if __name__ == '__main__':
    folder_path = r"C:\Users\User\Dropbox\DO & INV\DO & INV 2025\Melvin - MOS Burger\For Customer\MOS DOC (OTL) - Format"
    process_excel_files(folder_path)
