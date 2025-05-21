import os
import openpyxl
import traceback

def update_excel_files(directory, debug=False):
    """
    批量更新 Excel 文件，提高效率并增强稳定性。
    """
    processed_count = 0
    error_files = []

    for root, _, files in os.walk(directory):
        for file in files:
            if "xx25" in file.lower() and file.lower().endswith(('.xlsx', '.xlsm')):
                file_path = os.path.join(root, file)
                if debug:
                    print(f"🔄 正在处理: {file_path}")

                workbook = None
                try:
                    workbook = openpyxl.load_workbook(file_path)
                    if debug:
                        print(f"✅ 已加载: {file_path}")
                        print(f"📄 工作表: {workbook.sheetnames}")

                    modified = False

                    # 遍历 "do" 和 "invoice" 工作表，不跳出循环，处理所有匹配行
                    for sheet_name in ["DO", "Invoice"]:
                        if sheet_name in workbook.sheetnames:
                            sheet = workbook[sheet_name]
                            if debug:
                                print(f"📌 处理工作表: {sheet_name}")
                            for row in sheet.iter_rows():
                                for cell in row:
                                    cell_value = str(cell.value).strip().lower() if cell.value is not None else ""
                                    # 模糊匹配 "tartar sauce"
                                    cell_name = "honey mustard sauce"
                                    
                                    if cell_name in cell_value:
                                        row_number = cell.row
                                        product_code = "SD2136"
                                        product_description = "Honey Mustard Sauce"
                                        pack_size = "(10 x 1kg)"
                                        qty = 7.00
                                        UOM = "PKTS"
                                        selling_price = f"=7.5*G{row_number}"
                                        if debug:
                                            print(f"🔍 匹配到 {cell_name} → Sheet: {sheet_name}, Row: {row_number}")
                                        if sheet_name == "DO":
                                            sheet[f"B{row_number}"].value = product_code
                                            sheet[f"C{row_number}"].value = product_description
                                            sheet[f"G{row_number}"].value = pack_size
                                            sheet[f"I{row_number}"].value = qty
                                            sheet[f"K{row_number}"].value = UOM
                                            if debug:
                                                print(f"✅ DO 修改成功: B{row_number}, C{row_number}, G{row_number}, I{row_number}, K{row_number}")
                                        elif sheet_name == "Invoice":
                                            sheet[f"B{row_number}"].value = product_code
                                            sheet[f"C{row_number}"].value = product_description
                                            sheet[f"F{row_number}"].value = pack_size
                                            sheet[f"G{row_number}"].value = qty
                                            sheet[f"H{row_number}"].value = UOM
                                            sheet[f"I{row_number}"].value = selling_price
                                            if debug:
                                                print(f"✅ Invoice 修改成功: B{row_number}, C{row_number}, F{row_number}, G{row_number}, H{row_number}, I{row_number}")
                                        modified = True
                                        # 不跳出循环，继续处理所有匹配单元格
                    if modified:
                        workbook.save(file_path)
                        processed_count += 1
                        if debug:
                            print(f"💾 已保存修改: {file_path}")
                    workbook.close()
                except Exception as e:
                    traceback.print_exc()
                    error_files.append(file_path)
                    if debug:
                        print(f"❌ 处理失败: {file_path}，错误: {e}")
                finally:
                    if workbook is not None:
                        workbook.close()

    print(f"✅ 处理完成，共修改 {processed_count} 个文件。")
    if error_files:
        print("⚠️ 以下文件处理失败：")
        for error_file in error_files:
            print(f"  ❌ {error_file}")

# 示例调用
directory_path = r"C:\Users\User\Dropbox\DO & INV\DO & INV 2025\Anthony - Select(Waker Chicken)"
update_excel_files(directory_path, debug=True)
