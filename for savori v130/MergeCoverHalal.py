import os
from PyPDF2 import PdfMerger

def merge_cover_to_all(input_folder, output_folder, cover_filename):
    # 获取封面完整路径
    cover_path = os.path.join(input_folder, cover_filename)

    # 检查封面是否存在
    if not os.path.exists(cover_path):
        print(f"❌ 找不到封面文件：{cover_path}")
        return

    # 创建输出文件夹（如不存在）
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    # 遍历所有 PDF
    for filename in os.listdir(input_folder):
        if filename.endswith(".pdf") and filename != cover_filename:
            target_path = os.path.join(input_folder, filename)
            output_path = os.path.join(output_folder, filename)

            try:
                merger = PdfMerger()
                merger.append(cover_path)
                merger.append(target_path)
                merger.write(output_path)
                merger.close()
                print(f"✅ 已合并 → {output_path}")
            except Exception as e:
                print(f"❌ 合并失败 → {filename}，错误：{e}")

# === 参数设置 ===
input_path = r"C:\Users\User\Downloads\Re_ Innofrsh_Savori request for updated Halal cert (Exp_ May & Jun'25)"
output_path = r"C:\Users\User\Downloads\halal cert"
cover_pdf_name = "A477-2005 ; Innofresh exp.03-05-2026 cover EN.pdf"

# === 执行函数 ===
merge_cover_to_all(input_path, output_path, cover_pdf_name)
