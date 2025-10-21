# === 新增：放在文件顶部 import 之后 ===
import os
import re
from PyPDF2 import PdfMerger

# 提取 3 位发票编号（与你现有的正则保持一致）
_INV_NUM_RE = re.compile(r'-\s*(\d{3})\s*-')

def _get_inv_key(filename: str):
    """从文件名提取用于判重的 INV 号码（3 位）。"""
    m = _INV_NUM_RE.search(filename)
    return m.group(1) if m else None

def _filter_and_dedupe(pdf_paths: list[str]) -> list[str]:
    """
    1) 丢弃包含 '(Cancel)' 的文件
    2) 按 INV 编号分组；若组内存在 '(Revised)'，只保留该(些) Revised；
       - 若多个 Revised，保留修改时间最新的一份
       - 若没有 Revised，保留修改时间最新的一份（普通版）
    3) 维持按 INV 编号从小到大输出（也可按你原始顺序需要调整）
    """
    # 先排除 Cancel（不区分大小写）
    kept = [p for p in pdf_paths if "(cancel)" not in os.path.basename(p).lower()]

    groups = {}
    for p in kept:
        key = _get_inv_key(os.path.basename(p))
        if key is None:
            # 没有编号的也可以按需求决定是否纳入；这里：纳入一个“no-key”分组避免漏掉
            key = f"_nokey_{os.path.basename(p)}"
        groups.setdefault(key, []).append(p)

    result = []
    for key in sorted(groups.keys()):
        candidates = groups[key]
        # 先挑 Revised（不区分大小写）
        revised = [p for p in candidates if "(revised)" in os.path.basename(p).lower()]
        pool = revised if revised else candidates
        # 选最近修改时间的那个
        best = max(pool, key=lambda p: os.path.getmtime(p))
        result.append(best)

    return result


def merge_INV_pdfs_in_range(directory, start_num, end_num):
    pdf_files = []
    pattern = re.compile(r'-\s*(\d{3})\s*-')
    for filename in os.listdir(directory):
        if filename.endswith(".pdf"):
            match = pattern.search(filename)
            if match:
                file_num = int(match.group(1))
                fname_low = filename.lower()
                if start_num <= file_num <= end_num and "inv" in fname_low:
                    pdf_files.append(os.path.join(directory, filename))
    # 新增：过滤 + 去重
    return _filter_and_dedupe(pdf_files)


def merge_INV_pdfs_by_keywords(directory, keywords):
    pdf_files = []
    for filename in os.listdir(directory):
        fname_low = filename.lower()
        kw_low = [k.lower() for k in keywords]

        # 用 fname_low 判断扩展名和 INV，避免大小写漏检
        if fname_low.endswith(".pdf") and "inv" in fname_low:
            for k in kw_low:
                if k in fname_low:
                    pdf_files.append(os.path.join(directory, filename))
                    break

    return _filter_and_dedupe(pdf_files)



def merge_all_inv_pdfs(directory):
    pdf_files = []
    for filename in os.listdir(directory):
        fname_low = filename.lower()
        if filename.endswith(".pdf") and "inv" in fname_low:
            pdf_files.append(os.path.join(directory, filename))

    return _filter_and_dedupe(pdf_files)  # 新增

def merge_INV_pdfs_by_numbers(directory, numbers):
    pdf_files = []
    pattern = re.compile(r'-\s*(\d{3})\s*-')
    for filename in os.listdir(directory):
        if filename.endswith(".pdf"):
            match = pattern.search(filename)
            if match:
                file_num = int(match.group(1))
                fname_low = filename.lower()
                if file_num in numbers and "inv" in fname_low:
                    pdf_files.append(os.path.join(directory, filename))

    return _filter_and_dedupe(pdf_files)  # 新增


def merge_pdfs(pdf_files, output_filename):
    try:
        merger = PdfMerger()
        for pdf in pdf_files:
            merger.append(pdf)
        
        output_dir = os.path.dirname(output_filename)
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)

        merger.write(output_filename)
        merger.close()
        print(f"合并后的PDF文件已保存为: {output_filename}")
    except Exception as e:
        print(f"合并PDF文件时出错: {e}")

def main():
    directory = input("请输入文件夹路径: ").strip()
    if not os.path.isdir(directory):
        print("无效的文件夹路径。")
        return

    choice = input("请选择功能选项：1. 按范围查找PDF 2. 按关键词查找PDF 3. 合并所有带有INV的文件 4. 按数字序列查找PDF\n请输入1、2、3或4: ").strip()

    if choice == '1':
        # 范围
        start_num = int(input("请输入起始数字序号: ").strip())
        end_num = int(input("请输入结束数字序号: ").strip())
        if start_num > end_num:
            print("起始序号不能大于结束序号。")
            return
        pdf_files = merge_INV_pdfs_in_range(directory, start_num, end_num)

    elif choice == '2':
        # 关键词
        keywords_input = input("请输入关键词，用逗号分隔（例如：stadium,park,arena）: ").strip()
        keywords = [k.strip() for k in keywords_input.split(',') if k.strip()]
        pdf_files = merge_INV_pdfs_by_keywords(directory, keywords)

    elif choice == '3':
        # 全部 INV
        pdf_files = merge_all_inv_pdfs(directory)

    elif choice == '4':
        # 指定编号序列
        numbers_input = input("请输入数字序列，用逗号分隔（例如：004,015,100）: ").strip()
        numbers = [int(n.strip()) for n in numbers_input.split(',') if n.strip()]
        pdf_files = merge_INV_pdfs_by_numbers(directory, numbers)

    else:
        print("无效的选项")
        return

    if not pdf_files:
        print("没有找到符合条件的PDF文件。")
        return

    output_filename = input("请输入输出合并PDF文件的路径和文件名（例如：output/merged.pdf）: ").strip()
    if not output_filename.endswith(".pdf"):
        print("输出文件名必须以 .pdf 结尾。")
        return

    merge_pdfs(pdf_files, output_filename)

if __name__ == "__main__":
    main()
