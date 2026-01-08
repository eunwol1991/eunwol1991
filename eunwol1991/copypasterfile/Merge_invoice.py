import os
import re
from PyPDF2 import PdfMerger

# 根目录路径
ROOT_DIR = r"C:\Users\jhunj\Dropbox\DO & INV\DO & INV 2026"

# 匹配发票 PDF 文件名
invoice_pattern = re.compile(
    r'^(.*?)\s*([0-9]{4})\s*-\s*([0-9]{3})\s*-\s*INV(.*)\.pdf$',
    re.IGNORECASE
)

# 全局变量（运行时再赋值）
TARGET_MONTH = None
folder_month_pattern = None


def contains_supplier(path: str) -> bool:
    skip_keywords = ['supplier', 'sarpino', 'canadian pizza', 'stuffd',
                     'cash sales', 'staff purchase', 'rite pizza',
                     'Alt PIzza', 'ICON Steak']
    parts = os.path.normpath(path).split(os.sep)
    return any(any(keyword in part.lower() for keyword in skip_keywords) for part in parts)


def get_month_from_folder(folder_name: str) -> str:
    m = folder_month_pattern.match(folder_name)
    return m.group(1).capitalize() if m else ""


def extract_prefix(pdf_filename: str) -> str:
    m = invoice_pattern.match(pdf_filename)
    return m.group(1).strip() if m else ""


def is_cancelled(pdf_name: str) -> bool:
    return bool(re.search(r'(?i)cancel', pdf_name))


def pick_latest_pdf(pdf_list):
    base_map = {}
    for pdf in pdf_list:
        base_name = re.sub(r'(?i)\(revised\)', '', pdf)
        base_name = base_name.replace('__', '_').replace('  ', ' ').strip()
        if base_name not in base_map or re.search(r'(?i)revised', pdf):
            base_map[base_name] = pdf
    return list(base_map.values())


def merge_pdfs_in_folder(folder_path: str, pdf_files: list, output_name: str):
    if len(pdf_files) < 2:
        print(f"⚠️ 只有 1 份 PDF，无需合并。")
        return

    merger = PdfMerger()
    pdf_files.sort()

    print(f"  🔗 合并以下 PDF：")
    for pdf in pdf_files:
        pdf_path = os.path.join(folder_path, pdf)
        print(f"    + {os.path.basename(pdf_path)}")
        try:
            merger.append(pdf_path)
        except Exception as e:
            print(f"    [ERROR] 合并出错 {os.path.basename(pdf_path)}: {e}")

    out_path = os.path.join(folder_path, output_name)
    try:
        merger.write(out_path)
        merger.close()
        print(f"✅ 合并完成：{output_name}\n")
    except Exception as e:
        print(f"[ERROR] 写入合并 PDF 失败: {e}")


def process_folder(folder_path: str):
    folder_name = os.path.basename(folder_path)
    month = get_month_from_folder(folder_name)
    if not month:
        print(f"❌ 跳过目录: '{folder_name}'，不符合 '{TARGET_MONTH}. Xxx' 格式。")
        return

    try:
        all_entries = os.listdir(folder_path)
    except Exception as e:
        print(f"⚠️ 无法读取目录 '{folder_path}': {e}")
        return

    matched_pdfs = [
        f for f in all_entries
        if invoice_pattern.match(f) and not is_cancelled(f)
    ]
    matched_pdfs = pick_latest_pdf(matched_pdfs)

    if not matched_pdfs:
        print(f"ℹ️ 目录 '{folder_path}' 无匹配发票 PDF。")
        return

    print(f"\n🗂️ 处理目录：{folder_path}")
    print(f"  📄 匹配到 {len(matched_pdfs)} 个 PDF：")
    for pdf in matched_pdfs:
        print(f"    · {pdf}")

    prefixes = {extract_prefix(f) for f in matched_pdfs}
    prefixes = {p for p in prefixes if p}

    if len(prefixes) != 1:
        print(f"❌ 多前缀冲突: {prefixes}，跳过该目录。")
        return

    prefix = list(prefixes)[0]
    prefix_safe = re.sub(r'[<>:"/\\|?*]', '_', prefix)

    output_name = f"{prefix_safe} INV - {month}'25.pdf".strip()
    merge_pdfs_in_folder(folder_path, matched_pdfs, output_name)


def recursive_search(current_dir: str):
    if contains_supplier(current_dir):
        print(f"[SKIP] '{current_dir}' (contains supplier-related keywords).")
        return

    process_folder(current_dir)

    try:
        with os.scandir(current_dir) as it:
            for entry in it:
                if entry.is_dir():
                    recursive_search(entry.path)
    except Exception as e:
        print(f"[ERROR] Scanning subdirs of '{current_dir}': {e}")


def main():
    global TARGET_MONTH, folder_month_pattern

    if not os.path.isdir(ROOT_DIR):
        print(f"[ERROR] '{ROOT_DIR}' is not a valid directory.")
        return

    try:
        month_input = input("请输入要处理的月份 (1-12): ").strip()
        TARGET_MONTH = int(month_input)
    except ValueError:
        print("❌ 输入无效，请输入 1–12 的数字。")
        return

    # 运行时再生成正则
    folder_month_pattern = re.compile(
        rf'^{TARGET_MONTH}\.\s*([A-Za-z]{{3}})$', re.IGNORECASE)

    print(f"[START] Searching from: {ROOT_DIR}")
    recursive_search(ROOT_DIR)
    print("[DONE] Finished.")


if __name__ == '__main__':
    main()
