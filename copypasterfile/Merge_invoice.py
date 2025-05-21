import os
import re
from PyPDF2 import PdfMerger

# 根目录路径
ROOT_DIR = r"C:\Users\User\Dropbox\DO & INV\DO & INV 2025"

# 👇 改这里控制合并哪一月
TARGET_MONTH = 5

# 匹配发票 PDF 文件名
invoice_pattern = re.compile(
    r'^(.*?)\s*([0-9]{4})\s*-\s*([0-9]{3})\s*-\s*INV(.*)\.pdf$',
    re.IGNORECASE
)

# 匹配 "3. Mar" 这样的文件夹名
folder_month_pattern = re.compile(rf'^{TARGET_MONTH}\.\s*([A-Za-z]{{3}})$', re.IGNORECASE)

def contains_supplier(path: str) -> bool:
    skip_keywords = ['supplier', 'sarpino', 'canadian pizza', 'stuffd', 'cash sales', 'staff purchase','rite pizza']
    parts = os.path.normpath(path).split(os.sep)
    return any(any(keyword in part.lower() for keyword in skip_keywords) for part in parts)

def get_month_from_folder(folder_name: str) -> str:
    m = folder_month_pattern.match(folder_name)
    return m.group(1).capitalize() if m else ""

def extract_prefix(pdf_filename: str) -> str:
    m = invoice_pattern.match(pdf_filename)
    return m.group(1).strip() if m else ""

def merge_pdfs_in_folder(folder_path: str, pdf_files: list, output_name: str):
    if len(pdf_files) < 2:
        print(f"  [INFO] Only one matched PDF in '{folder_path}', skip merging.")
        return

    merger = PdfMerger()
    pdf_files.sort()

    for pdf in pdf_files:
        pdf_path = os.path.join(folder_path, pdf)
        print(f"    + {pdf_path}")
        try:
            merger.append(pdf_path)
        except Exception as e:
            print(f"    [ERROR] Merging {pdf_path}: {e}")

    out_path = os.path.join(folder_path, output_name)
    try:
        merger.write(out_path)
        merger.close()
        print(f"[DONE] Created => {out_path}")
    except Exception as e:
        print(f"[ERROR] Writing merged PDF: {e}")

def process_folder(folder_path: str):
    folder_name = os.path.basename(folder_path)
    month = get_month_from_folder(folder_name)
    if not month:
        print(f"[SKIP] Folder '{folder_name}' does not match '{TARGET_MONTH}. Xxx' format.")
        return

    try:
        all_entries = os.listdir(folder_path)
    except Exception as e:
        print(f"[ERROR] Cannot list folder '{folder_path}': {e}")
        return

    matched_pdfs = [f for f in all_entries if invoice_pattern.match(f)]

    if not matched_pdfs:
        print(f"[INFO] No invoice PDF matched in '{folder_path}'.")
        return

    print(f"\n[INFO] In '{folder_path}' => matched PDFs: {matched_pdfs}")

    # 检查 prefix 是否一致
    prefixes = {extract_prefix(f) for f in matched_pdfs}
    prefixes = {p for p in prefixes if p}

    if len(prefixes) != 1:
        print(f"  [ERROR] Multiple prefixes found: {prefixes} — skipping.")
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
    if not os.path.isdir(ROOT_DIR):
        print(f"[ERROR] '{ROOT_DIR}' is not a valid directory.")
        return

    print(f"[START] Searching from: {ROOT_DIR}")
    recursive_search(ROOT_DIR)
    print("[DONE] Finished.")

if __name__ == '__main__':
    main()
