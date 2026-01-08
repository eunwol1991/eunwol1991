import os
import re
from collections import defaultdict
import fitz  # pip install pymupdf
from colorama import Fore, Style, init
from difflib import SequenceMatcher
init(autoreset=True)


BASE_DIR = r"C:\Users\jhunj\Dropbox\DO & INV\DO & INV 2026"
valid_tags = {"INV", "DO & INV"}
pattern = re.compile(
    r"^(.+?)\s*(\d{4})\s*-\s*(\d{3})\s*-\s*([A-Z &]+)", re.IGNORECASE
)
target_month = "0126"  # 可编辑：筛选包含该字符串的月份标识（如 0825/0925）


def is_cancelled(fname):
    return bool(re.search(r'(?i)cancel', fname))


files = defaultdict(lambda: defaultdict(
    lambda: defaultdict(lambda: {"xlsx": [], "pdf": []})))
invalid_files = []


def extract_invoice_number_from_pdf(pdf_path):
    try:
        doc = fitz.open(pdf_path)
        text = ""
        for page in doc[:2]:
            text += page.get_text()
        match = re.search(
            r'Invoice No[\s:]*([\w. ]+)\s*(\d{4})\s*-\s*(\d{3})', text, re.IGNORECASE)
        if match:
            prefix, year, num = match.groups()
            return f"{prefix.strip()} {year.strip()} - {num.strip()}"
        match2 = re.search(r'([\w. ]+)\s*(\d{4})\s*-\s*(\d{3})', text)
        if match2:
            prefix, year, num = match2.groups()
            return f"{prefix.strip()} {year.strip()} - {num.strip()}"
    except Exception as e:
        print(f"⚠️ 读取 PDF 内容失败: {pdf_path} ({e})")
    return None


def print_report_header(title, char="─", width=60):
    print()
    print(f"{title} ".ljust(width, char))


def shorten_path(path, keep=3):
    parts = path.split(os.sep)
    if len(parts) <= keep + 1:
        return path
    return "…" + os.sep + os.sep.join(parts[-keep-1:])


def highlight_diff(a, b):
    """只高亮不同部分"""
    matcher = SequenceMatcher(None, a, b)
    result_a = ""
    result_b = ""
    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        if tag == 'equal':
            result_a += a[i1:i2]
            result_b += b[j1:j2]
        else:
            # 内容编号不同部分：黄色；文件名编号不同部分：青色
            result_a += f"{Fore.YELLOW}{a[i1:i2]}{Style.RESET_ALL}"
            result_b += f"{Fore.CYAN}{b[j1:j2]}{Style.RESET_ALL}"
    return result_a, result_b

# =================== 文件扫描与归档 ===================


for root, _, filenames in os.walk(BASE_DIR):
    for filename in filenames:
        # 跳过 cancel 文件（excel和pdf）
        if is_cancelled(filename):
            continue

        ext = os.path.splitext(filename)[1].lower()
        if ext not in [".xlsx", ".pdf"]:
            continue

        # 既支持文件名中包含（如 xxx_0825.pdf），也支持父级文件夹名包含（如 ...\\0825\\xxx.pdf）
        if target_month:
            if (target_month not in filename) and (target_month not in root):
                continue

        match = pattern.match(filename)
        if not match:
            continue  # 跳过无编号或无类型的文件

        prefix, year, num, doc_type_raw = match.groups()
        doc_type = doc_type_raw.strip().upper()
        number = int(num)
        prefix = prefix.strip()
        year = year.strip()
        path = os.path.join(root, filename)

        if doc_type not in valid_tags:
            if "INV" in doc_type or "NV" in doc_type or "IN" == doc_type:
                invalid_files.append(path)
            continue

        if ext == ".xlsx" and doc_type == "DO & INV":
            files[prefix][year][number]["xlsx"].append(path)
        elif ext == ".pdf" and doc_type == "INV":
            files[prefix][year][number]["pdf"].append(path)

# ========== 输出美化报告 ==========

print("\n📂 文件检查报告 v2.6 (PyMuPDF+cancel过滤+彩色高亮)")
print("=" * 64)
if target_month:
    print(f"筛选月份关键词: {target_month}")

# 1. 命名错误
print_report_header("❌ 命名错误的文件（例如 INV 拼成 NV）")
if invalid_files:
    for f in invalid_files:
        print(f"  - {shorten_path(f)}")
else:
    print("  ✅ 无")

# 2. 缺配对
print_report_header("🔗 缺配对文件（只检查有 .xlsx 的 prefix）")
unpaired = False
for prefix, year_map in files.items():
    if prefix.upper() == "C.P":
        continue
    for year, num_map in year_map.items():
        for number, f in num_map.items():
            if f["xlsx"] and not f["pdf"]:
                unpaired = True
                print(f"  - [{prefix}] {year} - {number:03d} 缺少 INV (.pdf)")
if not unpaired:
    print("  ✅ 所有 DO & INV 文件均配对 INV")

# 3. 重复编号
print_report_header("🔁 重复编号检查")
duplicate = False
for prefix, year_map in files.items():
    for year, num_map in year_map.items():
        for number, f in num_map.items():
            if len(f["xlsx"]) > 1 or len(f["pdf"]) > 1:
                duplicate = True
                print(f"  - [{prefix}] {year} - {number:03d}")
                if len(f["xlsx"]) > 1:
                    print("    多个 DO & INV (.xlsx):")
                    for path in f["xlsx"]:
                        print(f"      • {shorten_path(path)}")
                if len(f["pdf"]) > 1:
                    print("    多个 INV (.pdf):")
                    for path in f["pdf"]:
                        print(f"      • {shorten_path(path)}")
if not duplicate:
    print("  ✅ 无重复编号")

# 4. 编号不连续
print_report_header("📉 编号不连续（以 .xlsx 为基准）")
gap = False
for prefix, year_map in files.items():
    for year, num_map in year_map.items():
        base = {n for n, v in num_map.items() if v["xlsx"]}
        if not base:
            continue
        sorted_nums = sorted(base)
        expected = set(range(min(sorted_nums), max(sorted_nums) + 1))
        missing = sorted(expected - base)
        if missing:
            gap = True
            print(
                f"  - [{prefix}] {year} 缺少编号：{', '.join(f'{n:03d}' for n in missing)}")
if not gap:
    print("  ✅ 所有编号连续")

# 5. 内容编号不符，高亮差异
print_report_header("📋 文件内容编号与文件名编号不符")
mismatch_files = []
for prefix, year_map in files.items():
    for year, num_map in year_map.items():
        for number, f in num_map.items():
            file_no = f"{prefix} {year} - {number:03d}"
            for pdf_path in f["pdf"]:
                content_no = extract_invoice_number_from_pdf(pdf_path)
                if not content_no:
                    mismatch_files.append((pdf_path, "无法识别内容编号"))
                elif content_no != file_no:
                    mismatch_files.append(
                        (pdf_path, f"内容编号: {content_no}，文件名: {file_no}"))
if mismatch_files:
    grouped = {}
    for f, reason in mismatch_files:
        basename = os.path.basename(f)
        folder = os.path.dirname(f).split(os.sep)[-2:]
        key = " / ".join(folder)
        grouped.setdefault(key, []).append((basename, reason))
    for folder, items in grouped.items():
        print(f"  - {folder}/")
        for basename, reason in items:
            # 只高亮差异
            if "内容编号:" in reason and "文件名:" in reason:
                pattern = r"内容编号: (.*?)[，,]文件名: (.*)"
                m = re.search(pattern, reason)
                if m:
                    content_no, file_no = m.groups()
                    diff_content_no, diff_file_no = highlight_diff(
                        content_no, file_no)
                    reason_colored = f"内容编号: {diff_content_no}，文件名: {diff_file_no}"
                else:
                    reason_colored = reason
            elif "无法识别" in reason:
                reason_colored = f"{Fore.RED}{reason}{Style.RESET_ALL}"
            else:
                reason_colored = reason
            print(f"      • {basename} [{reason_colored}]")
else:
    print("  ✅ 所有 PDF 文件内容编号与文件名一致")

print("\n🎯 检查完成。\n" + "=" * 64)
