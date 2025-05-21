import os
import re
from collections import defaultdict

BASE_DIR = r"C:\Users\User\Dropbox\DO & INV\DO & INV 2025"

# 精准匹配 INV 或 DO & INV（大小写不敏感）
valid_tags = {"INV", "DO & INV"}

pattern = re.compile(
    r"^(.+?)\s*(\d{4})\s*-\s*(\d{3})\s*-\s*([A-Z &]+)", re.IGNORECASE
)

files = defaultdict(lambda: defaultdict(lambda: defaultdict(lambda: {"xlsx": [], "pdf": []})))
invalid_files = []

for root, _, filenames in os.walk(BASE_DIR):
    for filename in filenames:
        ext = os.path.splitext(filename)[1].lower()
        if ext not in [".xlsx", ".pdf"]:
            continue

        match = pattern.match(filename)
        # 只保留指定月份年份的文件（例如 "0425"）
        target_month = "0525" #editable variable
        if target_month not in filename:
            continue

        if not match:
            continue  # 跳过无编号或无类型的文件

        prefix, year, num, doc_type_raw = match.groups()
        doc_type = doc_type_raw.strip().upper()
        number = int(num)
        prefix = prefix.strip()
        year = year.strip()
        path = os.path.join(root, filename)

        # 跳过无 INV / DO & INV 类型的命名（如只是 xxx - 001）
        if doc_type not in valid_tags:
            # 但如果 doc_type 是拼错的（如 NV），则视为命名错误
            if "INV" in doc_type or "NV" in doc_type or "IN" == doc_type:
                invalid_files.append(path)
            continue

        if ext == ".xlsx" and doc_type == "DO & INV":
            files[prefix][year][number]["xlsx"].append(path)
        elif ext == ".pdf" and doc_type == "INV":
            files[prefix][year][number]["pdf"].append(path)

# === 输出报告 ===
print("📂 文件检查报告 v2.5")
print("=" * 50)

# 1. 命名错误文件
print("\n❌ 命名错误的文件（例如 INV 拼成 NV）：")
if invalid_files:
    for f in invalid_files:
        print(f" -", f)
else:
    print(" ✅ 无")

# 2. 缺配对的文件（非 C.P）
print("\n🔗 缺配对文件（只检查有 .xlsx 的 prefix）：")
unpaired = False
for prefix, year_map in files.items():
    if prefix.upper() == "C.P":
        continue  # C.P 不检查配对
    for year, num_map in year_map.items():
        for number, f in num_map.items():
            if f["xlsx"] and not f["pdf"]:
                unpaired = True
                print(f" - [{prefix}] {year} - {number:03d} 缺少 INV (.pdf)")
if not unpaired:
    print(" ✅ 所有 DO & INV 文件均配对 INV")

# 3. 重复编号（xlsx / pdf 超过一个）
print("\n🔁 重复编号检查：")
duplicate = False
for prefix, year_map in files.items():
    for year, num_map in year_map.items():
        for number, f in num_map.items():
            if len(f["xlsx"]) > 1:
                duplicate = True
                print(f" - [{prefix}] {year} - {number:03d} 有多个 DO & INV (.xlsx):")
                for path in f["xlsx"]:
                    print(f"    • {path}")
            if len(f["pdf"]) > 1:
                duplicate = True
                print(f" - [{prefix}] {year} - {number:03d} 有多个 INV (.pdf):")
                for path in f["pdf"]:
                    print(f"    • {path}")
if not duplicate:
    print(" ✅ 无重复编号")

# 4. 编号不连续（只看 .xlsx）
print("\n📉 编号不连续（以 .xlsx 为基准）：")
gap = False
for prefix, year_map in files.items():
    for year, num_map in year_map.items():
        if prefix.upper() == "C.P":
            base = {n for n, v in num_map.items() if v["xlsx"]}
        else:
            base = {n for n, v in num_map.items() if v["xlsx"]}

        if not base:
            continue
        sorted_nums = sorted(base)
        expected = set(range(min(sorted_nums), max(sorted_nums) + 1))
        missing = sorted(expected - base)
        if missing:
            gap = True
            print(f" - [{prefix}] {year} 缺少编号：{', '.join(f'{n:03d}' for n in missing)}")
if not gap:
    print(" ✅ 所有编号连续")

print("\n🎯 检查完成。")
