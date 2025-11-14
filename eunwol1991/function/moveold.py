import os
import re
import shutil
from collections import defaultdict

# === 配置 ===
ROOT_FOLDER = r"C:\Users\jhunj\Dropbox\Halal Update\Product Spec\Griffith"
HISTORY_FOLDER = r"C:\Users\jhunj\Dropbox\Halal Update\Product Spec\Griffith\History"
DRY_RUN = False  # 先设 True 仅查看将会移动哪些文件；确认无误后改为 False 执行移动

# 匹配示例：
# "CAJUN MARINADE PF9642A , REV.0 - SAVORI PTE LTD.pdf"
# "GARLIC BUTTER FLAVOUR MARINADE PF9644A, REV.3 - SAVORI PTE LTD.docx"
# 核心：以逗号前作为“系列名”，逗号后找 "REV.<数字>"
REV_REGEX = re.compile(
    r"^(?P<base>.+?),\s*REV\.?\s*(?P<rev>\d+)\b", re.IGNORECASE)


def scan_files(root):
    """返回目录下所有普通文件的绝对路径列表（不递归子文件夹）。"""
    files = []
    for name in os.listdir(root):
        path = os.path.join(root, name)
        if os.path.isfile(path):
            files.append(path)
    return files


def parse_rev_info(file_path):
    """
    从文件名解析 (base_key, rev_num)。
    base_key 用于分组：即逗号前的部分，去除首尾空白并压缩内部空格。
    若不匹配 REV，返回 (None, None)。
    """
    filename = os.path.basename(file_path)
    name_no_ext, ext = os.path.splitext(filename)
    m = REV_REGEX.match(name_no_ext)
    if not m:
        return None, None, None

    raw_base = m.group("base")
    rev_num = int(m.group("rev"))

    # 规范化 base 作为分组 key（忽略大小写差异与多余空格）
    base_norm = " ".join(raw_base.strip().split())
    base_key = base_norm.lower()

    # 将扩展名也纳入分组，避免不同类型文件互相影响（例如 PDF 与 DOCX）
    return (base_key, ext.lower()), rev_num, filename


def ensure_dir(path):
    if not os.path.exists(path):
        os.makedirs(path, exist_ok=True)


def main():
    ensure_dir(HISTORY_FOLDER)
    files = scan_files(ROOT_FOLDER)

    # 分组：key -> 列表[(rev, abs_path, filename)]
    groups = defaultdict(list)

    for abs_path in files:
        key, rev, filename = parse_rev_info(abs_path)
        if key is None:
            # 忽略不含 REV 的文件
            continue
        groups[key].append((rev, abs_path, filename))

    to_move = []

    for key, items in groups.items():
        # 找到该组最高 REV
        max_rev = max(r for r, _, _ in items)
        # 需要移动的是非最高的
        for r, abs_path, filename in items:
            if r < max_rev:
                to_move.append((r, abs_path, filename, max_rev, key))

    # 执行移动（或仅打印）
    if not to_move:
        print("没有需要移动的旧版本。")
        return

    print(f"共发现需要移动的旧版本文件：{len(to_move)} 个\n")

    # 为避免重名，若 History 中已存在同名文件，则在文件名后加 (n)
    def unique_dest(dest_dir, filename):
        base, ext = os.path.splitext(filename)
        candidate = os.path.join(dest_dir, filename)
        i = 1
        while os.path.exists(candidate):
            candidate = os.path.join(dest_dir, f"{base} ({i}){ext}")
            i += 1
        return candidate

    for r, src, filename, max_rev, key in sorted(to_move):
        dst = unique_dest(HISTORY_FOLDER, filename)
        rel = os.path.relpath(src, ROOT_FOLDER)
        print(
            f"[MOVE] {rel}  | 当前REV={r}, 最高REV={max_rev} -> {os.path.relpath(dst, ROOT_FOLDER)}")
        if not DRY_RUN:
            shutil.move(src, dst)

    if DRY_RUN:
        print("\nDRY_RUN 模式：未实际移动。确认输出无误后，将 DRY_RUN = False 再运行。")
    else:
        print("\n完成：旧版本已移动到 History。")


if __name__ == "__main__":
    main()
