#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
INV PDF 合并工具（支持按编号分组 + Revised 变体优先）

使用规则：
1) 丢弃包含 "(Cancel)" 的文件（大小写不敏感）
2) 以文件名里的 3 位 INV 编号（形如 "- 172 -"）分组
   - 若组内存在 "(Revised)"，只保留该组所有 Revised 变体（A/B/...）
   - 若组内不存在 "(Revised)"，保留该组所有变体（A/B/...）
3) 输出顺序：按 INV 编号升序；组内按文件名升序
4) 无法解析出 3 位编号的文件，会放在最后（按文件名升序）

依赖：PyPDF2
安装：pip install PyPDF2
"""

import os
import re
import sys
from typing import List
from PyPDF2 import PdfMerger

# 提取 3 位编号（例如文件名里出现 "- 172 -"）
_INV_NUM_RE = re.compile(r'-\s*(\d{3})\s*-')


def _get_inv_key(filename: str):
    """从文件名提取用于分组的 3 位 INV 号码。找不到则返回 None。"""
    m = _INV_NUM_RE.search(filename)
    return m.group(1) if m else None


def _filter_and_dedupe(pdf_paths: List[str]) -> List[str]:
    """
    规则：
    1) 丢弃含 '(Cancel)' 的文件（不区分大小写）
    2) 按 3 位 INV 编号分组；若组内存在 '(Revised)'，只保留该组所有 Revised 变体；
       若无 Revised，则保留该组所有变体（A/B/...）
    3) 输出顺序：按 INV 编号升序；组内按文件名升序
    4) 无法解析出 3 位编号的文件，统一归入 '_nokey' 组，置于最后，按文件名升序
    """
    # 1) 先排除 Cancel
    kept = [
        p for p in pdf_paths if "(cancel)" not in os.path.basename(p).lower()]

    # 2) 分组
    groups: dict[str, List[str]] = {}
    nokey: List[str] = []
    for p in kept:
        base = os.path.basename(p)
        key = _get_inv_key(base)
        if key is None:
            nokey.append(p)  # 无编号
        else:
            groups.setdefault(key, []).append(p)

    # 3) 生成输出（按规则挑选）
    result: List[str] = []

    # 有编号的组：按编号升序
    for key in sorted(groups.keys(), key=lambda x: int(x)):
        candidates = groups[key]
        revised = [
            p for p in candidates if "(revised)" in os.path.basename(p).lower()]
        pool = revised if revised else candidates
        pool_sorted = sorted(pool, key=lambda p: os.path.basename(p).lower())
        result.extend(pool_sorted)

    # 无编号的放最后
    if nokey:
        nokey_sorted = sorted(nokey, key=lambda p: os.path.basename(p).lower())
        result.extend(nokey_sorted)

    return result


def merge_INV_pdfs_in_range(directory: str, start_num: int, end_num: int) -> List[str]:
    """按编号范围选择 INV PDF 并返回经过滤后的路径列表。"""
    pdf_files: List[str] = []
    pattern = re.compile(r'-\s*(\d{3})\s*-')
    for filename in os.listdir(directory):
        fname_low = filename.lower()
        if fname_low.endswith(".pdf"):
            match = pattern.search(filename)
            if match:
                file_num = int(match.group(1))
                if start_num <= file_num <= end_num and "inv" in fname_low:
                    pdf_files.append(os.path.join(directory, filename))
    return _filter_and_dedupe(pdf_files)


def merge_INV_pdfs_by_keywords(directory: str, keywords: List[str]) -> List[str]:
    """按关键词匹配文件名（包含任一关键词）"""
    pdf_files: List[str] = []
    kw_low = [k.lower() for k in keywords]
    for filename in os.listdir(directory):
        fname_low = filename.lower()
        if fname_low.endswith(".pdf") and "inv" in fname_low:
            if any(k in fname_low for k in kw_low):
                pdf_files.append(os.path.join(directory, filename))
    return _filter_and_dedupe(pdf_files)


def merge_all_inv_pdfs(directory: str) -> List[str]:
    """目录下所有文件名包含 'inv' 的 PDF 全部纳入后过滤。"""
    pdf_files: List[str] = []
    for filename in os.listdir(directory):
        fname_low = filename.lower()
        if fname_low.endswith(".pdf") and "inv" in fname_low:
            pdf_files.append(os.path.join(directory, filename))
    return _filter_and_dedupe(pdf_files)


def merge_INV_pdfs_by_numbers(directory: str, numbers: List[int]) -> List[str]:
    """按给定编号序列（整数列表）选择 INV PDF。"""
    pdf_files: List[str] = []
    pattern = re.compile(r'-\s*(\d{3})\s*-')
    target_set = set(numbers)
    for filename in os.listdir(directory):
        fname_low = filename.lower()
        if fname_low.endswith(".pdf"):
            match = pattern.search(filename)
            if match:
                file_num = int(match.group(1))
                if file_num in target_set and "inv" in fname_low:
                    pdf_files.append(os.path.join(directory, filename))
    return _filter_and_dedupe(pdf_files)


def merge_pdfs(pdf_files: List[str], output_filename: str) -> None:
    """把给定的 PDF 列表按顺序合并到 output_filename。"""
    if not pdf_files:
        print("没有可合并的 PDF。")
        return

    try:
        merger = PdfMerger()
        for pdf in pdf_files:
            merger.append(pdf)

        out_dir = os.path.dirname(os.path.abspath(output_filename)) or "."
        if not os.path.exists(out_dir):
            os.makedirs(out_dir, exist_ok=True)

        merger.write(output_filename)
        merger.close()
        print(f"合并完成，共 {len(pdf_files)} 个文件 → {output_filename}")
    except Exception as e:
        print(f"合并 PDF 时出错：{e}")


def _ask_int(prompt: str) -> int:
    while True:
        s = input(prompt).strip()
        try:
            return int(s)
        except ValueError:
            print("请输入整数。")


def main():
    print("=== INV PDF 合并工具 ===")
    directory = input("请输入文件夹路径: ").strip()
    if not os.path.isdir(directory):
        print("无效的文件夹路径。")
        sys.exit(1)

    print("功能选项：")
    print("  1) 按编号范围")
    print("  2) 按关键词")
    print("  3) 合并全部 INV")
    print("  4) 按编号序列")
    choice = input("请输入 1 / 2 / 3 / 4: ").strip()

    if choice == "1":
        start_num = _ask_int("请输入起始编号（三位整数，如 1~999 都可，程序会自动比较）: ")
        end_num = _ask_int("请输入结束编号: ")
        if start_num > end_num:
            print("起始编号不能大于结束编号。")
            sys.exit(1)
        pdf_files = merge_INV_pdfs_in_range(directory, start_num, end_num)

    elif choice == "2":
        keywords_input = input("请输入关键词，以逗号分隔（例如：mos,canadian,172）: ").strip()
        keywords = [k.strip() for k in keywords_input.split(",") if k.strip()]
        if not keywords:
            print("未提供关键词。")
            sys.exit(1)
        pdf_files = merge_INV_pdfs_by_keywords(directory, keywords)

    elif choice == "3":
        pdf_files = merge_all_inv_pdfs(directory)

    elif choice == "4":
        numbers_input = input("请输入编号列表，以逗号分隔（例如：4,15,172）: ").strip()
        try:
            numbers = [int(n.strip())
                       for n in numbers_input.split(",") if n.strip()]
        except ValueError:
            print("编号必须是整数（形如 4, 15, 172）。")
            sys.exit(1)
        if not numbers:
            print("未提供编号。")
            sys.exit(1)
        pdf_files = merge_INV_pdfs_by_numbers(directory, numbers)

    else:
        print("无效选项。")
        sys.exit(1)

    if not pdf_files:
        print("没有找到符合条件的 PDF 文件。")
        sys.exit(0)

    print("\n即将合并以下文件（顺序同输出）：")
    for i, p in enumerate(pdf_files, 1):
        print(f"{i:03d}. {os.path.basename(p)}")

    output_filename = input("\n请输入输出文件路径（如 output/merged_inv.pdf）: ").strip()
    if not output_filename.lower().endswith(".pdf"):
        print("输出文件名必须以 .pdf 结尾。")
        sys.exit(1)

    merge_pdfs(pdf_files, output_filename)


if __name__ == "__main__":
    main()
