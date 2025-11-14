import os
import re
import threading
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from tkinter import filedialog, messagebox, StringVar

# ====== 你现成的生成模块（Excel -> PDF）======
# DO
from convert_do_pdf import (
    convert_all_excels as convert_all_do_excels,
    convert_range_excels as convert_range_do_excels,
)
# INV
from convert_inv_pdf import (
    convert_all_excels as convert_all_inv_excels,
    convert_range_excels as convert_range_inv_excels,
    convert_keyword_excels as convert_keyword_inv_excels,  # 不再暴露到 UI，但保留导入以兼容
)

# ====== 合并用 ======
from PyPDF2 import PdfMerger

# ------------------------
# 解析编号/范围/后缀：返回 [(num:int, suffix:Optional[str])]
# 支持：2,5,14 / 12-18 / 173-A / 173 - B；中英文逗号；忽略空项
# ------------------------


def parse_id_list(text: str):
    cleaned = (text or "").replace("，", ",")
    parts = [p.strip() for p in cleaned.split(",")]

    ids = []
    seen = set()

    def push(n: int, sfx: str | None):
        key = (n, sfx or None)
        if key not in seen:
            seen.add(key)
            ids.append(key)

    for p in parts:
        if not p:
            continue

        # 数字范围：12-18
        m = re.fullmatch(r"\s*(\d+)\s*-\s*(\d+)\s*", p)
        if m:
            a, b = int(m.group(1)), int(m.group(2))
            if a > b:
                a, b = b, a
            for n in range(a, b + 1):
                push(n, None)
            continue

        # 单个编号 + 可选字母后缀：173、173A、173-A、173 - A
        m = re.fullmatch(r"\s*(\d+)\s*(?:[-_ ]\s*)?([A-Za-z])?\s*", p)
        if m:
            n = int(m.group(1))
            sfx = m.group(2).upper() if m.group(2) else None
            push(n, sfx)
            continue

        messagebox.showerror("错误", f"无效的编号片段：{p}\n示例：2,5,14 或 12-18 或 173-A")
        return None

    if not ids:
        messagebox.showerror("错误", "请输入至少一个编号，例如：2,5,14 或 12-18 或 173-A")
        return None

    # 去重后按编号、后缀排序（A 在前，None 在后）
    ids.sort(key=lambda t: (t[0], t[1] is None, t[1] or ""))
    return ids

# ------------------------
# 在目录中按 “(数字, 可选字母后缀)” 找 PDF
# kind: 'INV' | 'DO' | 'ALL'
# ------------------------


def find_pdfs_by_ids(root_dir: str, ids, kind: str):
    pdfs = []
    all_pdfs = []

    # 收集候选 PDF
    for dirpath, _, files in os.walk(root_dir):
        for f in files:
            if not f.lower().endswith(".pdf"):
                continue
            path = os.path.join(dirpath, f)
            name = f.lower()
            if kind == "INV" and "inv" not in name:
                continue
            if kind == "DO" and "do" not in name:
                continue
            all_pdfs.append(path)

    for (n, sfx) in ids:
        n3 = f"{n:03d}"
        found = None

        if sfx:
            # 明确后缀
            patterns = [
                rf"(?<!\d){n3}\s*[-_ ]\s*{sfx}(?![A-Za-z0-9])",
                rf"(?<!\d){n}\s*[-_ ]\s*{sfx}(?![A-Za-z0-9])",
                rf"(?<!\d){n3}{sfx}(?![A-Za-z0-9])",
                rf"(?<!\d){n}{sfx}(?![A-Za-z0-9])",
            ]
        else:
            # 不带后缀：允许纯数字或任意后缀
            patterns = [
                rf"(?<!\d){n3}(?!\d)(?:\s*[-_ ]\s*[A-Za-z])?",
                rf"(?<!\d){n}(?!\d)(?:\s*[-_ ]\s*[A-Za-z])?",
            ]

        for p in all_pdfs:
            base = os.path.basename(p).upper()
            if any(re.search(pat, base) for pat in patterns):
                found = p
                break

        pdfs.append(((n, sfx), found if found else None))

    return pdfs

# ------------------------
# 合并 PDF（按给定顺序）
# ------------------------


def merge_pdfs(selected_paths, output_path):
    merger = PdfMerger()
    try:
        for p in selected_paths:
            merger.append(p)
        os.makedirs(os.path.dirname(output_path), exist_ok=True)
        merger.write(output_path)
    finally:
        merger.close()

# ------------------------
# 将离散 id 压缩为若干连续区间，供 *range* 型转换函数调用
# 对含后缀的 id，按其数字部分作为单点区间处理
# ------------------------


def collapse_ids_to_ranges(ids):
    nums = sorted({n for (n, _s) in ids})
    if not nums:
        return []
    ranges = []
    start = prev = nums[0]
    for x in nums[1:]:
        if x == prev + 1:
            prev = x
            continue
        ranges.append((start, prev))
        start = prev = x
    ranges.append((start, prev))
    return ranges

# ------------------------
# GUI 应用（合并“生成参数”和“合并参数”）
# ------------------------


class AllInOneApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF 工具（生成/合并统一版）")
        self.root.geometry("1040x720")
        self.create_widgets()

    def create_widgets(self):
        font = ("Helvetica", 12)

        # 目录
        ttk.Label(self.root, text="工作目录（Excel/PDF所在）",
                  font=font).grid(row=0, column=0, padx=10, pady=10, sticky="w")
        self.dir_entry = ttk.Entry(self.root, width=60, font=font)
        self.dir_entry.grid(row=0, column=1, padx=10, pady=10, sticky="ew")
        ttk.Button(self.root, text="浏览", command=self.browse_dir,
                   bootstyle=PRIMARY).grid(row=0, column=2, padx=10, pady=10)

        # 输出目录（仅生成时使用）
        ttk.Label(self.root, text="输出目录（用于保存生成的 PDF）", font=font).grid(
            row=1, column=0, padx=10, pady=10, sticky="w")
        self.out_entry = ttk.Entry(self.root, width=60, font=font)
        self.out_entry.grid(row=1, column=1, padx=10, pady=10, sticky="ew")
        ttk.Button(self.root, text="浏览", command=self.browse_out,
                   bootstyle=PRIMARY).grid(row=1, column=2, padx=10, pady=10)

        # 操作模式：生成 / 合并
        ttk.Label(self.root, text="操作", font=font).grid(
            row=2, column=0, padx=10, pady=10, sticky="e")
        self.mode_var = StringVar(value="GENERATE")
        ttk.Radiobutton(self.root, text="生成 PDF", variable=self.mode_var, value="GENERATE",
                        bootstyle="success").grid(row=2, column=1, padx=(0, 10), pady=10, sticky="w")
        ttk.Radiobutton(self.root, text="合并 PDF", variable=self.mode_var, value="MERGE",
                        bootstyle="info").grid(row=2, column=1, padx=(180, 10), pady=10, sticky="w")

        # —— 统一参数区 ——
        uni = ttk.Labelframe(self.root, text="参数", bootstyle="secondary")
        uni.grid(row=3, column=0, columnspan=3, padx=10, pady=10, sticky="ew")

        ttk.Label(uni, text="匹配类型", font=font).grid(
            row=0, column=0, padx=10, pady=8, sticky="w")
        self.match_kind = StringVar(value="ALL")  # 统一：INV / DO / ALL
        ttk.Combobox(uni, textvariable=self.match_kind, state="readonly", values=[
                     "INV", "DO", "ALL"], font=font, width=10).grid(row=0, column=1, padx=10, pady=8, sticky="w")

        ttk.Label(uni, text="编号/范围/后缀（逗号分隔）", font=font).grid(row=0,
                                                              column=2, padx=10, pady=8, sticky="e")
        self.ids_entry = ttk.Entry(uni, font=font)
        self.ids_entry.grid(row=0, column=3, padx=10, pady=8, sticky="ew")

        uni.grid_columnconfigure(3, weight=1)

        # 列表显示（用于合并预览 & 生成时也可显示将处理的编号）
        self.tree = ttk.Treeview(self.root, columns=(
            "#1", "#2"), show="headings", height=16, bootstyle="info")
        self.tree.heading("#1", text="编号")
        self.tree.heading("#2", text="匹配文件 / 说明")
        self.tree.column("#1", width=160, anchor="center")
        self.tree.column("#2", anchor="w")
        self.tree.grid(row=4, column=0, columnspan=3,
                       padx=10, pady=8, sticky="nsew")
        self.tree.tag_configure("miss", foreground="#FF5555")

        # 状态 & 动作
        self.status = ttk.Label(
            self.root, text="", font=font, bootstyle="secondary")
        self.status.grid(row=5, column=0, columnspan=3,
                         padx=10, pady=(0, 10), sticky="w")

        ttk.Button(self.root, text="开始执行", command=self.on_run, bootstyle=SUCCESS).grid(
            row=6, column=0, padx=10, pady=12, sticky="ew")
        ttk.Button(self.root, text="退出", command=self.root.quit, bootstyle=SECONDARY).grid(
            row=6, column=2, padx=10, pady=12, sticky="ew")

        # 自适应
        for c in range(3):
            self.root.grid_columnconfigure(c, weight=1)
        self.root.grid_rowconfigure(4, weight=1)

    def browse_dir(self):
        d = filedialog.askdirectory()
        if d:
            self.dir_entry.delete(0, "end")
            self.dir_entry.insert(0, d)

    def browse_out(self):
        d = filedialog.askdirectory()
        if d:
            self.out_entry.delete(0, "end")
            self.out_entry.insert(0, d)

    # ---------------- 生成（按逗号分隔的 id / 范围） ----------------
    def do_generate(self, directory, output_directory):
        kind = self.match_kind.get()          # INV / DO / ALL

        ids = parse_id_list(self.ids_entry.get())
        if ids is None:
            return

        # 读取所有 Excel 文件名（传给现成函数所需）
        excel_files = []
        for filename in os.listdir(directory):
            if (filename.endswith(".xlsx") or filename.endswith(".xls")) and not filename.startswith("~$"):
                excel_files.append(filename)

        # 将离散编号压缩为连续区间，供 *range* 转换函数使用
        ranges = collapse_ids_to_ranges(ids)
        if not ranges:
            messagebox.showerror("错误", "未能解析有效编号")
            return

        def run_one(side: str):
            # 逐区间调用 range 转换；如需严格只转部分编号，范围函数内部应支持过滤
            for (s, e) in ranges:
                if side == "DO":
                    convert_range_do_excels(
                        directory, output_directory, s, e, excel_files)
                else:  # INV
                    convert_range_inv_excels(
                        directory, output_directory, s, e, excel_files)

        try:
            if kind in ("DO", "ALL"):
                run_one("DO")
            if kind in ("INV", "ALL"):
                run_one("INV")
            messagebox.showinfo("完成", f"生成完成：{kind} / 共 {len(ranges)} 段")
        except Exception as e:
            messagebox.showerror("错误", f"生成失败：{e}")

    # ---------------- 合并 ----------------
    def do_merge(self, directory):
        ids = parse_id_list(self.ids_entry.get())
        if ids is None:
            return
        kind = self.match_kind.get()

        # 预览列表
        self.tree.delete(*self.tree.get_children())
        pairs = find_pdfs_by_ids(directory, ids, kind)
        selected_paths, missing = [], []
        for (n, sfx), p in pairs:
            disp = f"{n:03d}" + (f"-{sfx}" if sfx else "")
            if p:
                self.tree.insert("", "end", values=(disp, os.path.basename(p)))
                selected_paths.append(p)
            else:
                self.tree.insert("", "end", values=(
                    disp, "未找到"), tags=("miss",))
                missing.append(disp)

        if not selected_paths:
            self.status.configure(text=f"未找到任何可合并的文件", bootstyle="danger")
            return

        if missing:
            self.status.configure(
                text=f"找到 {len(selected_paths)} 个；未找到 {len(missing)} 个：{', '.join(missing)}", bootstyle="warning")
        else:
            self.status.configure(
                text=f"全部找到，共 {len(selected_paths)} 个文件可合并", bootstyle="success")

        out_pdf = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf")],
            title="保存合并后的 PDF 为..."
        )
        if not out_pdf:
            return

        def task():
            try:
                merge_pdfs(selected_paths, out_pdf)
                messagebox.showinfo(
                    "完成", f"已合并 {len(selected_paths)} 个文件到：\n{out_pdf}")
            except Exception as e:
                messagebox.showerror("错误", f"合并失败：{e}")

        threading.Thread(target=task, daemon=True).start()

    # ---------------- 主入口 ----------------
    def on_run(self):
        directory = (self.dir_entry.get() or "").strip()
        if not directory or not os.path.isdir(directory):
            messagebox.showerror("错误", "请选择有效的工作目录")
            return

        mode = self.mode_var.get()

        if mode == "GENERATE":
            output_directory = (self.out_entry.get() or "").strip()
            if not output_directory:
                messagebox.showerror("错误", "请选择有效的输出目录（生成PDF保存到此）")
                return
            threading.Thread(target=self.do_generate, args=(
                directory, output_directory), daemon=True).start()
        else:
            threading.Thread(target=self.do_merge, args=(
                directory,), daemon=True).start()


if __name__ == "__main__":
    root = ttk.Window(themename="solar")
    app = AllInOneApp(root)
    root.mainloop()
