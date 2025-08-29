import os
import re
import threading
from collections import defaultdict
from difflib import SequenceMatcher
try:
    import fitz  # pip install pymupdf
except Exception:  # pragma: no cover
    fitz = None
# 仅用于命令行彩色输出；GUI 中不会使用
try:
    from colorama import Fore, Style, init as colorama_init
    colorama_init(autoreset=True)
except Exception:  # pragma: no cover
    class _Dummy:
        RESET_ALL = ""
    class _Fore:
        RED = YELLOW = CYAN = ""
    Fore = _Fore()
    Style = _Dummy()
# ======== 可编辑默认值 ========
BASE_DIR = r"C:\\Users\\User\\Dropbox\\DO & INV\\DO & INV 2025"
TARGET_MONTH = "0925"  # 如：0825/0925
VALID_TAGS = {"INV", "DO & INV"}
FILENAME_PATTERN = re.compile(r"^(.+?)\s*(\d{4})\s*-\s*(\d{3})\s*-\s*([A-Z &]+)", re.IGNORECASE)
def is_cancelled(fname: str) -> bool:
    return bool(re.search(r"(?i)cancel", fname))
def shorten_path(path: str, keep: int = 3) -> str:
    parts = path.split(os.sep)
    if len(parts) <= keep + 1:
        return path
    return "…" + os.sep + os.sep.join(parts[-keep - 1 :])
def extract_invoice_number_from_pdf(pdf_path: str) -> str | None:
    if not fitz:
        return None
    try:
        doc = fitz.open(pdf_path)
        text = ""
        for page in doc[:2]:
            text += page.get_text() or ""
        match = re.search(r"Invoice No[\s:]*([\w. ]+)\s*(\d{4})\s*-\s*(\d{3})", text, re.IGNORECASE)
        if match:
            prefix, year, num = match.groups()
            return f"{prefix.strip()} {year.strip()} - {num.strip()}"
        match2 = re.search(r"([\w. ]+)\s*(\d{4})\s*-\s*(\d{3})", text)
        if match2:
            prefix, year, num = match2.groups()
            return f"{prefix.strip()} {year.strip()} - {num.strip()}"
    except Exception:
        return None
    return None
def diff_strings(a: str, b: str) -> tuple[str, str]:
    matcher = SequenceMatcher(None, a, b)
    ra, rb = "", ""
    for tag, i1, i2, j1, j2 in matcher.get_opcodes():
        if tag == "equal":
            ra += a[i1:i2]
            rb += b[j1:j2]
        else:
            ra += a[i1:i2]
            rb += b[j1:j2]
    return ra, rb
def scan_files(base_dir: str, target_month: str):
    """扫描目录，返回结构化结果，供 CLI/GUI 共用。"""
    files = defaultdict(lambda: defaultdict(lambda: defaultdict(lambda: {"xlsx": [], "pdf": []})))
    invalid_files: list[str] = []
    for root, _, filenames in os.walk(base_dir):
        for filename in filenames:
            if is_cancelled(filename):
                continue
            ext = os.path.splitext(filename)[1].lower()
            if ext not in [".xlsx", ".pdf"]:
                continue
            if target_month:
                if (target_month not in filename) and (target_month not in root):
                    continue
            m = FILENAME_PATTERN.match(filename)
            if not m:
                continue
            prefix, year, num, doc_type_raw = m.groups()
            doc_type = doc_type_raw.strip().upper()
            try:
                number = int(num)
            except ValueError:
                continue
            prefix = prefix.strip()
            year = year.strip()
            path = os.path.join(root, filename)
            if doc_type not in VALID_TAGS:
                if ("INV" in doc_type) or ("NV" in doc_type) or (doc_type == "IN"):
                    invalid_files.append(path)
                continue
            if ext == ".xlsx" and doc_type == "DO & INV":
                files[prefix][year][number]["xlsx"].append(path)
            elif ext == ".pdf" and doc_type == "INV":
                files[prefix][year][number]["pdf"].append(path)
    # 组装报告
    unpaired: list[tuple[str, str, int]] = []
    duplicates: list[dict] = []
    gaps: list[tuple[str, str, list[int]]] = []
    mismatches: list[tuple[str, str]] = []  # (pdf_path, reason)
    # 缺配对 & 重复
    for prefix, year_map in files.items():
        if prefix.upper() == "C.P":
            continue
        for year, num_map in year_map.items():
            # 重复、缺配
            for number, f in num_map.items():
                if f["xlsx"] and not f["pdf"]:
                    unpaired.append((prefix, year, number))
                if len(f["xlsx"]) > 1 or len(f["pdf"]) > 1:
                    duplicates.append(
                        {
                            "prefix": prefix,
                            "year": year,
                            "number": number,
                            "xlsx": list(f["xlsx"]),
                            "pdf": list(f["pdf"]),
                        }
                    )
            # 不连续（以 .xlsx 为基准）
            base = {n for n, v in num_map.items() if v["xlsx"]}
            if base:
                sorted_nums = sorted(base)
                expected = set(range(min(sorted_nums), max(sorted_nums) + 1))
                missing = sorted(expected - base)
                if missing:
                    gaps.append((prefix, year, missing))
            # 内容编号不符
            for number, f in num_map.items():
                file_no = f"{prefix} {year} - {number:03d}"
                for pdf_path in f["pdf"]:
                    content_no = extract_invoice_number_from_pdf(pdf_path)
                    if not content_no:
                        mismatches.append((pdf_path, "无法识别内容编号"))
                    elif content_no != file_no:
                        mismatches.append((pdf_path, f"内容编号: {content_no}，文件名: {file_no}"))
    return {
        "invalid_files": invalid_files,
        "unpaired": sorted(unpaired, key=lambda x: (x[0], x[1], x[2])),
        "duplicates": sorted(duplicates, key=lambda x: (x["prefix"], x["year"], x["number"])) ,
        "gaps": sorted(gaps, key=lambda x: (x[0], x[1])),
        "mismatches": mismatches,
    }
def _print_cli_report(base_dir: str, target_month: str):
    print("\n📂 文件检查报告 v3.0 (GUI/CLI 共用)")
    print("=" * 64)
    if target_month:
        print(f"筛选月份关键词: {target_month}")
    r = scan_files(base_dir, target_month)
    def header(title: str, char: str = "─", width: int = 60):
        print("\n" + f"{title} ".ljust(width, char))
    header("① 命名错误（INV 拼写类）")
    if r["invalid_files"]:
        for f in r["invalid_files"]:
            print(f"  - {shorten_path(f)}")
    else:
        print("  ✓ 无")
    header("② 缺配对（有 .xlsx 但缺 .pdf）")
    if r["unpaired"]:
        for prefix, year, num in r["unpaired"]:
            print(f"  - [{prefix}] {year} - {num:03d} 缺少 INV (.pdf)")
    else:
        print("  ✓ 全部配对")
    header("③ 重复编号")
    if r["duplicates"]:
        for item in r["duplicates"]:
            print(f"  - [{item['prefix']}] {item['year']} - {item['number']:03d}")
            if len(item["xlsx"]) > 1:
                print("    多个 DO & INV (.xlsx):")
                for p in item["xlsx"]:
                    print(f"      · {shorten_path(p)}")
            if len(item["pdf"]) > 1:
                print("    多个 INV (.pdf):")
                for p in item["pdf"]:
                    print(f"      · {shorten_path(p)}")
    else:
        print("  ✓ 无重复")
    header("④ 编号不连续（以 .xlsx 为基准）")
    if r["gaps"]:
        for prefix, year, miss in r["gaps"]:
            print(f"  - [{prefix}] {year} 缺少编号：{', '.join(f'{n:03d}' for n in miss)}")
    else:
        print("  ✓ 连续")
    header("⑤ 内容编号与文件名不符")
    if r["mismatches"]:
        grouped = {}
        for f, reason in r["mismatches"]:
            basename = os.path.basename(f)
            folder = os.path.dirname(f).split(os.sep)[-2:]
            key = " / ".join(folder)
            grouped.setdefault(key, []).append((basename, reason))
        for folder, items in grouped.items():
            print(f"  - {folder}/")
            for basename, reason in items:
                print(f"      · {basename} [{reason}]")
    else:
        print("  ✓ 一致")
    print("\n🎯 检查完成。\n" + "=" * 64)
# ================= GUI =================
def launch_gui():
    import tkinter as tk
    from tkinter import ttk, messagebox, filedialog
    class App(tk.Tk):
        def __init__(self):
            super().__init__()
            self.title("文件检查 - 可视化 v3.0")
            self.configure(bg="#F4F5F7")
            self.geometry("980x680")
            # 可访问性：大字体、柔和配色
            base_font = ("Segoe UI", 12)
            self.option_add("*Font", base_font)
            self.style = ttk.Style(self)
            # 软色主题
            self.style.configure("TLabel", background="#F4F5F7", foreground="#111827")
            self.style.configure("TFrame", background="#F4F5F7")
            self.style.configure("TButton", padding=8)
            self.style.configure("Treeview", rowheight=28, font=base_font)
            self.style.configure("Treeview.Heading", font=("Segoe UI", 12, "bold"))
            self._build_ui()
        def _build_ui(self):
            top = ttk.Frame(self)
            top.pack(fill=tk.X, padx=16, pady=10)
            # 行1：目录选择 + 月份
            ttk.Label(top, text="根目录:").grid(row=0, column=0, sticky=tk.W, padx=(0, 8))
            self.dir_var = tk.StringVar(value=BASE_DIR)
            self.dir_entry = ttk.Entry(top, textvariable=self.dir_var, width=70)
            self.dir_entry.grid(row=0, column=1, sticky=tk.W)
            ttk.Button(top, text="浏览…", command=self._choose_dir).grid(row=0, column=2, padx=8)
            ttk.Label(top, text="月份筛选:").grid(row=0, column=3, sticky=tk.E, padx=(16, 8))
            self.month_var = tk.StringVar(value=TARGET_MONTH)
            self.month_entry = ttk.Entry(top, textvariable=self.month_var, width=10)
            self.month_entry.grid(row=0, column=4, sticky=tk.W)
            self.run_btn = ttk.Button(top, text="开始检查", command=self._run_scan)
            self.run_btn.grid(row=0, column=5, padx=(16, 0))
            self.status_var = tk.StringVar(value="就绪")
            self.status = ttk.Label(self, textvariable=self.status_var)
            self.status.pack(anchor=tk.W, padx=16)
            # Notebook 区域
            self.nb = ttk.Notebook(self)
            self.nb.pack(fill=tk.BOTH, expand=True, padx=12, pady=8)
            self.tabs = {}
            for key, title, cols in [
                ("invalid", "命名错误", ("路径",)),
                ("unpaired", "缺配对", ("前缀", "年份", "编号")),
                ("duplicates", "重复编号", ("前缀", "年份", "编号", "类型", "路径")),
                ("gaps", "编号不连续", ("前缀", "年份", "缺少编号")),
                ("mismatch", "内容不符", ("文件名", "原因")),
            ]:
                frame = ttk.Frame(self.nb)
                self.nb.add(frame, text=title)
                tv = ttk.Treeview(frame, columns=cols, show="headings")
                for c in cols:
                    tv.heading(c, text=c)
                    tv.column(c, width=180, stretch=True)
                vsb = ttk.Scrollbar(frame, orient="vertical", command=tv.yview)
                tv.configure(yscrollcommand=vsb.set)
                tv.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
                vsb.pack(side=tk.RIGHT, fill=tk.Y)
                tv.bind("<Double-1>", lambda e, tree=tv: self._open_selected(tree))
                self.tabs[key] = tv
            # 放大/缩小按钮（便于近视/青光眼）
            zoom_bar = ttk.Frame(self)
            zoom_bar.pack(fill=tk.X, padx=12, pady=(0, 8))
            ttk.Label(zoom_bar, text="字体大小:").pack(side=tk.LEFT)
            for size in (12, 14, 16, 18):
                ttk.Button(zoom_bar, text=str(size), command=lambda s=size: self._set_font(s)).pack(
                    side=tk.LEFT, padx=4
                )
        def _set_font(self, size: int):
            base_font = ("Segoe UI", size)
            self.option_add("*Font", base_font)
            self.style.configure("Treeview", font=base_font, rowheight=int(size * 2.2))
            self.style.configure("Treeview.Heading", font=("Segoe UI", size, "bold"))
        def _choose_dir(self):
            path = filedialog.askdirectory(initialdir=self.dir_var.get() or os.getcwd())
            if path:
                self.dir_var.set(path)
        def _run_scan(self):
            base = self.dir_var.get().strip()
            month = self.month_var.get().strip()
            if not os.path.isdir(base):
                messagebox.showerror("错误", "请选择有效的根目录")
                return
            if fitz is None:
                messagebox.showwarning("提示", "未安装 PyMuPDF，无法检查 PDF 内容编号。可先执行: pip install pymupdf")
            self.run_btn.configure(state=tk.DISABLED)
            self.status_var.set("正在扫描，请稍候...")
            def worker():
                try:
                    result = scan_files(base, month)
                except Exception as e:
                    result = e
                self.after(0, lambda: self._fill_result(result))
            threading.Thread(target=worker, daemon=True).start()
        def _fill_result(self, result):
            self.run_btn.configure(state=tk.NORMAL)
            if isinstance(result, Exception):
                from tkinter import messagebox
                messagebox.showerror("错误", str(result))
                self.status_var.set("出错")
                return
            # 清空
            for tv in self.tabs.values():
                for i in tv.get_children():
                    tv.delete(i)
            # 1 命名错误
            tv = self.tabs["invalid"]
            for p in result["invalid_files"]:
                tv.insert("", "end", values=(shorten_path(p),), tags=(p,))
            # 2 缺配对
            tv = self.tabs["unpaired"]
            for prefix, year, num in result["unpaired"]:
                tv.insert("", "end", values=(prefix, year, f"{num:03d}"))
            # 3 重复
            tv = self.tabs["duplicates"]
            for item in result["duplicates"]:
                if len(item["xlsx"]) > 1:
                    for p in item["xlsx"]:
                        tv.insert(
                            "",
                            "end",
                            values=(item["prefix"], item["year"], f"{item['number']:03d}", "xlsx", shorten_path(p)),
                            tags=(p,),
                        )
                if len(item["pdf"]) > 1:
                    for p in item["pdf"]:
                        tv.insert(
                            "",
                            "end",
                            values=(item["prefix"], item["year"], f"{item['number']:03d}", "pdf", shorten_path(p)),
                            tags=(p,),
                        )
            # 4 缺号
            tv = self.tabs["gaps"]
            for prefix, year, missing in result["gaps"]:
                tv.insert("", "end", values=(prefix, year, ", ".join(f"{n:03d}" for n in missing)))
            # 5 内容不符
            tv = self.tabs["mismatch"]
            for p, reason in result["mismatches"]:
                tv.insert("", "end", values=(os.path.basename(p), reason), tags=(p,))
            self.status_var.set("完成")
        def _open_selected(self, tree):
            sel = tree.selection()
            if not sel:
                return
            item_id = sel[0]
            tags = tree.item(item_id, "tags") or [""]
            path = tags[0]
            if path and os.path.exists(path):
                try:
                    os.startfile(path)
                except Exception:
                    pass
    app = App()
    app.mainloop()
if __name__ == "__main__":
    # 默认启动 GUI；如需命令行输出，设置环境变量 CF_CLI=1
    if os.environ.get("CF_CLI", "0") == "1":
        _print_cli_report(BASE_DIR, TARGET_MONTH)
    else:
        launch_gui()