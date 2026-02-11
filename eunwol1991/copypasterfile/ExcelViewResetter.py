import os
import json
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import xlwings as xw

# -------------------- 配置保存/读取 --------------------


def get_config_path():
    appdata = os.environ.get("APPDATA") or os.path.expanduser("~")
    cfg_dir = os.path.join(appdata, "ExcelViewResetter")
    os.makedirs(cfg_dir, exist_ok=True)
    return os.path.join(cfg_dir, "config.json")


def load_config():
    path = get_config_path()
    if os.path.isfile(path):
        try:
            with open(path, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return {}
    return {}


def save_config(cfg: dict):
    path = get_config_path()
    try:
        with open(path, "w", encoding="utf-8") as f:
            json.dump(cfg, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


def add_recent_path(cfg: dict, new_path: str, limit: int = 8):
    recent = cfg.get("recent_paths", [])
    new_path = os.path.normpath(new_path)
    if new_path in recent:
        recent.remove(new_path)
    recent.insert(0, new_path)
    cfg["recent_paths"] = recent[:limit]
    save_config(cfg)

# -------------------- Dropbox 根目录探测 --------------------


def detect_dropbox_roots():
    # 返回可能的 Dropbox 根目录列表（按优先顺序）。
    # 读取 %APPDATA%\Dropbox\info.json 或 ~/AppData/Local/Dropbox/info.json

    candidates = []
    # Roaming
    p1 = os.path.join(os.environ.get("APPDATA", ""), "Dropbox", "info.json")
    # Local
    p2 = os.path.join(os.environ.get("LOCALAPPDATA", ""),
                      "Dropbox", "info.json")
    for p in [p1, p2]:
        if os.path.isfile(p):
            try:
                with open(p, "r", encoding="utf-8") as f:
                    info = json.load(f)
                for k in ("personal", "business"):
                    if k in info and "path" in info[k]:
                        root = info[k]["path"]
                        if os.path.isdir(root):
                            candidates.append(root)
            except Exception:
                pass
    # 去重
    out = []
    for c in candidates:
        if c not in out:
            out.append(c)
    return out

# -------------------- 你的核心处理逻辑 --------------------


def process_excel_files(
    directory: str,
    filename_substring: str = "xx26",
    visible: bool = True,
    debug: bool = True,
):
    if not os.path.isdir(directory):
        print(f'路径不存在或不是文件夹：{directory}')
        return

    app = xw.App(visible=visible)
    app.display_alerts = False
    app.screen_updating = False

    processed_count = 0
    skipped_count = 0
    error_files = []

    targets = {
        'do':      ('A1:K11', 'K11'),
        'invoice': ('A1:I11', 'I11'),
    }

    try:
        for root, _, files in os.walk(directory):
            for file in files:
                low = file.lower()
                if (filename_substring.lower() in low) and low.endswith(('.xlsx', '.xlsm', '.xlsb', '.xls')):
                    file_path = os.path.join(root, file)
                    if debug:
                        print(f"🔄 处理：{file_path}")

                    try:
                        wb = app.books.open(file_path)
                        sheet_map = {s.name.lower(): s for s in wb.sheets}

                        has_do = 'do' in sheet_map
                        has_invoice = 'invoice' in sheet_map
                        if not has_do and not has_invoice:
                            skipped_count += 1
                            if debug:
                                print(f"⚠️ 跳过（无 DO/Invoice 工作表）：{file_path}")
                            wb.close()
                            continue

                        for key in ('do', 'invoice'):
                            if key in sheet_map:
                                sheet = sheet_map[key]
                                rng, focus = targets[key]
                                try:
                                    sheet.activate()
                                    # 强制视图回左上
                                    wb.app.api.ActiveWindow.ScrollColumn = 1
                                    wb.app.api.ActiveWindow.ScrollRow = 1
                                    # 选择区域并定位
                                    sheet.api.Application.Goto(
                                        sheet.range(rng).api, True)
                                    sheet.range(focus).select()
                                    # 再次回左上，抵消选择导致的横移
                                    wb.app.api.ActiveWindow.ScrollColumn = 1
                                    wb.app.api.ActiveWindow.ScrollRow = 1
                                    if debug:
                                        print(
                                            f"✅ {sheet.name}: 选中 {rng}，定位 {focus}，并已强制左上")
                                except Exception as e:
                                    print(f"⚠️ 处理工作表 '{sheet.name}' 出错：{e}")

                        # 按 DO→Invoice 激活一次，确保打开时可见
                        for name in ('do', 'invoice'):
                            if name in sheet_map:
                                sheet_map[name].activate()
                                wb.app.api.ActiveWindow.ScrollColumn = 1
                                wb.app.api.ActiveWindow.ScrollRow = 1
                                if debug:
                                    print(f"🪄 已激活：{name.upper()}")

                        wb.save()
                        wb.close()
                        processed_count += 1
                        if debug:
                            print(f"💾 已保存：{file_path}")

                    except Exception as e:
                        error_files.append(file_path)
                        if debug:
                            print(f"❌ 处理失败：{file_path}；错误：{e}")

    finally:
        app.quit()

    print(f"✅ 完成：修改 {processed_count} 个，跳过 {skipped_count} 个。")
    if error_files:
        print("⚠️ 失败文件：")
        for p in error_files:
            print(f"  ❌ {p}")

# -------------------- 极简 GUI --------------------


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("ExcelViewResetter")
        self.geometry("680x220")

        self.cfg = load_config()
        self.dir_var = tk.StringVar()

        # 尝试用最近路径/Dropbox 预填
        recent = self.cfg.get("recent_paths", [])
        if recent:
            self.dir_var.set(recent[0])
        else:
            roots = detect_dropbox_roots()
            if roots:
                self.dir_var.set(roots[0])

        frm = ttk.Frame(self, padding=12)
        frm.pack(fill=tk.BOTH, expand=True)

        ttk.Label(frm, text="目标文件夹：").grid(row=0, column=0, sticky="w")
        ttk.Entry(frm, textvariable=self.dir_var, width=70).grid(
            row=0, column=1, padx=6, sticky="w")
        ttk.Button(frm, text="选择…", command=self.choose_dir).grid(
            row=0, column=2, padx=2)

        # 最近使用下拉
        ttk.Label(frm, text="最近使用：").grid(
            row=1, column=0, sticky="w", pady=(8, 0))
        self.recent_cb = ttk.Combobox(
            frm, state="readonly", values=recent, width=68)
        self.recent_cb.grid(row=1, column=1, sticky="w", pady=(8, 0))
        ttk.Button(frm, text="载入", command=self.load_recent).grid(
            row=1, column=2, padx=2, pady=(8, 0))

        # 过滤子串
        self.sub_var = tk.StringVar(value="xx26")
        ttk.Label(frm, text="文件名需包含：").grid(
            row=2, column=0, sticky="w", pady=(8, 0))
        ttk.Entry(frm, textvariable=self.sub_var, width=20).grid(
            row=2, column=1, sticky="w", pady=(8, 0))
        ttk.Label(frm, text="（忽略大小写）").grid(
            row=2, column=2, sticky="w", pady=(8, 0))

        # 可见/日志
        self.visible_var = tk.BooleanVar(value=True)
        self.debug_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(frm, text="显示 Excel 窗口", variable=self.visible_var).grid(
            row=3, column=0, sticky="w", pady=(8, 0))
        ttk.Checkbutton(frm, text="打印调试日志", variable=self.debug_var).grid(
            row=3, column=1, sticky="w", pady=(8, 0))

        # 操作
        ttk.Button(frm, text="开始处理", command=self.run).grid(
            row=4, column=0, columnspan=3, pady=14, sticky="we")

        for i in range(3):
            frm.grid_columnconfigure(i, weight=1)

    def choose_dir(self):
        # 起始目录优先用当前输入、其次 Dropbox 根
        initdir = self.dir_var.get() if os.path.isdir(self.dir_var.get()) else None
        if not initdir:
            roots = detect_dropbox_roots()
            if roots:
                initdir = roots[0]
        path = filedialog.askdirectory(title="选择目标文件夹", initialdir=initdir)
        if path:
            self.dir_var.set(path)

    def load_recent(self):
        val = self.recent_cb.get()
        if val:
            self.dir_var.set(val)

    def run(self):
        path = self.dir_var.get().strip()
        if not path or not os.path.isdir(path):
            messagebox.showerror("错误", "请选择有效的文件夹路径")
            return
        add_recent_path(self.cfg, path)

        filename_substring = self.sub_var.get().strip() or "xx25"
        visible = self.visible_var.get()
        debug = self.debug_var.get()

        try:
            process_excel_files(
                directory=path,
                filename_substring=filename_substring,
                visible=visible,
                debug=debug
            )
            messagebox.showinfo("完成", "处理结束，详情见控制台日志。")
        except Exception as e:
            messagebox.showerror("异常", str(e))


if __name__ == "__main__":
    App().mainloop()
