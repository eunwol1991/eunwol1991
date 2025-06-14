import os
import threading
import tkinter as tk
from tkinter import ttk, scrolledtext, font, filedialog

BASE_DIR = r"C:\\Users\\User\\Dropbox\\DO & INV\\DO & INV 2025"

MONTH_MAP = {
    1: "1. Jan",
    2: "2. Feb",
    3: "3. Mar",
    4: "4. Apr",
    5: "5. May",
    6: "6. Jun",
    7: "7. Jul",
    8: "8. Aug",
    9: "9. Sep",
    10: "10. Oct",
    11: "11. Nov",
    12: "12. Dec",
}

def search_month_folder(month: int, logger):
    month_name = MONTH_MAP.get(month)
    if not month_name:
        logger("Invalid month number")
        return
    found = False
    for root, dirs, _ in os.walk(BASE_DIR):
        for d in dirs:
            if d.lower() == month_name.lower():
                logger(f"Found: {os.path.join(root, d)}")
                found = True
    if not found:
        logger("未找到")

def clean_empty_month_folders(logger):
    deleted = skipped = failed = 0
    for root, dirs, _ in os.walk(BASE_DIR):
        for name in MONTH_MAP.values():
            path = os.path.join(root, name)
            if os.path.isdir(path):
                if not os.listdir(path):
                    try:
                        os.rmdir(path)
                        logger(f"✅ 已删除: {path}", "success")
                        deleted += 1
                    except Exception as e:
                        logger(f"❌ 删除失败: {path} -> {e}", "fail")
                        failed += 1
                else:
                    logger(f"⚠️ 保留: {path} (非空)", "skip")
                    skipped += 1
    logger(f"汇总：本次删除 {deleted}，保留 {skipped}，失败 {failed}", "info")

def create_month_folders(month: int, logger):
    month_name = MONTH_MAP.get(month)
    if not month_name:
        logger("Invalid month number")
        return
    created = skipped = failed = 0
    month_dirs_lower = {m.lower() for m in MONTH_MAP.values()}
    for root, dirs, _ in os.walk(BASE_DIR):
        if root == BASE_DIR:
            # 跳过根目录
            continue

        # 仅当当前目录已经包含任意月份文件夹时才补齐
        has_month = any(d.lower() in month_dirs_lower for d in dirs)
        if not has_month:
            continue

        month_path = os.path.join(root, month_name)
        if os.path.exists(month_path):
            logger(f"⚠️ 已跳过: {month_path} (文件夹已存在)", "skip")
            skipped += 1
        else:
            try:
                os.makedirs(month_path)
                logger(f"✅ 已创建: {month_path}", "success")
                created += 1
            except Exception as e:
                logger(f"❌ 创建失败: {month_path} -> {e}", "fail")
                failed += 1
    logger(f"汇总：本次创建 {created}，跳过 {skipped}，失败 {failed}", "info")

class InfoApp:
    def __init__(self, master: tk.Tk):
        self.master = master
        self.option = tk.IntVar(value=1)
        self.month_var = tk.StringVar()
        self.status_var = tk.StringVar(value="Ready")

        self.title_font = font.Font(family="Segoe UI", size=16, weight="bold")
        self.label_font = font.Font(family="Segoe UI", size=10)
        self.log_font = font.Font(family="Segoe UI", size=9)

        self.style = ttk.Style()
        self.style.theme_use("clam")

        master.title("Info File Utility")
        master.geometry("800x500")
        master.minsize(600, 400)

        master.columnconfigure(1, weight=1)
        master.rowconfigure(1, weight=1)

        ttk.Label(master, text="Info File Utility", font=self.title_font).grid(
            row=0, column=0, columnspan=2, pady=(10, 5))

        control = ttk.Frame(master)
        control.grid(row=1, column=0, sticky="nw", padx=10, pady=10)

        log_frame = ttk.Frame(master)
        log_frame.grid(row=1, column=1, sticky="nsew", padx=10, pady=10)

        tk.Radiobutton(control, text="1. 子文件夹智能搜索",
                       variable=self.option, value=1,
                       font=self.label_font, anchor="w").grid(row=0, column=0, sticky="w", pady=2)
        tk.Radiobutton(control, text="2. 自动清理空月份文件夹",
                       variable=self.option, value=2,
                       font=self.label_font, anchor="w").grid(row=1, column=0, sticky="w", pady=2)
        tk.Radiobutton(control, text="3. 智能创建月份文件夹",
                       variable=self.option, value=3,
                       font=self.label_font, anchor="w").grid(row=2, column=0, sticky="w", pady=2)

        tk.Label(control, text="月份:", font=self.label_font).grid(row=3, column=0, sticky="w", pady=(10,2))
        tk.Entry(control, textvariable=self.month_var, width=10,
                 font=self.label_font).grid(row=4, column=0, sticky="w")

        tk.Button(control, text="运行", command=self.start, width=15,
                  font=self.label_font).grid(row=5, column=0, pady=(15, 0))

        tk.Button(control, text="导出日志", command=self.export_log, width=15,
                  font=self.label_font).grid(row=6, column=0, pady=(5, 0))

        self.log_widget = scrolledtext.ScrolledText(log_frame, state=tk.DISABLED,
                                                    font=self.log_font, background="#f8f8f8")
        self.log_widget.tag_config("success", foreground="green")
        self.log_widget.tag_config("fail", foreground="red")
        self.log_widget.tag_config("skip", foreground="orange")
        self.log_widget.tag_config("info", foreground="blue")
        self.log_widget.pack(fill=tk.BOTH, expand=True)

        self.status_label = ttk.Label(master, textvariable=self.status_var,
                                      font=self.label_font, anchor="w")
        self.status_label.grid(row=2, column=0, columnspan=2, sticky="ew", padx=10, pady=(0,5))

        self.progress = ttk.Progressbar(master, mode="indeterminate")
        self.progress.grid(row=3, column=0, columnspan=2, sticky="ew", padx=10, pady=(0,5))

        master.bind("<Configure>", self.on_resize)

    def log(self, message: str, tag: str = None):
        self.log_widget.after(0, self._append_log, message, tag)

    def _append_log(self, message: str, tag: str = None):
        self.log_widget.config(state=tk.NORMAL)
        if tag:
            self.log_widget.insert(tk.END, message + "\n", tag)
        else:
            self.log_widget.insert(tk.END, message + "\n")
        self.log_widget.see(tk.END)
        self.log_widget.config(state=tk.DISABLED)

    def start(self):
        self.status_var.set("正在处理...")
        self.progress.start()
        self.log_widget.config(state=tk.NORMAL)
        self.log_widget.delete(1.0, tk.END)
        self.log_widget.config(state=tk.DISABLED)
        threading.Thread(target=self.execute).start()

    def export_log(self):
        file_path = filedialog.asksaveasfilename(
            defaultextension=".txt",
            filetypes=[("Text Files", "*.txt")]
        )
        if file_path:
            with open(file_path, "w", encoding="utf-8") as f:
                f.write(self.log_widget.get("1.0", tk.END))
            self.status_var.set("日志已导出")

    def execute(self):
        try:
            month = int(self.month_var.get()) if self.month_var.get() else None
        except ValueError:
            month = None

        if self.option.get() in (1, 3) and month is None:
            self.log("请输入月份数字")
            self.status_var.set("等待输入")
            return

        try:
            if self.option.get() == 1:
                search_month_folder(month, self.log)
            elif self.option.get() == 2:
                clean_empty_month_folders(self.log)
            else:
                create_month_folders(month, self.log)
            self.status_var.set("完成")
            self.progress.stop()
        except Exception as e:
            self.log(f"发生异常: {e}", "fail")
            self.status_var.set("遇到异常")
            self.progress.stop()

    def on_resize(self, event):
        width = max(event.width, 600)
        base = max(9, int(width / 80))
        self.title_font.configure(size=base + 8)
        self.label_font.configure(size=base)
        self.log_font.configure(size=max(8, base - 1))


def main():
    root = tk.Tk()
    InfoApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
