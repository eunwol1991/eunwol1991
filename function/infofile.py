import os
import re
import threading
import tkinter as tk
from tkinter import ttk, scrolledtext, font

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
    for root, dirs, _ in os.walk(BASE_DIR):
        for name in MONTH_MAP.values():
            path = os.path.join(root, name)
            if os.path.isdir(path):
                if not os.listdir(path):
                    try:
                        os.rmdir(path)
                        logger(f"删除空文件夹: {path}")
                    except Exception as e:
                        logger(f"删除失败: {path} -> {e}")
                else:
                    logger(f"未删除: {path} (非空)")

def matches_excel(name: str) -> bool:
    pattern1 = re.compile(r"\d{2}25 - .*\.xls.*", re.IGNORECASE)
    pattern2 = re.compile(r"\d{4} - .*\.xls.*", re.IGNORECASE)
    return bool(pattern1.match(name) or pattern2.match(name))

def create_month_folders(month: int, logger):
    month_name = MONTH_MAP.get(month)
    if not month_name:
        logger("Invalid month number")
        return
    for root, dirs, files in os.walk(BASE_DIR):
        dirs[:] = [d for d in dirs if 'history' not in d.lower()]
        for file in files:
            if matches_excel(file):
                file_dir = os.path.join(root)
                target_dir = file_dir
                if os.path.basename(file_dir).lower().find('format') != -1:
                    target_dir = os.path.dirname(file_dir)
                base = os.path.basename(target_dir).lower()
                if 'format' in base or 'history' in base:
                    continue
                month_path = os.path.join(target_dir, month_name)
                if not os.path.exists(month_path):
                    try:
                        os.makedirs(month_path, exist_ok=True)
                        logger(f"创建文件夹: {month_path}")
                    except Exception as e:
                        logger(f"创建失败: {month_path} -> {e}")
                else:
                    logger(f"跳过已存在: {month_path}")

class InfoApp:
    def __init__(self, master: tk.Tk):
        self.master = master
        self.option = tk.IntVar(value=1)
        self.month_var = tk.StringVar()
        self.status_var = tk.StringVar(value="Ready")

        self.title_font = font.Font(family="Segoe UI", size=16, weight="bold")
        self.label_font = font.Font(family="Segoe UI", size=10)
        self.log_font = font.Font(family="Segoe UI", size=9)

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

        ttk.Radiobutton(control, text="1. 子文件夹智能搜索",
                        variable=self.option, value=1,
                        font=self.label_font).grid(row=0, column=0, sticky="w", pady=2)
        ttk.Radiobutton(control, text="2. 自动清理空月份文件夹",
                        variable=self.option, value=2,
                        font=self.label_font).grid(row=1, column=0, sticky="w", pady=2)
        ttk.Radiobutton(control, text="3. 智能创建月份文件夹",
                        variable=self.option, value=3,
                        font=self.label_font).grid(row=2, column=0, sticky="w", pady=2)

        ttk.Label(control, text="月份:", font=self.label_font).grid(row=3, column=0, sticky="w", pady=(10,2))
        ttk.Entry(control, textvariable=self.month_var, width=10,
                  font=self.label_font).grid(row=4, column=0, sticky="w")

        ttk.Button(control, text="运行", command=self.start, width=15,
                   padding=5).grid(row=5, column=0, pady=(15, 0))

        self.log_widget = scrolledtext.ScrolledText(log_frame, state=tk.DISABLED,
                                                    font=self.log_font)
        self.log_widget.pack(fill=tk.BOTH, expand=True)

        self.status_label = ttk.Label(master, textvariable=self.status_var,
                                      font=self.label_font, anchor="w")
        self.status_label.grid(row=2, column=0, columnspan=2, sticky="ew", padx=10, pady=(0,5))

        master.bind("<Configure>", self.on_resize)

    def log(self, message: str):
        self.log_widget.after(0, self._append_log, message)

    def _append_log(self, message: str):
        self.log_widget.config(state=tk.NORMAL)
        self.log_widget.insert(tk.END, message + "\n")
        self.log_widget.see(tk.END)
        self.log_widget.config(state=tk.DISABLED)

    def start(self):
        self.status_var.set("正在处理...")
        self.log_widget.config(state=tk.NORMAL)
        self.log_widget.delete(1.0, tk.END)
        self.log_widget.config(state=tk.DISABLED)
        threading.Thread(target=self.execute).start()

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
        except Exception as e:
            self.log(f"发生异常: {e}")
            self.status_var.set("遇到异常")

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
