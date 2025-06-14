import os
import re
import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox

BASE_DIR = r"C:\Users\User\Dropbox\DO & INV\DO & INV 2025"

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

def log(text_widget, message):
    text_widget.config(state=tk.NORMAL)
    text_widget.insert(tk.END, message + "\n")
    text_widget.see(tk.END)
    text_widget.config(state=tk.DISABLED)


def search_month_folder(month, log_widget):
    month_name = MONTH_MAP.get(month)
    if not month_name:
        log(log_widget, "Invalid month number")
        return
    found_paths = []
    for root, dirs, _ in os.walk(BASE_DIR):
        for d in dirs:
            if d.lower() == month_name.lower():
                found_paths.append(os.path.join(root, d))
    if found_paths:
        for p in found_paths:
            log(log_widget, f"Found: {p}")
    else:
        log(log_widget, "未找到")


def clean_empty_month_folders(log_widget):
    for root, dirs, _ in os.walk(BASE_DIR):
        for name in MONTH_MAP.values():
            path = os.path.join(root, name)
            if os.path.isdir(path):
                if not os.listdir(path):
                    try:
                        os.rmdir(path)
                        log(log_widget, f"删除空文件夹: {path}")
                    except Exception as e:
                        log(log_widget, f"删除失败: {path} -> {e}")
                else:
                    log(log_widget, f"未删除: {path} (非空)")


def matches_excel(name):
    pattern1 = re.compile(r"\d{2}25 - .*\.xls.*", re.IGNORECASE)
    pattern2 = re.compile(r"\d{4} - .*\.xls.*", re.IGNORECASE)
    return pattern1.match(name) or pattern2.match(name)


def create_month_folders(month, log_widget):
    month_name = MONTH_MAP.get(month)
    if not month_name:
        log(log_widget, "Invalid month number")
        return
    for root, dirs, files in os.walk(BASE_DIR):
        # Skip directories containing 'history'
        dirs[:] = [d for d in dirs if 'history' not in d.lower()]
        for file in files:
            if matches_excel(file):
                file_dir = os.path.join(root)
                target_dir = file_dir
                if os.path.basename(file_dir).lower().find('format') != -1:
                    target_dir = os.path.dirname(file_dir)
                if ('format' in os.path.basename(target_dir).lower() or
                        'history' in os.path.basename(target_dir).lower()):
                    continue
                month_path = os.path.join(target_dir, month_name)
                if not os.path.exists(month_path):
                    try:
                        os.makedirs(month_path, exist_ok=True)
                        log(log_widget, f"创建文件夹: {month_path}")
                    except Exception as e:
                        log(log_widget, f"创建失败: {month_path} -> {e}")
                else:
                    log(log_widget, f"跳过已存在: {month_path}")


def run_selected(option, month_var, log_widget):
    log_widget.config(state=tk.NORMAL)
    log_widget.delete(1.0, tk.END)
    log_widget.config(state=tk.DISABLED)

    try:
        month = int(month_var.get()) if month_var.get() else None
    except ValueError:
        month = None

    if option.get() == 1:
        if month is None:
            messagebox.showerror("Error", "请输入月份数字")
            return
        search_month_folder(month, log_widget)
    elif option.get() == 2:
        clean_empty_month_folders(log_widget)
    elif option.get() == 3:
        if month is None:
            messagebox.showerror("Error", "请输入月份数字")
            return
        create_month_folders(month, log_widget)


def main():
    window = tk.Tk()
    window.title("Info File Utility")
    window.geometry("600x400")

    option = tk.IntVar(value=1)
    month_var = tk.StringVar()

    frame = ttk.Frame(window)
    frame.pack(pady=10)

    ttk.Radiobutton(frame, text="1. 子文件夹智能搜索", variable=option, value=1).grid(row=0, column=0, sticky='w')
    ttk.Radiobutton(frame, text="2. 自动清理空月份文件夹", variable=option, value=2).grid(row=1, column=0, sticky='w')
    ttk.Radiobutton(frame, text="3. 智能创建月份文件夹", variable=option, value=3).grid(row=2, column=0, sticky='w')

    ttk.Label(frame, text="月份:").grid(row=3, column=0, sticky='e')
    ttk.Entry(frame, textvariable=month_var, width=10).grid(row=3, column=1, sticky='w')

    log_widget = scrolledtext.ScrolledText(window, width=70, height=15, state=tk.DISABLED)
    log_widget.pack(padx=10, pady=10)

    ttk.Button(window, text="运行", command=lambda: run_selected(option, month_var, log_widget)).pack()

    window.mainloop()


if __name__ == "__main__":
    main()
