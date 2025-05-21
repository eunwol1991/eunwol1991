import os
import shutil
import tempfile
from pathlib import Path
import tkinter as tk
from tkinter import messagebox, scrolledtext
import threading

# 要清理的临时路径
TEMP_DIRS = [
    tempfile.gettempdir(),
    r"C:\Windows\Temp",
    str(Path.home() / "AppData" / "Local" / "Temp")
]

# 获取文件夹大小
def get_directory_size(path):
    total_size = 0
    for dirpath, dirnames, filenames in os.walk(path):
        for f in filenames:
            try:
                fp = os.path.join(dirpath, f)
                if os.path.exists(fp):
                    total_size += os.path.getsize(fp)
            except:
                pass
    return total_size

# 清理目录
def clean_directory(path):
    removed_files = 0
    freed_space = 0
    for root, dirs, files in os.walk(path):
        for name in files:
            try:
                file_path = os.path.join(root, name)
                size = os.path.getsize(file_path)
                os.remove(file_path)
                removed_files += 1
                freed_space += size
            except:
                pass
        for name in dirs:
            try:
                dir_path = os.path.join(root, name)
                shutil.rmtree(dir_path, ignore_errors=True)
            except:
                pass
    return removed_files, freed_space

# 执行清理任务（多线程避免卡界面）
def start_cleanup():
    log_text.delete('1.0', tk.END)
    def cleanup_task():
        for path in TEMP_DIRS:
            log_text.insert(tk.END, f"🔍 正在扫描：{path}\n")
            if not os.path.exists(path):
                log_text.insert(tk.END, f"⛔ 路径不存在，跳过。\n\n")
                continue

            size_before = get_directory_size(path)
            removed, freed = clean_directory(path)
            size_after = get_directory_size(path)

            log_text.insert(tk.END, f"📦 清理前大小：{round(size_before / (1024*1024), 2)} MB\n")
            log_text.insert(tk.END, f"🧾 删除文件数：{removed}\n")
            log_text.insert(tk.END, f"💨 释放空间：{round(freed / (1024*1024), 2)} MB\n")
            log_text.insert(tk.END, f"📉 清理后大小：{round(size_after / (1024*1024), 2)} MB\n\n")

        log_text.insert(tk.END, "🎉 清理完成！\n")
        messagebox.showinfo("完成", "清理完毕！")
    threading.Thread(target=cleanup_task).start()

# 创建GUI界面
window = tk.Tk()
window.title("C宝高端 Temp Cleaner 💻🧹")
window.geometry("620x500")

title_label = tk.Label(window, text="临时文件清理工具", font=("Helvetica", 18, "bold"))
title_label.pack(pady=10)

desc_label = tk.Label(window, text="点击下方按钮开始清理系统临时文件", font=("Helvetica", 12))
desc_label.pack()

start_button = tk.Button(window, text="🚀 开始清理", font=("Helvetica", 14), command=start_cleanup)
start_button.pack(pady=10)

log_text = scrolledtext.ScrolledText(window, width=75, height=20, font=("Courier", 10))
log_text.pack(padx=10, pady=10)

window.mainloop()
