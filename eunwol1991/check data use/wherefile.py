import os
import time
import tkinter as tk
from tkinter import ttk, messagebox
from fuzzywuzzy import fuzz

# ==== 全局变量 ====
BASE_DIR = r"C:\Users\jhunj\Dropbox"  # 改成你的路径
search_cache = []
current_results = []
copy_mapping = {}
history_keywords = []
eyecare_mode = False
sensitive_keywords = ["confidential", "invoice", "password", "secret"]

# 用于存储搜索状态的堆栈，每个状态为 (results, keyword)
results_stack = []

# ==== 搜索函数（初次搜索）====


def search_files(keyword):
    global search_cache, current_results, results_stack
    matches = []
    for root, dirs, files in os.walk(BASE_DIR):
        for file in files:
            if fuzz.partial_ratio(keyword.lower(), file.lower()) >= 80:
                full_path = os.path.join(BASE_DIR, os.path.relpath(
                    os.path.join(root, file), BASE_DIR))
                try:
                    size = os.path.getsize(full_path)
                    mtime = os.path.getmtime(full_path)
                    matches.append((full_path, size, mtime))
                except:
                    continue
    search_cache = matches
    current_results = matches.copy()
    # 重置状态堆栈，每次初次搜索都新建状态
    results_stack.clear()
    results_stack.append((current_results.copy(), keyword))
    return matches

# ==== 二次筛选（过滤）====


def filter_results(keyword):
    global current_results, results_stack
    # 保存当前状态到堆栈，便于倒退
    results_stack.append((current_results.copy(), keyword))
    filtered = []
    for file in current_results:
        if fuzz.partial_ratio(keyword.lower(), os.path.basename(file[0]).lower()) >= 80:
            filtered.append(file)
    current_results = filtered
    update_results(filtered, keyword)

# ==== 排序函数 ====


def sort_results(results, key="mtime", reverse=True):
    key_index = {
        "mtime": 2,
        "size": 1,
        "name": lambda x: os.path.basename(x[0]).lower()
    }
    k = key_index.get(key, 2)
    return sorted(results, key=k if callable(k) else lambda x: x[k], reverse=reverse)

# ==== 更新显示结果 ====


def update_results(results, keyword):
    result_text.config(state=tk.NORMAL)
    result_text.delete(1.0, tk.END)
    copy_mapping.clear()

    if not results:
        result_text.insert(tk.END, f"❌ 没有找到包含「{keyword}」的文件。\n")
        result_text.config(state=tk.DISABLED)
        return

    result_text.insert(tk.END, f"✅ 找到 {len(results)} 个文件包含「{keyword}」：\n\n")

    for i, (full_path, size, mtime) in enumerate(results):
        filename = os.path.basename(full_path)
        folder = os.path.relpath(os.path.dirname(full_path), BASE_DIR)
        size_kb = f"{size/1024:.1f} KB"
        time_str = time.strftime("%Y-%m-%d %H:%M", time.localtime(mtime))
        tag = f"path_{i}"

        if any(s in filename.lower() for s in sensitive_keywords):
            result_text.insert(tk.END, "⚠️ 敏感词提醒：文件名中包含敏感关键词！\n", "warning")

        result_text.insert(tk.END, f"📄 文件名: ", "label")
        insert_highlight(filename, keyword)
        result_text.insert(tk.END, f"\n📁 路径: {folder}\n", tag)
        result_text.insert(
            tk.END, f"📏 大小: {size_kb}    🕒 修改: {time_str}\n\n", "info")

        result_text.tag_bind(tag, "<Button-1>", lambda e,
                             p=full_path: ask_open_file(p))
        result_text.tag_config(tag, foreground="#7FDBFF", underline=True)
        copy_mapping[tag] = full_path

    result_text.config(state=tk.DISABLED)

# ==== 高亮关键词 ====


def insert_highlight(text, keyword):
    idx = 0
    while idx < len(text):
        i = text.lower().find(keyword.lower(), idx)
        if i == -1:
            result_text.insert(tk.END, text[idx:])
            break
        result_text.insert(tk.END, text[idx:i])
        result_text.insert(tk.END, text[i:i+len(keyword)], "highlight")
        idx = i + len(keyword)

# ==== 打开文件或文件夹 ====


def ask_open_file(path):
    choice = messagebox.askyesnocancel("打开选项", f"你想打开这个文件还是它的文件夹？\n\n{path}")
    if choice is True:
        os.startfile(path)
    elif choice is False:
        os.startfile(os.path.dirname(path))

# ==== 搜索按钮 ====


def on_search():
    keyword = entry.get().strip()
    if not keyword:
        messagebox.showwarning("提示", "请输入关键词")
        return
    if keyword not in history_keywords:
        history_keywords.append(keyword)
        history_cb['values'] = history_keywords
    results = search_files(keyword)
    sorted_results = sort_results(
        results, key=sort_key_map.get(sort_cb.get(), "mtime"))
    update_results(sorted_results, keyword)

# ==== 过滤按钮 ====


def on_filter():
    keyword = entry.get().strip()
    if not keyword:
        messagebox.showwarning("提示", "请输入关键词")
        return
    filter_results(keyword)

# ==== 排序逻辑 ====


def on_sort_change(event=None):
    if not current_results:
        return
    keyword = entry.get().strip()
    sorted_results = sort_results(
        current_results, key=sort_key_map.get(sort_cb.get(), "mtime"))
    update_results(sorted_results, keyword)

# ==== 倒退按钮功能 ====


def go_back():
    global current_results, results_stack
    if len(results_stack) <= 1:
        messagebox.showinfo("提示", "已经是最初状态，无法倒退。")
        return
    # 弹出当前状态，回退到上一个状态
    results_stack.pop()
    prev_results, prev_keyword = results_stack[-1]
    current_results = prev_results.copy()
    # 更新显示
    update_results(current_results, prev_keyword)
    # 同时将历史关键词回填到输入框中
    entry.delete(0, tk.END)
    entry.insert(0, prev_keyword)

# ==== 切换护眼模式 ====


def toggle_eyecare_mode():
    global eyecare_mode
    eyecare_mode = not eyecare_mode

    if eyecare_mode:
        root.configure(bg="#2b3e2e")
        result_text.configure(bg="#2b3e2e", fg="#CFE5CF",
                              insertbackground="white", font=("Consolas", 14))
        result_text.tag_config(
            "highlight", background="#A8E6A1", foreground="black")
        result_text.tag_config("label", foreground="#A4FFB1",
                               font=("Consolas", 12, "bold"))
        result_text.tag_config(
            "warning", foreground="#FF8888", font=("Consolas", 12, "bold"))
        result_text.tag_config("info", foreground="#DDDDDD")
    else:
        root.configure(bg="#1e1e1e")
        result_text.configure(bg="#1e1e1e", fg="#FFFFFF",
                              insertbackground="white", font=("Consolas", 11))
        result_text.tag_config(
            "highlight", background="#FFD700", foreground="black")
        result_text.tag_config("label", foreground="#00FF99",
                               font=("Consolas", 11, "bold"))
        result_text.tag_config(
            "warning", foreground="#FF6F61", font=("Consolas", 11, "bold"))
        result_text.tag_config("info", foreground="#AAAAAA")


# ==== UI 构建 ====
root = tk.Tk()
root.title("📁 文件搜索器 Pro v2.3")
root.geometry("960x620")
root.configure(bg="#1e1e1e")

style = ttk.Style()
style.theme_use("default")
style.configure("TLabel", foreground="#00BFFF",
                background="#1e1e1e", font=("Arial", 14))
style.configure("TEntry", font=("Arial", 12))
style.configure("TButton", font=("Arial", 12))
style.configure("TCombobox", font=("Arial", 12))

ttk.Label(root, text="请输入关键词：").pack(pady=(20, 5))
entry = ttk.Entry(root, width=50)
entry.pack()
entry.bind("<Return>", lambda e: on_search())

btn_row = tk.Frame(root, bg="#1e1e1e")
btn_row.pack(pady=10)
ttk.Button(btn_row, text="🔍 初次搜索", command=on_search).pack(
    side=tk.LEFT, padx=10)
ttk.Button(btn_row, text="🔁 继续过滤", command=on_filter).pack(
    side=tk.LEFT, padx=10)
ttk.Button(btn_row, text="⏪ 倒退", command=go_back).pack(side=tk.LEFT, padx=10)
ttk.Button(btn_row, text="🧿 护眼模式", command=toggle_eyecare_mode).pack(
    side=tk.LEFT, padx=10)

ttk.Label(btn_row, text="排序：").pack(side=tk.LEFT)
sort_cb = ttk.Combobox(
    btn_row, values=["按时间", "按大小", "按文件名"], width=10, state="readonly")
sort_cb.set("按时间")
sort_key_map = {"按时间": "mtime", "按大小": "size", "按文件名": "name"}
sort_cb.bind("<<ComboboxSelected>>", on_sort_change)
sort_cb.pack(side=tk.LEFT, padx=5)

ttk.Label(btn_row, text="历史：").pack(side=tk.LEFT)
history_cb = ttk.Combobox(btn_row, values=[], width=20, state="readonly")
history_cb.pack(side=tk.LEFT)
history_cb.bind("<<ComboboxSelected>>", lambda e: entry.delete(
    0, tk.END) or entry.insert(0, history_cb.get()))

# ==== 文本结果显示区 ====
frame = tk.Frame(root, bg="#1e1e1e")
frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)
xscroll = tk.Scrollbar(frame, orient=tk.HORIZONTAL)
xscroll.pack(side=tk.BOTTOM, fill=tk.X)
yscroll = tk.Scrollbar(frame, orient=tk.VERTICAL)
yscroll.pack(side=tk.RIGHT, fill=tk.Y)
result_text = tk.Text(frame, wrap=tk.NONE, bg="#1e1e1e", fg="#FFFFFF", insertbackground="white",
                      font=("Consolas", 11), xscrollcommand=xscroll.set, yscrollcommand=yscroll.set)
result_text.pack(fill=tk.BOTH, expand=True)
result_text.config(state=tk.DISABLED)
xscroll.config(command=result_text.xview)
yscroll.config(command=result_text.yview)

result_text.tag_config("highlight", background="#FFD700", foreground="black")
result_text.tag_config("warning", foreground="#FF6F61",
                       font=("Consolas", 11, "bold"))
result_text.tag_config("label", foreground="#00FF99",
                       font=("Consolas", 11, "bold"))
result_text.tag_config("info", foreground="#AAAAAA")

root.mainloop()
