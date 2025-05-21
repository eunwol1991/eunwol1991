import os
import re
import shutil
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from tkinter import filedialog, messagebox
from ttkbootstrap import Treeview

def main():
    file_info_list = []
    selected_files = []
    last_directory = [os.getcwd()]

    file_pattern_1 = re.compile(r"^[A-Z]+\s+xx25\s*-\s*00x\s*-\s*DO & INV\s+\((.*?)\)", re.IGNORECASE)
    file_pattern_2 = re.compile(r"^C\.P\s+xx25\s*-\s*00x\s*-\s*DO & INV\s+\((.*?)\)", re.IGNORECASE)

    try:
        with open("last_dir.txt", "r") as f:
            last_directory[0] = f.read().strip()
    except:
        pass

    app = ttk.Window(themename="darkly")
    app.title("🗂️ 文件选择器")
    app.geometry("920x740")
    app.resizable(True, True)

    ttk.Label(app, text="📂 请选择源目录并匹配文件", font=("Arial", 18), bootstyle="info").pack(pady=15)

    file_frame = ttk.Frame(app)
    file_frame.pack(fill=BOTH, expand=True, padx=10, pady=10)

    file_tree = Treeview(file_frame, columns=("#1"), show="headings", height=10, bootstyle="info")
    file_tree.heading("#1", text="文件")
    file_tree.pack(side=LEFT, fill=BOTH, expand=True, padx=(0, 10))
    file_tree.column("#1", anchor="w")

    scrollbar = ttk.Scrollbar(file_frame, orient="vertical", command=file_tree.yview)
    scrollbar.pack(side=RIGHT, fill=Y)
    file_tree.config(yscrollcommand=scrollbar.set)

    file_tree.bind("<Double-1>", lambda e: add_to_selected())

    selected_frame = ttk.Frame(app)
    selected_frame.pack(fill=BOTH, expand=True, padx=10, pady=(0, 10))

    ttk.Label(selected_frame, text=" 已选择的文件顺序", font=("Arial", 14), bootstyle="info").pack(anchor="w")
    selected_tree = Treeview(selected_frame, columns=("#1"), show="headings", height=5, bootstyle="warning")
    selected_tree.heading("#1", text="已选择文件")
    selected_tree.pack(fill=BOTH, expand=True)
    selected_tree.column("#1", anchor="w")

    entry_frame = ttk.Frame(app)
    entry_frame.pack(pady=10)
    ttk.Label(entry_frame, text="请输入发票起始编号 (例如: 0525 - 001)", font=("Arial", 12)).pack()
    invoice_entry = ttk.Entry(entry_frame, font=("Arial", 12), width=30)
    invoice_entry.pack(pady=5)

    def browse_source():
        source_dir = filedialog.askdirectory(initialdir=last_directory[0])
        if not source_dir:
            return
        last_directory[0] = os.path.dirname(source_dir)
        with open("last_dir.txt", "w") as f:
            f.write(source_dir)

        file_info_list.clear()
        file_tree.delete(*file_tree.get_children())

        for root_dir, dirs, files in os.walk(source_dir):
            for file in files:
                full_path = os.path.join(root_dir, file)
                if os.path.isfile(full_path):
                    filename_no_ext = os.path.splitext(file)[0]
                    match_1 = file_pattern_1.match(filename_no_ext)
                    match_2 = file_pattern_2.match(filename_no_ext)
                    if match_1 or match_2:
                        name = match_1.group(1).strip() if match_1 else match_2.group(1).strip()
                        display = f"{len(file_info_list) + 1}. {name}"
                        file_info_list.append({'display_name': display, 'file_path': full_path})
                        file_tree.insert("", END, values=(display,))

        if not file_info_list:
            messagebox.showinfo("提示", "未找到符合条件的文件。")

    def add_to_selected():
        selected_item = file_tree.selection()
        if not selected_item:
            messagebox.showwarning("⚠️ 警告", "请先选择一个文件。")
            return
        index = file_tree.index(selected_item)
        if index in [i['index'] for i in selected_files]:
            return
        selected_files.append({'index': index, 'file_info': file_info_list[index]})
        update_selected_listbox()

    def update_selected_listbox():
        selected_tree.delete(*selected_tree.get_children())
        for i, item in enumerate(selected_files):
            selected_tree.insert("", END, values=(f"{i + 1}. {item['file_info']['display_name']}",))

    def delete_selected():
        selected_item = selected_tree.selection()
        if not selected_item:
            return
        index = selected_tree.index(selected_item)
        del selected_files[index]
        update_selected_listbox()

    def clear_selected():
        selected_files.clear()
        update_selected_listbox()

    def copy_files():
        target_dir = r"C:\Users\User\Dropbox\for jj\Doc to print - JJ"
        if not os.path.exists(target_dir):
            messagebox.showerror("❌ 错误", "目标目录不存在。")
            return
        if not selected_files:
            messagebox.showwarning("⚠️ 警告", "请先选择要复制的文件。")
            return

        invoice_start = invoice_entry.get().strip()
        if not re.match(r"\d{4}\s*-\s*\d{3}", invoice_start):
            messagebox.showwarning("⚠️ 警告", "请输入有效的发票起始编号，例如：0325 - 001")
            return

        invoice_prefix, invoice_number = invoice_start.split("-")
        invoice_prefix = invoice_prefix.strip()
        invoice_number = int(invoice_number.strip())

        for item in selected_files:
            src_path = item['file_info']['file_path']
            filename = os.path.basename(src_path)
            new_filename = re.sub(r"xx25\s*-\s*00x", f"{invoice_prefix} - {invoice_number:03d}", filename)
            invoice_number += 1
            dst_path = os.path.join(target_dir, new_filename)

            if os.path.exists(dst_path):
                base, ext = os.path.splitext(new_filename)
                count = 1
                while os.path.exists(dst_path):
                    new_filename = f"{base}_{count}{ext}"
                    dst_path = os.path.join(target_dir, new_filename)
                    count += 1

            shutil.copy2(src_path, dst_path)

        messagebox.showinfo("✅ 成功", "文件已复制并重命名。")

    btn_frame = ttk.Frame(app)
    btn_frame.pack(pady=10)
    ttk.Button(btn_frame, text="📁 选择源目录", command=browse_source, bootstyle=INFO).grid(row=0, column=0, padx=10)
    ttk.Button(btn_frame, text="➕ 添加文件", command=add_to_selected, bootstyle=SUCCESS).grid(row=0, column=1, padx=10)
    ttk.Button(btn_frame, text="🗑️ 清空选择", command=clear_selected, bootstyle=WARNING).grid(row=0, column=2, padx=10)
    ttk.Button(btn_frame, text="❌ 删除项目", command=delete_selected, bootstyle=DANGER).grid(row=0, column=3, padx=10)
    ttk.Button(btn_frame, text="📤 执行复制", command=copy_files, bootstyle=PRIMARY).grid(row=0, column=4, padx=10)
    ttk.Button(btn_frame, text="🚪 退出程序", command=app.quit, bootstyle=SECONDARY).grid(row=0, column=5, padx=10)

    app.mainloop()

if __name__ == "__main__":
    main()
