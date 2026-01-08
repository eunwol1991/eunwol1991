import os
import re
import shutil
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from tkinter import filedialog, messagebox
from ttkbootstrap import Treeview
from tkinter import font as tkfont
from tkinter import DoubleVar, TclError


def main():
    file_info_list = []
    selected_files = []
    last_directory = [os.getcwd()]

    # ====== 基线窗口大小 ======
    BASE_WIDTH, BASE_HEIGHT = 920, 740

    # ====== 先创建 root，再做字体与样式 ======
    app = ttk.Window(themename="darkly")
    app.title("🗂️ 文件选择器")
    app.geometry(f"{BASE_WIDTH}x{BASE_HEIGHT}")
    app.resizable(True, True)

    style = ttk.Style()

    # ====== 顶栏 ======
    topbar = ttk.Frame(app)
    topbar.pack(fill=X, padx=10, pady=(10, 0))
    title_lbl = ttk.Label(
        topbar, text="📂 请选择源目录并匹配文件", font=("Arial", 18), bootstyle="info"
    )
    title_lbl.pack(side=LEFT, padx=(0, 10))

    # ====== 定义专用样式与命名字体 ======
    # 缩放显示标签：固定白字，不受主题覆盖
    style.configure("Zoom.TLabel", foreground="#FFFFFF")

    # 两把命名字体，专供两个 Treeview 使用（缩放只改这两把）
    base_font_size = 12
    tree_font_main = tkfont.Font(family="Arial", size=base_font_size)
    tree_font_chosen = tkfont.Font(family="Arial", size=base_font_size)

    # 两个 Treeview 的专用样式名（避免被主题默认样式干扰）
    style.configure("Main.Treeview", font=tree_font_main)
    style.configure("Chosen.Treeview", font=tree_font_chosen)

    # ====== 计算行高：用真实行距 + 上下间隙，确保不重叠 ======
    def rowheight_for(font_obj, factor: float) -> int:
        try:
            ls = int(font_obj.metrics("linespace"))
        except Exception:
            ls = int(14 * factor)  # 兜底
        gap = max(2, int(4 * factor))  # 上下留白
        return max(24, ls + gap * 2 + 2)

    # ====== UI 缩放 ======
    scale_var = DoubleVar(value=1.0)

    def apply_scale(value):
        try:
            factor = float(value)
        except Exception:
            factor = 1.0

        # DPI 缩放（可用则生效）
        try:
            app.tk.call("tk", "scaling", factor)
        except Exception:
            pass

        # 改两把命名字体的字号
        new_size = max(8, int(round(base_font_size * factor)))
        tree_font_main.configure(size=new_size)
        tree_font_chosen.configure(size=new_size)

        # 用同一把字体的行距来算 rowheight，分别设置
        rh_main = rowheight_for(tree_font_main, factor)
        rh_chosen = rowheight_for(tree_font_chosen, factor)
        style.configure("Main.Treeview", rowheight=rh_main)
        style.configure("Chosen.Treeview", rowheight=rh_chosen)

        # 表头 padding 也随比例调整
        style.configure("Treeview.Heading", padding=(max(4, int(6 * factor)),))

        # 窗口几何一起缩放
        app.geometry(f"{int(BASE_WIDTH*factor)}x{int(BASE_HEIGHT*factor)}")

        # 缩放显示文本（白色样式已设置）
        zoom_label.configure(text=f"UI 缩放：{int(round(factor*100))}%")

    # 缩放控件与显示
    zoom_label = ttk.Label(topbar, text="UI 缩放：100%", style="Zoom.TLabel")
    zoom_label.pack(side=RIGHT, padx=(10, 0))
    zoom_slider = ttk.Scale(
        topbar, from_=0.8, to=1.6, orient=HORIZONTAL,
        variable=scale_var, command=apply_scale, bootstyle=INFO, length=220
    )
    zoom_slider.pack(side=RIGHT)

    # ====== 文件列表（主 Treeview）======
    file_frame = ttk.Frame(app)
    file_frame.pack(fill=BOTH, expand=True, padx=10, pady=10)

    file_tree = Treeview(
        file_frame, columns=("#1"), show="headings",
        height=10, style="Main.Treeview", bootstyle="info"
    )
    file_tree.heading("#1", text="文件")
    file_tree.column("#1", anchor="w")
    file_tree.pack(side=LEFT, fill=BOTH, expand=True, padx=(0, 10))

    scrollbar = ttk.Scrollbar(
        file_frame, orient="vertical", command=file_tree.yview)
    scrollbar.pack(side=RIGHT, fill=Y)
    file_tree.config(yscrollcommand=scrollbar.set)

    # ====== 已选择列表（次 Treeview）======
    selected_frame = ttk.Frame(app)
    selected_frame.pack(fill=BOTH, expand=True, padx=10, pady=(0, 10))

    ttk.Label(selected_frame, text=" 已选择的文件顺序", font=(
        "Arial", 14), bootstyle="info").pack(anchor="w")
    selected_tree = Treeview(
        selected_frame, columns=("#1"), show="headings",
        height=5, style="Chosen.Treeview", bootstyle="warning"
    )
    selected_tree.heading("#1", text="已选择文件")
    selected_tree.column("#1", anchor="w")
    selected_tree.pack(fill=BOTH, expand=True)

    # ====== 发票起始编号输入 ======
    entry_frame = ttk.Frame(app)
    entry_frame.pack(pady=10)
    ttk.Label(entry_frame, text="请输入发票起始编号 (例如: 0126 - 001)",
              font=("Arial", 12)).pack()
    invoice_entry = ttk.Entry(entry_frame, font=("Arial", 12), width=30)
    invoice_entry.pack(pady=5)

    # ====== 文件名匹配模式 ======
    file_pattern = re.compile(
        r"""^(?P<prefix>[A-Z0-9._ \-]+?)          # 开头前缀，如 P17 / C.P / ABC-123
            \s+xx26\s*[-–—]\s*00x                # xx25 - 00x（连字符可为-、–、—）
            (?:\s*[-–—]\s*DO\s*&\s*INV)?         # 可选的 - DO & INV
            (?:\s*\((?P<name>.+)\))?             # 可选的 (名称，允许内嵌括号)
        """,
        re.IGNORECASE | re.VERBOSE
    )

    # ====== 读取上次目录 ======
    try:
        with open("last_dir.txt", "r") as f:
            last_directory[0] = f.read().strip()
    except:
        pass

    # ====== 业务逻辑 ======
    def browse_source():
        source_dir = filedialog.askdirectory(initialdir=last_directory[0])
        if not source_dir:
            return
        last_directory[0] = os.path.dirname(source_dir)
        with open("last_dir.txt", "w") as f:
            f.write(source_dir)

        file_info_list.clear()
        file_tree.delete(*file_tree.get_children())

        idx = 0
        for root_dir, dirs, files in os.walk(source_dir):
            if "history" in root_dir.lower():
                continue
            for file in files:
                full_path = os.path.join(root_dir, file)
                if not os.path.isfile(full_path):
                    continue
                filename_no_ext = os.path.splitext(file)[0]
                m = file_pattern.match(filename_no_ext)
                if not m:
                    continue

                name = (m.group("name") or m.group("prefix")).strip()
                display = f"{len(file_info_list) + 1}. {name}"
                file_info_list.append(
                    {'display_name': display, 'file_path': full_path})

                # 斑马纹行：odd/even
                tag = ("oddrow",) if (idx % 2 == 0) else ("evenrow",)
                file_tree.insert("", END, values=(display,), tags=tag)
                idx += 1

        # 配置斑马纹颜色（与 darkly 主题相容）
        file_tree.tag_configure("oddrow", background="#2B2B2B")
        file_tree.tag_configure("evenrow", background="#242424")

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
        selected_files.append(
            {'index': index, 'file_info': file_info_list[index]})
        update_selected_listbox()

    def update_selected_listbox():
        selected_tree.delete(*selected_tree.get_children())
        for i, item in enumerate(selected_files):
            tag = ("oddrow",) if (i % 2 == 0) else ("evenrow",)
            selected_tree.insert("", END, values=(
                f"{i + 1}. {item['file_info']['display_name']}",), tags=tag)
        # 斑马纹
        selected_tree.tag_configure("oddrow", background="#2B2B2B")
        selected_tree.tag_configure("evenrow", background="#242424")

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
        target_dir = r"C:\Users\jhunj\Dropbox\for jj\Doc to print - JJ"
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

            new_filename = re.sub(
                r"xx26\s*-\s*00x",
                f"{invoice_prefix} - {invoice_number:03d}",
                filename,
                flags=re.IGNORECASE
            )
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

    # 事件绑定与按钮区
    file_tree.bind("<Double-1>", lambda e: add_to_selected())

    btn_frame = ttk.Frame(app)
    btn_frame.pack(pady=10)
    ttk.Button(btn_frame, text="📁 选择源目录", command=browse_source,
               bootstyle=INFO).grid(row=0, column=0, padx=10)
    ttk.Button(btn_frame, text="➕ 添加文件", command=add_to_selected,
               bootstyle=SUCCESS).grid(row=0, column=1, padx=10)
    ttk.Button(btn_frame, text="🗑️ 清空选择", command=clear_selected,
               bootstyle=WARNING).grid(row=0, column=2, padx=10)
    ttk.Button(btn_frame, text="❌ 删除项目", command=delete_selected,
               bootstyle=DANGER).grid(row=0, column=3, padx=10)
    ttk.Button(btn_frame, text="📤 执行复制", command=copy_files,
               bootstyle=PRIMARY).grid(row=0, column=4, padx=10)
    ttk.Button(btn_frame, text="🚪 退出程序", command=app.quit,
               bootstyle=SECONDARY).grid(row=0, column=5, padx=10)

    # ====== 初始化一次缩放（让行高立即生效）======
    apply_scale(scale_var.get())

    app.mainloop()


if __name__ == "__main__":
    main()
