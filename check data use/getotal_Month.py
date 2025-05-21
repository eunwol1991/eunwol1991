import fitz  # PyMuPDF
import re
import pandas as pd
import os
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime
import logging
import queue

class InvoiceExtractorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Invoice Information Extractor")
        self.root.geometry("800x800")
        self.root.resizable(False, False)

        # 配置日志记录
        logging.basicConfig(filename='logs.txt', level=logging.INFO,
                            format='%(asctime)s - %(levelname)s - %(message)s')

        # 初始化变量
        self.original_df = pd.DataFrame()
        self.results_df = pd.DataFrame()
        self.filters = {}
        self.Month = ''
        self.sort_column = None
        self.sort_descending = False
        self.progress_var = tk.DoubleVar()
        self.gui_queue = queue.Queue()
        self.data_displayed = False  # 标记数据是否已显示

        # 构建 GUI 界面
        self.build_gui()

        # 开始处理 GUI 队列
        self.root.after(100, self.process_gui_queue)

    def build_gui(self):
        """构建 GUI 界面组件"""
        style = ttk.Style()
        style.configure("TLabel", font=("Helvetica", 12))
        style.configure("TButton", font=("Helvetica", 12))
        style.configure("TEntry", font=("Helvetica", 12))

        self.folder_path = tk.StringVar()
        self.search_in_month_var = tk.BooleanVar()
        self.month_combobox = ttk.Combobox()
        self.year_combobox = ttk.Combobox()

        frame = ttk.Frame(self.root, padding="20")
        frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        ttk.Label(frame, text="选择文件夹:").grid(row=0, column=0, padx=10, pady=5, sticky=tk.W)
        ttk.Entry(frame, textvariable=self.folder_path, width=50).grid(row=0, column=1, padx=10, pady=5, sticky=tk.W)
        ttk.Button(frame, text="浏览", command=self.browse_folder).grid(row=0, column=2, padx=10, pady=5)

        check_button = ttk.Checkbutton(frame, text="仅搜索选定月份的子文件夹", variable=self.search_in_month_var)
        check_button.grid(row=1, column=0, columnspan=3, padx=10, pady=10)

        ttk.Label(frame, text="选择月份:").grid(row=2, column=0, padx=10, pady=5, sticky=tk.W)
        self.month_combobox = ttk.Combobox(frame, values=[str(i).zfill(2) for i in range(1, 13)], width=5)
        self.month_combobox.grid(row=2, column=1, padx=10, pady=5, sticky=tk.W)

        ttk.Label(frame, text="选择年份:").grid(row=3, column=0, padx=10, pady=5, sticky=tk.W)
        self.year_combobox = ttk.Combobox(frame, values=[str(i) for i in range(2000, datetime.now().year + 1)], width=10)
        self.year_combobox.grid(row=3, column=1, padx=10, pady=5, sticky=tk.W)

        ttk.Button(frame, text="搜索发票", command=self.search_and_display_results).grid(row=4, column=0, columnspan=3, padx=10, pady=20)
        ttk.Button(self.root, text="检查缺号", command=self.check_missing_invoice_numbers).grid(row=11, column=0, padx=20, pady=10)

        self.progress_bar = ttk.Progressbar(self.root, orient="horizontal", length=400, mode="determinate", variable=self.progress_var)
        self.progress_bar.grid(row=5, column=0, padx=20, pady=10)

        tree_frame = ttk.Frame(self.root)
        tree_frame.grid(row=6, column=0, padx=20, pady=20, sticky=(tk.W, tk.E, tk.N, tk.S))
        # 增加了“File”列，方便显示出错的文件名
        self.tree = ttk.Treeview(tree_frame, columns=("Invoice Date", "Invoice No", "Total", "Account", "File"), show='headings')
        self.tree.heading("Invoice Date", text="Invoice Date", command=lambda: self.sort_treeview_column("Invoice Date"))
        self.tree.heading("Invoice No", text="Invoice No", command=lambda: self.sort_treeview_column("Invoice No"))
        self.tree.heading("Total", text="Total", command=lambda: self.sort_treeview_column("Total"))
        self.tree.heading("Account", text="Account", command=lambda: self.sort_treeview_column("Account"))
        self.tree.heading("File", text="File")
        self.tree.column("Invoice Date", width=100)
        self.tree.column("Invoice No", width=150)
        self.tree.column("Total", width=100)
        self.tree.column("Account", width=100)
        self.tree.column("File", width=200)
        # ✅ 先定义 scrollbar，再 grid
        scrollbar = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)

        # ✅ 正确顺序：先定义再布局
        self.tree.grid(row=0, column=0, sticky='nsew')
        scrollbar.grid(row=0, column=1, sticky='ns')

        # ✅ 让 TreeView 自动撑满 Frame
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)



        self.total_label = ttk.Label(self.root, text="Total: 0.00", font=("Helvetica", 12))
        self.total_label.grid(row=7, column=0, padx=20, pady=10, sticky=tk.W)

        ttk.Label(frame, text="按账户过滤:").grid(row=8, column=0, padx=10, pady=5, sticky=tk.W)
        self.account_combobox = ttk.Combobox(frame, values=["All", "Anthony", "Joshua", "Melvin"], width=10)
        self.account_combobox.grid(row=8, column=1, padx=10, pady=5, sticky=tk.W)
        self.account_combobox.current(0)
        self.account_combobox.bind("<<ComboboxSelected>>", self.filter_by_combobox_selection)

        ttk.Label(frame, text="关键字:").grid(row=9, column=0, padx=10, pady=5, sticky=tk.W)
        self.keyword_entry = ttk.Entry(frame, width=50)
        self.keyword_entry.grid(row=9, column=1, padx=10, pady=5, sticky=tk.W)
        ttk.Button(frame, text="按关键字过滤", command=self.keyword_filter).grid(row=9, column=2, padx=10, pady=5)

        ttk.Button(self.root, text="导出到 Excel", command=self.export_to_excel).grid(row=10, column=0, padx=20, pady=10)

        self.tree.tag_configure("error", background="pink")  # 可改为 foreground="red"
       # self.tree.bind("<ButtonRelease-1>", self.on_heading_click)
        self.tree.bind("<Double-1>", self.open_selected_pdf)
        self.setup_context_menu()


    def browse_folder(self):
        """浏览文件夹"""
        folder_selected = filedialog.askdirectory()
        self.folder_path.set(folder_selected)

    def search_and_display_results(self):
        """开始搜索并显示结果"""
        folder = self.folder_path.get()

        if not folder or not os.path.exists(folder):
            messagebox.showwarning("警告", "请选择有效的文件夹。")
            return

        selected_month = self.month_combobox.get()
        selected_year = self.year_combobox.get()[-2:]
        if not selected_month or not selected_year:
            messagebox.showwarning("警告", "请选择月份和年份。")
            return

        self.Month = f"{selected_month}{selected_year}"

        # 重置变量
        self.original_df = pd.DataFrame()
        self.results_df = pd.DataFrame()
        self.filters = {}
        self.sort_column = None
        self.sort_descending = False
        self.data_displayed = False

        self.progress_var.set(0)
        self.gui_queue = queue.Queue()

        # 在单独的线程中开始搜索发票，防止阻塞 GUI
        self.executor = ThreadPoolExecutor(max_workers=4)
        self.executor.submit(self.search_invoices, folder)

    def search_invoices(self, folder):
        """搜索发票"""
        extracted_data = []
        pdf_files = []

        month_map = {
            '01': 'jan', '02': 'feb', '03': 'mar', '04': 'apr',
            '05': 'may', '06': 'jun', '07': 'jul', '08': 'aug',
            '09': 'sep', '10': 'oct', '11': 'nov', '12': 'dec'
        }

        selected_month_name = month_map.get(self.month_combobox.get(), '').lower()

        for root, dirs, files in os.walk(folder):
            if self.search_in_month_var.get() and selected_month_name in root.lower():
                for file in files:
                    if file.lower().endswith('.pdf') and self.Month.lower() in file.lower() and 'inv' in file.lower() and 'invoice' not in file.lower():
                        pdf_path = os.path.join(root, file)
                        pdf_files.append(pdf_path)
                        logging.info(f"在 '{selected_month_name}' 子文件夹中找到 PDF: {pdf_path}")
            elif not self.search_in_month_var.get():
                for file in files:
                    if file.lower().endswith('.pdf') and self.Month.lower() in file.lower() and 'inv' in file.lower() and 'invoice' not in file.lower():
                        pdf_path = os.path.join(root, file)
                        pdf_files.append(pdf_path)
                        logging.info(f"在任意文件夹中找到 PDF: {pdf_path}")

        pdf_files = list(set(pdf_files))
        total_files = len(pdf_files)

        if total_files == 0:
            self.gui_queue.put(('no_invoices', None))
            return

        batch_size = 30
        for start_idx in range(0, total_files, batch_size):
            batch_files = pdf_files[start_idx:start_idx + batch_size]
            batch_data = self.process_batch(batch_files, total_files, start_idx)
            extracted_data.extend(batch_data)

        self.original_df = pd.DataFrame(extracted_data)
        self.results_df = self.original_df.copy()
        if self.original_df.empty:
            self.gui_queue.put(('no_invoices', None))
            logging.info("未找到任何发票")
        else:
            # ✅ 在这里设置默认排序方向（放这里刚刚好）
            self.sort_column = "Invoice No"
            self.sort_descending = False
            self.results_df = self.original_df.sort_values(by=self.sort_column, ascending=True)
            self.gui_queue.put(('display_results', self.results_df))


    def process_batch(self, pdf_files, total_files, start_idx):
        """处理一批 PDF 文件"""
        extracted_data = []
        with ThreadPoolExecutor() as executor:
            results = executor.map(self.extract_invoice_info_v2_3, pdf_files)
            for i, result in enumerate(results):
                if result is not None:
                    extracted_data.append(result)
                progress = (start_idx + i + 1) / total_files * 100
                self.gui_queue.put(('update_progress', progress))
                logging.info(f"已处理 {start_idx + i + 1}/{total_files} 个文件")
        return extracted_data
    

    def extract_invoice_info_v2_3(self, pdf_path):
        """从 PDF 中提取发票信息"""
        try:
            doc = fitz.open(pdf_path)
            text = ""

            for page_num in range(len(doc)):
                page = doc.load_page(page_num)
                text += page.get_text()

            # 正则表达式模式
            invoice_no_patterns = [
                re.compile(r"Invoice No\s*([A-Z0-9]{2,}\s*\d{4}\s*-\s*\d{3})"),
                re.compile(r"Invoice No\s*([A-Z0-9.]{2,}\s*\d{4}\s*-\s*\d{3})"),
            ]
            date_patterns = [
                re.compile(r"Invoice Date\s*(\d{2}/\d{2}/\d{4})"),
                re.compile(r"Date\s*(\d{2}/\d{2}/\d{4})")
            ]
            total_patterns = [
                re.compile(r"Total\s*\$([\d,\.]+)"),
                re.compile(r"Total\s*([\d,\.]+)\s*\$"),
                re.compile(r"(\d{1,3}(?:,\d{3})*\.\d{2})\s*\$"),
                re.compile(r"(\d{1,3}(?:,\d{3})*\.\d{2})")
            ]

            # 从文件路径中提取账户信息
            account = self.extract_account_from_path(pdf_path)

            invoice_no = "Not found"
            date = "Not found"
            total = "Not found"
            if "abr" in pdf_path.lower():
                total = self.extract_total_from_abr(doc, pdf_path)
            else:
                # 原本的万能通用逻辑
                for pattern in total_patterns:
                    matches = pattern.findall(text)
                    if matches:
                        total = max(matches, key=lambda x: float(x.replace(",", "").replace("$", "")))
                        break



            for pattern in invoice_no_patterns:
                match = pattern.search(text)
                if match:
                    invoice_no = match.group(1).strip()
                    break

            for pattern in date_patterns:
                match = pattern.search(text)
                if match:
                    date = match.group(1).strip()
                    break

            # 如果有任一字段未找到，记录日志提醒你哪一个文件出了问题
            if invoice_no == "Not found" or date == "Not found" or total == "Not found":
                logging.info(f"【错误】文件 {pdf_path} 缺少信息：Invoice No: {invoice_no}, Invoice Date: {date}, Total: {total}")

            extracted_info = {
                "Invoice Date": date,
                "Invoice No": invoice_no,
                "Total": total,
                "Account": account,
                "File": os.path.basename(pdf_path)  # 新增：显示文件名
            }
            return extracted_info

        except Exception as e:
            logging.error(f"处理 {pdf_path} 时出错: {e}")
            return None
        
    def check_missing_invoice_numbers(self):
        grouped = {}
        for no in self.original_df["Invoice No"]:
            if isinstance(no, str):
                match = re.match(r"([A-Z]+\s*\d+)\s*-\s*(\d+)", no)
                if match:
                    prefix = match.group(1).replace(" ", "")
                    number = int(match.group(2))
                    grouped.setdefault(prefix, []).append(number)

        result_text = ""
        for prefix, numbers in grouped.items():
            numbers = sorted(set(numbers))
            full_range = list(range(numbers[0], numbers[-1] + 1))
            missing = sorted(set(full_range) - set(numbers))
            if missing:
                missing_str = ", ".join(str(n).zfill(3) for n in missing)
                result_text += f"[ {prefix} ]\n缺漏: {missing_str}\n\n"
            else:
                result_text += f"[ {prefix} ]\n缺漏: 无\n\n"

        self.show_missing_numbers_window(result_text.strip())

    def show_missing_numbers_window(self, text):
        window = tk.Toplevel(self.root)
        window.title("发票缺号检查结果")
        window.geometry("500x400")

        frame = ttk.Frame(window, padding=10)
        frame.pack(fill=tk.BOTH, expand=True)

        text_widget = tk.Text(frame, wrap=tk.WORD, font=("Consolas", 11), borderwidth=2, relief="sunken")
        lines = text.split("\n")
        for line in lines:
            if "缺漏:" in line and "无" not in line:
                text_widget.insert(tk.END, line + "\n", "missing")
            else:
                text_widget.insert(tk.END, line + "\n")

        text_widget.tag_configure("missing", background="pink", font=("Consolas", 11, "bold"))
        text_widget.config(state=tk.DISABLED)
        text_widget.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        scrollbar = ttk.Scrollbar(frame, orient=tk.VERTICAL, command=text_widget.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        text_widget.config(yscrollcommand=scrollbar.set)

        def copy_text():
            self.root.clipboard_clear()
            self.root.clipboard_append(text)
            messagebox.showinfo("复制成功", "缺号清单已复制到剪贴板！")

        copy_btn = ttk.Button(window, text="复制内容", command=copy_text)
        copy_btn.pack(pady=10)


    def extract_total_from_abr(self, doc, pdf_path):
        try:
            blocks = []
            for page in doc:
                for block in page.get_text("dict")["blocks"]:
                    for line in block.get("lines", []):
                        for span in line.get("spans", []):
                            blocks.append({
                                "text": span["text"],
                                "x": span["bbox"][0],
                                "y": span["bbox"][1]
                            })

            sorted_blocks = sorted(blocks, key=lambda b: (-b["y"], -b["x"]))

            for b in sorted_blocks:
                match = re.search(r"(\d{1,3}(?:,\d{3})*\.\d{2})\s*\$?", b["text"])
                if match:
                    value = match.group(1)
                    logging.info(f"[ABR右下角暴力解法] 文件 {pdf_path} 抓到 Total：{value} at y={b['y']}, x={b['x']}")
                    return value

            logging.warning(f"[ABR右下角暴力解法] 文件 {pdf_path} 没有找到任何金额")
            return "Not found"

        except Exception as e:
            logging.error(f"[ABR右下角暴力解法] 抓取失败：{e}")
            return "Not found"



    def extract_account_from_path(self, pdf_path):
        """从文件路径中提取账户信息"""
        account_names = ["Anthony", "Joshua", "Melvin"]
        path_parts = pdf_path.split(os.sep)
        account = "Unknown"
        for part in path_parts:
            for name in account_names:
                if name.lower() in part.lower():
                    account = name
                    return account
        return account

    def apply_filters(self):
        """应用过滤器"""
        filtered_df = self.original_df.copy()
        for col, selected_values in self.filters.items():
            filtered_df = filtered_df[filtered_df[col].isin(selected_values)]
        self.results_df = filtered_df
        self.display_results(self.results_df)

    def display_results(self, df):
        """在主线程中更新显示结果"""
        self.root.after(0, self._display_results, df.copy())  # 使用副本

    def _display_results(self, df):
        """实际更新结果显示"""
        for i in self.tree.get_children():
            self.tree.delete(i)

        for index, row in df.iterrows():
            tags = ()
            if "Not found" in [row["Invoice Date"], row["Invoice No"], row["Total"]]:
                tags = ("error",)
            self.tree.insert("", tk.END, values=(
                row["Invoice Date"],
                row["Invoice No"],
                row["Total"],
                row["Account"],
                row["File"]
            ), tags=tags)



        # 不修改 df，直接计算总和
        total_sum = df['Total'].apply(lambda x: float(str(x).replace(',', '').replace('$', ''))).sum()
        self.total_label.config(text=f"Total: {total_sum:.2f}")

    def sort_treeview_column(self, col):
        if col == self.sort_column:
            self.sort_descending = not self.sort_descending
        else:
            self.sort_column = col
            self.sort_descending = False

        for c in self.tree["columns"]:
            self.tree.heading(c, text=c)

        arrow = " ↓" if self.sort_descending else " ↑"
        self.tree.heading(col, text=col + arrow)

        # 如果是 Total，转换为 float 再排序
        if col == "Total":
            self.results_df = self.results_df.copy()
            self.results_df["Total_sortable"] = self.results_df["Total"].apply(
                lambda x: float(str(x).replace(",", "").replace("$", "")) if pd.notnull(x) else 0
            )
            self.results_df = self.results_df.sort_values(
                by="Total_sortable", ascending=not self.sort_descending
            ).drop(columns=["Total_sortable"])
        else:
            self.results_df = self.results_df.sort_values(by=col, ascending=not self.sort_descending)

        self.display_results(self.results_df)



    def filter_by_combobox_selection(self, event):
        """按账户过滤"""
        selected_account = self.account_combobox.get()
        if selected_account == "All":
            if 'Account' in self.filters:
                del self.filters['Account']
        else:
            self.filters['Account'] = [selected_account]
        self.apply_filters()

    def keyword_filter(self):
        """按关键字过滤"""
        keyword = self.keyword_entry.get()
        if keyword:
            filtered_data = self.original_df[self.original_df.apply(lambda row: row.astype(str).str.contains(keyword, case=False).any(), axis=1)]
            self.results_df = filtered_data
            if filtered_data.empty:
                messagebox.showinfo("结果", "未找到匹配关键字的发票。")
                logging.info(f"未找到关键字 '{keyword}' 的发票")
            else:
                self.display_results(filtered_data)
        else:
            self.apply_filters()

    def export_to_excel(self):
        """导出结果到 Excel"""
        if self.results_df.empty:
            messagebox.showwarning("警告", "没有可导出的数据。")
        else:
            file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel 文件", "*.xlsx")])
            if file_path:
                self.results_df.to_excel(file_path, index=False)
                messagebox.showinfo("导出成功", f"数据已成功导出到 {file_path}")
                logging.info(f"数据已导出到 {file_path}")

    def on_heading_click(self, event):
        region = self.tree.identify("region", event.x, event.y)
        if region == "heading":
            col_index = int(self.tree.identify_column(event.x).replace("#", "")) - 1
            col_name = self.tree["columns"][col_index]
            self.sort_treeview_column(col_name)

        elif region == "cell":
            pass  # 如有需要，可在此处理单元格点击事件
    def open_selected_pdf(self, event):
        """双击行后打开 PDF 文件"""
        item = self.tree.identify_row(event.y)
        if not item:
            return
        file_name = self.tree.item(item, "values")[4]  # File 列在第五位（index=4）
        folder = self.folder_path.get()

        # 遍历文件夹下所有子路径，尝试找到这个文件名
        for root, _, files in os.walk(folder):
            if file_name in files:
                full_path = os.path.join(root, file_name)
                try:
                    os.startfile(full_path)  # Windows 专用
                except AttributeError:
                    import subprocess, sys
                    if sys.platform == "darwin":  # macOS
                        subprocess.call(["open", full_path])
                    else:  # Linux
                        subprocess.call(["xdg-open", full_path])
                break

    def setup_context_menu(self):
        """设置右键菜单"""
        self.menu = tk.Menu(self.root, tearoff=0)
        self.menu.add_command(label="Filter", command=lambda: self.filter_column(self.sort_column))
        self.tree.bind("<Button-3>", self.show_context_menu)

    def show_context_menu(self, event):
        """显示右键菜单"""
        region = self.tree.identify("region", event.x, event.y)
        if region == "heading":
            col = self.tree.identify_column(event.x)
            self.sort_column = self.tree.heading(col, "text")
            self.menu.post(event.x_root, event.y_root)

    def filter_column(self, col):
        """过滤列数据"""
        values = list(set(self.results_df[col]))
        filter_window = tk.Toplevel(self.root)
        filter_window.title(f"Filter {col}")

        selected_values = tk.StringVar(value=values)
        listbox = tk.Listbox(filter_window, listvariable=selected_values, selectmode=tk.MULTIPLE)
        listbox.pack(fill=tk.BOTH, expand=True)

        def select_all():
            listbox.select_set(0, tk.END)

        def deselect_all():
            listbox.select_clear(0, tk.END)

        def apply_filter():
            selected = [listbox.get(i) for i in listbox.curselection()]
            if selected:
                self.filters[col] = selected
            elif col in self.filters:
                del self.filters[col]
            self.apply_filters()
            filter_window.destroy()

        ttk.Button(filter_window, text="全选", command=select_all).pack(pady=5)
        ttk.Button(filter_window, text="全不选", command=deselect_all).pack(pady=5)
        ttk.Button(filter_window, text="应用过滤器", command=apply_filter).pack(pady=10)

    def process_gui_queue(self):
        """处理 GUI 队列"""
        try:
            while True:
                task = self.gui_queue.get_nowait()
                if task[0] == 'update_progress':
                    self.progress_var.set(task[1])
                elif task[0] == 'display_results':
                    # 仅在数据未显示时调用 display_results
                    if not self.data_displayed:
                        self.display_results(task[1])
                        self.data_displayed = True  # 标记数据已显示
                elif task[0] == 'no_invoices':
                    messagebox.showinfo("结果", "未找到符合条件的发票。")
                    logging.info("未找到任何发票")
        except queue.Empty:
            pass
        self.root.after(100, self.process_gui_queue)

if __name__ == "__main__":
    root = tk.Tk()
    app = InvoiceExtractorApp(root)
    root.mainloop()
