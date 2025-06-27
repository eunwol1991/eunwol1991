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
        self.keyword = ""  # 当前的关键词过滤
        self.Month = ''
        self.sort_column = None
        self.sort_descending = False
        self.progress_var = tk.DoubleVar()
        self.gui_queue = queue.Queue()
        self.data_displayed = False  # 标记数据是否已显示
        self.context_column = None  # 右键菜单当前列
        self.context_value = None   # 右键菜单选中的值

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
@@ -95,119 +98,118 @@ class InvoiceExtractorApp:
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
        ttk.Button(self.root, text="清除过滤", command=self.clear_filters).grid(row=12, column=0, padx=20, pady=10)

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
        self.keyword = ""
        self.keyword_entry.delete(0, tk.END)
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

        for root, _, files in os.walk(folder):
            if self.search_in_month_var.get() and selected_month_name not in root.lower():
                continue
            for file in files:
                if file.lower().endswith('.pdf') and self.Month.lower() in file.lower() and 'inv' in file.lower() and 'invoice' not in file.lower():
                    pdf_path = os.path.join(root, file)
                    pdf_files.append(pdf_path)
                    context = f"'{selected_month_name}' 子文件夹" if self.search_in_month_var.get() else '任意文件夹'
                    logging.info(f"在 {context} 中找到 PDF: {pdf_path}")

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
@@ -277,50 +279,53 @@ class InvoiceExtractorApp:
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
        finally:
            if 'doc' in locals() and doc:
                doc.close()
        
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
@@ -382,50 +387,57 @@ class InvoiceExtractorApp:
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
        if self.keyword:
            filtered_df = filtered_df[
                filtered_df.apply(
                    lambda row: row.astype(str).str.contains(self.keyword, case=False).any(),
                    axis=1,
                )
            ]
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


@@ -453,145 +465,170 @@ class InvoiceExtractorApp:
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
        self.keyword = self.keyword_entry.get().strip()
        self.apply_filters()
        if self.results_df.empty and self.keyword:
            messagebox.showinfo("结果", "未找到匹配关键字的发票。")
            logging.info(f"未找到关键字 '{self.keyword}' 的发票")

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

    def clear_filters(self):
        """清除所有过滤条件"""
        self.filters.clear()
        self.keyword = ""
        self.keyword_entry.delete(0, tk.END)
        self.account_combobox.current(0)
        self.apply_filters()

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
        self.header_menu = tk.Menu(self.root, tearoff=0)
        self.header_menu.add_command(label="Filter...", command=lambda: self.filter_column(self.context_column))

        self.cell_menu = tk.Menu(self.root, tearoff=0)
        self.cell_menu.add_command(label="Filter by value", command=self.filter_by_cell_value)

        self.tree.bind("<Button-3>", self.show_context_menu)

    def show_context_menu(self, event):
        """显示右键菜单"""
        region = self.tree.identify("region", event.x, event.y)
        if region == "heading":
            col_id = int(self.tree.identify_column(event.x).replace("#", "")) - 1
            self.context_column = self.tree["columns"][col_id]
            self.header_menu.post(event.x_root, event.y_root)
        elif region == "cell":
            row_id = self.tree.identify_row(event.y)
            col_id = int(self.tree.identify_column(event.x).replace("#", "")) - 1
            self.context_column = self.tree["columns"][col_id]
            values = self.tree.item(row_id, "values")
            self.context_value = values[col_id]
            label = f"Filter {self.context_column}: {self.context_value}"
            self.cell_menu.entryconfigure(0, label=label)
            self.cell_menu.post(event.x_root, event.y_root)

    def filter_column(self, col):
        """过滤列数据"""
        col = col.split()[0]  # remove sorting arrow if present
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

    def filter_by_cell_value(self):
        """根据所选单元格的值进行过滤"""
        if self.context_column and self.context_value is not None:
            current = self.filters.get(self.context_column, [])
            if self.context_value not in current:
                current.append(self.context_value)
            self.filters[self.context_column] = current
            self.apply_filters()

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
