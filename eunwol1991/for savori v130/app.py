import os
import threading
import ttkbootstrap as ttk
from ttkbootstrap.constants import *
from tkinter import filedialog, messagebox, StringVar, Tk

from convert_do_pdf import convert_all_excels as convert_all_do_excels, convert_range_excels as convert_range_do_excels
from convert_inv_pdf import convert_all_excels as convert_all_inv_excels, convert_range_excels as convert_range_inv_excels, convert_keyword_excels as convert_keyword_inv_excels
from merge_inv_pdfs import merge_INV_pdfs_in_range, merge_INV_pdfs_by_keywords, merge_all_inv_pdfs, merge_INV_pdfs_by_numbers, merge_pdfs as merge_inv_pdfs
from merge_do_pdfs import merge_do_pdfs_in_range, merge_pdfs as merge_do_pdfs


class PDFExcelApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF and Excel Tool")
        self.base_font_size = 12  # 基础字体大小

        # 设置初始窗口大小为1920 x 1080的1/4
        self.root.geometry("960x540")

        self.create_widgets()

    def create_widgets(self):
        self.font = ('Helvetica', self.base_font_size)

        ttk.Label(self.root, text="选择文件夹:", font=self.font).grid(
            row=0, column=0, padx=10, pady=10, sticky='w')

        self.file_path_entry = ttk.Entry(self.root, width=50, font=self.font)
        self.file_path_entry.grid(
            row=0, column=1, padx=10, pady=10, sticky='ew')

        self.browse_button = ttk.Button(
            self.root, text="浏览", command=self.browse_directory, bootstyle=PRIMARY)
        self.browse_button.grid(row=0, column=2, padx=10, pady=10, sticky='ew')

        ttk.Label(self.root, text="选择输出文件夹:", font=self.font).grid(
            row=1, column=0, padx=10, pady=10, sticky='w')

        self.output_path_entry = ttk.Entry(self.root, width=50, font=self.font)
        self.output_path_entry.grid(
            row=1, column=1, padx=10, pady=10, sticky='ew')

        self.browse_output_button = ttk.Button(
            self.root, text="浏览", command=self.browse_output_directory, bootstyle=PRIMARY)
        self.browse_output_button.grid(
            row=1, column=2, padx=10, pady=10, sticky='ew')

        ttk.Label(self.root, text="选择操作:", font=self.font).grid(
            row=2, column=0, padx=10, pady=10, sticky='w')

        self.options = [
            "1. 转换全部",
            "2. 输入数字序号只转换范围内的文件",
            "3. 输入关键词转换相关的文件",
            "4. 合并所有带有INV的文件",
            "5. 按数字序列查找PDF"
        ]
        self.option_var = StringVar(self.root)
        self.option_var.set(self.options[0])
        self.option_menu = ttk.Combobox(
            self.root, textvariable=self.option_var, values=self.options, state='readonly', font=self.font)
        self.option_menu.grid(row=2, column=1, padx=10, pady=10, sticky='ew')

        ttk.Label(self.root, text="起始数字序号:", font=self.font).grid(
            row=3, column=0, padx=10, pady=10, sticky='w')
        self.start_num_var = StringVar()
        self.start_num_entry = ttk.Entry(
            self.root, textvariable=self.start_num_var, font=self.font)
        self.start_num_entry.grid(
            row=3, column=1, padx=10, pady=10, sticky='ew')

        ttk.Label(self.root, text="结束数字序号:", font=self.font).grid(
            row=4, column=0, padx=10, pady=10, sticky='w')
        self.end_num_var = StringVar()
        self.end_num_entry = ttk.Entry(
            self.root, textvariable=self.end_num_var, font=self.font)
        self.end_num_entry.grid(row=4, column=1, padx=10, pady=10, sticky='ew')

        ttk.Label(self.root, text="关键词:", font=self.font).grid(
            row=5, column=0, padx=10, pady=10, sticky='w')
        self.keyword_var = StringVar()
        self.keyword_entry = ttk.Entry(
            self.root, textvariable=self.keyword_var, font=self.font)
        self.keyword_entry.grid(row=5, column=1, padx=10, pady=10, sticky='ew')

        style = ttk.Style()
        style.configure('TButton', font=self.font)

        self.convert_do_button = ttk.Button(
            self.root, text="Convert DO Excel to PDF", command=self.run_convert_do, bootstyle=SUCCESS, style='TButton')
        self.convert_do_button.grid(
            row=6, column=0, padx=10, pady=10, sticky='ew')

        self.convert_inv_button = ttk.Button(
            self.root, text="Convert INV Excel to PDF", command=self.run_convert_inv, bootstyle=SUCCESS, style='TButton')
        self.convert_inv_button.grid(
            row=6, column=1, padx=10, pady=10, sticky='ew')

        self.merge_inv_button = ttk.Button(
            self.root, text="Merge INV PDFs", command=self.run_merge_inv, bootstyle=SUCCESS, style='TButton')
        self.merge_inv_button.grid(
            row=7, column=0, padx=10, pady=10, sticky='ew')

        self.merge_do_button = ttk.Button(
            self.root, text="Merge DO PDFs", command=self.run_merge_do, bootstyle=SUCCESS, style='TButton')
        self.merge_do_button.grid(
            row=7, column=1, padx=10, pady=10, sticky='ew')

        # Adjust column weights to allow resizing
        for i in range(3):
            self.root.grid_columnconfigure(i, weight=1)

    def browse_directory(self):
        directory = filedialog.askdirectory()
        self.file_path_entry.delete(0, 'end')
        self.file_path_entry.insert(0, directory)

    def browse_output_directory(self):
        directory = filedialog.askdirectory()
        self.output_path_entry.delete(0, 'end')
        self.output_path_entry.insert(0, directory)

    def run_convert_do(self):
        directory = os.path.abspath(self.file_path_entry.get())
        output_directory = os.path.abspath(self.output_path_entry.get())
        excel_files = self.read_excel_files(directory)
        choice = self.option_var.get().split('.')[0]
        if choice == "1":
            self.run_in_background(convert_all_do_excels,
                                   directory, output_directory, excel_files)
        elif choice == "2":
            start_num = int(self.start_num_var.get())
            end_num = int(self.end_num_var.get())
            self.run_in_background(convert_range_do_excels, directory,
                                   output_directory, start_num, end_num, excel_files)
        else:
            messagebox.showerror("错误", "无效的选项")

    def run_convert_inv(self):
        directory = os.path.abspath(self.file_path_entry.get())
        output_directory = os.path.abspath(self.output_path_entry.get())
        excel_files = self.read_excel_files(directory)
        choice = self.option_var.get().split('.')[0]
        if choice == "1":
            self.run_in_background(
                self.convert_inv_with_debug, convert_all_inv_excels, directory, output_directory, excel_files)
        elif choice == "2":
            start_num = int(self.start_num_var.get())
            end_num = int(self.end_num_var.get())
            self.run_in_background(self.convert_inv_with_debug, convert_range_inv_excels,
                                   directory, output_directory, start_num, end_num, excel_files)
        elif choice == "3":
            keyword = self.keyword_var.get()
            self.run_in_background(self.convert_inv_with_debug, convert_keyword_inv_excels,
                                   directory, output_directory, keyword, excel_files)
        else:
            messagebox.showerror("错误", "无效的选项")

    def run_merge_inv(self):
        directory = os.path.abspath(self.file_path_entry.get())
        choice = self.option_var.get().split('.')[0]
        if choice == "2":
            start_num = int(self.start_num_var.get())
            end_num = int(self.end_num_var.get())
            pdf_files = merge_INV_pdfs_in_range(directory, start_num, end_num)
        elif choice == "3":
            keywords = self.keyword_var.get().split(',')
            pdf_files = merge_INV_pdfs_by_keywords(directory, keywords)
        elif choice == "4":
            pdf_files = merge_all_inv_pdfs(directory)
        elif choice == "5":
            numbers = [int(num.strip())
                       for num in self.start_num_var.get().split(',')]
            pdf_files = merge_INV_pdfs_by_numbers(directory, numbers)
        else:
            messagebox.showerror("错误", "无效的选项")
            return

        if not pdf_files:
            messagebox.showerror("错误", "没有找到符合条件的PDF文件。")
            return

        output_filename = filedialog.asksaveasfilename(
            defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])
        if not output_filename:
            return

        self.run_in_background(merge_inv_pdfs, pdf_files, output_filename)

    def run_merge_do(self):
        directory = os.path.abspath(self.file_path_entry.get())
        choice = self.option_var.get().split('.')[0]
        if choice == "2":
            start_num = int(self.start_num_var.get())
            end_num = int(self.end_num_var.get())
            pdf_files = merge_do_pdfs_in_range(directory, start_num, end_num)
        elif choice == "5":
            numbers = [int(num.strip())
                       for num in self.start_num_var.get().split(',')]
            pdf_files = merge_do_pdfs_in_range(
                directory, numbers[0], numbers[-1])
        else:
            messagebox.showerror("错误", "无效的选项")
            return

        if not pdf_files:
            messagebox.showerror("错误", "没有找到符合条件的PDF文件。")
            return

        output_filename = filedialog.asksaveasfilename(
            defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")])
        if not output_filename:
            return

        self.run_in_background(merge_do_pdfs, pdf_files, output_filename)

    def convert_inv_with_debug(self, func, *args):
        try:
            func(*args)
            print("Conversion successful.")
        except Exception as e:
            print(f"Conversion failed: {e}")

    def run_in_background(self, func, *args):
        # 在运行任务之前重新导入 win32com.client
        import win32com.client as win32
        thread = threading.Thread(target=func, args=args)
        thread.start()

    def read_excel_files(self, directory):
        excel_files = []
        for filename in os.listdir(directory):
            if (filename.endswith(".xlsx") or filename.endswith(".xls")) and not filename.startswith("~$"):
                excel_files.append(filename)
        return excel_files


if __name__ == "__main__":
    root = ttk.Window(themename="solar")
    app = PDFExcelApp(root)
    root.mainloop()
