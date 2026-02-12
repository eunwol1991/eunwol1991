import os
import sys
import threading
import importlib

from PyQt6.QtGui import QFont
from PyQt6.QtWidgets import (
    QApplication,
    QComboBox,
    QFileDialog,
    QGridLayout,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QVBoxLayout,
    QWidget,
)

from merge_inv_pdfs import (
    merge_INV_pdfs_in_range,
    merge_INV_pdfs_by_keywords,
    merge_all_inv_pdfs,
    merge_INV_pdfs_by_numbers,
    merge_pdfs as merge_inv_pdfs,
)
from merge_do_pdfs import merge_do_pdfs_in_range, merge_pdfs as merge_do_pdfs


def _is_wsl() -> bool:
    if os.name == "nt":
        return False
    if os.environ.get("WSL_DISTRO_NAME"):
        return True
    try:
        with open("/proc/version", "r", encoding="utf-8") as f:
            return "microsoft" in f.read().lower()
    except OSError:
        return False


def _platform_drive_root() -> str:
    if os.name == "nt":
        return "c:/"
    if _is_wsl():
        return "/mnt/c"
    return "/"


def _default_base_dir() -> str:
    root = _platform_drive_root()
    if root.endswith("/"):
        return f"{root}Users/jhunj/Dropbox/DO & INV/DO & INV 2026"
    return f"{root}/Users/jhunj/Dropbox/DO & INV/DO & INV 2026"


BASE_DIR_DEFAULT = _default_base_dir()


TOKYONIGHT_QSS = """
QWidget {
    background-color: #1a1b26;
    color: #c0caf5;
    font-size: 13px;
}

QLabel {
    color: #c0caf5;
    font-weight: 600;
}

QLineEdit,
QComboBox {
    background-color: #24283b;
    color: #c0caf5;
    border: 1px solid #414868;
    border-radius: 8px;
    padding: 8px 10px;
    selection-background-color: #33467c;
    min-height: 22px;
}

QLineEdit:focus,
QComboBox:focus {
    border: 1px solid #7aa2f7;
}

QPushButton {
    background-color: #7aa2f7;
    color: #1a1b26;
    border: none;
    border-radius: 8px;
    padding: 9px 12px;
    font-weight: 700;
    min-height: 24px;
}

QPushButton:hover {
    background-color: #89b4fa;
}

QPushButton:pressed {
    background-color: #5f82ce;
}
"""


class PDFExcelApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("PDF and Excel Tool")
        self.resize(1080, 680)
        self.setup_ui()

    def _load_do_converter(self):
        try:
            mod = importlib.import_module("convert_do_pdf")
            return mod.convert_all_excels, mod.convert_range_excels
        except ModuleNotFoundError as exc:
            if getattr(exc, "name", "") == "win32com":
                QMessageBox.critical(
                    self,
                    "Error",
                    "Missing dependency: pywin32 (win32com).\n"
                    "This Excel-to-PDF feature requires Windows Python + Microsoft Excel.\n"
                    "Install on Windows env: pip install pywin32",
                )
                return None, None
            raise

    def _load_inv_converter(self):
        try:
            mod = importlib.import_module("convert_inv_pdf")
            return (
                mod.convert_all_excels,
                mod.convert_range_excels,
                mod.convert_keyword_excels,
            )
        except ModuleNotFoundError as exc:
            if getattr(exc, "name", "") == "win32com":
                QMessageBox.critical(
                    self,
                    "Error",
                    "Missing dependency: pywin32 (win32com).\n"
                    "This Excel-to-PDF feature requires Windows Python + Microsoft Excel.\n"
                    "Install on Windows env: pip install pywin32",
                )
                return None, None, None
            raise

    def setup_ui(self):
        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(18, 18, 18, 18)
        main_layout.setSpacing(12)

        grid = QGridLayout()
        grid.setHorizontalSpacing(12)
        grid.setVerticalSpacing(12)

        row = 0
        grid.addWidget(QLabel("Select folder:"), row, 0)
        self.file_path_entry = QLineEdit()
        self.file_path_entry.setText(
            BASE_DIR_DEFAULT if os.path.isdir(BASE_DIR_DEFAULT) else ""
        )
        grid.addWidget(self.file_path_entry, row, 1)
        browse_btn = QPushButton("Browse")
        browse_btn.clicked.connect(self.browse_directory)
        grid.addWidget(browse_btn, row, 2)

        row += 1
        grid.addWidget(QLabel("Select output folder:"), row, 0)
        self.output_path_entry = QLineEdit()
        self.output_path_entry.setText(
            BASE_DIR_DEFAULT if os.path.isdir(BASE_DIR_DEFAULT) else ""
        )
        grid.addWidget(self.output_path_entry, row, 1)
        browse_out_btn = QPushButton("Browse")
        browse_out_btn.clicked.connect(self.browse_output_directory)
        grid.addWidget(browse_out_btn, row, 2)

        row += 1
        grid.addWidget(QLabel("Select action:"), row, 0)
        self.options = [
            "1. Convert all",
            "2. Convert by numeric range",
            "3. Convert by keyword",
            "4. Merge all files with INV",
            "5. Find PDFs by number list",
        ]
        self.option_menu = QComboBox()
        self.option_menu.addItems(self.options)
        grid.addWidget(self.option_menu, row, 1)

        row += 1
        grid.addWidget(QLabel("Start number:"), row, 0)
        self.start_num_entry = QLineEdit()
        grid.addWidget(self.start_num_entry, row, 1)

        row += 1
        grid.addWidget(QLabel("End number:"), row, 0)
        self.end_num_entry = QLineEdit()
        grid.addWidget(self.end_num_entry, row, 1)

        row += 1
        grid.addWidget(QLabel("Keyword:"), row, 0)
        self.keyword_entry = QLineEdit()
        grid.addWidget(self.keyword_entry, row, 1)

        main_layout.addLayout(grid)

        button_row_1 = QHBoxLayout()
        button_row_1.setSpacing(10)
        convert_do_btn = QPushButton("Convert DO Excel to PDF")
        convert_do_btn.clicked.connect(self.run_convert_do)
        button_row_1.addWidget(convert_do_btn)

        convert_inv_btn = QPushButton("Convert INV Excel to PDF")
        convert_inv_btn.clicked.connect(self.run_convert_inv)
        button_row_1.addWidget(convert_inv_btn)
        main_layout.addLayout(button_row_1)

        button_row_2 = QHBoxLayout()
        button_row_2.setSpacing(10)
        merge_inv_btn = QPushButton("Merge INV PDFs")
        merge_inv_btn.clicked.connect(self.run_merge_inv)
        button_row_2.addWidget(merge_inv_btn)

        merge_do_btn = QPushButton("Merge DO PDFs")
        merge_do_btn.clicked.connect(self.run_merge_do)
        button_row_2.addWidget(merge_do_btn)
        main_layout.addLayout(button_row_2)

    def browse_directory(self):
        default_dir = (
            BASE_DIR_DEFAULT
            if os.path.isdir(BASE_DIR_DEFAULT)
            else _platform_drive_root()
        )
        directory = QFileDialog.getExistingDirectory(
            self, "Select source folder", default_dir
        )
        if directory:
            self.file_path_entry.setText(directory)

    def browse_output_directory(self):
        default_dir = (
            BASE_DIR_DEFAULT
            if os.path.isdir(BASE_DIR_DEFAULT)
            else _platform_drive_root()
        )
        directory = QFileDialog.getExistingDirectory(
            self, "Select output folder", default_dir
        )
        if directory:
            self.output_path_entry.setText(directory)

    def _choice(self):
        return self.option_menu.currentText().split(".")[0]

    def _parse_int(self, text: str, field_name: str):
        try:
            return int(text.strip())
        except ValueError:
            QMessageBox.warning(self, "Error", f"{field_name} must be an integer.")
            return None

    def run_convert_do(self):
        directory = os.path.abspath(self.file_path_entry.text().strip())
        output_directory = os.path.abspath(self.output_path_entry.text().strip())
        excel_files = self.read_excel_files(directory)
        choice = self._choice()

        if choice == "1":
            convert_all_do_excels, _ = self._load_do_converter()
            if not convert_all_do_excels:
                return
            self.run_in_background(
                convert_all_do_excels, directory, output_directory, excel_files
            )
        elif choice == "2":
            _, convert_range_do_excels = self._load_do_converter()
            if not convert_range_do_excels:
                return
            start_num = self._parse_int(self.start_num_entry.text(), "Start number")
            end_num = self._parse_int(self.end_num_entry.text(), "End number")
            if start_num is None or end_num is None:
                return
            self.run_in_background(
                convert_range_do_excels,
                directory,
                output_directory,
                start_num,
                end_num,
                excel_files,
            )
        else:
            QMessageBox.warning(self, "Error", "Invalid option")

    def run_convert_inv(self):
        directory = os.path.abspath(self.file_path_entry.text().strip())
        output_directory = os.path.abspath(self.output_path_entry.text().strip())
        excel_files = self.read_excel_files(directory)
        choice = self._choice()

        if choice == "1":
            convert_all_inv_excels, _, _ = self._load_inv_converter()
            if not convert_all_inv_excels:
                return
            self.run_in_background(
                self.convert_inv_with_debug,
                convert_all_inv_excels,
                directory,
                output_directory,
                excel_files,
            )
        elif choice == "2":
            _, convert_range_inv_excels, _ = self._load_inv_converter()
            if not convert_range_inv_excels:
                return
            start_num = self._parse_int(self.start_num_entry.text(), "Start number")
            end_num = self._parse_int(self.end_num_entry.text(), "End number")
            if start_num is None or end_num is None:
                return
            self.run_in_background(
                self.convert_inv_with_debug,
                convert_range_inv_excels,
                directory,
                output_directory,
                start_num,
                end_num,
                excel_files,
            )
        elif choice == "3":
            _, _, convert_keyword_inv_excels = self._load_inv_converter()
            if not convert_keyword_inv_excels:
                return
            keyword = self.keyword_entry.text()
            self.run_in_background(
                self.convert_inv_with_debug,
                convert_keyword_inv_excels,
                directory,
                output_directory,
                keyword,
                excel_files,
            )
        else:
            QMessageBox.warning(self, "Error", "Invalid option")

    def run_merge_inv(self):
        directory = os.path.abspath(self.file_path_entry.text().strip())
        choice = self._choice()

        if choice == "2":
            start_num = self._parse_int(self.start_num_entry.text(), "Start number")
            end_num = self._parse_int(self.end_num_entry.text(), "End number")
            if start_num is None or end_num is None:
                return
            pdf_files = merge_INV_pdfs_in_range(directory, start_num, end_num)
        elif choice == "3":
            keywords = self.keyword_entry.text().split(",")
            pdf_files = merge_INV_pdfs_by_keywords(directory, keywords)
        elif choice == "4":
            pdf_files = merge_all_inv_pdfs(directory)
        elif choice == "5":
            try:
                numbers = [
                    int(num.strip()) for num in self.start_num_entry.text().split(",")
                ]
            except ValueError:
                QMessageBox.warning(self, "Error", "Number list must contain integers.")
                return
            pdf_files = merge_INV_pdfs_by_numbers(directory, numbers)
        else:
            QMessageBox.warning(self, "Error", "Invalid option")
            return

        if not pdf_files:
            QMessageBox.warning(self, "Error", "No matching PDF files found.")
            return

        output_filename, _ = QFileDialog.getSaveFileName(
            self,
            "Save merged INV PDF",
            "",
            "PDF files (*.pdf)",
        )
        if not output_filename:
            return
        if not output_filename.lower().endswith(".pdf"):
            output_filename += ".pdf"

        self.run_in_background(merge_inv_pdfs, pdf_files, output_filename)

    def run_merge_do(self):
        directory = os.path.abspath(self.file_path_entry.text().strip())
        choice = self._choice()

        if choice == "2":
            start_num = self._parse_int(self.start_num_entry.text(), "Start number")
            end_num = self._parse_int(self.end_num_entry.text(), "End number")
            if start_num is None or end_num is None:
                return
            pdf_files = merge_do_pdfs_in_range(directory, start_num, end_num)
        elif choice == "5":
            try:
                numbers = [
                    int(num.strip()) for num in self.start_num_entry.text().split(",")
                ]
            except ValueError:
                QMessageBox.warning(self, "Error", "Number list must contain integers.")
                return
            if not numbers:
                QMessageBox.warning(self, "Error", "Number list is empty.")
                return
            pdf_files = merge_do_pdfs_in_range(directory, numbers[0], numbers[-1])
        else:
            QMessageBox.warning(self, "Error", "Invalid option")
            return

        if not pdf_files:
            QMessageBox.warning(self, "Error", "No matching PDF files found.")
            return

        output_filename, _ = QFileDialog.getSaveFileName(
            self,
            "Save merged DO PDF",
            "",
            "PDF files (*.pdf)",
        )
        if not output_filename:
            return
        if not output_filename.lower().endswith(".pdf"):
            output_filename += ".pdf"

        self.run_in_background(merge_do_pdfs, pdf_files, output_filename)

    def convert_inv_with_debug(self, func, *args):
        try:
            func(*args)
            print("Conversion successful.")
        except Exception as exc:
            print(f"Conversion failed: {exc}")

    def run_in_background(self, func, *args):
        thread = threading.Thread(target=func, args=args, daemon=True)
        thread.start()

    def read_excel_files(self, directory):
        excel_files = []
        if not os.path.isdir(directory):
            QMessageBox.warning(self, "Error", f"Invalid folder path: {directory}")
            return excel_files
        for filename in os.listdir(directory):
            if (
                filename.endswith(".xlsx") or filename.endswith(".xls")
            ) and not filename.startswith("~$"):
                excel_files.append(filename)
        return excel_files


if __name__ == "__main__":
    qt_app = QApplication(sys.argv)
    qt_app.setStyle("Fusion")
    qt_app.setFont(QFont("Segoe UI", 12))
    qt_app.setStyleSheet(TOKYONIGHT_QSS)

    window = PDFExcelApp()
    window.show()
    sys.exit(qt_app.exec())
