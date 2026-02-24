import os
import sys
import threading
import importlib
import json
import re
import shutil
import subprocess

from PyQt6.QtCore import pyqtSignal
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
    QProgressBar,
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


def _windows_path_to_wsl(path: str) -> str:
    text = (path or "").strip().strip('"').strip("'")
    m = re.match(r"^([A-Za-z]):[\\/](.*)$", text)
    if not m:
        return text
    drive = m.group(1).lower()
    rest = m.group(2).replace("\\", "/")
    return f"/mnt/{drive}/{rest}"


def _wsl_path_to_windows(path: str) -> str:
    text = (path or "").strip().strip('"').strip("'")
    m = re.match(r"^/mnt/([a-zA-Z])/(.*)$", text)
    if not m:
        drive_match = re.match(r"^([A-Za-z]):[\\/](.*)$", text)
        if not drive_match:
            return text
        drive = drive_match.group(1).upper()
        rest = drive_match.group(2).replace("/", "\\")
        return f"{drive}:\\{rest}"
    drive = m.group(1).upper()
    rest = m.group(2).replace("/", "\\")
    return f"{drive}:\\{rest}"


def _normalize_user_path(path: str) -> str:
    text = (path or "").strip().strip('"').strip("'")
    if _is_wsl():
        text = _windows_path_to_wsl(text)
    return os.path.abspath(text)


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
    task_done = pyqtSignal(str, bool)

    def __init__(self):
        super().__init__()
        self.setWindowTitle("PDF and Excel Tool")
        self.resize(1080, 680)
        self.setup_ui()
        self.task_done.connect(self._on_task_done)

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

    def _find_windows_python(self) -> str | None:
        env_override = os.environ.get("WIN_PYTHON_EXE", "").strip()
        candidates = [
            env_override,
            "py",
            "python.exe",
            shutil.which("py.exe") or "",
            shutil.which("py") or "",
            shutil.which("python.exe") or "",
            "/mnt/c/Windows/py.exe",
            "/mnt/c/Windows/System32/py.exe",
            "/mnt/c/Users/jhunj/AppData/Local/Programs/Python/Python313/python.exe",
            "/mnt/c/Users/jhunj/AppData/Local/Programs/Python/Python312/python.exe",
            "/mnt/c/Users/jhunj/AppData/Local/Programs/Python/Python311/python.exe",
        ]
        seen = set()
        windows_only_candidates = []
        for exe in candidates:
            if not exe or exe in seen:
                continue
            seen.add(exe)
            if not os.path.exists(exe) and not shutil.which(exe):
                continue
            cmd = [exe, "-c", "import sys; print(sys.platform)"]
            exe_lower = exe.lower()
            if exe_lower.endswith("py.exe") or exe_lower == "py":
                cmd = [exe, "-3", "-c", "import sys; print(sys.platform)"]
            try:
                proc = subprocess.run(cmd, capture_output=True, text=True, timeout=8)
            except Exception:
                continue
            if not (
                proc.returncode == 0 and (proc.stdout or "").strip().startswith("win")
            ):
                continue

            windows_only_candidates.append(exe)

            check_cmd = [exe, "-c", "import win32com.client; print('ok')"]
            if exe_lower.endswith("py.exe") or exe_lower == "py":
                check_cmd = [exe, "-3", "-c", "import win32com.client; print('ok')"]
            try:
                check = subprocess.run(
                    check_cmd,
                    capture_output=True,
                    text=True,
                    timeout=10,
                )
            except Exception:
                continue
            if check.returncode == 0:
                return exe

        if windows_only_candidates:
            return windows_only_candidates[0]
        return None

    def _run_windows_converter(self, module_name: str, function_name: str, args: list):
        win_py = self._find_windows_python()
        if not win_py:
            raise RuntimeError(
                "Cannot find Windows Python. Set WIN_PYTHON_EXE or ensure py.exe is accessible in WSL."
            )

        preflight_cmd = [win_py, "-c", "import win32com.client; print('ok')"]
        if win_py.lower().endswith("py.exe") or win_py.lower() == "py":
            preflight_cmd = [
                win_py,
                "-3",
                "-c",
                "import win32com.client; print('ok')",
            ]
        preflight = subprocess.run(preflight_cmd, capture_output=True, text=True)
        if preflight.returncode != 0:
            install_hint = (
                f"{win_py} -m pip install pywin32"
                if not (win_py.lower().endswith("py.exe") or win_py.lower() == "py")
                else f"{win_py} -3 -m pip install pywin32"
            )
            raise RuntimeError(
                "Windows Python is missing pywin32 (win32com). "
                f"Install with: {install_hint}"
            )

        script_dir = os.path.dirname(os.path.abspath(__file__))
        worker_dir = _wsl_path_to_windows(script_dir) if _is_wsl() else script_dir
        payload = json.dumps(
            {
                "module": module_name,
                "function": function_name,
                "args": args,
            },
            ensure_ascii=True,
        )
        code = (
            "import json, os, sys; "
            "os.chdir(sys.argv[1]); "
            "p=json.loads(sys.argv[2]); "
            "m=__import__(p['module']); "
            "f=getattr(m,p['function']); "
            "f(*p['args'])"
        )

        cmd = [win_py, "-c", code, worker_dir, payload]
        if win_py.lower().endswith("py.exe") or win_py.lower() == "py":
            cmd = [win_py, "-3", "-c", code, worker_dir, payload]

        proc = subprocess.run(cmd, capture_output=True, text=True)
        if proc.returncode != 0:
            message = (proc.stderr or proc.stdout or "Windows converter failed").strip()
            raise RuntimeError(message)

    def _source_directory(self) -> str:
        return _normalize_user_path(self.file_path_entry.text())

    def _output_directory(self) -> str:
        return _normalize_user_path(self.output_path_entry.text())

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

        self.status_label = QLabel("Ready")
        main_layout.addWidget(self.status_label)

        self.progress = QProgressBar()
        self.progress.setVisible(False)
        self.progress.setTextVisible(False)
        main_layout.addWidget(self.progress)

    def _set_busy(self, note: str):
        self.progress.setVisible(True)
        self.progress.setRange(0, 0)
        self.status_label.setText(note)
        QApplication.processEvents()

    def _on_task_done(self, message: str, is_error: bool):
        self.progress.setRange(0, 1)
        self.progress.setValue(0)
        self.progress.setVisible(False)
        self.status_label.setText(message)
        if is_error:
            QMessageBox.critical(self, "Task failed", message)

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
        directory = self._source_directory()
        output_directory = self._output_directory()
        excel_files = self.read_excel_files(directory)
        choice = self._choice()

        if choice == "1":
            if _is_wsl():
                self._set_busy("Converting DO files...")
                self.run_in_background(
                    self._run_windows_converter,
                    "convert_do_pdf",
                    "convert_all_excels",
                    [
                        _wsl_path_to_windows(directory),
                        _wsl_path_to_windows(output_directory),
                        excel_files,
                    ],
                )
                return
            convert_all_do_excels, _ = self._load_do_converter()
            if not convert_all_do_excels:
                return
            self._set_busy("Converting DO files...")
            self.run_in_background(
                convert_all_do_excels, directory, output_directory, excel_files
            )
        elif choice == "2":
            start_num = self._parse_int(self.start_num_entry.text(), "Start number")
            end_num = self._parse_int(self.end_num_entry.text(), "End number")
            if start_num is None or end_num is None:
                return
            if _is_wsl():
                self._set_busy("Converting DO files by range...")
                self.run_in_background(
                    self._run_windows_converter,
                    "convert_do_pdf",
                    "convert_range_excels",
                    [
                        _wsl_path_to_windows(directory),
                        _wsl_path_to_windows(output_directory),
                        start_num,
                        end_num,
                        excel_files,
                    ],
                )
                return
            _, convert_range_do_excels = self._load_do_converter()
            if not convert_range_do_excels:
                return
            self._set_busy("Converting DO files by range...")
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
        directory = self._source_directory()
        output_directory = self._output_directory()
        excel_files = self.read_excel_files(directory)
        choice = self._choice()

        if choice == "1":
            if _is_wsl():
                self._set_busy("Converting INV files...")
                self.run_in_background(
                    self._run_windows_converter,
                    "convert_inv_pdf",
                    "convert_all_excels",
                    [
                        _wsl_path_to_windows(directory),
                        _wsl_path_to_windows(output_directory),
                        excel_files,
                    ],
                )
                return
            convert_all_inv_excels, _, _ = self._load_inv_converter()
            if not convert_all_inv_excels:
                return
            self._set_busy("Converting INV files...")
            self.run_in_background(
                self.convert_inv_with_debug,
                convert_all_inv_excels,
                directory,
                output_directory,
                excel_files,
            )
        elif choice == "2":
            start_num = self._parse_int(self.start_num_entry.text(), "Start number")
            end_num = self._parse_int(self.end_num_entry.text(), "End number")
            if start_num is None or end_num is None:
                return
            if _is_wsl():
                self._set_busy("Converting INV files by range...")
                self.run_in_background(
                    self._run_windows_converter,
                    "convert_inv_pdf",
                    "convert_range_excels",
                    [
                        _wsl_path_to_windows(directory),
                        _wsl_path_to_windows(output_directory),
                        start_num,
                        end_num,
                        excel_files,
                    ],
                )
                return
            _, convert_range_inv_excels, _ = self._load_inv_converter()
            if not convert_range_inv_excels:
                return
            self._set_busy("Converting INV files by range...")
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
            keyword = self.keyword_entry.text()
            if _is_wsl():
                self._set_busy("Converting INV files by keyword...")
                self.run_in_background(
                    self._run_windows_converter,
                    "convert_inv_pdf",
                    "convert_keyword_excels",
                    [
                        _wsl_path_to_windows(directory),
                        _wsl_path_to_windows(output_directory),
                        keyword,
                        excel_files,
                    ],
                )
                return
            _, _, convert_keyword_inv_excels = self._load_inv_converter()
            if not convert_keyword_inv_excels:
                return
            self._set_busy("Converting INV files by keyword...")
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
        directory = self._source_directory()
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

        self._set_busy("Merging INV PDFs...")
        self.run_in_background(merge_inv_pdfs, pdf_files, output_filename)

    def run_merge_do(self):
        directory = self._source_directory()
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

        self._set_busy("Merging DO PDFs...")
        self.run_in_background(merge_do_pdfs, pdf_files, output_filename)

    def convert_inv_with_debug(self, func, *args):
        try:
            func(*args)
            print("Conversion successful.")
        except Exception as exc:
            print(f"Conversion failed: {exc}")
            raise

    def run_in_background(self, func, *args):
        def runner():
            try:
                func(*args)
                self.task_done.emit("Done", False)
            except Exception as exc:
                error_text = str(exc).strip() or repr(exc)
                print(f"Background task failed: {error_text}")
                self.task_done.emit(error_text, True)

        thread = threading.Thread(target=runner, daemon=True)
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
