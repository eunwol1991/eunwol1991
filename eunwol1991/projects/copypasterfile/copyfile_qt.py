import os
import re
import shutil
import sys
from datetime import date, timedelta

try:
    from PySide6.QtCore import QTimer, Qt
    from PySide6.QtGui import QFont
    from PySide6.QtWidgets import (
        QApplication,
        QCheckBox,
        QDialog,
        QFileDialog,
        QGridLayout,
        QHBoxLayout,
        QLabel,
        QLineEdit,
        QListWidget,
        QListWidgetItem,
        QMainWindow,
        QMessageBox,
        QPushButton,
        QSlider,
        QVBoxLayout,
        QWidget,
    )
except ModuleNotFoundError as exc:
    raise SystemExit(
        "PySide6 is required. Install with: python -m pip install PySide6"
    ) from exc


def _windows_path_to_wsl(path: str) -> str:
    text = (path or "").strip()
    m = re.match(r"^([A-Za-z]):\\(.*)$", text)
    if not m:
        return text
    drive = m.group(1).lower()
    rest = m.group(2).replace("\\", "/")
    return f"/mnt/{drive}/{rest}"


def _first_existing_path(candidates: list[str]) -> str:
    for c in candidates:
        if c and os.path.isdir(c):
            return c
    return ""


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


def _default_dropbox_base() -> str:
    root = _platform_drive_root()
    if root.endswith("/"):
        return f"{root}Users/jhunj/Dropbox"
    return f"{root}/Users/jhunj/Dropbox"


FILE_PATTERN = re.compile(
    r"""^(?P<prefix>[A-Z0-9._ \-]+?)
        \s+xx26\s*[-\u2013\u2014]\s*00x
        (?:\s*[-\u2013\u2014]\s*DO\s*&\s*INV)?
        (?:\s*\((?P<name>.+)\))?
    """,
    re.IGNORECASE | re.VERBOSE,
)

WEF_PATTERN = re.compile(
    r"\bWEF\b[^0-9]{0,12}(\d{1,2})[./-](\d{1,2})[./-](\d{2,4})\b",
    re.IGNORECASE,
)


TOKYONIGHT_QSS_TEMPLATE = """
QWidget {
    background-color: #1a1b26;
    color: #c0caf5;
    font-size: {font_px}px;
}

QLabel {
    color: #c0caf5;
    font-weight: 600;
}

QLineEdit,
QComboBox,
QListWidget {
    background-color: #24283b;
    color: #c0caf5;
    border: 1px solid #414868;
    border-radius: 8px;
    padding: 8px 10px;
    selection-background-color: #33467c;
    min-height: 22px;
}

QLineEdit:focus,
QComboBox:focus,
QListWidget:focus {
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


def build_tokyonight_qss(font_px: int) -> str:
    return TOKYONIGHT_QSS_TEMPLATE.replace("{font_px}", str(max(10, int(font_px))))


def get_activation_cutoff(ref_date: date) -> date:
    if ref_date.weekday() == 4:
        return ref_date + timedelta(days=4)
    if ref_date.weekday() == 5:
        return ref_date + timedelta(days=2)
    if ref_date.weekday() == 6:
        return ref_date + timedelta(days=1)
    return ref_date + timedelta(days=1)


def classify_wef_status(
    wef_date: date | None, ref_date: date, cutoff_date: date
) -> str:
    if wef_date is None:
        return "no-wef"
    if wef_date < ref_date:
        return "past"
    if wef_date <= cutoff_date:
        return "current"
    return "future"


def extract_wef_date_from_path(path: str) -> date | None:
    latest = None
    folder_path = os.path.dirname(path or "")
    for segment in os.path.normpath(folder_path).split(os.sep):
        if not segment:
            continue
        for m in WEF_PATTERN.finditer(segment):
            day = int(m.group(1))
            month = int(m.group(2))
            year = int(m.group(3))
            if year < 100:
                year += 2000
            try:
                parsed = date(year, month, day)
            except ValueError:
                continue
            if latest is None or parsed > latest:
                latest = parsed
    return latest


def get_next_invoice_number(search_dir: str, invoice_prefix: str) -> int:
    pattern = re.compile(
        rf"(?<!\d){re.escape(invoice_prefix)}\s*-\s*(\d{{3}})(?!\d)", re.IGNORECASE
    )
    max_number = 0
    for root_dir, dirs, files in os.walk(search_dir):
        dirs[:] = [d for d in dirs if d.lower() != "history"]
        for name in files:
            for m in pattern.finditer(name):
                max_number = max(max_number, int(m.group(1)))
    return max_number + 1


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("File Selector (Qt)")
        self.resize(1080, 680)

        dropbox_base = _default_dropbox_base()
        source_default = f"{dropbox_base}/DO & INV/DO & INV 2026"
        target_default = f"{dropbox_base}/for jj/Doc to print - JJ"

        self.source_dir = _first_existing_path(
            [
                source_default,
                os.getcwd(),
            ]
        )
        self.target_dir = _first_existing_path(
            [
                target_default,
            ]
        )

        self.file_info_list: list[dict] = []
        self.filtered_indices: list[int] = []
        self.selected_indices: list[int] = []
        self.base_font = 12
        self.default_browse_root = self.source_dir or _platform_drive_root()

        root = QWidget(self)
        self.setCentralWidget(root)
        v = QVBoxLayout(root)
        v.setContentsMargins(18, 18, 18, 18)
        v.setSpacing(12)

        top = QHBoxLayout()
        self.title_label = QLabel("Choose Source Folder and Match Files")
        self.title_label.setStyleSheet("font-weight: 700;")
        top.addWidget(self.title_label)
        top.addStretch(1)
        self.scale_label = QLabel("UI Scale: 140%")
        top.addWidget(self.scale_label)
        self.scale_slider = QSlider(Qt.Orientation.Horizontal)
        self.scale_slider.setRange(80, 180)
        self.scale_slider.setValue(140)
        self.scale_slider.valueChanged.connect(self.apply_scale)
        self.scale_slider.setFixedWidth(220)
        top.addWidget(self.scale_slider)
        v.addLayout(top)

        source_row = QHBoxLayout()
        self.source_edit = QLineEdit(self.source_dir)
        source_row.addWidget(QLabel("Source:"))
        source_row.addWidget(self.source_edit)
        btn_source = QPushButton("Browse Source")
        btn_source.clicked.connect(self.browse_source)
        source_row.addWidget(btn_source)
        v.addLayout(source_row)

        filter_row = QHBoxLayout()
        filter_row.addWidget(QLabel("Filter:"))
        self.filter_edit = QLineEdit()
        self.filter_edit.setPlaceholderText("Type keywords to filter matched files")
        self.filter_edit.textChanged.connect(self.apply_file_filter)
        filter_row.addWidget(self.filter_edit)
        btn_clear_filter = QPushButton("Clear Filter")
        btn_clear_filter.clicked.connect(lambda: self.filter_edit.setText(""))
        filter_row.addWidget(btn_clear_filter)
        v.addLayout(filter_row)

        lists = QGridLayout()
        self.file_list = QListWidget()
        self.file_list.itemDoubleClicked.connect(lambda _x: self.add_selected())
        self.selected_list = QListWidget()
        lists.addWidget(QLabel("Matched Files"), 0, 0)
        lists.addWidget(QLabel("Selected Files Order"), 0, 1)
        lists.addWidget(self.file_list, 1, 0)
        lists.addWidget(self.selected_list, 1, 1)
        v.addLayout(lists)

        action_row = QHBoxLayout()
        for text, cb in [
            ("Refresh", self.refresh_source),
            ("Add", self.add_selected),
            ("Delete", self.delete_selected),
            ("Clear", self.clear_selected),
            ("Up", self.move_up),
            ("Down", self.move_down),
        ]:
            b = QPushButton(text)
            b.clicked.connect(cb)
            action_row.addWidget(b)
        v.addLayout(action_row)

        invoice_row = QHBoxLayout()
        invoice_row.addWidget(QLabel("Invoice Start:"))
        self.invoice_edit = QLineEdit()
        self.invoice_edit.setPlaceholderText("0326 - 001 or 0326 (auto mode)")
        self.invoice_edit.setFixedWidth(220)
        invoice_row.addWidget(self.invoice_edit)
        self.auto_checkbox = QCheckBox("Auto number by MMYY")
        invoice_row.addWidget(self.auto_checkbox)
        invoice_row.addStretch(1)
        btn_copy = QPushButton("Copy")
        btn_copy.clicked.connect(self.copy_files)
        invoice_row.addWidget(btn_copy)
        v.addLayout(invoice_row)

        self.status = QLabel("")
        v.addWidget(self.status)

        self.apply_scale(self.scale_slider.value())
        self.refresh_source()

    def apply_scale(self, percent: int):
        factor = max(0.8, float(percent) / 100.0)
        app = QApplication.instance()
        if isinstance(app, QApplication):
            f = QFont("Segoe UI", max(8, int(round(self.base_font * factor))))
            app.setFont(f)
            app.setStyleSheet(build_tokyonight_qss(int(round(13 * factor))))
        title_size = max(12, int(round(18 * factor)))
        self.title_label.setStyleSheet(f"font-weight: 700; font-size: {title_size}px;")
        self.scale_label.setText(f"UI Scale: {int(percent)}%")

    def browse_source(self):
        start = self.default_browse_root
        path = self.ask_directory_quick(start)
        if not path:
            return
        self.source_edit.setText(path)
        self.refresh_source()

    def ask_directory_quick(self, initial_dir: str) -> str:
        class QuickFolderPickerDialog(QDialog):
            def __init__(self, parent, start_dir: str):
                super().__init__(parent)
                self.setWindowTitle("Quick Folder Picker")
                self.resize(920, 620)
                self.selected_path = ""
                self.indexed_dirs: list[dict[str, str]] = []
                self.filtered_dirs: list[dict[str, str]] = []

                base_dir = os.path.abspath(start_dir or os.getcwd())
                if not os.path.isdir(base_dir):
                    base_dir = os.getcwd()

                root_layout = QVBoxLayout(self)

                top_row = QHBoxLayout()
                top_row.addWidget(QLabel("Index root folder:"))
                self.root_edit = QLineEdit(base_dir)
                top_row.addWidget(self.root_edit)
                btn_change_root = QPushButton("Change Root")
                btn_change_root.clicked.connect(self.change_root)
                top_row.addWidget(btn_change_root)
                root_layout.addLayout(top_row)

                filter_row = QHBoxLayout()
                filter_row.addWidget(QLabel("Keyword filter:"))
                self.keyword_edit = QLineEdit()
                self.keyword_edit.setPlaceholderText(
                    "Type keywords (space-separated) to filter folders"
                )
                self.keyword_edit.textChanged.connect(self.apply_filter)
                self.keyword_edit.returnPressed.connect(self.confirm)
                filter_row.addWidget(self.keyword_edit)
                btn_clear = QPushButton("Clear")
                btn_clear.clicked.connect(lambda: self.keyword_edit.setText(""))
                filter_row.addWidget(btn_clear)
                root_layout.addLayout(filter_row)

                self.folder_list = QListWidget()
                self.folder_list.itemDoubleClicked.connect(lambda _item: self.confirm())
                root_layout.addWidget(self.folder_list)

                self.status_label = QLabel("Waiting to build folder index...")
                root_layout.addWidget(self.status_label)

                btn_row = QHBoxLayout()
                btn_system = QPushButton("System Picker")
                btn_system.clicked.connect(self.use_system_picker)
                btn_row.addWidget(btn_system)
                btn_row.addStretch(1)
                btn_cancel = QPushButton("Cancel")
                btn_cancel.clicked.connect(self.reject)
                btn_row.addWidget(btn_cancel)
                btn_confirm = QPushButton("Confirm")
                btn_confirm.clicked.connect(self.confirm)
                btn_row.addWidget(btn_confirm)
                root_layout.addLayout(btn_row)

                self.rebuild_index()
                QTimer.singleShot(0, self.keyword_edit.setFocus)

            def change_root(self):
                start = self.root_edit.text().strip() or os.getcwd()
                picked = QFileDialog.getExistingDirectory(
                    self, "Select root folder", start
                )
                if not picked:
                    return
                self.root_edit.setText(picked)
                self.rebuild_index()

            def use_system_picker(self):
                start = self.root_edit.text().strip() or os.getcwd()
                picked = QFileDialog.getExistingDirectory(
                    self, "Select source folder", start
                )
                if not picked:
                    return
                self.selected_path = picked
                self.accept()

            def rebuild_index(self):
                selected_root = os.path.abspath(
                    self.root_edit.text().strip() or os.getcwd()
                )
                if not os.path.isdir(selected_root):
                    QMessageBox.warning(self, "Warning", "Root folder does not exist.")
                    return

                self.status_label.setText("Scanning folders, please wait...")
                QApplication.processEvents()

                self.indexed_dirs.clear()
                self.indexed_dirs.append({"path": selected_root, "relative": "."})
                for root_dir, dirs, _ in os.walk(selected_root):
                    dirs[:] = [d for d in dirs if d.lower() != "history"]
                    for name in dirs:
                        full_path = os.path.join(root_dir, name)
                        relative = os.path.relpath(full_path, selected_root)
                        self.indexed_dirs.append(
                            {"path": full_path, "relative": relative}
                        )
                self.indexed_dirs.sort(key=lambda x: x["relative"].lower())
                self.apply_filter()

            def apply_filter(self):
                query = self.keyword_edit.text().strip().lower()
                terms = [term for term in query.split() if term]

                self.filtered_dirs = []
                for item in self.indexed_dirs:
                    haystack = (
                        f"{item['relative']} {os.path.basename(item['path'])}".lower()
                    )
                    if terms and not all(term in haystack for term in terms):
                        continue
                    self.filtered_dirs.append(item)

                self.folder_list.clear()
                for entry in self.filtered_dirs:
                    folder_name = os.path.basename(entry["path"]) or entry["path"]
                    text = f"{folder_name}    [{entry['relative']}]"
                    list_item = QListWidgetItem(text)
                    list_item.setData(Qt.ItemDataRole.UserRole, entry["path"])
                    self.folder_list.addItem(list_item)

                if self.folder_list.count() > 0:
                    self.folder_list.setCurrentRow(0)

                self.status_label.setText(
                    f"Indexed {len(self.indexed_dirs)} folder(s), matched {len(self.filtered_dirs)}."
                )

            def confirm(self):
                item = self.folder_list.currentItem()
                if item is None:
                    QMessageBox.warning(
                        self, "Warning", "Please choose one folder first."
                    )
                    return
                path = item.data(Qt.ItemDataRole.UserRole)
                if not path or not os.path.isdir(path):
                    QMessageBox.warning(
                        self, "Warning", "Selected folder is not valid."
                    )
                    return
                self.selected_path = path
                self.accept()

        dialog = QuickFolderPickerDialog(self, initial_dir)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            return dialog.selected_path
        return ""

    def refresh_source(self):
        source_dir = self.source_edit.text().strip()
        if not os.path.isdir(source_dir):
            QMessageBox.warning(self, "Warning", "Source folder does not exist.")
            return
        self.source_dir = source_dir
        self.file_info_list.clear()
        self.filtered_indices.clear()
        self.selected_indices.clear()
        self.file_list.clear()
        self.selected_list.clear()

        ref = date.today()
        cutoff = get_activation_cutoff(ref)
        hidden_future = 0
        for root_dir, _, files in os.walk(source_dir):
            if "history" in root_dir.lower():
                continue
            for name in files:
                full_path = os.path.join(root_dir, name)
                if not os.path.isfile(full_path):
                    continue
                no_ext = os.path.splitext(name)[0]
                m = FILE_PATTERN.match(no_ext)
                if not m:
                    continue
                wef_date = extract_wef_date_from_path(full_path)
                status = classify_wef_status(wef_date, ref, cutoff)
                if status == "future":
                    hidden_future += 1
                    continue
                friendly = (m.group("name") or m.group("prefix")).strip()
                if wef_date:
                    friendly = (
                        f"{friendly} [WEF {wef_date:%d-%m-%Y} | {status.upper()}]"
                    )
                else:
                    friendly = f"{friendly} [NO WEF]"
                self.file_info_list.append(
                    {"display_name": friendly, "file_path": full_path}
                )
        self.apply_file_filter()
        self.status.setText(
            f"Loaded {len(self.file_info_list)} file(s). Hidden FUTURE: {hidden_future}."
        )

    def apply_file_filter(self):
        query = self.filter_edit.text().strip().lower()
        terms = [t for t in query.split() if t]

        self.filtered_indices.clear()
        self.file_list.clear()

        for idx, info in enumerate(self.file_info_list):
            haystack = (
                f"{info['display_name']} {os.path.basename(info['file_path'])}".lower()
            )
            if terms and not all(term in haystack for term in terms):
                continue
            self.filtered_indices.append(idx)

        for visible_pos, real_idx in enumerate(self.filtered_indices, start=1):
            info = self.file_info_list[real_idx]
            self.file_list.addItem(f"{visible_pos}. {info['display_name']}")

        selected_visible = [
            idx for idx in self.selected_indices if idx in set(self.filtered_indices)
        ]
        self.status.setText(
            f"Matched {len(self.filtered_indices)} / {len(self.file_info_list)} file(s). "
            f"Selected {len(self.selected_indices)} total ({len(selected_visible)} visible)."
        )

    def add_selected(self):
        row = self.file_list.currentRow()
        if row < 0 or row >= len(self.filtered_indices):
            return
        real_index = self.filtered_indices[row]
        if real_index in self.selected_indices:
            return
        self.selected_indices.append(real_index)
        self.refresh_selected_list()

    def delete_selected(self):
        row = self.selected_list.currentRow()
        if row < 0 or row >= len(self.selected_indices):
            return
        del self.selected_indices[row]
        self.refresh_selected_list()

    def clear_selected(self):
        self.selected_indices.clear()
        self.refresh_selected_list()

    def move_up(self):
        row = self.selected_list.currentRow()
        if row <= 0:
            return
        self.selected_indices[row - 1], self.selected_indices[row] = (
            self.selected_indices[row],
            self.selected_indices[row - 1],
        )
        self.refresh_selected_list(select_row=row - 1)

    def move_down(self):
        row = self.selected_list.currentRow()
        if row < 0 or row >= len(self.selected_indices) - 1:
            return
        self.selected_indices[row + 1], self.selected_indices[row] = (
            self.selected_indices[row],
            self.selected_indices[row + 1],
        )
        self.refresh_selected_list(select_row=row + 1)

    def refresh_selected_list(self, select_row: int | None = None):
        self.selected_list.clear()
        for i, idx in enumerate(self.selected_indices, start=1):
            info = self.file_info_list[idx]
            self.selected_list.addItem(QListWidgetItem(f"{i}. {info['display_name']}"))
        if select_row is not None and 0 <= select_row < self.selected_list.count():
            self.selected_list.setCurrentRow(select_row)

    def copy_files(self):
        if not self.target_dir:
            QMessageBox.critical(self, "Error", "Target folder is not configured.")
            return
        try:
            os.makedirs(self.target_dir, exist_ok=True)
        except OSError as exc:
            QMessageBox.critical(self, "Error", f"Cannot access target folder: {exc}")
            return
        if not self.selected_indices:
            QMessageBox.warning(self, "Warning", "Please choose files to copy first.")
            return

        invoice_input = self.invoice_edit.text().strip()
        if self.auto_checkbox.isChecked():
            if not re.fullmatch(r"\d{4}", invoice_input):
                QMessageBox.warning(
                    self, "Warning", "Auto mode expects 4-digit prefix (MMYY)."
                )
                return
            invoice_prefix = invoice_input
            invoice_number = get_next_invoice_number(self.source_dir, invoice_prefix)
        else:
            if not re.fullmatch(r"\d{4}\s*-\s*\d{3}", invoice_input):
                QMessageBox.warning(self, "Warning", "Use format: 0326 - 001")
                return
            invoice_prefix, num_text = invoice_input.split("-")
            invoice_prefix = invoice_prefix.strip()
            invoice_number = int(num_text.strip())

        if invoice_number > 999:
            QMessageBox.critical(
                self,
                "Error",
                f"Prefix {invoice_prefix} already reached {invoice_number:03d}.",
            )
            return

        copied = 0
        for idx in self.selected_indices:
            src_path = self.file_info_list[idx]["file_path"]
            filename = os.path.basename(src_path)
            new_filename = re.sub(
                r"xx26\s*[-\u2013\u2014]\s*00x",
                f"{invoice_prefix} - {invoice_number:03d}",
                filename,
                flags=re.IGNORECASE,
            )
            invoice_number += 1

            dst_path = os.path.join(self.target_dir, new_filename)
            if os.path.exists(dst_path):
                base_name, ext = os.path.splitext(new_filename)
                suffix = 1
                while os.path.exists(dst_path):
                    dst_path = os.path.join(
                        self.target_dir, f"{base_name}_{suffix}{ext}"
                    )
                    suffix += 1

            shutil.copy2(src_path, dst_path)
            copied += 1

        QMessageBox.information(self, "Success", f"Copied {copied} file(s).")


def main():
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    app.setFont(QFont("Segoe UI", 12))
    app.setStyleSheet(build_tokyonight_qss(13))
    w = MainWindow()
    w.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
