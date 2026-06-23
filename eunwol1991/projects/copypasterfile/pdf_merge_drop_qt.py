import os
import re
import subprocess
import sys
from typing import override

from PyPDF2 import PdfMerger

try:
    from PySide6.QtGui import QDragEnterEvent, QDropEvent, QFont
    from PySide6.QtWidgets import (
        QApplication,
        QHBoxLayout,
        QFileDialog,
        QLabel,
        QLineEdit,
        QListWidget,
        QListWidgetItem,
        QMainWindow,
        QMessageBox,
        QPushButton,
        QVBoxLayout,
        QWidget,
    )
except ModuleNotFoundError as exc:
    raise SystemExit(
        "PySide6 is required. Install with: python -m pip install PySide6"
    ) from exc


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
QListWidget:focus {
    border: 1px solid #7aa2f7;
}

QListWidget#DropList {
    border: 1px dashed #7aa2f7;
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


def windows_path_to_platform(path: str) -> str:
    match = re.match(r"^/?([A-Za-z]):[\\/](.*)$", path)
    if match is None:
        return path
    if os.name == "nt":
        return path
    drive = match.group(1).lower()
    tail = match.group(2).replace("\\", "/")
    return f"/mnt/{drive}/{tail}"


def normalize_pdf_path(path: str) -> str:
    return os.path.normpath(windows_path_to_platform((path or "").strip()))


def normalize_folder_path(path: str) -> str:
    return os.path.normpath(windows_path_to_platform((path or "").strip()))


def platform_path_to_windows(path: str) -> str:
    normalized = os.path.normpath((path or "").strip())
    match = re.match(r"^/mnt/([A-Za-z])/(.*)$", normalized)
    if match is None:
        return normalized
    drive = match.group(1).upper()
    tail = match.group(2).replace("/", "\\")
    return f"{drive}:\\{tail}"


def escape_powershell_single_quoted(text: str) -> str:
    return text.replace("'", "''")


def open_windows_folder_dialog(initial_dir: str) -> str:
    start = platform_path_to_windows(initial_dir) if initial_dir else ""
    start_literal = escape_powershell_single_quoted(start)
    script = f"""
Add-Type -AssemblyName System.Windows.Forms | Out-Null
$dlg = New-Object System.Windows.Forms.FolderBrowserDialog
$dlg.Description = 'Select PDF folder'
if ('{start_literal}' -and [System.IO.Directory]::Exists('{start_literal}')) {{
    $dlg.SelectedPath = '{start_literal}'
}}
if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {{
    $dlg.SelectedPath
}}
"""
    try:
        completed = subprocess.run(
            ["powershell.exe", "-NoProfile", "-STA", "-Command", script],
            check=False,
            capture_output=True,
            text=True,
            timeout=120,
        )
    except (OSError, subprocess.SubprocessError):
        return ""
    if completed.returncode != 0:
        return ""
    lines = [line.strip() for line in completed.stdout.splitlines() if line.strip()]
    return lines[-1] if lines else ""


def pdf_group_key(path: str) -> str:
    name = os.path.basename(path).lower()
    stem, ext = os.path.splitext(name)
    stem = re.sub(r"\s*\(revised\)\s*", " ", stem, flags=re.IGNORECASE)
    stem = " ".join(stem.split())
    return f"{stem}{ext}"


def is_cancel_pdf(path: str) -> bool:
    return "(cancel)" in os.path.basename(path).lower()


def prefer_revised_pdfs(paths: list[str]) -> list[str]:
    groups: dict[str, list[str]] = {}
    order: list[str] = []
    for path in paths:
        if is_cancel_pdf(path):
            continue
        key = pdf_group_key(path)
        if key not in groups:
            groups[key] = []
            order.append(key)
        groups[key].append(path)

    result: list[str] = []
    for key in order:
        candidates = groups[key]
        revised = [p for p in candidates if "(revised)" in os.path.basename(p).lower()]
        result.extend(revised if revised else candidates)
    return result


def resolve_output_path(output_text: str, first_pdf_path: str) -> str:
    output_path = (output_text or "").strip()
    output_path = windows_path_to_platform(output_path)
    if os.path.dirname(output_path):
        return os.path.normpath(output_path)
    first_dir = os.path.dirname(os.path.abspath(first_pdf_path))
    return os.path.normpath(os.path.join(first_dir, output_path))


class DropPdfList(QListWidget):
    def __init__(self, owner: "MainWindow"):
        super().__init__()
        self.owner = owner
        self.setAcceptDrops(True)

    @override
    def dragEnterEvent(self, event: QDragEnterEvent):
        self.owner.handle_drag_enter(event)

    @override
    def dropEvent(self, event: QDropEvent):
        self.owner.handle_drop(event)


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Drop PDF Merger")
        self.resize(760, 520)
        self.setAcceptDrops(True)
        self.pdf_paths: list[str] = []

        root = QWidget(self)
        self.setCentralWidget(root)
        layout = QVBoxLayout(root)
        layout.setContentsMargins(18, 18, 18, 18)
        layout.setSpacing(12)

        self.title_label: QLabel = QLabel("Drop PDFs, Keep Their Order")
        self.title_label.setStyleSheet("font-weight: 700; font-size: 20px;")
        layout.addWidget(self.title_label)

        guide = QLabel("Drag local PDF files here. The list below is the merge order.")
        guide.setStyleSheet("color: #a9b1d6;")
        layout.addWidget(guide)

        folder_row = QHBoxLayout()
        folder_row.addWidget(QLabel("Folder:"))
        self.folder_edit: QLineEdit = QLineEdit()
        self.folder_edit.setPlaceholderText(
            r"Paste Windows folder path, e.g. C:\Users\...\6. Jun"
        )
        folder_row.addWidget(self.folder_edit)
        self.add_folder_button: QPushButton = QPushButton("Add Folder PDFs")
        _ = self.add_folder_button.clicked.connect(self.add_folder_pdfs)
        folder_row.addWidget(self.add_folder_button)
        self.browse_folder_button: QPushButton = QPushButton("Browse Folder")
        _ = self.browse_folder_button.clicked.connect(self.browse_folder)
        folder_row.addWidget(self.browse_folder_button)
        layout.addLayout(folder_row)

        self.pdf_list: DropPdfList = DropPdfList(self)
        self.pdf_list.setObjectName("DropList")
        self.pdf_list.setAlternatingRowColors(False)
        self.pdf_list.setMinimumHeight(260)
        layout.addWidget(self.pdf_list)

        output_row = QHBoxLayout()
        output_row.addWidget(QLabel("Output:"))
        self.output_edit: QLineEdit = QLineEdit()
        self.output_edit.setPlaceholderText("merged.pdf or /path/to/merged.pdf")
        output_row.addWidget(self.output_edit)
        layout.addLayout(output_row)

        action_row = QHBoxLayout()
        self.status: QLabel = QLabel("Waiting for PDF files.")
        self.status.setStyleSheet("color: #a9b1d6;")
        action_row.addWidget(self.status)
        action_row.addStretch(1)
        self.merge_button: QPushButton = QPushButton("Merge PDFs")
        _ = self.merge_button.clicked.connect(self.merge_pdfs)
        action_row.addWidget(self.merge_button)
        layout.addLayout(action_row)

    def add_pdf_paths(self, paths: list[str]):
        for path in paths:
            path = normalize_pdf_path(path)
            if os.path.splitext(path)[1].lower() != ".pdf":
                continue
            self.pdf_paths.append(path)
        self.pdf_paths = prefer_revised_pdfs(self.pdf_paths)
        self.refresh_pdf_list()

    def refresh_pdf_list(self):
        self.pdf_list.clear()
        for index, path in enumerate(self.pdf_paths, start=1):
            self.pdf_list.addItem(QListWidgetItem(f"{index}. {os.path.basename(path)}"))
        count = len(self.pdf_paths)
        if count:
            self.status.setText(f"Ready to merge {count} PDF file(s).")
        else:
            self.status.setText("Waiting for PDF files.")

    def add_folder_pdfs(self):
        folder = normalize_folder_path(self.folder_edit.text())
        if not folder or not os.path.isdir(folder):
            _ = QMessageBox.warning(self, "Warning", "Please enter a valid folder path.")
            return
        pdf_paths = [
            os.path.join(folder, name)
            for name in os.listdir(folder)
            if os.path.isfile(os.path.join(folder, name))
            and os.path.splitext(name)[1].lower() == ".pdf"
        ]
        self.add_pdf_paths(pdf_paths)

    def browse_folder(self):
        start = normalize_folder_path(self.folder_edit.text())
        if not start or not os.path.isdir(start):
            start = os.getcwd()
        folder = open_windows_folder_dialog(start)
        if not folder:
            folder = QFileDialog.getExistingDirectory(self, "Select PDF folder", start)
        if not folder:
            return
        self.folder_edit.setText(folder)
        self.add_folder_pdfs()

    def handle_drag_enter(self, event: QDragEnterEvent):
        mime_data = event.mimeData()
        if not mime_data.hasUrls():
            event.ignore()
            return
        if any(url.isLocalFile() for url in mime_data.urls()):
            event.acceptProposedAction()
        else:
            event.ignore()

    def handle_drop(self, event: QDropEvent):
        mime_data = event.mimeData()
        if not mime_data.hasUrls():
            event.ignore()
            return
        paths = [url.toLocalFile() for url in mime_data.urls() if url.isLocalFile()]
        self.add_pdf_paths(paths)
        event.acceptProposedAction()

    @override
    def dragEnterEvent(self, event: QDragEnterEvent):
        self.handle_drag_enter(event)

    @override
    def dropEvent(self, event: QDropEvent):
        self.handle_drop(event)

    def merge_pdfs(self):
        if not self.pdf_paths:
            _ = QMessageBox.warning(self, "Warning", "Please drop PDF files first.")
            return

        output_text = self.output_edit.text().strip()
        if not output_text:
            _ = QMessageBox.warning(self, "Warning", "Please enter an output PDF name.")
            return

        output_path = resolve_output_path(output_text, self.pdf_paths[0])
        output_dir = os.path.dirname(output_path)

        merger = PdfMerger()
        try:
            for pdf_path in self.pdf_paths:
                merger.append(pdf_path)
            if output_dir:
                os.makedirs(output_dir, exist_ok=True)
            merger.write(output_path)
        except Exception as exc:
            _ = QMessageBox.critical(self, "Error", f"Could not merge PDFs: {exc}")
            return
        finally:
            merger.close()

        self.status.setText(f"Merged {len(self.pdf_paths)} PDF file(s) to {output_path}")
        _ = QMessageBox.information(self, "Success", f"Merged PDF saved to:\n{output_path}")


def main():
    app = QApplication(sys.argv)
    _ = app.setStyle("Fusion")
    app.setFont(QFont("Segoe UI", 12))
    app.setStyleSheet(TOKYONIGHT_QSS)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
