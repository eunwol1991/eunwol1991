import os
import re
import subprocess
import sys
from difflib import SequenceMatcher

try:
    from PySide6.QtCore import Qt, QTimer, QStringListModel
    from PySide6.QtGui import QFont
    from PySide6.QtWidgets import (
        QApplication,
        QCompleter,
        QFileDialog,
        QFormLayout,
        QHBoxLayout,
        QLabel,
        QLineEdit,
        QMainWindow,
        QMessageBox,
        QProgressBar,
        QPushButton,
        QSlider,
        QTableWidget,
        QTableWidgetItem,
        QVBoxLayout,
        QWidget,
    )
except ModuleNotFoundError as exc:
    raise SystemExit(
        "PySide6 is required. Install with: python -m pip install PySide6"
    ) from exc

from .service import apply_insert, load_context, preview_insert, suggest


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
QTableWidget {
    background-color: #24283b;
    color: #c0caf5;
    border: 1px solid #414868;
    border-radius: 8px;
    padding: 8px 10px;
    selection-background-color: #33467c;
    min-height: 22px;
}

QLineEdit:focus,
QTableWidget:focus {
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


def _windows_path_to_wsl(path: str) -> str:
    text = (path or "").strip()
    m = re.match(r"^([A-Za-z]):\\(.*)$", text)
    if not m:
        return text
    drive = m.group(1).lower()
    rest = m.group(2).replace("\\", "/")
    return f"/mnt/{drive}/{rest}"


def _wsl_path_to_windows(path: str) -> str:
    text = (path or "").strip()
    m = re.match(r"^/mnt/([a-zA-Z])/(.*)$", text)
    if not m:
        return text
    drive = m.group(1).upper()
    rest = m.group(2).replace("/", "\\")
    return f"{drive}:\\{rest}"


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


def _open_windows_file_dialog(initial_dir: str) -> str:
    start = _wsl_path_to_windows(initial_dir)
    ps_script = (
        "Add-Type -AssemblyName System.Windows.Forms | Out-Null;"
        "$dlg = New-Object System.Windows.Forms.OpenFileDialog;"
        "$dlg.Filter = 'Excel files (*.xlsx;*.xlsm)|*.xlsx;*.xlsm|All files (*.*)|*.*';"
        f"$dlg.InitialDirectory = '{start.replace("'", "''")}';"
        "if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { $dlg.FileName }"
    )
    try:
        proc = subprocess.run(
            ["powershell.exe", "-NoProfile", "-Command", ps_script],
            capture_output=True,
            text=True,
            timeout=120,
        )
    except Exception:
        return ""
    if proc.returncode != 0:
        return ""
    selected = (proc.stdout or "").strip().splitlines()
    if not selected:
        return ""
    return _windows_path_to_wsl(selected[-1].strip())


def _normalize_text(value: str) -> str:
    return re.sub(r"\s+", " ", str(value or "").strip().lower())


def _score(query: str, value: str) -> float:
    nq = _normalize_text(query)
    nv = _normalize_text(value)
    if not nq:
        return 1.0
    if not nv:
        return 0.0
    if nq in nv or nv in nq:
        return 0.95
    return SequenceMatcher(None, nq, nv).ratio()


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Delivery Details Entry Assistant (Qt)")
        self.resize(1080, 720)

        self.context = None
        self.last_plan = None
        self.base_font = 12
        self._live_timer = QTimer(self)
        self._live_timer.setSingleShot(True)
        self._live_timer.timeout.connect(self._refresh_live)
        self._suspend_live = False
        dropbox_base = _default_dropbox_base()
        source_default = f"{dropbox_base}/DO & INV/DO & INV 2026"
        self.default_browse_root = _first_existing_path(
            [
                source_default,
                os.getcwd(),
                _platform_drive_root(),
            ]
        )

        root = QWidget(self)
        self.setCentralWidget(root)
        v = QVBoxLayout(root)
        v.setContentsMargins(18, 18, 18, 18)
        v.setSpacing(12)
        self.root_layout = v

        top = QHBoxLayout()
        self.top_layout = top
        self.title_label = QLabel("Delivery Details Entry Assistant")
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

        file_row = QHBoxLayout()
        self.file_row_layout = file_row
        file_row.addWidget(QLabel("Excel:"))
        self.file_edit = QLineEdit()
        self.file_edit.setPlaceholderText("Select .xlsx/.xlsm file")
        file_row.addWidget(self.file_edit)
        btn_browse = QPushButton("Browse")
        btn_browse.clicked.connect(self.browse_file)
        file_row.addWidget(btn_browse)
        self.btn_browse = btn_browse
        v.addLayout(file_row)

        form = QFormLayout()
        self.form_layout = form
        self.desc_edit = QLineEdit()
        self.product_code_edit = QLineEdit()
        self.customer_edit = QLineEdit()
        self.outlet_edit = QLineEdit()
        self.qty_pcs_edit = QLineEdit()
        self.qty_ctns_edit = QLineEdit()
        self.invoice_edit = QLineEdit()
        form.addRow("Description", self.desc_edit)
        form.addRow("Product Code", self.product_code_edit)
        form.addRow("Customer", self.customer_edit)
        form.addRow("Outlet", self.outlet_edit)
        form.addRow("Qty in Pcs", self.qty_pcs_edit)
        form.addRow("Qty in Ctns", self.qty_ctns_edit)
        form.addRow("Invoice #", self.invoice_edit)
        v.addLayout(form)

        self.desc_model = QStringListModel(self)
        self.code_model = QStringListModel(self)
        self.customer_model = QStringListModel(self)
        self.outlet_model = QStringListModel(self)
        self.desc_completer = QCompleter(self.desc_model, self)
        self.code_completer = QCompleter(self.code_model, self)
        self.customer_completer = QCompleter(self.customer_model, self)
        self.outlet_completer = QCompleter(self.outlet_model, self)
        for c in [
            self.desc_completer,
            self.code_completer,
            self.customer_completer,
            self.outlet_completer,
        ]:
            c.setCaseSensitivity(Qt.CaseSensitivity.CaseInsensitive)
            c.setFilterMode(Qt.MatchFlag.MatchContains)
            c.setCompletionMode(QCompleter.CompletionMode.PopupCompletion)
        self.desc_edit.setCompleter(self.desc_completer)
        self.product_code_edit.setCompleter(self.code_completer)
        self.customer_edit.setCompleter(self.customer_completer)
        self.outlet_edit.setCompleter(self.outlet_completer)
        self.desc_edit.textChanged.connect(self._schedule_live_refresh)
        self.product_code_edit.textChanged.connect(self._schedule_live_refresh)
        self.customer_edit.textChanged.connect(self._schedule_live_refresh)
        self.outlet_edit.textChanged.connect(self._schedule_live_refresh)
        self.desc_completer.activated.connect(self._on_desc_selected)
        self.code_completer.activated.connect(self._on_code_selected)
        self.customer_completer.activated.connect(self._on_customer_selected)
        self.outlet_completer.activated.connect(self._on_outlet_selected)

        actions = QHBoxLayout()
        self.actions_layout = actions
        self.action_buttons = []
        for text, cb in [
            ("Suggest", self.on_suggest),
            ("Preview", self.on_preview),
            ("Insert", self.on_insert),
        ]:
            b = QPushButton(text)
            b.clicked.connect(cb)
            actions.addWidget(b)
            self.action_buttons.append(b)
        actions.addStretch(1)
        v.addLayout(actions)

        self.table = QTableWidget(0, 4)
        self.table.setHorizontalHeaderLabels(
            ["Matched Date", "Customer", "Outlet", "Score"]
        )
        self.table.horizontalHeader().setStretchLastSection(True)
        v.addWidget(self.table)

        self.status = QLabel("")
        v.addWidget(self.status)

        self.progress = QProgressBar()
        self.progress.setVisible(False)
        self.progress.setTextVisible(False)
        v.addWidget(self.progress)

        self.apply_scale(140)

    def apply_scale(self, percent: int):
        factor = max(0.8, float(percent) / 100.0)
        app = QApplication.instance()
        if isinstance(app, QApplication):
            app.setFont(QFont("Segoe UI", max(8, int(round(self.base_font * factor)))))
            app.setStyleSheet(build_tokyonight_qss(int(round(13 * factor))))
        margin = max(10, int(round(18 * factor)))
        root_spacing = max(8, int(round(12 * factor)))
        row_spacing = max(6, int(round(8 * factor)))
        self.root_layout.setContentsMargins(margin, margin, margin, margin)
        self.root_layout.setSpacing(root_spacing)
        self.top_layout.setSpacing(row_spacing)
        self.file_row_layout.setSpacing(row_spacing)
        self.actions_layout.setSpacing(row_spacing)
        self.form_layout.setHorizontalSpacing(max(8, int(round(10 * factor))))
        self.form_layout.setVerticalSpacing(max(6, int(round(8 * factor))))
        input_h = max(28, int(round(32 * factor)))
        button_h = max(30, int(round(34 * factor)))
        for w in [
            self.file_edit,
            self.desc_edit,
            self.product_code_edit,
            self.customer_edit,
            self.outlet_edit,
            self.qty_pcs_edit,
            self.qty_ctns_edit,
            self.invoice_edit,
        ]:
            w.setMinimumHeight(input_h)
        self.btn_browse.setMinimumHeight(button_h)
        for b in self.action_buttons:
            b.setMinimumHeight(button_h)
        self.table.verticalHeader().setDefaultSectionSize(
            max(24, int(round(30 * factor)))
        )
        self.title_label.setStyleSheet(
            f"font-weight: 700; font-size: {max(12, int(round(18 * factor)))}px;"
        )
        self.scale_label.setText(f"UI Scale: {int(percent)}%")

    def browse_file(self):
        path = ""
        if _is_wsl():
            path = _open_windows_file_dialog(self.default_browse_root)
        if not path:
            path, _ = QFileDialog.getOpenFileName(
                self,
                "Select Excel",
                self.default_browse_root,
                "Excel Files (*.xlsx *.xlsm);;All Files (*)",
            )
        if path:
            self.file_edit.setText(path)
            self.context = None
            self.last_plan = None
            self._refresh_live(load_if_needed=True)

    def _set_busy(self, is_busy: bool, note: str = ""):
        if is_busy:
            self.progress.setVisible(True)
            self.progress.setRange(0, 0)
            self.status.setText(note or "Loading workbook...")
        else:
            self.progress.setRange(0, 1)
            self.progress.setValue(0)
            self.progress.setVisible(False)
        QApplication.processEvents()

    def ensure_context(self) -> bool:
        path = self.file_edit.text().strip()
        if _is_wsl():
            path = _windows_path_to_wsl(path)
        if not path:
            QMessageBox.warning(
                self, "Missing file", "Please select an Excel file first."
            )
            return False
        if os.path.isdir(path):
            QMessageBox.warning(
                self,
                "Missing file",
                "Current path is a folder. Please browse and select an Excel file.",
            )
            return False
        if self.context and self.context.get("file_path") == path:
            return True
        try:
            self._set_busy(True, "Reading workbook data...")
            self.context = load_context(path)
            self.last_plan = None
            self._rebuild_completer_models()
            return True
        except Exception as exc:
            self.context = None
            QMessageBox.critical(self, "Load failed", str(exc))
            return False
        finally:
            self._set_busy(False)

    def _top_values(
        self, records: list[dict], key: str, query: str, limit: int = 20
    ) -> list[str]:
        values = sorted(
            {
                str(r.get(key) or "").strip()
                for r in records
                if str(r.get(key) or "").strip()
            }
        )
        if not query:
            return values[:limit]
        ranked = sorted(values, key=lambda v: _score(query, v), reverse=True)
        return ranked[:limit]

    def _rebuild_completer_models(self):
        ctx = self.context
        if ctx is None:
            return
        records = ctx.get("records", [])
        d_query = self.desc_edit.text().strip()
        p_query = self.product_code_edit.text().strip()
        c_query = self.customer_edit.text().strip()
        o_query = self.outlet_edit.text().strip()

        top_desc = self._top_values(records, "description", d_query, limit=24)
        filtered = records
        if d_query:
            top_desc_set = set(top_desc)
            filtered = [
                r for r in records if str(r.get("description") or "") in top_desc_set
            ]

        top_code = self._top_values(filtered, "product_code", p_query, limit=24)
        if p_query:
            top_code_set = set(top_code)
            filtered = [
                r for r in filtered if str(r.get("product_code") or "") in top_code_set
            ]

        top_customer = self._top_values(filtered, "customer", c_query, limit=24)
        if c_query:
            top_customer_set = set(top_customer)
            filtered = [
                r for r in filtered if str(r.get("customer") or "") in top_customer_set
            ]

        top_outlet = self._top_values(filtered, "outlet", o_query, limit=24)

        self.desc_model.setStringList(top_desc)
        self.code_model.setStringList(top_code)
        self.customer_model.setStringList(top_customer)
        self.outlet_model.setStringList(top_outlet)

    def _schedule_live_refresh(self):
        if self._suspend_live:
            return
        self._live_timer.start(180)

    def _refresh_live(self, load_if_needed: bool = False):
        if not self.file_edit.text().strip():
            return
        if load_if_needed and not self.ensure_context():
            return
        if self.context is None:
            return
        self._rebuild_completer_models()
        self.on_suggest()

    def _on_desc_selected(self, value: str):
        self._suspend_live = True
        self.desc_edit.setText(value)
        self._suspend_live = False
        self._refresh_live()

    def _on_customer_selected(self, value: str):
        self._suspend_live = True
        self.customer_edit.setText(value)
        self._suspend_live = False
        self._refresh_live()

    def _on_code_selected(self, value: str):
        self._suspend_live = True
        self.product_code_edit.setText(value)
        self._suspend_live = False
        self._refresh_live()

    def _on_outlet_selected(self, value: str):
        self._suspend_live = True
        self.outlet_edit.setText(value)
        self._suspend_live = False
        self._refresh_live()

    def on_suggest(self):
        if not self.ensure_context():
            return
        context = self.context
        if context is None:
            return
        ranked = suggest(
            context,
            {
                "description": self.desc_edit.text().strip(),
                "product_code": self.product_code_edit.text().strip(),
                "customer": self.customer_edit.text().strip(),
                "outlet": self.outlet_edit.text().strip(),
            },
            limit=8,
        )
        self.table.setRowCount(0)
        for rec in ranked:
            row = self.table.rowCount()
            self.table.insertRow(row)
            data = rec["record"]
            rec_date = data.get("record_date")
            date_text = rec_date.strftime("%d/%m/%Y") if rec_date else ""
            self.table.setItem(row, 0, QTableWidgetItem(date_text))
            self.table.setItem(row, 1, QTableWidgetItem(str(data.get("customer", ""))))
            self.table.setItem(row, 2, QTableWidgetItem(str(data.get("outlet", ""))))
            self.table.setItem(row, 3, QTableWidgetItem(f"{rec['score']:.1f}"))
        self.status.setText(f"Suggestions: {len(ranked)}")

    def on_preview(self):
        if not self.ensure_context():
            return
        context = self.context
        if context is None:
            return
        try:
            self.last_plan = preview_insert(
                context,
                {
                    "description": self.desc_edit.text().strip(),
                    "product_code": self.product_code_edit.text().strip(),
                    "customer": self.customer_edit.text().strip(),
                    "outlet": self.outlet_edit.text().strip(),
                    "qty_pcs": int((self.qty_pcs_edit.text() or "0").strip()),
                    "qty_ctns": int((self.qty_ctns_edit.text() or "0").strip()),
                    "invoice": self.invoice_edit.text().strip(),
                },
            )
        except ValueError:
            QMessageBox.warning(
                self, "Invalid input", "Qty in Pcs / Qty in Ctns must be integers."
            )
            return

        lines = [f"Anchor row: {self.last_plan.get('anchor_row')}"]
        lines.append(f"Source row: {self.last_plan.get('source_row')}")
        lines.append(f"Insert row: {self.last_plan['insert_row']}")
        for col_idx, value in sorted(self.last_plan["user_values"].items()):
            lines.append(f"Column {col_idx}: {value}")
        QMessageBox.information(self, "Preview", "\n".join(lines))

    def on_insert(self):
        if self.last_plan is None:
            self.on_preview()
            if self.last_plan is None:
                return
        context = self.context
        if context is None:
            QMessageBox.warning(self, "Missing context", "Please load workbook first.")
            return
        if (
            QMessageBox.question(self, "Confirm", "Insert this row into workbook?")
            != QMessageBox.StandardButton.Yes
        ):
            return
        try:
            backup_path = apply_insert(context, self.last_plan)
        except Exception as exc:
            QMessageBox.critical(self, "Insert failed", str(exc))
            return
        QMessageBox.information(self, "Done", f"Row inserted. Backup:\n{backup_path}")


def launch():
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    app.setFont(QFont("Segoe UI", 12))
    app.setStyleSheet(build_tokyonight_qss(13))
    w = MainWindow()
    w.show()
    sys.exit(app.exec())
