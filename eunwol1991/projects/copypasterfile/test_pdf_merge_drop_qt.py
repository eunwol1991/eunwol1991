import importlib
import os
import sys
import tempfile
import unittest
from unittest.mock import Mock, patch

if "QT_QPA_PLATFORM" not in os.environ:
    os.environ["QT_QPA_PLATFORM"] = "offscreen"

try:
    from PySide6.QtCore import QUrl
    from PySide6.QtWidgets import QApplication as QtApplication
except ModuleNotFoundError:
    QtApplication = None
    QUrl = None

pdf_merge_drop_qt = importlib.import_module("pdf_merge_drop_qt") if QtApplication else None


def _window():
    qt_application = QtApplication
    if qt_application is None or pdf_merge_drop_qt is None:
        raise RuntimeError("PySide6 is not installed")
    app = qt_application.instance() or qt_application(sys.argv)
    window = pdf_merge_drop_qt.MainWindow()
    window.show()
    app.processEvents()
    return app, window


@unittest.skipIf(QtApplication is None, "PySide6 is not installed")
class ResolveOutputPathTest(unittest.TestCase):
    def test_filename_only_output_saves_beside_first_pdf(self):
        assert pdf_merge_drop_qt is not None
        result = pdf_merge_drop_qt.resolve_output_path(
            "merged.pdf", "/tmp/invoices/first.pdf"
        )

        self.assertEqual(result, os.path.normpath("/tmp/invoices/merged.pdf"))

    def test_path_output_is_used_as_given(self):
        assert pdf_merge_drop_qt is not None
        result = pdf_merge_drop_qt.resolve_output_path(
            "/tmp/out/merged.pdf", "/tmp/invoices/first.pdf"
        )

        self.assertEqual(result, os.path.normpath("/tmp/out/merged.pdf"))

    def test_windows_path_output_is_converted_for_wsl(self):
        assert pdf_merge_drop_qt is not None
        result = pdf_merge_drop_qt.resolve_output_path(
            r"C:\Users\jhunj\Desktop\merged.pdf", "/tmp/invoices/first.pdf"
        )

        if os.name == "nt":
            self.assertEqual(result, os.path.normpath(r"C:\Users\jhunj\Desktop\merged.pdf"))
        else:
            self.assertEqual(result, "/mnt/c/Users/jhunj/Desktop/merged.pdf")


@unittest.skipIf(QtApplication is None, "PySide6 is not installed")
class PdfListTest(unittest.TestCase):
    def test_add_pdf_paths_filters_pdf_case_insensitively_and_preserves_order(self):
        _app, window = _window()

        window.add_pdf_paths(
            [
                "/tmp/second.PDF",
                "/tmp/notes.txt",
                "/tmp/first.pdf",
                "/tmp/third.PdF",
            ]
        )

        self.assertEqual(
            window.pdf_paths,
            ["/tmp/second.PDF", "/tmp/first.pdf", "/tmp/third.PdF"],
        )
        self.assertEqual(
            [window.pdf_list.item(i).text() for i in range(window.pdf_list.count())],
            ["1. second.PDF", "2. first.pdf", "3. third.PdF"],
        )

    def test_add_pdf_paths_converts_windows_paths_for_wsl_display_and_merge(self):
        _app, window = _window()

        window.add_pdf_paths([r"C:\Users\jhunj\Desktop\invoice.pdf"])

        if os.name == "nt":
            self.assertEqual(window.pdf_paths, [r"C:\Users\jhunj\Desktop\invoice.pdf"])
        else:
            self.assertEqual(window.pdf_paths, ["/mnt/c/Users/jhunj/Desktop/invoice.pdf"])
        self.assertEqual(window.pdf_list.item(0).text(), "1. invoice.pdf")

    def test_pdf_list_is_a_visible_drop_target(self):
        _app, window = _window()

        self.assertTrue(window.pdf_list.acceptDrops())

    def test_add_pdf_paths_ignores_cancel_files(self):
        _app, window = _window()

        window.add_pdf_paths(
            [
                "/tmp/customer 0326 - 001 - INV.pdf",
                "/tmp/customer 0326 - 002 - INV (Cancel).pdf",
            ]
        )

        self.assertEqual(window.pdf_paths, ["/tmp/customer 0326 - 001 - INV.pdf"])

    def test_add_pdf_paths_prefers_revised_when_same_normalized_name_exists(self):
        _app, window = _window()

        window.add_pdf_paths(
            [
                "/tmp/customer 0326 - 001 - INV.pdf",
                "/tmp/customer 0326 - 001 - INV (Revised).pdf",
                "/tmp/customer 0326 - 002 - INV.pdf",
            ]
        )

        self.assertEqual(
            window.pdf_paths,
            [
                "/tmp/customer 0326 - 001 - INV (Revised).pdf",
                "/tmp/customer 0326 - 002 - INV.pdf",
            ],
        )
        self.assertEqual(
            [window.pdf_list.item(i).text() for i in range(window.pdf_list.count())],
            ["1. customer 0326 - 001 - INV (Revised).pdf", "2. customer 0326 - 002 - INV.pdf"],
        )


@unittest.skipIf(QtApplication is None, "PySide6 is not installed")
class FolderInputTest(unittest.TestCase):
    def test_normalize_folder_path_converts_windows_folder_for_wsl(self):
        assert pdf_merge_drop_qt is not None

        result = pdf_merge_drop_qt.normalize_folder_path(
            r"C:\Users\jhunj\Dropbox\DO & INV\DO & INV 2026\Supplier - PO to Soup Spoon\6. Jun"
        )

        if os.name == "nt":
            self.assertEqual(
                result,
                os.path.normpath(
                    r"C:\Users\jhunj\Dropbox\DO & INV\DO & INV 2026\Supplier - PO to Soup Spoon\6. Jun"
                ),
            )
        else:
            self.assertEqual(
                result,
                "/mnt/c/Users/jhunj/Dropbox/DO & INV/DO & INV 2026/Supplier - PO to Soup Spoon/6. Jun",
            )

    def test_add_folder_loads_pdf_files_and_applies_cancel_revised_rules(self):
        _app, window = _window()

        with tempfile.TemporaryDirectory() as folder:
            for name in (
                "Customer 0326 - 001 - INV.pdf",
                "Customer 0326 - 001 - INV (Revised).pdf",
                "Customer 0326 - 002 - INV (Cancel).pdf",
                "notes.txt",
            ):
                with open(os.path.join(folder, name), "wb") as f:
                    _ = f.write(b"placeholder")

            window.folder_edit.setText(folder)
            window.add_folder_pdfs()

        self.assertEqual(len(window.pdf_paths), 1)
        self.assertTrue(window.pdf_paths[0].endswith("Customer 0326 - 001 - INV (Revised).pdf"))
        self.assertEqual(window.pdf_list.item(0).text(), "1. Customer 0326 - 001 - INV (Revised).pdf")

    def test_windows_folder_picker_returns_last_non_empty_stdout_line(self):
        assert pdf_merge_drop_qt is not None

        completed = Mock()
        completed.returncode = 0
        completed.stdout = "\r\nC:\\Users\\jhunj\\Dropbox\\DO & INV\r\n"
        completed.stderr = ""

        with patch("pdf_merge_drop_qt.subprocess.run", return_value=completed):
            result = pdf_merge_drop_qt.open_windows_folder_dialog("/mnt/c/Users/jhunj")

        self.assertEqual(result, r"C:\Users\jhunj\Dropbox\DO & INV")

    def test_browse_folder_uses_windows_picker_before_qt_fallback(self):
        _app, window = _window()

        with tempfile.TemporaryDirectory() as folder:
            pdf_path = os.path.join(folder, "Customer 0326 - 001 - INV.pdf")
            with open(pdf_path, "wb") as f:
                _ = f.write(b"placeholder")

            with patch("pdf_merge_drop_qt.open_windows_folder_dialog", return_value=folder):
                with patch("pdf_merge_drop_qt.QFileDialog.getExistingDirectory") as qt_dialog:
                    window.browse_folder()

        qt_dialog.assert_not_called()
        self.assertEqual(len(window.pdf_paths), 1)
        self.assertTrue(window.folder_edit.text().endswith(folder))

    def test_browse_folder_falls_back_to_qt_dialog_when_windows_picker_returns_empty(self):
        _app, window = _window()

        with tempfile.TemporaryDirectory() as folder:
            pdf_path = os.path.join(folder, "Customer 0326 - 001 - INV.pdf")
            with open(pdf_path, "wb") as f:
                _ = f.write(b"placeholder")

            with patch("pdf_merge_drop_qt.open_windows_folder_dialog", return_value=""):
                with patch("pdf_merge_drop_qt.QFileDialog.getExistingDirectory", return_value=folder):
                    window.browse_folder()

        self.assertEqual(len(window.pdf_paths), 1)
        self.assertTrue(window.folder_edit.text().endswith(folder))


@unittest.skipIf(QtApplication is None, "PySide6 is not installed")
class DropEventTest(unittest.TestCase):
    def test_drop_event_adds_local_file_urls_only(self):
        if QUrl is None:
            self.skipTest("PySide6 is not installed")
        qurl = QUrl
        _app, window = _window()

        class MimeData:
            def hasUrls(self):
                return True

            def urls(self):
                return [
                    qurl.fromLocalFile("/tmp/local-one.pdf"),
                    qurl("https://example.test/remote.pdf"),
                    qurl.fromLocalFile("/tmp/local-two.PDF"),
                ]

        class DropEvent:
            def __init__(self):
                self.accepted = False

            def mimeData(self):
                return MimeData()

            def acceptProposedAction(self):
                self.accepted = True

        event = DropEvent()

        window.dropEvent(event)

        self.assertTrue(event.accepted)
        self.assertEqual(window.pdf_paths, ["/tmp/local-one.pdf", "/tmp/local-two.PDF"])


@unittest.skipIf(QtApplication is None, "PySide6 is not installed")
class MergeFlowTest(unittest.TestCase):
    def test_merge_warns_when_no_pdfs(self):
        _app, window = _window()
        window.output_edit.setText("merged.pdf")

        with patch("pdf_merge_drop_qt.QMessageBox.warning") as warning:
            window.merge_pdfs()

        warning.assert_called_once()

    def test_merge_warns_when_output_name_is_empty(self):
        _app, window = _window()
        window.add_pdf_paths(["/tmp/first.pdf"])
        window.output_edit.setText("   ")

        with patch("pdf_merge_drop_qt.QMessageBox.warning") as warning:
            window.merge_pdfs()

        warning.assert_called_once()

    def test_merge_appends_paths_in_list_order_and_writes_resolved_output(self):
        _app, window = _window()
        window.add_pdf_paths(["/tmp/z-last.pdf", "/tmp/a-first.pdf", "/tmp/middle.pdf"])
        window.output_edit.setText("merged.pdf")

        merger = Mock()
        with patch("pdf_merge_drop_qt.PdfMerger", return_value=merger) as merger_class:
            with patch("pdf_merge_drop_qt.QMessageBox.information") as information:
                window.merge_pdfs()

        merger_class.assert_called_once_with()
        self.assertEqual(
            [call.args[0] for call in merger.append.call_args_list],
            ["/tmp/z-last.pdf", "/tmp/a-first.pdf", "/tmp/middle.pdf"],
        )
        merger.write.assert_called_once_with(os.path.normpath("/tmp/merged.pdf"))
        merger.close.assert_called_once_with()
        information.assert_called_once()


if __name__ == "__main__":
    unittest.main()
