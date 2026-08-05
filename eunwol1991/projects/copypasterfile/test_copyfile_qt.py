import importlib
import os
import sys
import tempfile
import unittest
from unittest.mock import patch

if "QT_QPA_PLATFORM" not in os.environ:
    os.environ["QT_QPA_PLATFORM"] = "offscreen"

try:
    from PySide6.QtCore import Qt
    from PySide6.QtTest import QTest
    from PySide6.QtWidgets import QApplication as QtApplication
except ModuleNotFoundError:
    QtApplication = None
    Qt = None
    QTest = None

copyfile_qt = importlib.import_module("copyfile_qt") if QtApplication else None


def _window():
    qt_application = QtApplication
    if qt_application is None or copyfile_qt is None:
        raise RuntimeError("PySide6 is not installed")
    app = qt_application.instance() or qt_application(sys.argv)
    window = copyfile_qt.MainWindow()
    window.file_list.setCurrentRow(0)
    window.show()
    app.processEvents()
    return app, window


@unittest.skipIf(QtApplication is None, "PySide6 is not installed")
class DuplicateSelectionTest(unittest.TestCase):
    def setUp(self):
        scan_patch = patch(
            "copyfile_qt.scan_source_files",
            return_value=(
                [{"display_name": "First File", "file_path": "/tmp/first.txt"}],
                0,
            ),
        )
        _ = scan_patch.start()
        self.addCleanup(scan_patch.stop)

    def test_checked_allow_duplicates_appends_same_real_index(self):
        _app, window = _window()

        window.allow_duplicates_checkbox.setChecked(True)

        window.add_selected()
        window.add_selected()

        self.assertEqual(window.selected_indices, [0, 0])

    def test_unchecked_allow_duplicates_preserves_single_selection(self):
        _app, window = _window()

        self.assertFalse(window.allow_duplicates_checkbox.isChecked())

        window.add_selected()
        window.add_selected()

        self.assertEqual(window.selected_indices, [0])

    def test_refresh_source_focuses_filter_edit(self):
        _app, window = _window()
        qt_application = QtApplication
        assert qt_application is not None

        window.file_list.setFocus()
        qt_application.processEvents()

        window.refresh_source()

        self.assertTrue(window.filter_edit.hasFocus())

    def test_tab_on_single_filtered_result_adds_selection(self):
        _app, window = _window()
        qt_application = QtApplication
        qtest = QTest
        qt = Qt
        assert qt_application is not None
        assert qtest is not None
        assert qt is not None

        window.filter_edit.setFocus()
        qt_application.processEvents()

        qtest.keyClick(window.filter_edit, qt.Key.Key_Tab)

        self.assertEqual(window.selected_indices, [0])

    def test_tab_on_single_filtered_result_reuses_duplicate_setting(self):
        _app, window = _window()
        qt_application = QtApplication
        qtest = QTest
        qt = Qt
        assert qt_application is not None
        assert qtest is not None
        assert qt is not None

        window.allow_duplicates_checkbox.setChecked(True)
        window.filter_edit.setFocus()
        qt_application.processEvents()

        qtest.keyClick(window.filter_edit, qt.Key.Key_Tab)
        qtest.keyClick(window.filter_edit, qt.Key.Key_Tab)

        self.assertEqual(window.selected_indices, [0, 0])


@unittest.skipIf(QtApplication is None, "PySide6 is not installed")
class MultiResultTabFlowTest(unittest.TestCase):
    def setUp(self):
        scan_patch = patch(
            "copyfile_qt.scan_source_files",
            return_value=(
                [
                    {"display_name": "First File", "file_path": "/tmp/first.txt"},
                    {"display_name": "Second File", "file_path": "/tmp/second.txt"},
                ],
                0,
            ),
        )
        _ = scan_patch.start()
        self.addCleanup(scan_patch.stop)

    def test_tab_on_filter_with_multiple_results_focuses_file_list_and_selects_first_row(self):
        _app, window = _window()
        qt_application = QtApplication
        qtest = QTest
        qt = Qt
        assert qt_application is not None
        assert qtest is not None
        assert qt is not None

        window.file_list.setCurrentRow(-1)
        window.filter_edit.setFocus()
        qt_application.processEvents()

        qtest.keyClick(window.filter_edit, qt.Key.Key_Tab)

        self.assertTrue(window.file_list.hasFocus())
        self.assertEqual(window.file_list.currentRow(), 0)
        self.assertEqual(window.selected_indices, [])

    def test_tab_on_file_list_with_current_row_adds_selected_file(self):
        _app, window = _window()
        qt_application = QtApplication
        qtest = QTest
        qt = Qt
        assert qt_application is not None
        assert qtest is not None
        assert qt is not None

        window.file_list.setCurrentRow(1)
        window.file_list.setFocus()
        qt_application.processEvents()

        qtest.keyClick(window.file_list, qt.Key.Key_Tab)

        self.assertEqual(window.selected_indices, [1])


@unittest.skipIf(QtApplication is None, "PySide6 is not installed")
class ShortcutKeyFlowTest(unittest.TestCase):
    def setUp(self):
        scan_patch = patch(
            "copyfile_qt.scan_source_files",
            return_value=(
                [{"display_name": "First File", "file_path": "/tmp/first.txt"}],
                0,
            ),
        )
        _ = scan_patch.start()
        self.addCleanup(scan_patch.stop)

    def test_escape_from_selected_list_returns_focus_to_filter(self):
        _app, window = _window()
        qt_application = QtApplication
        qtest = QTest
        qt = Qt
        assert qt_application is not None
        assert qtest is not None
        assert qt is not None

        window.selected_list.setFocus()
        qt_application.processEvents()

        qtest.keyClick(window.selected_list, qt.Key.Key_Escape)

        self.assertTrue(window.filter_edit.hasFocus())

    def test_enter_from_filter_runs_copy_files(self):
        _app, window = _window()
        qt_application = QtApplication
        qtest = QTest
        qt = Qt
        assert qt_application is not None
        assert qtest is not None
        assert qt is not None

        with patch.object(window, "copy_files") as copy_files:
            window.filter_edit.setFocus()
            qt_application.processEvents()

            qtest.keyClick(window.filter_edit, qt.Key.Key_Return)

        copy_files.assert_called_once_with()

    def test_ctrl_b_from_filter_opens_browse_source(self):
        _app, window = _window()
        qt_application = QtApplication
        qtest = QTest
        qt = Qt
        assert qt_application is not None
        assert qtest is not None
        assert qt is not None

        with patch.object(window, "browse_source") as browse_source:
            window.filter_edit.setFocus()
            qt_application.processEvents()

            qtest.keyClick(
                window.filter_edit,
                qt.Key.Key_B,
                qt.KeyboardModifier.ControlModifier,
            )

        browse_source.assert_called_once_with()


@unittest.skipIf(QtApplication is None, "PySide6 is not installed")
class InvoiceNumberTest(unittest.TestCase):
    def test_next_invoice_number_ignores_cn_files(self):
        assert copyfile_qt is not None

        with tempfile.TemporaryDirectory() as folder:
            open(os.path.join(folder, "ACME 0326 - 010 - CN.pdf"), "w").close()
            open(os.path.join(folder, "ACME 0326 - 002 - DO & INV.pdf"), "w").close()

            next_number = copyfile_qt.get_next_invoice_number(folder, "0326")

        self.assertEqual(next_number, 3)


if __name__ == "__main__":
    unittest.main()
