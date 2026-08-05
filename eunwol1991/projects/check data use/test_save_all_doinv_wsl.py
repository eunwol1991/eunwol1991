from __future__ import annotations

import importlib.util
import logging
import sys
import tempfile
import types
import unittest
from pathlib import Path


SCRIPT_PATH = Path(__file__).with_name("SaveAllDOINV.py")


def load_module():
    pythoncom = types.ModuleType("pythoncom")
    setattr(pythoncom, "CoInitialize", lambda: None)
    setattr(pythoncom, "CoUninitialize", lambda: None)

    win32com = types.ModuleType("win32com")
    win32com_client = types.ModuleType("win32com.client")
    setattr(win32com, "client", win32com_client)

    pypdf = types.ModuleType("pypdf")
    setattr(pypdf, "PdfReader", object)
    setattr(pypdf, "PdfWriter", object)

    original_modules = {
        name: sys.modules.get(name)
        for name in ("pythoncom", "win32com", "win32com.client", "pypdf")
    }
    sys.modules.update(
        {
            "pythoncom": pythoncom,
            "win32com": win32com,
            "win32com.client": win32com_client,
            "pypdf": pypdf,
        }
    )

    try:
        spec = importlib.util.spec_from_file_location("save_all_doinv", SCRIPT_PATH)
        if spec is None or spec.loader is None:
            raise RuntimeError("Unable to load SaveAllDOINV.py")
        module = importlib.util.module_from_spec(spec)
        sys.modules[spec.name] = module
        spec.loader.exec_module(module)
        return module
    finally:
        _ = sys.modules.pop("save_all_doinv", None)
        for name, original in original_modules.items():
            if original is None:
                _ = sys.modules.pop(name, None)
            else:
                sys.modules[name] = original


class WslCompatibilityTests(unittest.TestCase):
    def test_converts_mnt_c_script_path_to_windows_path(self) -> None:
        module = load_module()

        converted = module.wsl_path_to_windows_path(
            Path("/mnt/c/work/Savori-WorkSpace/eunwol1991/projects/check data use/SaveAllDOINV.py")
        )

        self.assertEqual(
            converted,
            r"C:\work\Savori-WorkSpace\eunwol1991\projects\check data use\SaveAllDOINV.py",
        )

    def test_builds_windows_python_command_for_wsl_relaunch(self) -> None:
        module = load_module()

        command = module.build_windows_python_command(
            Path("/mnt/c/work/Savori-WorkSpace/eunwol1991/projects/check data use/SaveAllDOINV.py")
        )

        self.assertEqual(
            command,
            [
                "/mnt/c/work/Savori-WorkSpace/eunwol1991/.venv/Scripts/python.exe",
                r"C:\work\Savori-WorkSpace\eunwol1991\projects\check data use\SaveAllDOINV.py",
            ],
        )

    def test_create_excel_application_continues_when_calculation_setting_fails(self) -> None:
        module = load_module()
        excel = FakeExcelApplication()
        win32com_client = types.ModuleType("win32com.client")
        setattr(win32com_client, "DispatchEx", lambda _name: excel)
        setattr(module, "win32com_client", win32com_client)

        with self.assertNoLogs(level=logging.WARNING):
            created = module.create_excel_application()

        self.assertIs(created, excel)
        self.assertFalse(excel.Visible)
        self.assertFalse(excel.DisplayAlerts)

    def test_prepare_workbook_for_export_recalculates_before_pdf_export(self) -> None:
        module = load_module()
        excel = FakeCalculationExcel()
        workbook = FakeWorkbook()
        sleeps: list[float] = []
        setattr(module.time, "sleep", sleeps.append)

        module.prepare_workbook_for_export(excel, workbook)

        self.assertEqual(excel.Calculation, module.XL_CALCULATION_AUTOMATIC)
        self.assertTrue(workbook.refreshed)
        self.assertEqual(excel.rebuild_count, 1)
        self.assertEqual(sleeps, [0.2])

    def test_process_workbook_saves_calculated_workbook_before_pdf_export(self) -> None:
        module = load_module()
        events: list[str] = []
        workbook = FakeProcessWorkbook(events)
        excel = FakeProcessExcel(workbook)
        setattr(module, "export_sheet_to_pdf", lambda _sheet, _path: events.append("export"))
        setattr(module.time, "sleep", lambda _seconds: None)

        result = module.process_workbook(
            excel,
            Path("/tmp/generated_invoice.xlsx"),
            Path("/tmp"),
            1,
        )

        self.assertIsNone(result.error)
        self.assertFalse(excel.open_kwargs["ReadOnly"])
        self.assertIn("save", events)
        self.assertLess(events.index("save"), events.index("export"))

    def test_process_workbook_skips_locked_excel_file_before_opening(self) -> None:
        module = load_module()
        events: list[str] = []
        workbook = FakeProcessWorkbook(events)
        excel = FakeProcessExcel(workbook)
        setattr(module, "export_sheet_to_pdf", lambda _sheet, _path: events.append("export"))

        with tempfile.TemporaryDirectory() as temp_name:
            temp_dir = Path(temp_name)
            workbook_path = temp_dir / "CFC 0726 - 019 - DO & INV.xlsx"
            lock_path = temp_dir / "~$CFC 0726 - 019 - DO & INV.xlsx"
            workbook_path.touch()
            lock_path.touch()

            result = module.process_workbook(excel, workbook_path, temp_dir, 2)

        self.assertEqual(
            result.error,
            "Workbook appears to be open in Excel: ~$CFC 0726 - 019 - DO & INV.xlsx",
        )
        self.assertEqual(excel.open_kwargs, {})
        self.assertEqual(events, [])


class FakeExcelApplication:
    Visible: bool
    DisplayAlerts: bool

    def __init__(self) -> None:
        self.Visible = True
        self.DisplayAlerts = True

    def __setattr__(self, name: str, value: object) -> None:
        if name == "Calculation":
            raise RuntimeError("Cannot set Calculation property")
        super().__setattr__(name, value)


class FakeCalculationExcel:
    Calculation: int

    def __init__(self) -> None:
        self.Calculation = 0
        self.rebuild_count = 0
        self._states = [1, 0]

    @property
    def CalculationState(self) -> int:
        return self._states.pop(0)

    def CalculateFullRebuild(self) -> None:
        self.rebuild_count += 1


class FakeWorkbook:
    def __init__(self) -> None:
        self.refreshed = False

    def RefreshAll(self) -> None:
        self.refreshed = True


class FakeSheet:
    def __init__(self, name: str) -> None:
        self.Name = name


class FakeProcessWorkbook(FakeWorkbook):
    def __init__(self, events: list[str]) -> None:
        super().__init__()
        self.Worksheets = [FakeSheet("DO"), FakeSheet("Invoice")]
        self._events = events

    def Save(self) -> None:
        self._events.append("save")

    def Close(self, SaveChanges: bool) -> None:
        self._events.append(f"close:{SaveChanges}")


class FakeWorkbooks:
    def __init__(self, workbook: FakeProcessWorkbook) -> None:
        self._workbook = workbook
        self.open_kwargs: dict[str, object] = {}

    def Open(self, **kwargs: object) -> FakeProcessWorkbook:
        self.open_kwargs = kwargs
        return self._workbook


class FakeProcessExcel(FakeCalculationExcel):
    def __init__(self, workbook: FakeProcessWorkbook) -> None:
        super().__init__()
        self._states = [0]
        self.Workbooks = FakeWorkbooks(workbook)

    @property
    def open_kwargs(self) -> dict[str, object]:
        return self.Workbooks.open_kwargs


if __name__ == "__main__":
    unittest.main()
