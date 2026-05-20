import importlib.util
import unittest
from collections.abc import Iterable
from pathlib import Path
from tempfile import TemporaryDirectory
from typing import Protocol, cast
from unittest.mock import patch


MODULE_PATH = Path(__file__).with_name("opensaveexcel.py")


class OpenSaveExcelModule(Protocol):
    def iter_excel_files(self, folder: Path) -> Iterable[Path]:
        ...

    def get_folder(self) -> Path:
        ...

    def process_workbook(self, app: object, workbook_path: Path) -> bool:
        ...


def load_module() -> OpenSaveExcelModule:
    spec = importlib.util.spec_from_file_location("opensaveexcel", MODULE_PATH)
    if spec is None or spec.loader is None:
        raise RuntimeError("Unable to load opensaveexcel.py")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return cast(OpenSaveExcelModule, cast(object, module))


class OpenSaveExcelTests(unittest.TestCase):
    def test_iter_excel_files_only_returns_xx26_excel_workbooks(self):
        module = load_module()

        with TemporaryDirectory() as temp_dir:
            folder = Path(temp_dir)
            included = [
                folder / "report_xx26.xlsx",
                folder / "REPORT_XX26.xlsm",
                folder / "client-xx26.xlsb",
                folder / "old_xx26.xls",
            ]
            excluded = [
                folder / "report_xx25.xlsx",
                folder / "xx26.txt",
                folder / "~$draft_xx26.xlsx",
            ]

            for path in included + excluded:
                _ = path.write_text("", encoding="utf-8")

            result = list(module.iter_excel_files(folder))

        self.assertEqual(result, sorted(included))

    def test_get_folder_reads_path_from_user_input(self):
        module = load_module()

        with TemporaryDirectory() as temp_dir:
            with patch("builtins.input", return_value=f'"{temp_dir}"'):
                result = module.get_folder()

        self.assertEqual(result, Path(temp_dir))

    def test_get_folder_rejects_shell_command_input(self):
        module = load_module()

        with patch(
            "builtins.input",
            return_value="source /mnt/c/Work/Savori-WorkSpace/.uv-venv/bin/activate",
        ):
            with self.assertRaisesRegex(ValueError, "请输入 Excel 文件夹路径，不要输入命令"):
                _ = module.get_folder()

    def test_process_workbook_disables_external_link_updates(self):
        module = load_module()
        opened: list[tuple[str, bool | None]] = []

        class FakeWorkbook:
            def save(self) -> None:
                pass

            def close(self) -> None:
                pass

        class FakeBooks:
            def open(self, path: str, *, update_links: bool | None = None) -> FakeWorkbook:
                opened.append((path, update_links))
                return FakeWorkbook()

        class FakeApp:
            books = FakeBooks()

        workbook_path = Path("/tmp/report_xx26.xlsx")

        result = module.process_workbook(FakeApp(), workbook_path)

        self.assertTrue(result)
        self.assertEqual(opened, [(str(workbook_path), False)])


if __name__ == "__main__":
    _ = unittest.main()
