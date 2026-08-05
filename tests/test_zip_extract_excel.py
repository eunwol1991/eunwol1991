import importlib.util
import zipfile
from pathlib import Path
from types import ModuleType
from typing import Callable, cast


def _load_zip_extract_excel() -> ModuleType:
    path = Path(__file__).resolve().parents[1] / "eunwol1991/projects/check data use/ZipExtractExcel.py"
    spec = importlib.util.spec_from_file_location("zip_extract_excel", path)
    assert spec is not None
    assert spec.loader is not None
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def test_process_zip_removes_trailing_parenthesized_copy_suffix(tmp_path: Path):
    module = _load_zip_extract_excel()
    target_dir = tmp_path / "target"
    setattr(module, "TARGET_DIR", target_dir)
    target_dir.mkdir()
    process_zip = cast(Callable[[Path], bool], getattr(module, "process_zip"))

    zip_path = tmp_path / "excel.zip"
    with zipfile.ZipFile(zip_path, "w") as archive:
        archive.writestr("Report (1).xlsx", b"excel-content")

    assert process_zip(zip_path)
    assert (target_dir / "Report.xlsx").read_bytes() == b"excel-content"
    assert not (target_dir / "Report (1).xlsx").exists()
    assert not zip_path.exists()
