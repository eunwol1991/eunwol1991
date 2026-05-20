"""Batch open and save selected Excel workbooks with Microsoft Excel.

This script is intended for Windows machines with Microsoft Excel installed.
It uses xlwings so Excel itself performs the open/save operation.
"""

from __future__ import annotations

import argparse
import importlib
import sys
from collections.abc import Callable, Iterable
from pathlib import Path
from typing import NamedTuple, Protocol, cast


EXCEL_EXTENSIONS = {".xlsx", ".xlsm", ".xlsb", ".xls"}
TARGET_NAME_FRAGMENT = "xx26"


class ProcessSummary(NamedTuple):
    succeeded: int
    failed: int
    skipped: int


class ParsedArgs(NamedTuple):
    visible: bool


class WorkbookProtocol(Protocol):
    def save(self) -> None:
        ...

    def close(self) -> None:
        ...


class BooksProtocol(Protocol):
    def open(self, path: str, *, update_links: bool = False) -> WorkbookProtocol:
        ...


class ExcelAppProtocol(Protocol):
    display_alerts: bool
    screen_updating: bool

    @property
    def books(self) -> BooksProtocol:
        ...

    def quit(self) -> None:
        ...


def parse_args(argv: list[str]) -> ParsedArgs:
    parser = argparse.ArgumentParser(
        description="批量用 Microsoft Excel 打开并重新保存名称包含 xx26 的 Excel 文件。",
    )
    _ = parser.add_argument(
        "--visible",
        action="store_true",
        help="显示 Excel 窗口，便于排查问题。默认隐藏运行。",
    )
    namespace = parser.parse_args(argv)
    return ParsedArgs(visible=bool(namespace.visible))


def get_folder() -> Path:
    folder_text = input("请输入 Excel 文件夹路径: ").strip().strip('"')
    if folder_text.startswith(("source ", "python ", "python3 ", "cd ")):
        raise ValueError("请输入 Excel 文件夹路径，不要输入命令。")
    folder = Path(folder_text).expanduser()
    if not folder.exists():
        raise FileNotFoundError(f"文件夹不存在: {folder}")
    if not folder.is_dir():
        raise NotADirectoryError(f"不是文件夹: {folder}")
    return folder


def iter_excel_files(folder: Path) -> Iterable[Path]:
    for path in sorted(folder.iterdir()):
        if not path.is_file():
            continue
        if path.name.startswith("~$"):
            continue
        if TARGET_NAME_FRAGMENT.lower() not in path.stem.lower():
            continue
        if path.suffix.lower() in EXCEL_EXTENSIONS:
            yield path


def process_workbook(app: ExcelAppProtocol, workbook_path: Path) -> bool:
    workbook = None
    try:
        print(f"处理中: {workbook_path.name}")
        workbook = app.books.open(str(workbook_path), update_links=False)
        workbook.save()
        print(f"已完成: {workbook_path.name}")
        return True
    except Exception as exc:
        print(f"失败: {workbook_path.name}")
        print(f"原因: {exc}")
        return False
    finally:
        if workbook is not None:
            try:
                workbook.close()
            except Exception as exc:
                print(f"关闭失败: {workbook_path.name}")
                print(f"原因: {exc}")


def count_skipped_files(folder: Path, excel_files: list[Path]) -> int:
    selected = set(excel_files)
    return sum(1 for path in folder.iterdir() if path.is_file() and path not in selected)


def process_folder(folder: Path, *, visible: bool = False) -> ProcessSummary:
    try:
        xw = importlib.import_module("xlwings")
    except ImportError as exc:
        raise RuntimeError("缺少 xlwings，请先运行: pip install xlwings") from exc

    excel_files = list(iter_excel_files(folder))
    skipped = count_skipped_files(folder, excel_files)

    if not excel_files:
        print("没有找到名称包含 xx26 的可处理 Excel 文件。")
        return ProcessSummary(succeeded=0, failed=0, skipped=skipped)

    app = None
    succeeded = 0
    failed = 0

    try:
        app_factory = cast(Callable[[bool], ExcelAppProtocol], getattr(xw, "App"))
        app = app_factory(visible)
        app.display_alerts = False
        app.screen_updating = False

        for workbook_path in excel_files:
            if process_workbook(app, workbook_path):
                succeeded += 1
            else:
                failed += 1
    finally:
        if app is not None:
            app.quit()

    return ProcessSummary(succeeded=succeeded, failed=failed, skipped=skipped)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(sys.argv[1:] if argv is None else argv)

    try:
        folder = get_folder()
        summary = process_folder(folder, visible=args.visible)
    except (FileNotFoundError, NotADirectoryError, RuntimeError, ValueError) as exc:
        print(f"错误: {exc}")
        return 1
    except KeyboardInterrupt:
        print("用户已取消。")
        return 130

    print("全部完成")
    print(f"成功: {summary.succeeded}，失败: {summary.failed}，跳过: {summary.skipped}")
    return 0 if summary.failed == 0 else 2


if __name__ == "__main__":
    raise SystemExit(main())
