# OpenSaveExcel Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Make `eunwol1991/projects/opensaveexcel.py` reliably batch open and save only Excel workbooks whose filenames contain `xx26` on Windows with Microsoft Excel.

**Architecture:** Keep this as a single focused script. Separate option parsing, interactive path input, dependency loading, workbook discovery, workbook processing, and CLI exit handling into small functions so failures are explicit and Excel cleanup is centralized.

**Tech Stack:** Python standard library (`argparse`, `pathlib`, `sys`), `xlwings`, Microsoft Excel COM automation through xlwings on Windows.

---

## File Structure

- Modify: `eunwol1991/projects/opensaveexcel.py` — standalone command-line utility for prompting for a folder path, validating input, launching Excel, processing `xx26` workbooks, and printing a summary.
- Create: `eunwol1991/projects/test_opensaveexcel.py` — unit tests for workbook filtering that do not require Microsoft Excel.

## Task 1: Replace hard-coded script with robust CLI utility

**Files:**
- Modify: `eunwol1991/projects/opensaveexcel.py`
- Create: `eunwol1991/projects/test_opensaveexcel.py`

- [ ] **Step 1: Replace the file with the implementation**

```python
"""Batch open and save Excel workbooks with Microsoft Excel.

This script is intended for Windows machines with Microsoft Excel installed.
It uses xlwings so Excel itself performs the open/save operation.
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path
from typing import Iterable, NamedTuple


EXCEL_EXTENSIONS = {".xlsx", ".xlsm", ".xlsb", ".xls"}
TARGET_NAME_FRAGMENT = "xx26"


class ProcessSummary(NamedTuple):
    succeeded: int
    failed: int
    skipped: int


def parse_args(argv: list[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="批量用 Microsoft Excel 打开并重新保存 Excel 文件。",
    )
    parser.add_argument(
        "folder",
        nargs="?",
        help="Excel 文件所在文件夹；不填写时会提示输入。",
    )
    parser.add_argument(
        "--visible",
        action="store_true",
        help="显示 Excel 窗口，便于排查问题。默认隐藏运行。",
    )
    return parser.parse_args(argv)


def get_folder(folder_arg: str | None) -> Path:
    folder_text = folder_arg or input("请输入 Excel 文件夹路径: ").strip().strip('"')
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
        if TARGET_NAME_FRAGMENT not in path.stem.lower():
            continue
        if path.suffix.lower() in EXCEL_EXTENSIONS:
            yield path


def process_workbook(app: object, workbook_path: Path) -> bool:
    workbook = None
    try:
        print(f"处理中: {workbook_path.name}")
        workbook = app.books.open(str(workbook_path))
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


def process_folder(folder: Path, *, visible: bool = False) -> ProcessSummary:
    try:
        import xlwings as xw
    except ImportError as exc:
        raise RuntimeError("缺少 xlwings，请先运行: pip install xlwings") from exc

    excel_files = list(iter_excel_files(folder))
    skipped = sum(1 for path in folder.iterdir() if path.is_file()) - len(excel_files)

    if not excel_files:
        print("没有找到可处理的 Excel 文件。")
        return ProcessSummary(succeeded=0, failed=0, skipped=max(skipped, 0))

    app = None
    succeeded = 0
    failed = 0

    try:
        app = xw.App(visible=visible)
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

    return ProcessSummary(succeeded=succeeded, failed=failed, skipped=max(skipped, 0))


def main(argv: list[str] | None = None) -> int:
    args = parse_args(sys.argv[1:] if argv is None else argv)

    try:
        folder = get_folder(args.folder)
        summary = process_folder(folder, visible=args.visible)
    except (FileNotFoundError, NotADirectoryError, RuntimeError) as exc:
        print(f"错误: {exc}")
        return 1
    except KeyboardInterrupt:
        print("用户已取消。")
        return 130

    print("全部完成")
    print(
        f"成功: {summary.succeeded}，失败: {summary.failed}，跳过: {summary.skipped}",
    )
    return 0 if summary.failed == 0 else 2


if __name__ == "__main__":
    raise SystemExit(main())
```

- [ ] **Step 2: Compile the script**

Run: `python3 -m py_compile eunwol1991/projects/opensaveexcel.py`

Expected: exit code `0` with no output.

- [ ] **Step 3: Verify CLI help does not require Excel**

Run: `python3 eunwol1991/projects/opensaveexcel.py --help`

Expected: exit code `0`, usage text includes `--visible` and the Chinese description.

- [ ] **Step 4: Verify invalid path handling**

Run: `python3 eunwol1991/projects/opensaveexcel.py /tmp/path-that-does-not-exist-for-opensaveexcel`

Expected: exit code `1`, output starts with `错误: 文件夹不存在:`.

## Self-Review

- Spec coverage: CLI path handling, interactive fallback, Excel processing, `xx26` filename filtering, supported extensions, temp file skipping, per-file failure handling, cleanup, and summary are all covered by Task 1.
- Placeholder scan: no TBD/TODO placeholders remain.
- Type consistency: `ProcessSummary`, `parse_args`, `get_folder`, `iter_excel_files`, `process_workbook`, `process_folder`, and `main` are defined before use.
