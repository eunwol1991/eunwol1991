#!/usr/bin/env python3

from __future__ import annotations

import os
import re
import shutil
import sys
import tempfile
import zipfile
from datetime import datetime
from pathlib import Path, PurePosixPath
from uuid import uuid4


# Windows:
# C:\Users\jhunj\Dropbox\for jj\Excel Zip
SOURCE_DIR = Path("/mnt/c/Users/jhunj/Dropbox/for jj/Excel Zip")

# Windows:
# C:\Users\jhunj\Dropbox\for jj\Doc to print - JJ
TARGET_DIR = Path("/mnt/c/Users/jhunj/Dropbox/for jj/Doc to print - JJ")

EXCEL_EXTENSIONS = {
    ".xlsx",
    ".xlsm",
    ".xls",
    ".xlsb",
}

COPY_SUFFIX_RE = re.compile(r"\s*\(\d+\)$")


def log(message: str) -> None:
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    print(f"[{timestamp}] {message}", flush=True)


def is_zip_symlink(info: zipfile.ZipInfo) -> bool:
    """检查 ZIP 项目是否为符号链接。"""
    file_type = (info.external_attr >> 16) & 0o170000
    return file_type == 0o120000


def safe_extract_zip(zip_path: Path, extract_dir: Path) -> None:
    """
    安全解压 ZIP。

    防止：
    1. 路径穿越，例如 ../../file.xlsx
    2. 绝对路径
    3. ZIP 内符号链接
    """
    with zipfile.ZipFile(zip_path, "r") as archive:
        bad_file = archive.testzip()

        if bad_file is not None:
            raise RuntimeError(f"ZIP 文件损坏，异常项目：{bad_file}")

        for info in archive.infolist():
            if info.is_dir():
                continue

            if is_zip_symlink(info):
                raise RuntimeError(
                    f"ZIP 内含符号链接，已停止处理：{info.filename}"
                )

            normalized_name = info.filename.replace("\\", "/")
            relative_path = PurePosixPath(normalized_name)

            if relative_path.is_absolute() or ".." in relative_path.parts:
                raise RuntimeError(
                    f"ZIP 内含不安全路径，已停止处理：{info.filename}"
                )

            output_path = extract_dir.joinpath(*relative_path.parts)
            output_path.parent.mkdir(parents=True, exist_ok=True)

            with archive.open(info, "r") as source_file:
                with output_path.open("wb") as output_file:
                    shutil.copyfileobj(source_file, output_file)


def find_excel_files(extract_dir: Path) -> list[Path]:
    """寻找解压目录中的所有 Excel 文件。"""
    excel_files: list[Path] = []

    for file_path in extract_dir.rglob("*"):
        if not file_path.is_file():
            continue

        if "__MACOSX" in file_path.parts:
            continue

        if file_path.name == ".DS_Store":
            continue

        if file_path.name.startswith("~$"):
            continue

        if file_path.suffix.lower() not in EXCEL_EXTENSIONS:
            continue

        excel_files.append(file_path)

    return sorted(excel_files, key=lambda path: path.name.casefold())


def normalize_excel_filename(file_name: str) -> str:
    path = Path(file_name)
    clean_stem = COPY_SUFFIX_RE.sub("", path.stem).rstrip()
    return f"{clean_stem}{path.suffix}"


def check_duplicate_filenames(excel_files: list[Path]) -> None:
    """
    因为所有 Excel 都会放到目标文件夹根目录，
    所以 ZIP 内不能出现两个相同文件名。
    """
    filename_map: dict[str, list[Path]] = {}

    for file_path in excel_files:
        key = normalize_excel_filename(file_path.name).casefold()
        filename_map.setdefault(key, []).append(file_path)

    duplicate_groups = {
        name: paths
        for name, paths in filename_map.items()
        if len(paths) > 1
    }

    if not duplicate_groups:
        return

    details: list[str] = []

    for paths in duplicate_groups.values():
        details.append(
            ", ".join(str(path) for path in paths)
        )

    raise RuntimeError(
        "ZIP 内出现重复文件名，无法安全放入同一个目标文件夹："
        + " | ".join(details)
    )


def atomic_copy_replace(source: Path, destination: Path) -> None:
    """
    先复制到目标文件夹内的临时文件，再用 os.replace 覆盖。

    这样可以降低复制到一半时留下损坏文件的风险。
    """
    destination.parent.mkdir(parents=True, exist_ok=True)

    temporary_file = destination.parent / (
        f".{destination.name}.{uuid4().hex}.tmp"
    )

    try:
        _ = shutil.copy2(source, temporary_file)
        os.replace(temporary_file, destination)
    finally:
        if temporary_file.exists():
            temporary_file.unlink()


def rollback_changes(
    completed_changes: list[tuple[Path, Path | None]],
) -> list[str]:
    """
    尝试还原当前 ZIP 已经进行的覆盖操作。

    backup_path 为 None 表示目标文件原本不存在，
    回滚时应删除新复制进去的文件。
    """
    rollback_errors: list[str] = []

    for destination, backup_path in reversed(completed_changes):
        try:
            if backup_path is not None and backup_path.exists():
                atomic_copy_replace(backup_path, destination)
            elif destination.exists():
                destination.unlink()
        except Exception as error:
            rollback_errors.append(
                f"{destination.name}: {error}"
            )

    return rollback_errors


def process_zip(zip_path: Path) -> bool:
    """处理单个 ZIP。"""
    log(f"开始处理：{zip_path.name}")

    completed_changes: list[tuple[Path, Path | None]] = []

    try:
        with tempfile.TemporaryDirectory(
            prefix=f"excel_zip_{zip_path.stem}_"
        ) as working_directory:
            working_dir = Path(working_directory)
            extract_dir = working_dir / "extracted"
            backup_dir = working_dir / "backup"

            extract_dir.mkdir(parents=True, exist_ok=True)
            backup_dir.mkdir(parents=True, exist_ok=True)

            safe_extract_zip(zip_path, extract_dir)

            excel_files = find_excel_files(extract_dir)

            if not excel_files:
                raise RuntimeError("ZIP 内没有找到 Excel 文件")

            check_duplicate_filenames(excel_files)

            log(f"找到 {len(excel_files)} 个 Excel 文件")

            for source_file in excel_files:
                destination_name = normalize_excel_filename(source_file.name)
                destination_file = TARGET_DIR / destination_name
                backup_file: Path | None = None

                if destination_file.exists():
                    backup_file = (
                        backup_dir
                        / f"{uuid4().hex}_{destination_file.name}"
                    )
                    _ = shutil.copy2(destination_file, backup_file)
                    action = "覆盖"
                else:
                    action = "新增"

                log(
                    f"{action}：{source_file.name} -> {destination_name}"
                )

                atomic_copy_replace(
                    source_file,
                    destination_file,
                )

                completed_changes.append(
                    (destination_file, backup_file)
                )

            zip_path.unlink()

            log(f"完成：{zip_path.name}，ZIP 已删除")

            return True

    except PermissionError as error:
        log(
            f"处理失败，可能有 Excel 文件正在打开：{error}"
        )

    except zipfile.BadZipFile:
        log(f"处理失败，ZIP 文件无效或已损坏：{zip_path.name}")

    except Exception as error:
        log(f"处理失败：{error}")

    if completed_changes:
        log("正在尝试还原本次已覆盖的文件")

        rollback_errors = rollback_changes(completed_changes)

        if rollback_errors:
            log("警告：部分文件回滚失败")
            for rollback_error in rollback_errors:
                log(f"回滚失败：{rollback_error}")
        else:
            log("已成功还原本次变更")

    log(f"ZIP 已保留：{zip_path.name}")
    return False


def main() -> int:
    log("Excel ZIP 自动替换程序启动")

    if not SOURCE_DIR.exists():
        log(f"来源文件夹不存在：{SOURCE_DIR}")
        return 1

    if not SOURCE_DIR.is_dir():
        log(f"来源路径不是文件夹：{SOURCE_DIR}")
        return 1

    TARGET_DIR.mkdir(parents=True, exist_ok=True)

    zip_files = sorted(
        SOURCE_DIR.glob("*.zip"),
        key=lambda path: (
            path.stat().st_mtime,
            path.name.casefold(),
        ),
    )

    if not zip_files:
        log("Excel Zip 文件夹内没有 ZIP 文件")
        return 0

    log(f"共找到 {len(zip_files)} 个 ZIP 文件")

    success_count = 0
    failed_count = 0

    for zip_path in zip_files:
        if process_zip(zip_path):
            success_count += 1
        else:
            failed_count += 1

    log(f"处理结束，成功 {success_count} 个，失败 {failed_count} 个")

    if failed_count > 0:
        return 1

    return 0


if __name__ == "__main__":
    sys.exit(main())
