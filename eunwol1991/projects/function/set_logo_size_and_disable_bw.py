import os
import re
from pathlib import Path


try:
    import openpyxl
except ImportError:
    openpyxl = None


EXCEL_EXTS = {".xlsx", ".xlsm", ".xltx", ".xltm"}
TARGET_HEIGHT_CM = 2.49
TARGET_WIDTH_CM = 8.85
PIXELS_PER_INCH = 96
CM_PER_INCH = 2.54


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


def normalize_input_path(raw: str) -> Path:
    value = (raw or "").strip().strip('"').strip("'")
    if not value:
        return Path(value)

    windows_drive_match = re.match(r"^([A-Za-z]):[\\/]", value)
    if windows_drive_match and _is_wsl():
        drive = windows_drive_match.group(1).lower()
        tail = value[2:].replace("\\", "/")
        return Path(f"/mnt/{drive}{tail}")

    return Path(value)


def cm_to_px(cm: float) -> int:
    return int(round((cm / CM_PER_INCH) * PIXELS_PER_INCH))


def iter_excel_files(path: Path):
    if path.is_file():
        if path.suffix.lower() in EXCEL_EXTS and not path.name.startswith("~$"):
            yield path
        return

    for root, _dirs, files in os.walk(path):
        for filename in files:
            if filename.startswith("~$"):
                continue
            file_path = Path(root) / filename
            if file_path.suffix.lower() in EXCEL_EXTS:
                yield file_path


def process_workbook(file_path: Path) -> tuple[int, int] | None:
    if openpyxl is None:
        return None

    ext = file_path.suffix.lower()
    keep_vba = ext == ".xlsm"

    try:
        wb = openpyxl.load_workbook(file_path, data_only=False, keep_vba=keep_vba)
    except Exception as exc:
        print(f"[WARN] Open failed: {file_path} -> {exc}")
        return None

    resized_images = 0
    bw_disabled = 0
    target_width_px = cm_to_px(TARGET_WIDTH_CM)
    target_height_px = cm_to_px(TARGET_HEIGHT_CM)

    for ws in wb.worksheets:
        images = list(getattr(ws, "_images", []))
        for img in images:
            changed = False
            if getattr(img, "width", None) != target_width_px:
                img.width = target_width_px
                changed = True
            if getattr(img, "height", None) != target_height_px:
                img.height = target_height_px
                changed = True
            if changed:
                resized_images += 1

        if ws.page_setup.blackAndWhite:
            ws.page_setup.blackAndWhite = False
            bw_disabled += 1

    if resized_images > 0 or bw_disabled > 0:
        try:
            wb.save(file_path)
        except Exception as exc:
            print(f"[WARN] Save failed: {file_path} -> {exc}")
            return None

    return resized_images, bw_disabled


def main() -> None:
    if openpyxl is None:
        print("Missing dependency: pip install openpyxl")
        return

    raw = input("Input folder or file path: ")
    if not raw:
        print("No path provided.")
        return

    target = normalize_input_path(raw)
    if not target.exists():
        print(f"Path not found: {target}")
        return

    files = list(iter_excel_files(target))
    if not files:
        print("No Excel files found.")
        return

    processed = 0
    changed_files = 0
    failed = 0
    total_resized = 0
    total_bw_disabled = 0

    for file_path in files:
        result = process_workbook(file_path)
        processed += 1
        if result is None:
            failed += 1
            continue

        resized_images, bw_disabled = result
        if resized_images > 0 or bw_disabled > 0:
            changed_files += 1
            total_resized += resized_images
            total_bw_disabled += bw_disabled
            print(
                f"[OK] {file_path} -> resized images: {resized_images}, "
                f"black-and-white disabled sheets: {bw_disabled}"
            )

    print(
        f"Done. Files: {processed}, Changed: {changed_files}, Failed: {failed}, "
        f"Total resized images: {total_resized}, "
        f"Total sheets with black-and-white disabled: {total_bw_disabled}"
    )


if __name__ == "__main__":
    main()
