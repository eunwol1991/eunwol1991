import os
from pathlib import Path

try:
    import openpyxl
    from openpyxl.drawing.image import Image as XLImage
except ImportError:
    openpyxl = None
    XLImage = None


LOGO_PATH = r"C:\Users\jhunj\Dropbox\for jj\Logo_Savori - Green.png"
NAME_SUBSTRING = "xx26"
EXCEL_EXTS = {".xlsx", ".xlsm", ".xltx", ".xltm"}
EXTRA_FILES = [
    r"C:\Users\jhunj\Dropbox\DO & INV\DO & INV 2026\Melvin - StuffD\SD 2026 - DO format (By Outlet).xlsx",
    r"C:\Users\jhunj\Dropbox\DO & INV\DO & INV 2026\Melvin - StuffD\SD 2026 - INV format (By Outlet).xlsx",
]


def iter_excel_files(path: Path, substring: str):
    if path.is_file():
        if should_process_file(path, substring):
            yield path
        return

    for root, _dirs, files in os.walk(path):
        for filename in files:
            if filename.startswith("~$"):
                continue
            file_path = Path(root) / filename
            if should_process_file(file_path, substring):
                yield file_path


def should_process_file(path: Path, substring: str) -> bool:
    if path.suffix.lower() not in EXCEL_EXTS:
        return False
    return substring.lower() in path.name.lower()


def replace_images_in_sheet(ws, logo_path: str) -> int:
    images = list(getattr(ws, "_images", []))
    if not images:
        return 0

    ws._images = []
    replaced = 0

    for old in images:
        new_img = XLImage(logo_path)
        old_width = getattr(old, "width", None)
        old_height = getattr(old, "height", None)
        if old_width and old_height:
            new_img.width = old_width
            new_img.height = old_height

        anchor = getattr(old, "anchor", None)
        added = False
        if anchor is not None:
            try:
                ws.add_image(new_img, anchor)
                added = True
            except Exception:
                added = False

        if not added:
            try:
                ws.add_image(new_img)
                added = True
            except Exception:
                added = False

        if not added:
            try:
                new_img.anchor = anchor
            except Exception:
                pass
            ws._images.append(new_img)

        replaced += 1

    return replaced


def process_workbook(file_path: Path):
    ext = file_path.suffix.lower()
    keep_vba = ext == ".xlsm"

    try:
        wb = openpyxl.load_workbook(file_path, data_only=False, keep_vba=keep_vba)
    except Exception as exc:
        print(f"[WARN] Open failed: {file_path} -> {exc}")
        return None

    replaced_total = 0
    for ws in wb.worksheets:
        replaced_total += replace_images_in_sheet(ws, LOGO_PATH)

    if replaced_total > 0:
        try:
            wb.save(file_path)
        except Exception as exc:
            print(f"[WARN] Save failed: {file_path} -> {exc}")
            return None

    return replaced_total


def main() -> None:
    if openpyxl is None or XLImage is None:
        print("Missing dependency: pip install openpyxl pillow")
        return

    if not os.path.isfile(LOGO_PATH):
        print(f"Logo not found: {LOGO_PATH}")
        return

    raw = input("Input folder or file path: ").strip().strip('"')
    if not raw:
        print("No path provided.")
        return

    target = Path(raw)
    if not target.exists():
        print(f"Path not found: {target}")
        return

    files = {p for p in iter_excel_files(target, NAME_SUBSTRING)}
    missing_extras = []
    for extra in EXTRA_FILES:
        extra_path = Path(extra)
        if extra_path.exists():
            if extra_path.suffix.lower() in EXCEL_EXTS:
                files.add(extra_path)
        else:
            missing_extras.append(extra_path)

    if not files:
        print(f"No files found with '{NAME_SUBSTRING}' in name and no extra files found.")
        return

    if missing_extras:
        for extra_path in missing_extras:
            print(f"[WARN] Extra file not found: {extra_path}")

    processed = 0
    changed_files = 0
    no_image = 0
    failed = 0
    total_images = 0

    for file_path in files:
        result = process_workbook(file_path)
        processed += 1
        if result is None:
            failed += 1
            continue
        if result == 0:
            no_image += 1
            continue
        changed_files += 1
        total_images += result
        print(f"[OK] {file_path} -> {result} images replaced")

    print(
        f"Done. Files: {processed}, Changed: {changed_files}, "
        f"No images: {no_image}, Failed: {failed}, "
        f"Total images replaced: {total_images}"
    )


if __name__ == "__main__":
    main()
