import os
import sys
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox

try:
    import openpyxl
except Exception:
    openpyxl = None


START_SHEET_NAME = "Jan End"
END_SHEET_NAME = "Dec 2025"


def rename_between_markers(file_path: Path) -> tuple[list[tuple[str, str]], list[str]]:
    if openpyxl is None:
        raise RuntimeError("openpyxl is not installed. Please run: pip install openpyxl")

    ext = file_path.suffix.lower()
    keep_vba = ext == ".xlsm"
    wb = openpyxl.load_workbook(file_path, data_only=False, keep_vba=keep_vba)

    try:
        names = list(wb.sheetnames)
        if START_SHEET_NAME not in names:
            raise ValueError(f"Cannot find sheet: {START_SHEET_NAME}")
        if END_SHEET_NAME not in names:
            raise ValueError(f"Cannot find sheet: {END_SHEET_NAME}")

        start_idx = names.index(START_SHEET_NAME)
        end_idx = names.index(END_SHEET_NAME)

        left = min(start_idx, end_idx)
        right = max(start_idx, end_idx)
        if right - left <= 1:
            return [], ["No sheets exist between marker sheets."]

        changed: list[tuple[str, str]] = []
        skipped: list[str] = []

        existing_names = set(names)
        targets = names[left + 1: right]

        for old_name in targets:
            if old_name.endswith("1"):
                skipped.append(f"{old_name} (already ends with '1')")
                continue

            new_name = f"{old_name}1"
            if len(new_name) > 31:
                skipped.append(f"{old_name} (target too long: {new_name})")
                continue

            if new_name in existing_names:
                skipped.append(f"{old_name} (target exists: {new_name})")
                continue

            ws = wb[old_name]
            ws.title = new_name
            changed.append((old_name, new_name))
            existing_names.remove(old_name)
            existing_names.add(new_name)

        if changed:
            wb.save(file_path)
        return changed, skipped
    finally:
        wb.close()


def pick_excel_file() -> Path | None:
    root = tk.Tk()
    root.withdraw()
    root.update()
    selected = filedialog.askopenfilename(
        title="Select Excel file",
        filetypes=[
            ("Excel files", "*.xlsx *.xlsm *.xltx *.xltm"),
            ("All files", "*.*"),
        ],
    )
    root.destroy()
    if not selected:
        return None
    return Path(selected)


def build_report(changed: list[tuple[str, str]], skipped: list[str]) -> str:
    lines: list[str] = []
    lines.append(f"Changed: {len(changed)}")
    for old_name, new_name in changed:
        lines.append(f"  - {old_name} -> {new_name}")
    lines.append(f"Skipped: {len(skipped)}")
    for item in skipped:
        lines.append(f"  - {item}")
    return "\n".join(lines)


def main() -> int:
    if openpyxl is None:
        print("Missing dependency: openpyxl")
        print("Install with: pip install openpyxl")
        return 1

    if len(sys.argv) > 1:
        file_path = Path(sys.argv[1]).expanduser()
    else:
        picked = pick_excel_file()
        if picked is None:
            print("No file selected.")
            return 0
        file_path = picked

    if not file_path.exists() or not file_path.is_file():
        print(f"File not found: {file_path}")
        return 1

    if file_path.suffix.lower() not in {".xlsx", ".xlsm", ".xltx", ".xltm"}:
        print("Unsupported file type. Use .xlsx/.xlsm/.xltx/.xltm")
        return 1

    try:
        changed, skipped = rename_between_markers(file_path)
    except Exception as exc:
        print(f"Failed: {exc}")
        try:
            messagebox.showerror("Rename Failed", str(exc))
        except Exception:
            pass
        return 1

    report = build_report(changed, skipped)
    print(report)
    try:
        messagebox.showinfo("Done", report)
    except Exception:
        pass
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
