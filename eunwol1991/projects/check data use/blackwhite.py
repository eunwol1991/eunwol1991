import os
import win32com.client as win32



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


def _platform_drive_root() -> str:
    if os.name == "nt":
        return "c:/"
    if _is_wsl():
        return "/mnt/c"
    return "/"


def _from_c(path_tail: str) -> str:
    tail = (path_tail or "").lstrip("/")
    root = _platform_drive_root()
    if root.endswith("/"):
        return f"{root}{tail}"
    return f"{root}/{tail}"
# 要处理的根目录
root_folder = _from_c("Users/jhunj/Dropbox/DO & INV/DO & INV 2025")
# 要跳过的目录
skip_folder = os.path.join(
    root_folder, "1. Order Summary & Sales Summary(Mth end)")

# 筛选关键字
keywords = ["xx25", "xx24"]


def set_book_bw(excel, file_path):
    try:
        wb = excel.Workbooks.Open(
            Filename=file_path, UpdateLinks=0, ReadOnly=False)
        count = 0
        for sheet in wb.Sheets:
            try:
                sheet.PageSetup.BlackAndWhite = True
                count += 1
            except Exception as e:
                print(f"    [SHEET-FAIL] {sheet.Name}: {e}")
        wb.Save()
        wb.Close(SaveChanges=True)
        print(f"[OK] {file_path}  — 设置为黑白的表：{count}")
    except Exception as e:
        print(f"[FAIL] {file_path} — {e}")


def main():
    excel = win32.Dispatch("Excel.Application")
    excel.Visible = False
    skipped_files = []

    for foldername, subfolders, filenames in os.walk(root_folder):
        if foldername.startswith(skip_folder):
            continue
        for filename in filenames:
            if filename.lower().endswith((".xls", ".xlsx", ".xlsm")):
                # 只处理包含 xx25 / xx24 的文件
                if any(kw in filename.lower() for kw in keywords):
                    file_path = os.path.join(foldername, filename)
                    set_book_bw(excel, file_path)
                else:
                    skipped_files.append(os.path.join(foldername, filename))

    excel.Quit()

    print("\n=== 已跳过的文件（不含 xx25 / xx24） ===")
    for f in skipped_files:
        print(f)


if __name__ == "__main__":
    main()
