import argparse
import os
import re
import subprocess
import sys
from pathlib import Path
from typing import Dict, List, Optional


EXCEL_EXTS = {".xlsx", ".xlsm", ".xls"}
RE_DOC_KEY = re.compile(r"\b([A-Za-z]{2,})\s*(\d{4})\s*-\s*(\d{3})\b", re.IGNORECASE)
RE_REVISED = re.compile(r"\b(?:rev(?:ise|ised|ision)?|reivse|r\d+)\b", re.IGNORECASE)


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
    return f"{root}{tail}" if root.endswith("/") else f"{root}/{tail}"


def _windows_path_to_wsl(path: str) -> str:
    text = (path or "").strip().strip('"').strip("'")
    m = re.match(r"^([A-Za-z]):[\\/](.*)$", text)
    if not m:
        return text
    drive = m.group(1).lower()
    rest = m.group(2).replace("\\", "/")
    return f"/mnt/{drive}/{rest}"


def _normalize_input_path(path: str) -> str:
    text = (path or "").strip().strip('"').strip("'")
    if re.match(r"^[A-Za-z]:(?![\\/])", text):
        text = text[:2] + "/" + text[2:]
    text = text.replace("\\", "/")
    return text


def _wsl_path_to_windows(path: str) -> str:
    text = (path or "").strip().strip('"').strip("'")
    m = re.match(r"^/mnt/([a-zA-Z])/(.*)$", text)
    if not m:
        return text
    drive = m.group(1).upper()
    rest = m.group(2).replace("/", "\\")
    return f"{drive}:\\{rest}"


def doc_key_from_name(name: str) -> Optional[str]:
    match = RE_DOC_KEY.search(name or "")
    if not match:
        return None
    prefix = match.group(1).upper()
    mmyy = match.group(2)
    seq = match.group(3)
    return f"{prefix}-{mmyy}-{seq}"


def is_revised_name(path: Path) -> bool:
    return bool(RE_REVISED.search(path.stem))


def iter_pdfs(root: Path) -> List[Path]:
    return [p for p in root.rglob("*.pdf") if p.is_file()]


def iter_excels(root: Path) -> List[Path]:
    files: List[Path] = []
    for p in root.rglob("*"):
        if not p.is_file():
            continue
        if p.suffix.lower() not in EXCEL_EXTS:
            continue
        if p.name.startswith("~$"):
            continue
        files.append(p)
    return files


def choose_pdf_targets(pdfs: List[Path]) -> List[Path]:
    grouped: Dict[tuple[Path, str], List[Path]] = {}
    passthrough: List[Path] = []

    for pdf in pdfs:
        key = doc_key_from_name(pdf.name)
        if not key:
            passthrough.append(pdf)
            continue
        grouped.setdefault((pdf.parent, key), []).append(pdf)

    selected: List[Path] = []
    for _, files in grouped.items():
        revised = [f for f in files if is_revised_name(f)]
        if revised:
            revised.sort(key=lambda p: p.name.lower())
            selected.append(revised[-1])
        else:
            files.sort(key=lambda p: p.name.lower())
            selected.append(files[-1])

    selected.extend(passthrough)
    return sorted(set(selected), key=lambda p: str(p).lower())


def build_excel_index(excels: List[Path]) -> Dict[str, List[Path]]:
    index: Dict[str, List[Path]] = {}
    for x in excels:
        key = doc_key_from_name(x.name)
        if not key:
            continue
        index.setdefault(key, []).append(x)

    for key in list(index.keys()):
        index[key] = sorted(index[key], key=lambda p: str(p).lower())
    return index


def _score_excel_for_pdf(pdf: Path, excel: Path) -> int:
    score = 0
    if excel.parent == pdf.parent:
        score += 200
    if "do & inv" in excel.name.lower() or "doinv" in excel.name.lower():
        score += 50
    if is_revised_name(pdf) and is_revised_name(excel):
        score += 30
    if pdf.stem.lower() in excel.stem.lower() or excel.stem.lower() in pdf.stem.lower():
        score += 20
    score -= len(str(excel))
    return score


def pick_excel_for_pdf(pdf: Path, excel_index: Dict[str, List[Path]]) -> Optional[Path]:
    key = doc_key_from_name(pdf.name)
    if not key:
        return None
    options = excel_index.get(key, [])
    if not options:
        return None
    return sorted(options, key=lambda x: _score_excel_for_pdf(pdf, x), reverse=True)[0]


def export_invoice_sheet_to_pdf(
    excel_path: Path, output_pdf_path: Path, dry_run: bool
) -> None:
    if dry_run:
        return

    excel_win = _wsl_path_to_windows(str(excel_path)) if _is_wsl() else str(excel_path)
    pdf_win = (
        _wsl_path_to_windows(str(output_pdf_path))
        if _is_wsl()
        else str(output_pdf_path)
    )

    ps = (
        "$ErrorActionPreference='Stop';"
        "$excel=New-Object -ComObject Excel.Application;"
        "$excel.Visible=$false;"
        "$excel.DisplayAlerts=$false;"
        f"$wb=$excel.Workbooks.Open('{excel_win.replace("'", "''")}');"
        "$ws=$wb.Worksheets.Item('Invoice');"
        f"$ws.ExportAsFixedFormat(0, '{pdf_win.replace("'", "''")}');"
        "$wb.Close($false);"
        "$excel.Quit();"
        "[System.Runtime.Interopservices.Marshal]::ReleaseComObject($ws)|Out-Null;"
        "[System.Runtime.Interopservices.Marshal]::ReleaseComObject($wb)|Out-Null;"
        "[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)|Out-Null;"
        "[GC]::Collect();[GC]::WaitForPendingFinalizers();"
    )
    subprocess.run(["powershell.exe", "-NoProfile", "-Command", ps], check=True)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Recursively regenerate Invoice PDFs from matching Excel files."
    )
    parser.add_argument("--root", default="", help="Root folder for recursive scan")
    parser.add_argument("--dry-run", action="store_true", help="Report only")
    parser.add_argument(
        "--no-prompt", action="store_true", help="Disable interactive prompt"
    )
    return parser.parse_args()


def _prompt_with_default(label: str, default_value: str) -> str:
    entered = input(f"{label} [{default_value}]: ").strip().strip('"')
    return entered if entered else default_value


def _prompt_yes_no(label: str, default_yes: bool) -> bool:
    hint = "Y/n" if default_yes else "y/N"
    entered = input(f"{label} ({hint}): ").strip().lower()
    if not entered:
        return default_yes
    return entered in {"y", "yes"}


def main() -> int:
    args = parse_args()
    default_root = _from_c("Users/jhunj/Dropbox/DO & INV/DO & INV 2026")

    root_input = args.root.strip().strip('"') if args.root else ""
    if not args.no_prompt:
        root_input = _prompt_with_default(
            "Input root folder", root_input or default_root
        )
        if "--dry-run" not in sys.argv:
            args.dry_run = _prompt_yes_no("Run dry-run only", default_yes=True)

    raw_root = _normalize_input_path(root_input or default_root)
    candidates = [Path(_windows_path_to_wsl(raw_root)), Path(raw_root)]
    root_path = None
    for c in candidates:
        if c.exists() and c.is_dir():
            root_path = c
            break

    if root_path is None:
        print(f"[ERROR] Root folder not found: {raw_root}")
        print("Tip: use one of these formats:")
        print("  /mnt/c/Users/jhunj/Dropbox/DO & INV/DO & INV 2026")
        print("  C:/Users/jhunj/Dropbox/DO & INV/DO & INV 2026")
        return 1

    pdf_all = iter_pdfs(root_path)
    excel_all = iter_excels(root_path)
    pdf_targets = choose_pdf_targets(pdf_all)
    excel_index = build_excel_index(excel_all)

    print(f"PDF found: {len(pdf_all)}")
    print(f"PDF selected after revise-preference: {len(pdf_targets)}")
    print(f"Excel found: {len(excel_all)}")

    processed = 0
    updated = 0
    missing_excel = 0
    failed = 0

    for pdf in pdf_targets:
        key = doc_key_from_name(pdf.name)
        if not key:
            continue
        processed += 1
        excel = pick_excel_for_pdf(pdf, excel_index)
        if excel is None:
            missing_excel += 1
            print(f"[MISS] {pdf} -> no matching excel")
            continue
        try:
            export_invoice_sheet_to_pdf(excel, pdf, args.dry_run)
            updated += 1
            tag = "[DRY-RUN]" if args.dry_run else "[UPDATED]"
            print(f"{tag} {pdf} <- {excel.name}")
        except Exception as exc:
            failed += 1
            print(f"[FAIL] {pdf} <- {excel.name} -> {exc}")

    print("\n=== Summary ===")
    print(f"Processed key-based PDFs: {processed}")
    print(f"Updated PDFs: {updated}")
    print(f"Missing matching Excel: {missing_excel}")
    print(f"Failed updates: {failed}")
    print(f"Mode: {'dry-run' if args.dry_run else 'write'}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
