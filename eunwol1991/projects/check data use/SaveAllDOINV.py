from __future__ import annotations

import gc
import importlib
import logging
import os
import shutil
import subprocess
import sys
import tempfile
import time
from dataclasses import dataclass
from collections.abc import Callable, Iterable
from pathlib import Path
from types import ModuleType
from typing import BinaryIO, Protocol, cast


class PdfReaderProtocol(Protocol):
    pages: list[object]


class PdfWriterProtocol(Protocol):
    def add_page(self, page: object) -> object: ...

    def write(self, stream: BinaryIO) -> object: ...

    def close(self) -> object: ...


class ExcelApplicationProtocol(Protocol):
    Visible: bool
    DisplayAlerts: bool
    ScreenUpdating: bool
    EnableEvents: bool
    AskToUpdateLinks: bool
    Calculation: int
    AutomationSecurity: int
    Workbooks: object

    def Quit(self) -> object: ...


PdfReaderFactory = Callable[[str], PdfReaderProtocol]
PdfWriterFactory = Callable[[], PdfWriterProtocol]

pythoncom: ModuleType | None = None
win32com_client: ModuleType | None = None
PdfReader: PdfReaderFactory | None = None
PdfWriter: PdfWriterFactory | None = None


# =========================
# Configuration
# =========================
TARGET_DIR = Path(r"C:\Users\jhunj\Dropbox\for jj\Doc to print - JJ")
RECURSIVE = False

DO_SHEET_NAME = "DO"
INVOICE_SHEET_NAME = "Invoice"

# The requested output names.
DO_OUTPUT_NAME = "today DO.pdf"
INVOICE_OUTPUT_NAME = "today invoice.pdf"
LOG_NAME = "export_do_invoice.log"

SUPPORTED_EXTENSIONS = {".xlsx", ".xlsm", ".xls"}

# Excel constants
XL_TYPE_PDF = 0
XL_QUALITY_STANDARD = 0
XL_CALCULATION_AUTOMATIC = -4105
XL_CALCULATION_DONE = 0
MSO_AUTOMATION_SECURITY_FORCE_DISABLE = 3


def is_wsl_environment() -> bool:
    return bool(os.environ.get("WSL_DISTRO_NAME") or os.environ.get("WSL_INTEROP"))


def wsl_path_to_windows_path(path: Path) -> str:
    resolved = path.resolve()
    parts = resolved.parts
    if len(parts) < 3 or parts[0] != "/" or parts[1] != "mnt" or len(parts[2]) != 1:
        raise RuntimeError(f"WSL path is not on a Windows drive: {resolved}")

    drive = parts[2].upper()
    rest = "\\".join(parts[3:])
    return f"{drive}:\\{rest}" if rest else f"{drive}:\\"


def build_windows_python_command(script_path: Path) -> list[str]:
    windows_script_path = wsl_path_to_windows_path(script_path)
    venv_python = find_windows_venv_python(script_path)
    if venv_python is not None:
        return [str(venv_python), windows_script_path]
    return ["py.exe", "-3", windows_script_path]


def find_windows_venv_python(script_path: Path) -> Path | None:
    for parent in script_path.resolve().parents:
        python_exe = parent / ".venv" / "Scripts" / "python.exe"
        if python_exe.exists():
            return python_exe
    return None


def relaunch_with_windows_python_if_needed(script_path: Path) -> int | None:
    if os.name == "nt":
        return None
    if not is_wsl_environment():
        return None

    command = build_windows_python_command(script_path)
    print("Detected WSL; relaunching with Windows Python for Excel automation...")
    try:
        completed = subprocess.run(command, check=False)
    except FileNotFoundError:
        print("Unable to find cmd.exe from WSL. Run this script with Windows Python instead.")
        return 2
    except RuntimeError as exc:
        print(str(exc))
        return 2
    return completed.returncode


def ensure_runtime_dependencies() -> None:
    global PdfReader, PdfWriter, pythoncom, win32com_client

    try:
        pythoncom_module = importlib.import_module("pythoncom")
        win32com_client_module = importlib.import_module("win32com.client")
    except ImportError as exc:
        raise SystemExit(
            "Missing dependency: pywin32. Install with: py -m pip install pywin32"
        ) from exc

    try:
        pypdf_module = importlib.import_module("pypdf")
    except ImportError as exc:
        raise SystemExit("Missing dependency: pypdf. Install with: py -m pip install pypdf") from exc

    pythoncom = pythoncom_module
    win32com_client = win32com_client_module
    PdfReader = cast(PdfReaderFactory, getattr(pypdf_module, "PdfReader"))
    PdfWriter = cast(PdfWriterFactory, getattr(pypdf_module, "PdfWriter"))


def get_pdf_reader() -> PdfReaderFactory:
    if PdfReader is None:
        raise RuntimeError("pypdf is not loaded")
    return PdfReader


def get_pdf_writer() -> PdfWriterFactory:
    if PdfWriter is None:
        raise RuntimeError("pypdf is not loaded")
    return PdfWriter


def get_pythoncom_call(name: str) -> Callable[[], object]:
    if pythoncom is None:
        raise RuntimeError("pythoncom is not loaded")
    return cast(Callable[[], object], getattr(pythoncom, name))


def get_dispatch_ex() -> Callable[[str], ExcelApplicationProtocol]:
    if win32com_client is None:
        raise RuntimeError("win32com.client is not loaded")
    return cast(Callable[[str], ExcelApplicationProtocol], getattr(win32com_client, "DispatchEx"))


@dataclass
class ExportResult:
    source: Path
    do_pdf: Path | None = None
    invoice_pdf: Path | None = None
    error: str | None = None


def configure_logging(log_path: Path) -> None:
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s | %(levelname)s | %(message)s",
        handlers=[
            logging.FileHandler(log_path, encoding="utf-8"),
            logging.StreamHandler(sys.stdout),
        ],
        force=True,
    )


def natural_sort_key(path: Path) -> tuple[str, str]:
    """Stable case-insensitive filename ordering."""
    return (path.name.casefold(), str(path).casefold())


def list_excel_files(folder: Path, recursive: bool) -> list[Path]:
    iterator: Iterable[Path] = folder.rglob("*") if recursive else folder.iterdir()
    files: list[Path] = []

    for path in iterator:
        if not path.is_file():
            continue
        if path.name.startswith("~$"):
            continue
        if path.suffix.casefold() not in SUPPORTED_EXTENSIONS:
            continue
        files.append(path)

    return sorted(files, key=natural_sort_key)


def safe_component(text: str, max_length: int = 120) -> str:
    invalid = '<>:"/\\|?*'
    cleaned = "".join("_" if ch in invalid else ch for ch in text).strip().rstrip(".")
    return (cleaned or "workbook")[:max_length]


def get_sheet(workbook, sheet_name: str):
    for sheet in workbook.Worksheets:
        if str(sheet.Name).casefold() == sheet_name.casefold():
            return sheet
    return None


def get_excel_lock_file(workbook_path: Path) -> Path:
    return workbook_path.with_name(f"~${workbook_path.name}")


def export_sheet_to_pdf(sheet, output_path: Path) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)

    sheet.ExportAsFixedFormat(
        Type=XL_TYPE_PDF,
        Filename=str(output_path),
        Quality=XL_QUALITY_STANDARD,
        IncludeDocProperties=True,
        IgnorePrintAreas=False,
        OpenAfterPublish=False,
    )

    validate_pdf(output_path)


def validate_pdf(pdf_path: Path) -> int:
    if not pdf_path.exists():
        raise RuntimeError(f"PDF was not created: {pdf_path.name}")
    if pdf_path.stat().st_size <= 0:
        raise RuntimeError(f"PDF is empty: {pdf_path.name}")

    reader = get_pdf_reader()(str(pdf_path))
    page_count = len(reader.pages)
    if page_count < 1:
        raise RuntimeError(f"PDF has no pages: {pdf_path.name}")
    return page_count


def merge_pdfs(pdf_paths: list[Path], output_path: Path) -> int:
    if not pdf_paths:
        raise RuntimeError(f"No PDFs available for {output_path.name}")

    writer = get_pdf_writer()()
    expected_pages = 0

    try:
        for pdf_path in pdf_paths:
            reader = get_pdf_reader()(str(pdf_path))
            expected_pages += len(reader.pages)
            for page in reader.pages:
                writer.add_page(page)

        temp_output = output_path.with_suffix(output_path.suffix + ".tmp")
        with temp_output.open("wb") as handle:
            writer.write(handle)

        # Validate before replacing the existing final file.
        actual_pages = validate_pdf(temp_output)
        if actual_pages != expected_pages:
            raise RuntimeError(
                f"Merged page count mismatch for {output_path.name}: "
                f"expected {expected_pages}, got {actual_pages}"
            )

        os.replace(temp_output, output_path)
        return actual_pages
    finally:
        writer.close()


def create_excel_application():
    excel = get_dispatch_ex()("Excel.Application")
    set_excel_application_property(excel, "Visible", False)
    set_excel_application_property(excel, "DisplayAlerts", False)
    set_excel_application_property(excel, "ScreenUpdating", False)
    set_excel_application_property(excel, "EnableEvents", False)
    set_excel_application_property(excel, "AskToUpdateLinks", False)
    set_excel_application_property(excel, "Calculation", XL_CALCULATION_AUTOMATIC)
    set_excel_application_property(
        excel,
        "AutomationSecurity",
        MSO_AUTOMATION_SECURITY_FORCE_DISABLE,
    )
    return excel


def set_excel_application_property(
    excel: ExcelApplicationProtocol,
    name: str,
    value: bool | int,
) -> None:
    try:
        setattr(excel, name, value)
    except Exception:  # noqa: BLE001 - Excel may reject optional app settings
        pass


def prepare_workbook_for_export(excel, workbook) -> None:
    set_excel_application_property(excel, "Calculation", XL_CALCULATION_AUTOMATIC)
    workbook.RefreshAll()
    excel.CalculateFullRebuild()
    while excel.CalculationState != XL_CALCULATION_DONE:
        time.sleep(0.2)


def process_workbook(excel, workbook_path: Path, temp_dir: Path, index: int) -> ExportResult:
    result = ExportResult(source=workbook_path)
    workbook = None

    try:
        lock_file = get_excel_lock_file(workbook_path)
        if lock_file.exists():
            result.error = f"Workbook appears to be open in Excel: {lock_file.name}"
            logging.warning("[%d] Skipping locked workbook: %s", index, workbook_path.name)
            return result

        logging.info("[%d] Opening: %s", index, workbook_path.name)
        workbook = excel.Workbooks.Open(
            Filename=str(workbook_path),
            UpdateLinks=0,
            ReadOnly=False,
            IgnoreReadOnlyRecommended=True,
            AddToMru=False,
            Notify=False,
        )
        prepare_workbook_for_export(excel, workbook)
        workbook.Save()

        stem = safe_component(workbook_path.stem)

        do_sheet = get_sheet(workbook, DO_SHEET_NAME)
        if do_sheet is not None:
            do_path = temp_dir / f"{index:04d}_{stem}_DO.pdf"
            export_sheet_to_pdf(do_sheet, do_path)
            result.do_pdf = do_path
            logging.info("[%d] DO exported", index)
        else:
            logging.warning("[%d] Missing DO sheet: %s", index, workbook_path.name)

        invoice_sheet = get_sheet(workbook, INVOICE_SHEET_NAME)
        if invoice_sheet is not None:
            invoice_path = temp_dir / f"{index:04d}_{stem}_Invoice.pdf"
            export_sheet_to_pdf(invoice_sheet, invoice_path)
            result.invoice_pdf = invoice_path
            logging.info("[%d] Invoice exported", index)
        else:
            logging.warning("[%d] Missing Invoice sheet: %s", index, workbook_path.name)

        if do_sheet is None and invoice_sheet is None:
            result.error = "Missing both DO and Invoice sheets"

    except Exception as exc:  # noqa: BLE001 - batch must continue
        result.error = f"{type(exc).__name__}: {exc}"
        logging.exception("[%d] Failed: %s", index, workbook_path.name)
    finally:
        if workbook is not None:
            try:
                workbook.Close(SaveChanges=False)
            except Exception:  # noqa: BLE001
                logging.exception("[%d] Failed to close workbook: %s", index, workbook_path.name)
        workbook = None
        gc.collect()

    return result


def restore_and_quit_excel(excel) -> None:
    if excel is None:
        return

    try:
        excel.EnableEvents = True
        excel.ScreenUpdating = True
        excel.DisplayAlerts = True
    except Exception:
        pass

    try:
        excel.Quit()
    except Exception:
        logging.exception("Excel did not quit cleanly")


def main() -> int:
    if os.name != "nt":
        relaunch_code = relaunch_with_windows_python_if_needed(Path(__file__))
        if relaunch_code is not None:
            return relaunch_code
        print("This script must run on Windows with Microsoft Excel installed.")
        return 2

    ensure_runtime_dependencies()

    if not TARGET_DIR.exists() or not TARGET_DIR.is_dir():
        print(f"Folder not found: {TARGET_DIR}")
        return 2

    configure_logging(TARGET_DIR / LOG_NAME)
    logging.info("Starting DO/Invoice PDF export")
    logging.info("Folder: %s", TARGET_DIR)

    excel_files = list_excel_files(TARGET_DIR, RECURSIVE)
    if not excel_files:
        logging.error("No Excel files found")
        return 1

    logging.info("Found %d Excel file(s)", len(excel_files))

    _ = get_pythoncom_call("CoInitialize")()
    excel = None
    temp_dir = Path(tempfile.mkdtemp(prefix="do_invoice_pdf_"))
    results: list[ExportResult] = []

    try:
        excel = create_excel_application()

        for index, workbook_path in enumerate(excel_files, start=1):
            results.append(process_workbook(excel, workbook_path, temp_dir, index))

        do_pdfs = [item.do_pdf for item in results if item.do_pdf is not None]
        invoice_pdfs = [item.invoice_pdf for item in results if item.invoice_pdf is not None]

        do_output = TARGET_DIR / DO_OUTPUT_NAME
        invoice_output = TARGET_DIR / INVOICE_OUTPUT_NAME

        if do_pdfs:
            do_pages = merge_pdfs(do_pdfs, do_output)
            logging.info("Created %s (%d pages)", do_output.name, do_pages)
        else:
            logging.warning("No DO PDF was created; existing %s was not overwritten", do_output.name)

        if invoice_pdfs:
            invoice_pages = merge_pdfs(invoice_pdfs, invoice_output)
            logging.info("Created %s (%d pages)", invoice_output.name, invoice_pages)
        else:
            logging.warning(
                "No Invoice PDF was created; existing %s was not overwritten",
                invoice_output.name,
            )

        failed = [item for item in results if item.error]
        logging.info(
            "Finished: %d workbook(s), %d DO export(s), %d Invoice export(s), %d issue(s)",
            len(results),
            len(do_pdfs),
            len(invoice_pdfs),
            len(failed),
        )

        if failed:
            logging.warning("Files with issues:")
            for item in failed:
                logging.warning("- %s: %s", item.source.name, item.error)

        return 0 if not failed else 1

    finally:
        restore_and_quit_excel(excel)
        excel = None
        gc.collect()
        _ = get_pythoncom_call("CoUninitialize")()
        shutil.rmtree(temp_dir, ignore_errors=True)
        time.sleep(0.2)


if __name__ == "__main__":
    raise SystemExit(main())
