from __future__ import annotations

import os
import re
import sys
from pathlib import Path

try:
    from pypdf import PdfReader, PdfWriter
except ImportError:
    print("Missing package: pypdf")
    print("Install it with: pip install pypdf")
    input("Press Enter to exit...")
    raise SystemExit(1)


INPUT_FOLDER = Path(r"C:\Users\jhunj\Dropbox\for jj\Doc to print - JJ")
OUTPUT_FILENAME = "Merged_All_PDFs.pdf"


def natural_sort_key(path: Path) -> list[object]:
    """Sort file names naturally, e.g. 2.pdf before 10.pdf."""
    parts = re.split(r"(\d+)", path.name.casefold())
    return [int(part) if part.isdigit() else part for part in parts]


def merge_pdfs(folder: Path, output_filename: str) -> None:
    if not folder.exists() or not folder.is_dir():
        raise FileNotFoundError(f"Folder not found: {folder}")

    output_path = folder / output_filename
    temp_output_path = folder / f".{output_filename}.tmp"

    pdf_files = sorted(
        (
            path
            for path in folder.iterdir()
            if path.is_file()
            and path.suffix.casefold() == ".pdf"
            and path.name.casefold() != output_filename.casefold()
            and path.name.casefold() != temp_output_path.name.casefold()
        ),
        key=natural_sort_key,
    )

    if not pdf_files:
        raise FileNotFoundError(f"No PDF files found in: {folder}")

    writer = PdfWriter()
    merged_files: list[str] = []
    skipped_files: list[str] = []
    total_pages = 0

    print(f"Found {len(pdf_files)} PDF file(s).")
    print("Merge order:")

    for index, pdf_path in enumerate(pdf_files, start=1):
        try:
            with pdf_path.open("rb") as source_file:
                reader = PdfReader(source_file, strict=False)

                if reader.is_encrypted:
                    try:
                        decrypt_result = reader.decrypt("")
                    except Exception:
                        decrypt_result = 0

                    if not decrypt_result:
                        raise PermissionError("PDF is password-protected")

                page_count = len(reader.pages)
                if page_count == 0:
                    raise ValueError("PDF contains no pages")

                for page in reader.pages:
                    writer.add_page(page)

            total_pages += page_count
            merged_files.append(pdf_path.name)
            print(f"[{index}/{len(pdf_files)}] OK   {pdf_path.name} ({page_count} page(s))")

        except Exception as exc:
            skipped_files.append(f"{pdf_path.name}: {exc}")
            print(f"[{index}/{len(pdf_files)}] SKIP {pdf_path.name} - {exc}")

    if not merged_files:
        raise RuntimeError("No PDF could be merged.")

    writer.add_metadata({
        "/Title": "Merged PDF",
        "/Creator": "Python pypdf",
    })

    try:
        with temp_output_path.open("wb") as output_file:
            writer.write(output_file)

        # Re-open the result before replacing the old output.
        verification_reader = PdfReader(str(temp_output_path), strict=False)
        verified_pages = len(verification_reader.pages)

        if verified_pages != total_pages:
            raise RuntimeError(
                f"Verification failed: expected {total_pages} pages, got {verified_pages} pages"
            )

        os.replace(temp_output_path, output_path)

    finally:
        if temp_output_path.exists():
            try:
                temp_output_path.unlink()
            except OSError:
                pass

    print("\nMerge completed.")
    print(f"Output: {output_path}")
    print(f"Merged files: {len(merged_files)}")
    print(f"Total pages: {total_pages}")

    if skipped_files:
        print("\nSkipped files:")
        for item in skipped_files:
            print(f"- {item}")


if __name__ == "__main__":
    try:
        merge_pdfs(INPUT_FOLDER, OUTPUT_FILENAME)
    except Exception as error:
        print(f"\nERROR: {error}")
        sys.exit_code = 1
    else:
        sys.exit_code = 0

    input("\nPress Enter to exit...")
    raise SystemExit(sys.exit_code)
    