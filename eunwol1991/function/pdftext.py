import os
import fitz  # PyMuPDF


def print_pdf_text(pdf_path: str) -> None:
    pdf_path = pdf_path.strip().strip('"')

    if not os.path.isfile(pdf_path):
        raise FileNotFoundError(f"PDF not found: {pdf_path}")

    doc = fitz.open(pdf_path)
    try:
        print("\n" + "=" * 80)
        print(f"FILE: {pdf_path}")
        print(f"PAGES: {doc.page_count}")
        print("=" * 80 + "\n")

        for i in range(doc.page_count):
            page = doc.load_page(i)
            text = page.get_text("text") or ""

            print(f"\n===== PAGE {i+1} / {doc.page_count} =====\n")
            print(text)

        print("\n" + "=" * 80)
        print("Done ✅\n")

    finally:
        doc.close()


if __name__ == "__main__":
    print("Paste PDF full path (type 'q' to quit).")
    while True:
        user_input = input("PDF path: ").strip()
        if user_input.lower() in ("q", "quit", "exit"):
            break
        try:
            print_pdf_text(user_input)
        except Exception as e:
            print(f"\nError: {e}\n")
