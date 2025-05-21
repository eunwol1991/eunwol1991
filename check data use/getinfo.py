import pdfplumber
import os
import re

def parse_pdf(file_path, output_path):
    try:
        # Open the specified PDF file
        print(f"Attempting to open PDF file: {file_path}")  # Debug: Print file path being processed
        with pdfplumber.open(file_path) as pdf:
            extracted_data = []
            # Iterate through each page of the PDF and extract relevant text
            for page_num, page in enumerate(pdf.pages):
                text = page.extract_text()
                print(f"Page {page_num + 1} content:\n{text}\n")  # Debug: Print the entire page content

                if text:
                    # Extract specific fields using regex with ranges
                    invoice_no_match = re.search(r"Invoice No\s*[:\-]?\s*([A-Z0-9\s\.-]+)\s*Invoice Date", text, re.IGNORECASE)
                    if invoice_no_match:
                        print(f"Invoice No found: {invoice_no_match.group(1).strip()}")  # Debug: Print matched Invoice No
                    else:
                        print("Invoice No not found")  # Debug: Print when Invoice No not found

                    invoice_date_match = re.search(r"Invoice Date\s*[:\-]?\s*(\d{2}/\d{2}/\d{4})\s*(Terms|Your P/O|Your PO)", text, re.IGNORECASE)
                    if invoice_date_match:
                        print(f"Invoice Date found: {invoice_date_match.group(1)}")  # Debug: Print matched Invoice Date
                    else:
                        print("Invoice Date not found")  # Debug: Print when Invoice Date not found

                    your_po_match = re.search(r"Your P[O/]\s*[:\-]?\s*(.+?)\s*(Bill To|Deliver To|Total)", text, re.IGNORECASE)
                    if your_po_match:
                        print(f"Your PO found: {your_po_match.group(2).strip()}")  # Debug: Print matched Your PO
                    else:
                        print("Your PO not found")  # Debug: Print when Your PO not found

                    total_match = re.search(r"Total\s*[:\-]?\s*\$\s*([\d\s,\.]+)\s*(E\. & O\.E|Terms & Conditions|$)", text, re.IGNORECASE)
                    if total_match:
                        print(f"Total found: {total_match.group(1).replace(' ', '')}")  # Debug: Print matched Total
                    else:
                        print("Total not found")  # Debug: Print when Total not found

                    invoice_no = invoice_no_match.group(1).strip() if invoice_no_match else "Not found"
                    invoice_date = invoice_date_match.group(1) if invoice_date_match else "Not found"
                    your_po = your_po_match.group(1).strip() if your_po_match else "Not found"
                    total = total_match.group(1).replace(' ', '') if total_match else "Not found"

                    # Append extracted data as a formatted string
                    extracted_data.append(f"Invoice No: {invoice_no}\nInvoice Date: {invoice_date}\nYour PO: {your_po}\nTotal: {total}\n\n")

        # Write the extracted data to a text file in a nicely formatted way
        print(f"Writing extracted data to output file: {output_path}")  # Debug: Print output file path
        with open(output_path, 'w', encoding='utf-8') as f:
            for data in extracted_data:
                f.write(data + "\n")

        print(f"Content written to {output_path}\n")

    except Exception as e:
        print(f"Error processing {file_path}: {e}")

if __name__ == "__main__":
    # Program Start
    print("Program Start")

    # Specify the directory and file to parse
    directory = r"C:\Users\User\Dropbox\DO & INV\DO & INV 2024\Joshua - ABR (Swensens)\10. Oct"
    file_name = "ABR 1024 - 001 - INV (Tampine Whse).pdf"
    file_path = os.path.join(directory, file_name)
    output_path = r"C:\Users\User\puython\check data use\text.txt"

    # Parse the specified PDF file and write output to a text file
    parse_pdf(file_path, output_path)

    # Program End
    print("Program End")