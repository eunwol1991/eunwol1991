import importlib.util
import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook


MODULE_PATH = Path(__file__).with_name("invoice_pdf_excel_check.py")


def _load_module():
    spec = importlib.util.spec_from_file_location(
        "invoice_pdf_excel_check", MODULE_PATH
    )
    if spec is None or spec.loader is None:
        raise RuntimeError(f"Unable to load module from {MODULE_PATH}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def _write_delivery_details_workbook(path: Path, rows: list[list[object]]) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Delivery details"
    ws.append([])
    ws.append([])
    ws.append([])
    ws.append(
        [
            "Invoice #",
            "Date",
            "Total Value Inclusive GST",
            "Customer",
            "Account",
            "Outlet",
        ]
    )
    for row in rows:
        ws.append(row)
    wb.save(path)


class InvoicePdfExcelCheckTests(unittest.TestCase):
    def test_load_excel_records_flags_same_customer_with_multiple_accounts_in_same_month(
        self,
    ):
        module = _load_module()

        with tempfile.TemporaryDirectory() as temp_dir:
            workbook_path = Path(temp_dir) / "delivery_details.xlsx"
            _write_delivery_details_workbook(
                workbook_path,
                [
                    [
                        "INV 0126-001",
                        "05/01/2026",
                        "10.90",
                        "Melvin Cafe",
                        "melvin",
                        "Outlet A",
                    ],
                    [
                        "INV 0126-002",
                        "08/01/2026",
                        "21.80",
                        "Melvin Cafe",
                        "anthony",
                        "Outlet B",
                    ],
                    [
                        "INV 0126-003",
                        "09/01/2026",
                        "15.00",
                        "Other Cafe",
                        "stable",
                        "Outlet C",
                    ],
                ],
            )

            records = module.load_excel_records(str(workbook_path), "0126")

        self.assertEqual(
            records["INV 0126-001"]["excel_status"],
            "Customer-Account mismatch within month",
        )
        self.assertEqual(
            records["INV 0126-002"]["excel_status"],
            "Customer-Account mismatch within month",
        )
        self.assertEqual(records["INV 0126-003"]["excel_status"], "OK")
        self.assertEqual(records["INV 0126-001"]["excel_accounts"], ["melvin"])
        self.assertEqual(records["INV 0126-002"]["excel_accounts"], ["anthony"])


if __name__ == "__main__":
    unittest.main()
