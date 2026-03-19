import sys
import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook
from streamlit.testing.v1 import AppTest


SCRIPT_PATH = Path(__file__).with_name("stock_datagrid.py")


def _create_workbook(path: Path) -> None:
    wb = Workbook()
    ws1 = wb.active
    ws1.title = "Stocks report"
    ws1.append([])
    ws1.append([])
    ws1.append(
        [
            "Supplier",
            "Brand",
            "Product Code",
            "Description",
            "Pack Size",
            "Unit",
            "Expiry Date",
            "Relabel To Date",
            "Daily Update",
            "Stock Qty",
        ]
    )
    ws1.append(
        [
            "Acme",
            "Sauces",
            "P001",
            "Aioli Sauce (Cold)",
            "24 x 50g",
            "ctn",
            "2026-04-01",
            "2026-03-20",
            "",
            10,
        ]
    )
    ws1.append(
        [
            "Acme",
            "Frozen",
            "P002",
            "Bobo Sliced Fish Cake",
            "12 x 100g",
            "ctn",
            "2026-05-01",
            "2026-03-22",
            "",
            8,
        ]
    )
    ws2 = wb.create_sheet("Lai Hock Whse")
    ws2.append([])
    ws2.append([])
    ws2.append(
        [
            "Supplier",
            "Brand",
            "Product Code",
            "Description",
            "Unused",
            "Pack Size",
            "Unit",
            "Expiry Date",
            "Relabel To Date",
            "Stocks Balance",
            "Stock Qty",
        ]
    )
    ws2.append(
        [
            "Acme",
            "Sauces",
            "P001",
            "Aioli Sauce (Cold)",
            "",
            "24 x 50g",
            "ctn",
            "2026-04-01",
            "2026-03-20",
            "",
            5,
        ]
    )
    ws2.append(
        [
            "Acme",
            "Sauces",
            "P003",
            "Brown Sauce (Hot)",
            "",
            "24 x 50g",
            "ctn",
            "2026-06-01",
            "2026-03-24",
            "",
            12,
        ]
    )
    wb.save(path)


def _create_wrapper(path: Path, workbook_path: Path) -> None:
    path.write_text(
        "import sys\n"
        "from pathlib import Path\n"
        "import streamlit as st\n"
        f"script = Path(r'{SCRIPT_PATH}')\n"
        "sys.path.insert(0, str(script.parent))\n"
        "import stock_datagrid\n"
        f"file_path = Path(r'{workbook_path}')\n"
        "st.session_state['stock_file_name'] = file_path.name\n"
        "st.session_state['stock_file_bytes'] = file_path.read_bytes()\n"
        "stock_datagrid.run_stock_page()\n",
        encoding="utf-8",
    )


class StockFilterUiTests(unittest.TestCase):
    def test_stock_filters_render_as_multiselects_like_supplier(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            temp_path = Path(temp_dir)
            workbook_path = temp_path / "stock.xlsx"
            wrapper_path = temp_path / "wrapper.py"
            _create_workbook(workbook_path)
            _create_wrapper(wrapper_path, workbook_path)

            sys.path.insert(0, str(SCRIPT_PATH.parent))
            at = AppTest.from_file(str(wrapper_path), default_timeout=60)
            at.run()

            multiselect_labels = [x.label for x in at.multiselect]
            text_input_labels = [x.label for x in at.text_input]

            self.assertIn("Description（去括号后）", multiselect_labels)
            self.assertIn("Product Code", multiselect_labels)
            self.assertIn("Remark（来自描述括号）", multiselect_labels)
            self.assertNotIn("Description（去括号后）", text_input_labels)
            self.assertNotIn("Product Code", text_input_labels)
            self.assertNotIn("Remark（来自描述括号）", text_input_labels)


if __name__ == "__main__":
    unittest.main()
