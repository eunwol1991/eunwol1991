import importlib.util
import sys
import unittest
from pathlib import Path

import pandas as pd


MODULE_PATH = Path(__file__).with_name("stock_datagrid.py")


def _load_module():
    sys.path.insert(0, str(MODULE_PATH.parent))
    spec = importlib.util.spec_from_file_location("stock_datagrid", MODULE_PATH)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"Unable to load module from {MODULE_PATH}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


class StockFilterQueryTests(unittest.TestCase):
    def test_description_selection_filters_by_base_description(self):
        module = _load_module()
        df = pd.DataFrame(
            {
                "description": [
                    "Aioli Sauce (Cold)",
                    "Bobo Sliced Fish Cake",
                    "Brown Sauce",
                ]
            }
        )

        filtered = module._apply_filter_state(df, {"f_desc": ["Brown Sauce"]})

        self.assertEqual(filtered["description"].tolist(), ["Brown Sauce"])

    def test_product_code_selection_filters_by_exact_choice(self):
        module = _load_module()
        df = pd.DataFrame({"product_code": ["AB-001", "AC-002", "ZX-100"]})

        filtered = module._apply_filter_state(df, {"f_code": ["AC-002"]})

        self.assertEqual(filtered["product_code"].tolist(), ["AC-002"])

    def test_remark_selection_can_exclude_matching_rows(self):
        module = _load_module()
        df = pd.DataFrame(
            {
                "description": [
                    "Aioli Sauce (Cold)",
                    "Brown Sauce (Hot)",
                    "Fish Cake",
                ]
            }
        )

        filtered = module._apply_filter_state(
            df,
            {"f_remark": ["Hot"], "f_remark_ex": True},
        )

        self.assertEqual(
            filtered["description"].tolist(), ["Aioli Sauce (Cold)", "Fish Cake"]
        )

    def test_brand_selection_can_exclude_matching_rows(self):
        module = _load_module()
        df = pd.DataFrame(
            {
                "brand": ["Alpha", "Beta", "Gamma"],
                "product_code": ["A-1", "B-1", "G-1"],
            }
        )

        filtered = module._apply_filter_state(
            df,
            {"f_brand": ["Beta"], "f_brand_ex": True},
        )

        self.assertEqual(filtered["brand"].tolist(), ["Alpha", "Gamma"])


if __name__ == "__main__":
    unittest.main()
