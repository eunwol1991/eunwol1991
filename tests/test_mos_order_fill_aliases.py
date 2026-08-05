import importlib.util
from pathlib import Path
from types import ModuleType
from typing import Callable, cast


def _load_mos_order_fill() -> ModuleType:
    path = Path(__file__).resolve().parents[1] / "eunwol1991/projects/check data use/mos_order_fill.py"
    spec = importlib.util.spec_from_file_location("mos_order_fill", path)
    assert spec is not None
    assert spec.loader is not None
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def test_braised_po_item_names_map_to_excel_columns():
    mos_order_fill = _load_mos_order_fill()
    map_item = cast(Callable[[str], str], getattr(mos_order_fill, "_map_item"))

    assert map_item("Braised sauce 500g") == "Braised Duck Sauce (12 x 500g)"
    assert map_item("Braised sauce sachet 20g") == "Braised Duck Sauce (300 x 20g)"
    assert map_item("Braised chilli sachet 15g") == "Chilli Sauce (300 x 15g)"
