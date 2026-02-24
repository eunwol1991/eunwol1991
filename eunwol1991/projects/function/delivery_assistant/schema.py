import re


def normalize_header(value: str) -> str:
    return re.sub(r"[^a-z0-9]+", " ", str(value).strip().lower()).strip()


EXPECTED_HEADERS = {
    "date",
    "month",
    "customer",
    "outlet",
    "product description",
    "product code",
    "qty in pcs",
    "qty in ctns",
    "invoice",
    "invoice #",
}


_ALIASES = {
    "date": {"date", "delivery date", "invoice date"},
    "month": {"month"},
    "year": {"year"},
    "account": {"account"},
    "customer": {"customer", "customer name", "client"},
    "outlet": {"outlet", "outlet name", "deliver to", "ship to"},
    "description": {"product description", "description", "item description"},
    "product_code": {"product code", "code", "item code", "sku"},
    "qty_pcs": {"qty in pcs", "quantity in pcs", "qty pcs"},
    "qty_ctns": {"qty in ctns", "quantity in ctns", "qty ctns"},
    "invoice": {"invoice #", "invoice no", "invoice", "invoice number"},
}


def build_column_map(ws, header_row: int) -> dict[str, int]:
    mapping: dict[str, int] = {}
    for col in range(1, ws.max_column + 1):
        val = ws.cell(row=header_row, column=col).value
        if val is None:
            continue
        norm = normalize_header(str(val))
        for key, aliases in _ALIASES.items():
            if key in mapping:
                continue
            if norm in {normalize_header(a) for a in aliases}:
                mapping[key] = col
                break
    return mapping
