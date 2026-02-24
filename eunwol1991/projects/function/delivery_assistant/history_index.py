from datetime import date, datetime


def parse_record_date(value):
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    text = str(value or "").strip()
    if not text:
        return None
    for fmt in ("%d/%m/%Y", "%d/%m/%y", "%Y-%m-%d", "%d-%m-%Y"):
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            continue
    return None


def build_history_records(
    ws,
    header_row: int,
    columns: dict,
    max_empty_streak: int = 200,
) -> list[dict]:
    records = []
    row = header_row + 1
    empty_streak = 0
    while row <= ws.max_row:
        cust = ws.cell(row=row, column=columns["customer"]).value
        outlet = ws.cell(row=row, column=columns["outlet"]).value
        desc = ws.cell(row=row, column=columns["description"]).value
        if not any([cust, outlet, desc]):
            empty_streak += 1
            if empty_streak >= max_empty_streak:
                break
            row += 1
            continue
        empty_streak = 0
        records.append(
            {
                "row_idx": row,
                "customer": str(cust or ""),
                "outlet": str(outlet or ""),
                "description": str(desc or ""),
                "product_code": str(
                    ws.cell(row=row, column=columns.get("product_code", 0)).value
                    if columns.get("product_code")
                    else ""
                ),
                "record_date": parse_record_date(
                    ws.cell(row=row, column=columns["date"]).value
                ),
            }
        )
        row += 1
    return records
