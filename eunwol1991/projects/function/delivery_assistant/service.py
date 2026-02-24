from datetime import date

from .excel_io import make_backup, open_workbook, open_workbook_for_read
from .history_index import build_history_records
from .matching import confidence_level, rank_candidates
from .row_insert import insert_with_template
from .schema import EXPECTED_HEADERS, build_column_map
from .sheet_locator import detect_header_row, find_sheet_ci
from .workflow import choose_source_row, find_color_anchor_row


def load_context(file_path: str) -> dict:
    wb = open_workbook_for_read(file_path)
    try:
        ws = find_sheet_ci(wb, "delivery details")
        header_row = detect_header_row(ws, EXPECTED_HEADERS)
        columns = build_column_map(ws, header_row)
        for required in (
            "customer",
            "outlet",
            "description",
            "qty_pcs",
            "qty_ctns",
            "invoice",
        ):
            if required not in columns:
                raise ValueError(f"Missing required column: {required}")
        records = build_history_records(
            ws,
            header_row,
            {
                "date": columns.get("date", columns["invoice"]),
                "customer": columns["customer"],
                "outlet": columns["outlet"],
                "description": columns["description"],
                "product_code": columns.get("product_code"),
            },
        )
        anchor_row = find_color_anchor_row(
            ws,
            start_row=header_row + 1,
            min_col=1,
            max_col=ws.max_column,
            min_colored_cells=2,
        )
        return {
            "file_path": file_path,
            "sheet_title": ws.title,
            "header_row": header_row,
            "columns": columns,
            "records": records,
            "anchor_row": anchor_row,
        }
    finally:
        wb.close()


def suggest(context: dict, query: dict, limit: int = 8) -> list[dict]:
    ranked = rank_candidates(query, context["records"], limit=limit)
    second = ranked[1]["score"] if len(ranked) > 1 else 0.0
    if ranked:
        ranked[0]["confidence"] = confidence_level(ranked[0]["score"], second)
    return ranked


def preview_insert(context: dict, values: dict) -> dict:
    header_row = context["header_row"]
    columns = context["columns"]
    today = date.today()
    anchor_row = context.get("anchor_row")
    source = choose_source_row(context.get("records", []), values, anchor_row)

    if anchor_row is not None:
        insert_row = int(anchor_row) + 1
    else:
        insert_row = header_row + 1

    if source is not None:
        template_row = int(source["row_idx"])
    else:
        template_row = insert_row + 1

    user_values = {
        columns["description"]: values["description"],
        columns["customer"]: values["customer"],
        columns["outlet"]: values["outlet"],
        columns["qty_pcs"]: values["qty_pcs"],
        columns["qty_ctns"]: values["qty_ctns"],
        columns["invoice"]: values["invoice"],
    }
    if "product_code" in columns:
        user_values[columns["product_code"]] = values.get("product_code", "")
    if "date" in columns:
        user_values[columns["date"]] = today.strftime("%d/%m/%Y")
    if "month" in columns:
        user_values[columns["month"]] = today.month
    if "year" in columns:
        user_values[columns["year"]] = today.year

    return {
        "insert_row": insert_row,
        "template_row": template_row,
        "user_values": user_values,
        "source_row": int(source["row_idx"]) if source is not None else None,
        "anchor_row": anchor_row,
    }


def apply_insert(context: dict, plan: dict) -> str:
    backup_path = make_backup(context["file_path"])
    wb = open_workbook(context["file_path"])
    try:
        ws = find_sheet_ci(wb, context.get("sheet_title", "delivery details"))
        result = insert_with_template(
            ws=ws,
            insert_row=plan["insert_row"],
            template_row=plan["template_row"],
            min_col=1,
            max_col=ws.max_column,
            user_values=plan["user_values"],
        )
        for col in result["formula_cols"]:
            val = ws.cell(row=plan["insert_row"], column=col).value
            if not (isinstance(val, str) and val.startswith("=")):
                raise ValueError(
                    "Formula safety check failed: detected value paste in formula column."
                )
        wb.save(context["file_path"])
    finally:
        wb.close()

    refreshed = load_context(context["file_path"])
    context.clear()
    context.update(refreshed)
    return backup_path
