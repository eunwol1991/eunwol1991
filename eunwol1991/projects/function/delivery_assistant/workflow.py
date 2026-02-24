import re


def _normalize(value: str) -> str:
    return re.sub(r"\s+", " ", str(value or "").strip().lower())


def _contains(a: str, b: str) -> bool:
    na = _normalize(a)
    nb = _normalize(b)
    if not na or not nb:
        return False
    return na in nb or nb in na


def _is_colored(cell) -> bool:
    fill = getattr(cell, "fill", None)
    if fill is None:
        return False
    if not getattr(fill, "patternType", None):
        return False
    fg = getattr(fill, "fgColor", None)
    if fg is None:
        return False
    if fg.type == "rgb":
        rgb = (fg.rgb or "").upper()
        return rgb not in {"00000000", "FFFFFFFF", "00FFFFFF", "000000"}
    if fg.type in {"indexed", "theme", "auto"}:
        return True
    return False


def find_color_anchor_row(
    ws,
    start_row: int,
    min_col: int = 1,
    max_col: int | None = None,
    min_colored_cells: int = 2,
):
    right = max_col or ws.max_column
    for row in range(start_row, ws.max_row + 1):
        colored = 0
        for col in range(min_col, right + 1):
            if _is_colored(ws.cell(row=row, column=col)):
                colored += 1
                if colored >= min_colored_cells:
                    return row
    return None


def choose_source_row(records: list[dict], values: dict, anchor_row: int | None):
    candidates = records
    if anchor_row is not None:
        candidates = [r for r in records if int(r.get("row_idx") or 0) < anchor_row]

    scored: list[tuple[int, dict]] = []
    for rec in candidates:
        score = 0
        if _contains(values.get("product_code", ""), rec.get("product_code", "")):
            score += 40
        if _contains(values.get("customer", ""), rec.get("customer", "")):
            score += 30
        if _contains(values.get("outlet", ""), rec.get("outlet", "")):
            score += 20
        if _contains(values.get("description", ""), rec.get("description", "")):
            score += 10
        if score > 0:
            scored.append((score, rec))

    if not scored:
        return None
    scored.sort(key=lambda x: (x[0], int(x[1].get("row_idx") or 0)), reverse=True)
    if anchor_row is not None and len(scored) > 1 and scored[0][0] == scored[1][0]:
        raise ValueError(
            "Ambiguous source rows. Please refine product code/customer/outlet."
        )
    return scored[0][1]
