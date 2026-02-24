from .schema import normalize_header


def find_sheet_ci(workbook, target_name: str):
    target = target_name.strip().casefold()
    matches = [
        ws for ws in workbook.worksheets if str(ws.title).strip().casefold() == target
    ]
    if not matches:
        raise KeyError(f"Sheet not found: {target_name}")
    if len(matches) > 1:
        raise KeyError(f"Ambiguous sheet name: {target_name}")
    return matches[0]


def detect_header_row(ws, expected_headers: set[str], scan_rows: int = 20) -> int:
    expected = {normalize_header(h) for h in expected_headers}
    best_row = 1
    best_score = -1
    for row_idx in range(1, scan_rows + 1):
        rows = ws.iter_rows(min_row=row_idx, max_row=row_idx)
        row_cells = next(rows, ())
        row_vals = [
            normalize_header(cell.value) for cell in row_cells if cell.value is not None
        ]
        if not row_vals:
            continue
        score = 0
        for raw in row_vals:
            if raw in expected:
                score += 2
                continue
            for hdr in expected:
                if raw == hdr or raw in hdr or hdr in raw:
                    score += 1
                    break
        if score > best_score:
            best_score = score
            best_row = row_idx
    return best_row
