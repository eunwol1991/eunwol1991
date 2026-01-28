import os
import re
import threading
import queue
from decimal import Decimal, InvalidOperation
from datetime import datetime, date
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinter.scrolledtext import ScrolledText

try:
    import pandas as pd
except Exception:  # pragma: no cover
    pd = None

try:
    import fitz  # PyMuPDF
except Exception:  # pragma: no cover
    fitz = None

BASE_DIR_DEFAULT = r"C:\Users\jhunj\Dropbox\DO & INV\DO & INV 2026"
SPECIAL_CANADIAN_DIR = r"C:\Users\jhunj\Dropbox\for jj\Outlets PDF"
SHEET_NAME = "Delivery details"
HEADER_ROW_INDEX = 3  # A4 -> 0-based row index
IGNORE_DIR_NAME = "Melvin - Stuff'd"

AMOUNT_RE = re.compile(r"([0-9]{1,3}(?:,[0-9]{3})*(?:\.[0-9]{2})?|[0-9]+(?:\.[0-9]{2})?)")
DATE_RE = re.compile(r"(\d{1,2}/\d{1,2}/\d{2,4})")
INVOICE_NO_PATTERNS = [
    re.compile(
        r"Invoice\s*No\s*(?:Invoice\s*)?[:#]?\s*([A-Z0-9.]{2,}\s*\d{4}\s*-\s*\d{3}(?:\s*-\s*[A-Za-z])?)",
        re.I,
    ),
]
DATE_PATTERNS = [
    re.compile(r"Invoice\s*Date\s*[:#]?\s*(\d{1,2}/\d{1,2}/\d{2,4})", re.I),
    re.compile(r"Date\s*[:#]?\s*(\d{1,2}/\d{1,2}/\d{2,4})", re.I),
]
SUBTOTAL_PATTERN = re.compile(
    r"Sub\s*Total\s*\$?\s*([0-9,]+(?:\.[0-9]{2})?)", re.I
)
GST_PATTERN = re.compile(
    r"(?:Add\s*GST|GST)\s*9%?\s*\$?\s*([0-9,]+(?:\.[0-9]{2})?)", re.I
)
MONEY_RE = re.compile(r"(\d{1,3}(?:,\d{3})*\.\d{2}|\d+\.\d{2})")
INV_NAME_HINT = re.compile(r"\b[A-Z]{2,}\s*\d{4}\s*-\s*\d{3}\b", re.I)


def normalize_col_name(name: str) -> str:
    return re.sub(r"[^a-z0-9]+", "", str(name).strip().lower())


def normalize_invoice_no(value: str) -> str:
    return re.sub(r"\s+", " ", value or "").strip().lower()


def compact_text(value: str) -> str:
    return re.sub(r"\s+", "", value or "").strip().lower()


def should_ignore_invoice(invoice_no: str) -> bool:
    text = re.sub(r"\s+", " ", invoice_no or "").strip().upper()
    if not text:
        return True
    if re.search(r"\bCN\b", text) or text.startswith("CN"):
        return True
    parts = re.findall(r"[A-Z0-9]+", text)
    if parts and parts[0] == "SD":
        return True
    return False


def to_decimal(value) -> Decimal | None:
    if value is None:
        return None
    if pd is not None and pd.isna(value):
        return None
    if isinstance(value, Decimal):
        return value
    if isinstance(value, (int, float)):
        return Decimal(str(value))
    text = str(value).strip()
    if not text or text.lower() == "nan":
        return None
    text = text.replace("$", "").replace(",", "")
    try:
        return Decimal(text)
    except InvalidOperation:
        return None


def format_money(value: Decimal | None) -> str:
    if value is None:
        return ""
    return f"{value:.2f}"


def parse_date_value(value) -> date | None:
    if value is None:
        return None
    if pd is not None and pd.isna(value):
        return None
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    text = str(value).strip()
    if not text:
        return None
    m = DATE_RE.search(text)
    if not m:
        return None
    d, mth, y = m.group(1).split("/")
    try:
        day = int(d)
        month = int(mth)
        year = int(y)
        if year < 100:
            year += 2000
        return date(year, month, day)
    except ValueError:
        return None


def format_date_value(value: date | None) -> str:
    if value is None:
        return ""
    return value.strftime("%d/%m/%Y")


def find_invoice_col(columns: list[str]) -> str | None:
    exact = {"invoice", "invoice#", "invoiceno", "invoicenumber"}
    for col in columns:
        if normalize_col_name(col) in exact:
            return col
    for col in columns:
        if "invoice" in normalize_col_name(col):
            return col
    return None


def find_date_col(columns: list[str]) -> str | None:
    preferred = ["date", "invoicedate", "deliverydate"]
    for pref in preferred:
        for col in columns:
            if normalize_col_name(col) == pref:
                return col
    for col in columns:
        if "date" in normalize_col_name(col):
            return col
    return None


def find_total_col(columns: list[str]) -> str | None:
    exact = {
        "totalvalueinclusivegst",
        "totalinclusivegst",
        "totalvalueinclgst",
        "totalvalueinclusivegst$",
    }
    for col in columns:
        if normalize_col_name(col) in exact:
            return col
    for col in columns:
        norm = normalize_col_name(col)
        if "total" in norm and "gst" in norm and "inclusive" in norm:
            return col
    for col in columns:
        norm = normalize_col_name(col)
        if "total" in norm and "gst" in norm:
            return col
    return None


def extract_amount_from_line(line: str) -> Decimal | None:
    text = line.strip()
    if re.search(r"\bFOC\b", text, re.I):
        return Decimal("0")
    if re.search(r"\$\s*-\b", text) or text == "-":
        return Decimal("0")
    matches = list(AMOUNT_RE.finditer(text))
    if not matches:
        return None
    for match in reversed(matches):
        start, end = match.span(1)
        if end < len(text) and text[end] == "%":
            continue
        return to_decimal(match.group(1))
    return None


def is_invoice_pdf_filename(name_lower: str) -> bool:
    if not name_lower.endswith(".pdf"):
        return False
    has_inv = bool(re.search(r"\b(inv|invoice)\b", name_lower)) or ("inv" in name_lower)
    has_do = bool(re.search(r"\bdo\b", name_lower)) or ("delivery order" in name_lower) or ("d/o" in name_lower)
    if has_do and not has_inv:
        return False
    if INV_NAME_HINT.search(name_lower):
        return True
    return has_inv


def build_invoice_matcher(invoice_no: str) -> re.Pattern:
    tokens = re.findall(r"[A-Za-z0-9]+", invoice_no)
    if not tokens:
        return re.compile(r"$^")
    pattern = r"(?<![A-Z0-9])" + re.escape(tokens[0])
    for tok in tokens[1:]:
        pattern += r"(?:[\s._-]*" + re.escape(tok) + r")"
    pattern += r"(?![A-Z0-9])"
    return re.compile(pattern, re.I)


def collect_money_amounts(lines: list[str]) -> list[Decimal]:
    amounts: list[Decimal] = []
    for line in lines:
        if re.search(r"\bFOC\b", line, re.I):
            amounts.append(Decimal("0"))
            continue
        for match in MONEY_RE.finditer(line):
            start, end = match.span(1)
            if end < len(line) and line[end] == "%":
                continue
            val = to_decimal(match.group(1))
            if val is not None:
                amounts.append(val)
    return amounts


def collect_money_amounts_from_layout(doc, min_x_ratio: float, exclude_y: list[float]) -> list[Decimal]:
    amounts: list[Decimal] = []
    for page in doc:
        page_width = float(page.rect.width)
        text_dict = page.get_text("dict")
        for block in text_dict.get("blocks", []):
            for line in block.get("lines", []):
                for span in line.get("spans", []):
                    text = (span.get("text") or "").strip()
                    if not text:
                        continue
                    x = float(span["bbox"][0])
                    y = float(span["bbox"][1])
                    if x < page_width * min_x_ratio:
                        continue
                    if any(abs(y - ey) <= 6 for ey in exclude_y):
                        continue
                    if re.search(r"\bFOC\b", text, re.I):
                        amounts.append(Decimal("0"))
                        continue
                    for match in MONEY_RE.finditer(text):
                        start, end = match.span(1)
                        if end < len(text) and text[end] == "%":
                            continue
                        val = to_decimal(match.group(1))
                        if val is not None:
                            amounts.append(val)
    return amounts


def infer_subtotal_gst(
    amounts: list[Decimal],
    min_value: Decimal | None = None,
) -> tuple[Decimal | None, Decimal | None, Decimal | None]:
    if min_value is not None:
        amounts = [a for a in amounts if a >= min_value]
    if not amounts:
        return None, None, None
    total_candidate = max(amounts)
    best = None
    for i, a in enumerate(amounts):
        for j, b in enumerate(amounts):
            if i == j:
                continue
            if abs((a + b) - total_candidate) <= Decimal("0.05"):
                subtotal = max(a, b)
                gst = min(a, b)
                best = (subtotal, gst, total_candidate)
                return best
    return None, None, total_candidate


def extract_amounts_from_layout(doc) -> dict:
    spans: list[dict] = []
    for pno, page in enumerate(doc):
        page_width = float(page.rect.width)
        text_dict = page.get_text("dict")
        for block in text_dict.get("blocks", []):
            for line in block.get("lines", []):
                for span in line.get("spans", []):
                    text = (span.get("text") or "").strip()
                    if not text:
                        continue
                    spans.append(
                        {
                            "text": text,
                            "x": float(span["bbox"][0]),
                            "y": float(span["bbox"][1]),
                            "page": pno,
                            "width": page_width,
                        }
                    )

    if not spans:
        return {}

    amounts: list[dict] = []
    for s in spans:
        for match in MONEY_RE.finditer(s["text"]):
            val = to_decimal(match.group(1))
            if val is None:
                continue
            amounts.append(
                {
                    "value": val,
                    "x": s["x"],
                    "y": s["y"],
                    "page": s["page"],
                    "width": s["width"],
                }
            )

    labels: dict[str, dict] = {}
    header_y: list[float] = []
    for s in spans:
        lower = s["text"].lower()
        if "subtotal" in lower:
            labels["subtotal"] = s
        elif "add gst" in lower or ("gst" in lower and "9" in lower):
            labels["gst"] = s
        elif lower.strip().startswith("total"):
            labels["total"] = s
        if any(k in lower for k in ["price", "per unit", "carton", "quantity", "amount"]):
            header_y.append(s["y"])

    def pick_amount(label_key: str) -> Decimal | None:
        label = labels.get(label_key)
        if not label:
            return None
        page_amounts = [a for a in amounts if a["page"] == label["page"]]
        if not page_amounts:
            return None
        right = [a for a in page_amounts if a["x"] >= label["width"] * 0.70]
        if not right:
            return None
        right = [a for a in right if a["x"] >= label["x"] - 2] or right
        right = [a for a in right if not any(abs(a["y"] - hy) <= 6 for hy in header_y)]

        span_right = [
            s for s in spans
            if s["page"] == label["page"]
            and s["x"] >= label["width"] * 0.70
            and abs(s["y"] - label["y"]) <= 10
        ]
        for s in span_right:
            text = s["text"].strip()
            if text in {"-", "$", "$-"} or re.fullmatch(r"\$\s*-\s*", text):
                return Decimal("0")
            if re.search(r"\bFOC\b", text, re.I):
                return Decimal("0")

        def best(cands: list[dict]) -> Decimal:
            chosen = min(cands, key=lambda a: (abs(a["y"] - label["y"]), -a["x"]))
            return chosen["value"]

        same_line = [a for a in right if abs(a["y"] - label["y"]) <= 3]
        if same_line:
            return best(same_line)
        near = [a for a in right if abs(a["y"] - label["y"]) <= 10]
        if near:
            return best(near)
        return None

    subtotal = pick_amount("subtotal")
    gst = pick_amount("gst")
    total = pick_amount("total")
    return {"subtotal": subtotal, "gst": gst, "total": total, "labels": labels, "header_y": header_y}


def _is_gst_line(line: str) -> bool:
    lower = line.lower()
    if "gst" not in lower:
        return False
    if "add gst" in lower:
        return True
    if "gst 9" in lower or "gst9" in lower or "9%" in lower:
        return True
    if "reg" in lower:
        return False
    return True


def parse_pdf_invoice(path: str) -> tuple[dict, str | None]:
    if not fitz:
        return {}, "PyMuPDF not installed"
    try:
        doc = fitz.open(path)
    except Exception as exc:
        return {}, str(exc)
    try:
        text = "\n".join(page.get_text("text") or "" for page in doc)
        lines = [ln.strip() for ln in text.splitlines() if ln.strip()]
        preview_lines = lines[:50]

        invoice_no = None
        pdf_date = None
        subtotal = None
        gst = None
        total_line = None

        for pattern in INVOICE_NO_PATTERNS:
            m = pattern.search(text)
            if m:
                invoice_no = m.group(1).strip()
                break

        for pattern in DATE_PATTERNS:
            m = pattern.search(text)
            if m:
                pdf_date = parse_date_value(m.group(1))
                break

        if subtotal is None:
            subtotal_matches = SUBTOTAL_PATTERN.findall(text)
            if subtotal_matches:
                subtotal = to_decimal(subtotal_matches[-1])

        if gst is None:
            gst_matches = GST_PATTERN.findall(text)
            if gst_matches:
                gst = to_decimal(gst_matches[-1])

        def scan_amount_after_label(idx: int, prefer_next_if_percent: bool, max_lookahead: int = 20) -> Decimal | None:
            for j in range(idx + 1, min(len(lines), idx + max_lookahead + 1)):
                amt = extract_amount_from_line(lines[j])
                if amt is not None:
                    return amt
                if prefer_next_if_percent and "%" in lines[j]:
                    continue
            return None

        for idx, line in enumerate(lines):
            lower = line.lower()
            if invoice_no is None and "invoice" in lower and "no" in lower:
                m = re.search(r"invoice\s*no\s*[:#]?\s*(.+)", line, re.I)
                if m and m.group(1).strip():
                    invoice_no = m.group(1).strip()
                elif idx + 1 < len(lines):
                    candidate = lines[idx + 1].strip()
                    if candidate.lower().startswith("invoice "):
                        candidate = candidate[8:].strip()
                    invoice_no = candidate

            if pdf_date is None and "date" in lower:
                m = DATE_RE.search(line)
                if m:
                    pdf_date = parse_date_value(m.group(1))
                elif idx + 1 < len(lines):
                    m2 = DATE_RE.search(lines[idx + 1])
                    if m2:
                        pdf_date = parse_date_value(m2.group(1))

            if subtotal is None and re.search(r"\bsub\s*total\b", lower):
                subtotal = scan_amount_after_label(idx, prefer_next_if_percent=False)

            if gst is None and _is_gst_line(line):
                gst = scan_amount_after_label(idx, prefer_next_if_percent=True)

            if total_line is None and lower.startswith("total"):
                total_line = extract_amount_from_line(line)

        if subtotal is None and gst is None and total_line is not None:
            if total_line == Decimal("0"):
                subtotal = Decimal("0")
                gst = Decimal("0")

        if invoice_no is None:
            m = re.search(r"invoice\s*no\s*[:#]?\s*(.+)", text, re.I)
            if m:
                invoice_no = m.group(1).strip()

        if pdf_date is None:
            m = re.search(r"date\s*[:#]?\s*(\d{1,2}/\d{1,2}/\d{2,4})", text, re.I)
            if m:
                pdf_date = parse_date_value(m.group(1))

        if subtotal is None:
            m = SUBTOTAL_PATTERN.search(text)
            if m:
                subtotal = to_decimal(m.group(1))

        if gst is None:
            m = GST_PATTERN.search(text)
            if m:
                gst = to_decimal(m.group(1))

        is_abr = "abr" in os.path.basename(path).lower() or bool(re.search(r"\bABR\b", text))
        layout_info = None
        if is_abr:
            layout_info = extract_amounts_from_layout(doc)
            if layout_info.get("subtotal") is not None:
                subtotal = layout_info.get("subtotal")
            if layout_info.get("gst") is not None:
                gst = layout_info.get("gst")
            if layout_info.get("total") is not None:
                total_line = layout_info.get("total")

            header_y = layout_info.get("header_y", []) if layout_info else []
            amounts = collect_money_amounts_from_layout(doc, 0.70, header_y)
            inferred_subtotal, inferred_gst, inferred_total = infer_subtotal_gst(
                amounts, min_value=Decimal("1.00")
            )
            if (subtotal is None or gst is None) and inferred_subtotal is not None and inferred_gst is not None:
                subtotal = inferred_subtotal
                gst = inferred_gst
            elif subtotal is not None and gst is not None and inferred_subtotal is not None and inferred_gst is not None:
                if inferred_total is not None and abs((subtotal + gst) - inferred_total) > Decimal("0.05"):
                    subtotal = inferred_subtotal
                    gst = inferred_gst
        else:
            amounts = collect_money_amounts(lines)
            inferred_subtotal, inferred_gst, inferred_total = infer_subtotal_gst(amounts)
            if (subtotal is None or gst is None) and inferred_subtotal is not None and inferred_gst is not None:
                subtotal = inferred_subtotal
                gst = inferred_gst
            elif subtotal is not None and gst is not None and inferred_subtotal is not None and inferred_gst is not None:
                if inferred_total is not None and abs((subtotal + gst) - inferred_total) > Decimal("0.05"):
                    subtotal = inferred_subtotal
                    gst = inferred_gst

        return {
            "invoice_no": invoice_no,
            "date": pdf_date,
            "subtotal": subtotal,
            "gst": gst,
            "total_line": total_line,
            "debug_lines": preview_lines,
            "debug_labels": list((layout_info or {}).get("labels", {}).values()) if layout_info else [],
        }, None
    finally:
        doc.close()


def load_excel_records(path: str, month_code: str) -> dict:
    if pd is None:
        raise RuntimeError("pandas not installed")
    df = pd.read_excel(path, sheet_name=SHEET_NAME, header=HEADER_ROW_INDEX)
    if df.empty:
        return {}

    columns = [str(c) for c in df.columns]
    invoice_col = find_invoice_col(columns)
    date_col = find_date_col(columns)
    total_col = find_total_col(columns)

    if not invoice_col or not date_col or not total_col:
        missing = []
        if not invoice_col:
            missing.append("Invoice #")
        if not date_col:
            missing.append("Date")
        if not total_col:
            missing.append("Total Value Inclusive GST")
        raise ValueError(f"Missing columns: {', '.join(missing)}. Columns: {columns}")

    records: dict[str, dict] = {}
    month_code_lower = month_code.lower()

    for _, row in df.iterrows():
        inv_raw = row.get(invoice_col)
        if pd.isna(inv_raw):
            continue
        inv = str(inv_raw).strip()
        if not inv or inv.lower() == "nan":
            continue
        if month_code_lower and month_code_lower not in inv.lower():
            continue
        if should_ignore_invoice(inv):
            continue

        excel_date = parse_date_value(row.get(date_col))
        excel_total = to_decimal(row.get(total_col)) or Decimal("0")

        rec = records.get(inv)
        if not rec:
            rec = {
                "invoice_no": inv,
                "date_values": set(),
                "excel_total": Decimal("0"),
                "line_count": 0,
            }
            records[inv] = rec

        if excel_date:
            rec["date_values"].add(excel_date)
        rec["excel_total"] += excel_total
        rec["line_count"] += 1

    for inv, rec in records.items():
        date_values = sorted(rec["date_values"])
        if len(date_values) == 1:
            rec["excel_date"] = date_values[0]
            rec["excel_status"] = "OK"
        else:
            rec["excel_date"] = date_values[0] if date_values else None
            rec["excel_status"] = "Excel Date mismatch"
        rec["excel_dates"] = date_values

    return records


def scan_candidate_files(base_dirs: list[str], month_code: str) -> list[tuple[str, str, str]]:
    candidates: list[tuple[str, str, str]] = []
    ignore_lower = IGNORE_DIR_NAME.lower()
    month_lower = month_code.lower()

    for base_dir in base_dirs:
        if not base_dir or not os.path.isdir(base_dir):
            continue
        for root, dirs, files in os.walk(base_dir):
            dirs[:] = [d for d in dirs if d.lower() != ignore_lower]
            for fname in files:
                fname_lower = fname.lower()
                if not is_invoice_pdf_filename(fname_lower):
                    continue
                if month_lower and month_lower not in fname_lower:
                    continue
                full_path = os.path.join(root, fname)
                candidates.append((fname_lower, compact_text(fname_lower), full_path))

    return candidates


def select_best_candidate(invoice_no: str, candidates: list[tuple[str, str, str]]):
    invoice_lower = normalize_invoice_no(invoice_no)
    invoice_compact = compact_text(invoice_no)
    matcher = build_invoice_matcher(invoice_no)

    matches = [
        c for c in candidates
        if matcher.search(c[0])
    ]
    if not matches:
        return None, False, False, []

    revised = [m for m in matches if "revised" in m[0]]
    multiple = len(matches) > 1

    def score(item):
        fname_lower, fname_compact, path = item
        base = os.path.splitext(os.path.basename(path))[0].lower()
        base_compact = compact_text(base)
        ext = os.path.splitext(path)[1].lower()
        score_val = 0
        if ext == ".pdf":
            score_val += 100
        if base == invoice_lower:
            score_val += 80
        if base_compact == invoice_compact:
            score_val += 70
        if invoice_lower in base:
            score_val += 30
        score_val -= len(base)
        return score_val

    if revised:
        chosen = sorted(revised, key=score, reverse=True)[0]
        return chosen[2], multiple, True, [m[2] for m in matches]

    chosen = sorted(matches, key=score, reverse=True)[0]
    return chosen[2], multiple, False, [m[2] for m in matches]


def build_results(excel_records: dict, candidates: list[tuple[str, str, str]], on_progress=None):
    results = []
    invoice_list = sorted(excel_records.keys())
    total = len(invoice_list)

    for idx, inv in enumerate(invoice_list, start=1):
        if on_progress:
            on_progress(total, idx, f"Checking {inv}")
        rec = excel_records[inv]
        excel_total = rec["excel_total"]
        excel_date = rec.get("excel_date")
        excel_status = rec.get("excel_status", "OK")

        path, multiple, revised_used, all_candidates = select_best_candidate(inv, candidates)
        pdf_found = bool(path)
        pdf_info = {}
        pdf_error = None
        pdf_total = None
        pdf_date = None
        pdf_invoice = None
        diff = None

        if pdf_found:
            pdf_info, pdf_error = parse_pdf_invoice(path)
            pdf_invoice = pdf_info.get("invoice_no")
            pdf_date = pdf_info.get("date")
            subtotal = pdf_info.get("subtotal")
            gst = pdf_info.get("gst")
            if subtotal is not None and gst is not None:
                pdf_total = subtotal + gst
            if pdf_total is not None:
                diff = pdf_total - excel_total

        status = "OK"
        if excel_status == "Excel Date mismatch":
            status = "Excel Date mismatch"
        elif not pdf_found:
            status = "Missing PDF"
        elif pdf_error or not pdf_info or pdf_invoice is None or pdf_date is None or pdf_total is None:
            status = "PDF parse fail"
        else:
            if normalize_invoice_no(pdf_invoice) != normalize_invoice_no(inv):
                status = "Invoice mismatch"
            elif excel_date and pdf_date and pdf_date != excel_date:
                status = "Date mismatch"
            elif pdf_total is not None and abs(pdf_total - excel_total) > Decimal("0.05"):
                status = "Total mismatch"

        results.append({
            "invoice_no": inv,
            "excel_date": excel_date,
            "excel_dates": rec.get("excel_dates", []),
            "excel_total": excel_total,
            "line_count": rec.get("line_count", 0),
            "excel_status": excel_status,
            "pdf_found": pdf_found,
            "pdf_path": path or "",
            "pdf_invoice": pdf_invoice,
            "pdf_date": pdf_date,
            "pdf_subtotal": pdf_info.get("subtotal") if pdf_info else None,
            "pdf_gst": pdf_info.get("gst") if pdf_info else None,
            "pdf_total": pdf_total,
            "pdf_total_line": pdf_info.get("total_line") if pdf_info else None,
            "diff": diff,
            "status": status,
            "candidates": all_candidates,
            "revised_used": revised_used,
            "debug_lines": pdf_info.get("debug_lines") if pdf_info else [],
            "debug_labels": pdf_info.get("debug_labels") if pdf_info else [],
            "pdf_error": pdf_error,
        })

    return results


class InvoiceValidatorApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Invoice PDF vs Excel Validator")
        self.geometry("1300x720")
        self.minsize(1100, 600)

        self.excel_path_var = tk.StringVar()
        self.month_var = tk.StringVar()
        self.status_filter_var = tk.StringVar(value="All")
        self.status_var = tk.StringVar(value="Ready")
        self.progress_var = tk.DoubleVar(value=0)

        self._queue = queue.Queue()
        self._running = False
        self._data = {}
        self._results = []
        self._debug_records = []

        self._build_ui()
        self.after(150, self._process_queue)

    def _build_ui(self):
        style = ttk.Style(self)
        style.configure("Treeview", rowheight=24)

        top = ttk.Frame(self, padding=10)
        top.pack(fill=tk.X)

        ttk.Label(top, text="Excel file:").grid(row=0, column=0, sticky=tk.W)
        ttk.Entry(top, textvariable=self.excel_path_var, width=80).grid(
            row=0, column=1, sticky=tk.W, padx=6
        )
        ttk.Button(top, text="Browse...", command=self._choose_excel).grid(
            row=0, column=2, padx=4
        )

        ttk.Label(top, text=f"Base dir (fixed): {BASE_DIR_DEFAULT}").grid(
            row=1, column=0, columnspan=3, sticky=tk.W, padx=(0, 6), pady=(6, 0)
        )
        ttk.Label(top, text=f"Canadian special dir: {SPECIAL_CANADIAN_DIR}").grid(
            row=2, column=0, columnspan=3, sticky=tk.W, padx=(0, 6), pady=(2, 0)
        )

        ttk.Label(top, text="Month code (e.g. 0126):").grid(
            row=3, column=0, sticky=tk.W, pady=(6, 0)
        )
        ttk.Entry(top, textvariable=self.month_var, width=12).grid(
            row=3, column=1, sticky=tk.W, padx=6, pady=(6, 0)
        )
        self.run_btn = ttk.Button(top, text="Run", command=self._start)
        self.run_btn.grid(row=3, column=2, padx=4, pady=(6, 0))

        top.columnconfigure(1, weight=1)

        filter_bar = ttk.Frame(self, padding=(10, 0, 10, 6))
        filter_bar.pack(fill=tk.X)
        ttk.Label(filter_bar, text="Status filter:").pack(side=tk.LEFT)
        self.status_filter = ttk.Combobox(
            filter_bar,
            textvariable=self.status_filter_var,
            width=32,
            state="readonly",
            values=["All"],
        )
        self.status_filter.pack(side=tk.LEFT, padx=(6, 12))
        self.status_filter.bind("<<ComboboxSelected>>", self._apply_filter)
        ttk.Button(filter_bar, text="Debug", command=self._open_debug).pack(side=tk.LEFT)

        bar = ttk.Frame(self, padding=(10, 0, 10, 6))
        bar.pack(fill=tk.X)

        self.progress = ttk.Progressbar(
            bar, orient="horizontal", length=300, mode="determinate", variable=self.progress_var
        )
        self.progress.pack(side=tk.LEFT)
        ttk.Label(bar, textvariable=self.status_var).pack(side=tk.LEFT, padx=10)

        tree_frame = ttk.Frame(self)
        tree_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 10))

        columns = (
            "invoice_no",
            "excel_date",
            "excel_total",
            "pdf_found",
            "pdf_path",
            "pdf_date",
            "pdf_total",
            "diff",
            "status",
        )
        self.tree = ttk.Treeview(tree_frame, columns=columns, show="headings")
        self.tree.heading("invoice_no", text="Invoice #")
        self.tree.heading("excel_date", text="Excel Date")
        self.tree.heading("excel_total", text="Excel Total Inc GST")
        self.tree.heading("pdf_found", text="PDF Found")
        self.tree.heading("pdf_path", text="PDF Path")
        self.tree.heading("pdf_date", text="PDF Date")
        self.tree.heading("pdf_total", text="PDF Total")
        self.tree.heading("diff", text="Diff")
        self.tree.heading("status", text="Status")

        self.tree.column("invoice_no", width=160, anchor=tk.W)
        self.tree.column("excel_date", width=110, anchor=tk.CENTER)
        self.tree.column("excel_total", width=140, anchor=tk.E)
        self.tree.column("pdf_found", width=80, anchor=tk.CENTER)
        self.tree.column("pdf_path", width=320, anchor=tk.W)
        self.tree.column("pdf_date", width=110, anchor=tk.CENTER)
        self.tree.column("pdf_total", width=120, anchor=tk.E)
        self.tree.column("diff", width=80, anchor=tk.E)
        self.tree.column("status", width=220, anchor=tk.W)

        self.tree.tag_configure("status_total", foreground="#B00020", background="#FFECEC")
        self.tree.tag_configure("status_date", foreground="#1E40AF", background="#E8F0FF")

        vsb = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=self.tree.yview)
        hsb = ttk.Scrollbar(tree_frame, orient=tk.HORIZONTAL, command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")

        tree_frame.rowconfigure(0, weight=1)
        tree_frame.columnconfigure(0, weight=1)

        self.tree.bind("<Double-1>", self._open_detail)

    def _choose_excel(self):
        path = filedialog.askopenfilename(
            title="Select Excel file",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
        )
        if path:
            self.excel_path_var.set(path)

    def _start(self):
        if self._running:
            return
        excel_path = self.excel_path_var.get().strip()
        month_code = self.month_var.get().strip()

        if not excel_path or not os.path.isfile(excel_path):
            messagebox.showerror("Error", "Please select a valid Excel file.")
            return
        if not re.fullmatch(r"\d{4}", month_code):
            messagebox.showerror("Error", "Month code must be 4 digits (e.g. 0126).")
            return
        if pd is None:
            messagebox.showerror("Error", "pandas is not installed.")
            return
        if fitz is None:
            messagebox.showerror("Error", "PyMuPDF (fitz) is not installed.")
            return

        self._running = True
        self.run_btn.configure(state=tk.DISABLED)
        self.progress_var.set(0)
        self.status_var.set("Loading Excel...")
        self._clear_tree()

        worker = threading.Thread(
            target=self._run_job,
            args=(excel_path, month_code),
            daemon=True,
        )
        worker.start()

    def _run_job(self, excel_path: str, month_code: str):
        try:
            excel_records = load_excel_records(excel_path, month_code)
            if not excel_records:
                self._queue.put(("error", "No invoices found for the given month code."))
                return
            self._queue.put(("status", "Scanning files..."))
            base_dirs = [BASE_DIR_DEFAULT]
            if SPECIAL_CANADIAN_DIR and os.path.isdir(SPECIAL_CANADIAN_DIR):
                base_dirs.append(SPECIAL_CANADIAN_DIR)
            candidates = scan_candidate_files(base_dirs, month_code)

            def on_progress(total, current, note):
                self._queue.put(("progress", total, current, note))

            results = build_results(excel_records, candidates, on_progress=on_progress)
            self._queue.put(("done", results))
        except Exception as exc:
            self._queue.put(("error", str(exc)))

    def _process_queue(self):
        try:
            while True:
                msg = self._queue.get_nowait()
                kind = msg[0]
                if kind == "status":
                    self.status_var.set(msg[1])
                elif kind == "progress":
                    total, current, note = msg[1], msg[2], msg[3]
                    self.progress.configure(maximum=total)
                    self.progress_var.set(current)
                    self.status_var.set(note)
                elif kind == "done":
                    self._running = False
                    self.run_btn.configure(state=tk.NORMAL)
                    results = msg[1]
                    self._load_results(results)
                    self.status_var.set(f"Done. {len(results)} invoices.")
                elif kind == "error":
                    self._running = False
                    self.run_btn.configure(state=tk.NORMAL)
                    self.status_var.set("Error")
                    messagebox.showerror("Error", msg[1])
        except queue.Empty:
            pass
        self.after(150, self._process_queue)

    def _clear_tree(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        self._data = {}

    def _load_results(self, results: list[dict]):
        self._results = results
        self._debug_records = [r for r in results if r.get("status") == "PDF parse fail"]
        self._refresh_status_filter()
        self._apply_filter()

    def _open_debug(self):
        top = tk.Toplevel(self)
        top.title("Parse Fail Debug")
        top.geometry("900x600")

        txt = ScrolledText(top, wrap=tk.WORD)
        txt.pack(fill=tk.BOTH, expand=True)

        lines = []
        if not self._debug_records:
            lines.append("No parse failures.")
        else:
            for rec in self._debug_records:
                lines.append("=" * 60)
                lines.append(f"PDF Path: {rec.get('pdf_path', '')}")
                lines.append(f"Invoice #: {rec.get('invoice_no', '')}")
                lines.append(f"PDF Invoice No: {rec.get('pdf_invoice') or ''}")
                lines.append(f"PDF Date: {format_date_value(rec.get('pdf_date'))}")
                lines.append(f"Subtotal: {format_money(rec.get('pdf_subtotal'))}")
                lines.append(f"GST: {format_money(rec.get('pdf_gst'))}")
                lines.append(f"Total line: {format_money(rec.get('pdf_total_line'))}")
                lines.append("")
                lines.append("Label spans:")
                for lbl in rec.get("debug_labels", []):
                    try:
                        lines.append(f"  p{lbl.get('page')} x={lbl.get('x'):.1f} y={lbl.get('y'):.1f} text={lbl.get('text')}")
                    except Exception:
                        lines.append(f"  {lbl}")
                lines.append("")
                lines.append("Text preview (first 50 lines):")
                for ln in rec.get("debug_lines", [])[:50]:
                    lines.append(ln)
                lines.append("")

        txt.insert(tk.END, "\n".join(lines))
        txt.configure(state="disabled")

    def _refresh_status_filter(self):
        statuses = sorted({rec.get("status", "") for rec in self._results if rec.get("status", "")})
        values = ["All"] + statuses
        self.status_filter.configure(values=values)
        if self.status_filter_var.get() not in values:
            self.status_filter_var.set("All")

    def _apply_filter(self, _event=None):
        self._clear_tree()
        selected = self.status_filter_var.get()
        for rec in self._results:
            status = rec.get("status", "")
            if selected != "All" and status != selected:
                continue
            tags = ()
            if status == "Total mismatch":
                tags = ("status_total",)
            elif status in ("Date mismatch", "Excel Date mismatch"):
                tags = ("status_date",)
            values = (
                rec["invoice_no"],
                format_date_value(rec.get("excel_date")),
                format_money(rec.get("excel_total")),
                "Y" if rec.get("pdf_found") else "N",
                rec.get("pdf_path", ""),
                format_date_value(rec.get("pdf_date")),
                format_money(rec.get("pdf_total")),
                format_money(rec.get("diff")),
                rec.get("status", ""),
            )
            item_id = self.tree.insert("", tk.END, values=values, tags=tags)
            self._data[item_id] = rec

    def _open_detail(self, _event=None):
        item = self.tree.selection()
        if not item:
            return
        rec = self._data.get(item[0])
        if not rec:
            return

        top = tk.Toplevel(self)
        top.title(f"Invoice Detail - {rec.get('invoice_no', '')}")
        top.geometry("720x520")

        txt = ScrolledText(top, wrap=tk.WORD)
        txt.pack(fill=tk.BOTH, expand=True)

        lines = []
        lines.append(f"Invoice #: {rec.get('invoice_no', '')}")
        lines.append(f"Excel Date: {format_date_value(rec.get('excel_date'))}")
        if rec.get("excel_dates"):
            dates = ", ".join(format_date_value(d) for d in rec.get("excel_dates", []))
            lines.append(f"Excel Date(s): {dates}")
        lines.append(f"Excel Total Inc GST: {format_money(rec.get('excel_total'))}")
        lines.append(f"Line count: {rec.get('line_count', 0)}")
        lines.append(f"Excel status: {rec.get('excel_status', '')}")
        lines.append(f"Excel Date consistent: {'Yes' if rec.get('excel_status') == 'OK' else 'No'}")
        lines.append("")
        lines.append(f"PDF Found: {'Y' if rec.get('pdf_found') else 'N'}")
        lines.append(f"PDF Path: {rec.get('pdf_path', '')}")
        lines.append(f"PDF Invoice No: {rec.get('pdf_invoice') or ''}")
        lines.append(f"PDF Date: {format_date_value(rec.get('pdf_date'))}")
        lines.append(f"PDF Subtotal: {format_money(rec.get('pdf_subtotal'))}")
        lines.append(f"PDF GST: {format_money(rec.get('pdf_gst'))}")
        lines.append(f"PDF Total (Subtotal+GST): {format_money(rec.get('pdf_total'))}")
        lines.append(f"PDF Total line (ref): {format_money(rec.get('pdf_total_line'))}")
        lines.append(f"Diff: {format_money(rec.get('diff'))}")
        lines.append(f"Status: {rec.get('status', '')}")
        lines.append(f"Revised chosen: {'Yes' if rec.get('revised_used') else 'No'}")
        if rec.get("status") and rec.get("status") != "OK":
            lines.append(f"Mismatch reason: {rec.get('status')}")
        if rec.get("pdf_error"):
            lines.append(f"PDF Error: {rec.get('pdf_error')}")
        if rec.get("status") == "PDF parse fail":
            lines.append("")
            lines.append("Parse preview (first 30 lines):")
            for ln in rec.get("debug_lines", [])[:30]:
                lines.append(ln)
            lines.append("")
            lines.append("Parse fields:")
            lines.append(f"invoice_no={rec.get('pdf_invoice') or ''}")
            lines.append(f"date={format_date_value(rec.get('pdf_date'))}")
            lines.append(f"subtotal={format_money(rec.get('pdf_subtotal'))}")
            lines.append(f"gst={format_money(rec.get('pdf_gst'))}")
            if rec.get("debug_labels"):
                lines.append("")
                lines.append("Label spans:")
                for lbl in rec.get("debug_labels", []):
                    try:
                        lines.append(f"p{lbl.get('page')} x={lbl.get('x'):.1f} y={lbl.get('y'):.1f} text={lbl.get('text')}")
                    except Exception:
                        lines.append(str(lbl))

        txt.insert(tk.END, "\n".join(lines))
        txt.configure(state="disabled")


if __name__ == "__main__":
    app = InvoiceValidatorApp()
    app.mainloop()
