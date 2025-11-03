"""
MOS Order → Excel 自动填表脚本（Savori三行版 + 关键字门店匹配 + 日期PO排除修正版）
依赖：pip install pymupdf openpyxl regex
"""

from pathlib import Path
import re
import sys
import shutil
import fitz                     # PyMuPDF
import openpyxl as xl
from openpyxl.styles import PatternFill
from openpyxl.utils import column_index_from_string

# === 1. 配置区 ============================================================ #
EXCEL_PATH = Path(
    r"C:\Users\jhunj\Dropbox\for jj\mos order\Order Summary For MOS - JJ.xlsx")
SHEET_NAME = "MOS Format"
PDF_DIR = Path(r"C:\Users\jhunj\Dropbox\for jj\mos_pdfs")  # 待处理 PDF 文件夹
PO_COL = "AK"
ROW_START = 39
ROW_END = 73
COL_START = "C"
COL_END = "AF"
HIGHLIGHT = PatternFill(fill_type="solid", fgColor="00FFFCD7")
ARCHIVE_DIR = Path(r"C:\Users\jhunj\Dropbox\for jj\mos order")
# —— PO 号下限（含）——
PO_MIN = 600000


# ======== 门店别名（键已全部小写） =======================================
STORE_ALIAS = {
    "mos burger 100 am (38)":                 "MOS - 100AM",
    "mos burger 18 tai seng (67)":            "MOS - 18 Tai Seng",
    "mos burger amk hub (24)":                "MOS - Ang Mo Kio Hub",
    "mos burger bedok mall (42)":             "MOS - Bedok Mall",
    "mos burger jewel changi airport (59)":   "MOS - Café Jewel Changi Airport*",
    "mos burger causeway point (33)":         "MOS - Causeway Point",
    "mos burger changi city point (80)":      "MOS - Changi City Point",
    "mos burger city square mall (69)":       "MOS - City Square Mall",
    "mos burger compass one (49)":            "MOS - Compass One",
    "mos burger eastpoint (64)":              "MOS - Eastpoint Mall",
    "mos burger bukit gombak mrt (71)":       "MOS - Express Bukit Gombak MRT",
    "mos burger holland village mrt (70)":    "MOS - Express Holland Village MRT",
    "mos burger fusionopolis one (55)":       "MOS - Fusionopolis One",
    "mos burger harbourfront centre (31)":    "MOS - Harbourfront Centre",
    "mos burger heartland mall kovan (58)":   "MOS - Heartland Mall Kovan",
    "mos burger hillion mall (56)":           "MOS - Hillion Mall",
    "mos burger hougang mall (63)":           "MOS - Hougang Mall",
    "mos burger ion orchard (79)":            "MOS - ION Orchard",
    "mos burger jem (39)":                    "MOS - JEM",
    "mos burger jurong point (14)":           "MOS - Jurong Point",
    "mos burger kampung admiralty (53)":      "MOS - Kampung Admiralty",
    "mos burger marina bay link mall (81)":   "MOS - Marina Bay Link Mall",
    "mos burger merlion park (72)":           "MOS - Merlion Park",
    "mos burger millenia walk (30)":          "MOS - Millenia Walk",
    "mos burger nex (36)":                    "MOS - NEX",
    "mos burger northpoint city (50)":        "MOS - Northpoint",
    "mos burger our tampines hub (57)":       "MOS - Our Tampines Hub",
    "mos burger raffles city (20)":           "MOS - Raffles City",
    "mos burger seletar mall (75)":           "MOS - Seletar Mall",
    "mos burger tampines mall (35)":          "MOS - Tampines Mall",
    "mos burger the centrepoint (51)":        "MOS - The Centrepoint",
    "mos burger tiong bahru plaza (48)":      "MOS - Tiong Bahru Plaza",
    "mos burger toa payoh hdb hub (12)":      "MOS - Toa Payoh HDB Hub",
    "mos burger waterway point (62)":         "MOS - Waterway Point",
    "mos burger white sands (47)":            "MOS - White Sands",
}

# ======== 物品别名（键已全部小写） =======================================
ITEM_ALIAS = {
    "ikeda japanese chicken cutlet":                "Ikeda Japanese Chicken Cutlet (6 x 1.1kg)",
    "japanese chicken katsu 70g":                   "CS TAY Japanese Chicken Katsu 70G (6 x 1kg)",
    "crispy fried 3 joint wing":                    "CS Tay Crispy Fried Chicken Wings 3 Joints (1132) (6 x 1.1kg)",
    "cs tay crispy chicken patty":                  "CS Tay Crispy Chicken Patty (4 x 2.5kg)",
    "eb kranch alaska pollock fingers":             "Kranch Alaska Pollock Finger (5 x 900g)",
    "battered natural onion rings":                 "Salud Battered Natural Onion Ring (6 x 1kg)",
    "rosti hashbrown":                              "Rosti Hashbrown (10 x 1kg)",
    "prawn noodle soup":                            "Prawn Noodle Soup (10 x 1kg)",
    "prawn noodle sauce":                           "Prawn Noodle Sauce (10 x 1kg)",
    "tomato capsicum soup":                         "Tomato Capsicum Soup (10 x 1kg)",
    "par cooked beef patty 120g":                   "Cooked Hamburg Beef Patty 120G (10 x 720g)",
    "demi glace sauce":                             "Demi Glace Sauce (10 x 1kg)",
    "hoisin sauce":                                 "Hoisin Sauce (10 x 1kg)",
    "seaweed egg drop soup":                        "Seaweed Egg Drop Soup (10 x 1kg)",
    "chilli crab sauce":                            "Chilli Crab Sauce (10 x 1kg)",
    "coconut sauce":                                "Coconut Sauce (12 x 500g)",
    "coleslaw":                                     "Coleslaw (5 x 2kg)",
    "anchor minidish unsalted butter 7g":           "Anchor Minidish Unsalted Butter (1 x 144 x 7g)",
    "mini mantou, plain":                           "Mini Mantou Plain (24 x 180g)",
    "bobo sliced fish cake 1kg":                    "BoBo Sliced Fish Cake (1kg)",
    "chicken honey baked ham sliced":               "Chicken Honey Baked Ham Sliced (12 x 1kg)",
    "mozzarella cheese 5g pearl iqf":               "Mozzerella Cheese 5G Pearls IQF (8 x 1kg)",
    "premium smoked salmon sliced r trout":         "Premium Smoked Salmon Sliced (R Trout) (10 x 1kg)",
    "hard boiled egg 10pcs only":                   "Hard Boiled Egg (18 x 10pcs)",
    "fried shallot 1kg":                            "Fried Shallot (10 x 1kg)",
    "jalapenos nachos sliced":                      "Jalapeno Pepper Sliced (6 x 2.95kg)",
    "parsley flakes":                               "Parsley Leaves (6 x 5gm)",
    "toasted coconut flakes":                       "Toasted Coconut Flakes (36 x 200g)",
    "red sauce":                                    "Red Sauce (10 x 1kg)",
    "sriracha mayo":                                "J-Lek Sriracha Hot Chilli Sauce (12 x 455ml)",
}

# ========================================================================== #
UNIT_RE = r"(?:ctn|ctns|pkt|pkts|tin|tins|can|cans|box|boxes|btl|btls|pc|pcs)"
_PLURAL_MAP = {"ctn": "ctns", "pkt": "pkts", "tin": "tins",
               "can": "cans", "box": "boxes", "btl": "btls", "pc": "pcs"}
_SINGULAR_MAP = {v: k for k, v in _PLURAL_MAP.items()}


def _normalize_uom(uom: str) -> str:
    u = (uom or "").strip().lower()
    if u in _SINGULAR_MAP:
        return _SINGULAR_MAP[u]
    return u


def _format_qty_uom(qty: int, uom: str) -> str:
    base = _normalize_uom(uom)
    if qty and qty > 1:
        plural = _PLURAL_MAP.get(base, base + "s")
        return f"{qty} {plural}"
    return f"{qty} {base}"


def _norm_spaces(s: str) -> str:
    return re.sub(r"\s+", " ", s).strip()


def _looks_like_date_fragment(s: str) -> bool:
    """粗判是否像日期：含 '/' 或出现 20xx 年份痕迹"""
    if not s:
        return False
    return ("/" in s) or bool(re.search(r"20\d{2}", s))

# ---------- Item / Store normalization ---------- #


def _norm_item_key(desc: str) -> str:
    s = desc.lower()
    s = re.sub(r"\(.*?\)", "", s)
    s = re.sub(r"[^a-z0-9\s,.-]+", " ", s)
    s = _norm_spaces(s)
    return s


def _map_item(desc: str) -> str:
    key = _norm_item_key(desc)
    if key in ITEM_ALIAS:
        return ITEM_ALIAS[key]
    k2 = key.replace(",", "")
    if k2 in ITEM_ALIAS:
        return ITEM_ALIAS[k2]
    for k in ITEM_ALIAS:
        if k in key or key in k:
            return ITEM_ALIAS[k]
    return desc


def _norm_store_text(s: str) -> str:
    s = s.lower()
    s = re.sub(r"\(.*?\)", "", s)
    s = s.replace("mos", " ").replace("burger", " ").replace("mosburger", " ")
    s = s.replace("mos -", " ").replace("mos–", " ").replace("-", " ")
    s = re.sub(r"[^a-z0-9\s]", " ", s)
    return _norm_spaces(s)


def resolve_store(pdf_store_line: str, store_row_keys: list[str]) -> str | None:
    alias_hit = STORE_ALIAS.get(pdf_store_line.lower())
    if alias_hit:
        return alias_hit
    pdf_norm = _norm_store_text(pdf_store_line)
    excel_norm_map = {_norm_store_text(k): k for k in store_row_keys}
    candidates = [(len(e), orig) for e, orig in excel_norm_map.items()
                  if pdf_norm in e or e in pdf_norm]
    if candidates:
        candidates.sort(reverse=True)
        return candidates[0][1]
    words = [w for w in pdf_norm.split() if len(w) >= 3]
    for e, orig in excel_norm_map.items():
        if all(w in e for w in words):
            return orig
    return None


def _ensure_dir(p: Path):
    try:
        p.mkdir(parents=True, exist_ok=True)
    except Exception:
        pass


def _safe_move_file(src: Path, dest_dir: Path):
    try:
        _ensure_dir(dest_dir)
        target = dest_dir / src.name
        if target.exists():
            stem, suf, i = src.stem, src.suffix, 1
            while (dest_dir / f"{stem} ({i}){suf}").exists():
                i += 1
            target = dest_dir / f"{stem} ({i}){suf}"
        shutil.move(str(src), str(target))
    except Exception as e:
        print(f"[!] 无法移动文件 {src} -> {dest_dir}: {e}", file=sys.stderr)

# ---------- 模糊列匹配（补齐缺失函数） ---------- #


def _tokenize_item(s: str) -> set[str]:
    """Tokenize item text for fuzzy matching: lowercased words, length>=3."""
    norm = _norm_item_key(s)
    toks = [w for w in norm.split() if len(w) >= 3]
    return set(toks)


def find_best_item_col(desc: str, item_col: dict[str, int]) -> tuple[int | None, str | None]:
    """
    Fuzzy match: choose Excel header with maximum token overlap.
    - Accept if unique max overlap >= 1.
    - If tie with max==1 -> ambiguous -> None.
    - If tie with max>=2, break tie by overlap char length; still tie -> None.
    """
    pdf_tokens = _tokenize_item(desc)
    if not pdf_tokens:
        return None, None

    # (score, overlap_chars, header, col)
    scored: list[tuple[int, int, str, int]] = []
    for header, col in item_col.items():
        header_tokens = _tokenize_item(header)
        overlap = pdf_tokens & header_tokens
        score = len(overlap)
        if score > 0:
            overlap_chars = sum(len(t) for t in overlap)
            scored.append((score, overlap_chars, header, col))

    if not scored:
        return None, None

    scored.sort(key=lambda x: (x[0], x[1]), reverse=True)
    top_score = scored[0][0]
    top = [s for s in scored if s[0] == top_score]

    if len(top) == 1 and top_score >= 1:
        return top[0][3], top[0][2]

    if top_score >= 2:
        top.sort(key=lambda x: x[1], reverse=True)
        if len(top) == 1 or top[0][1] > top[1][1]:
            return top[0][3], top[0][2]
        return None, None

    return None, None


def _extract_po_from_lines(lines: list[str]) -> str:
    joined = " \n ".join(lines)

    # 1) 'Approved by' 附近
    m = re.search(
        r"(?is)approved\s*by\s*[:\-]?[^\d]{0,80}?(\d\D?\d\D?\d\D?\d\D?\d\D?\d)", joined)
    if m:
        raw = m.group(1)
        if not _looks_like_date_fragment(raw):
            digits = re.sub(r"\D", "", raw)
            if len(digits) == 6 and re.fullmatch(r"\d{6}", digits) and int(digits) >= PO_MIN:
                return digits

    # 2) 逐行窗口
    for i, ln in enumerate(lines):
        if re.search(r"(?i)approved\s*by", ln):
            window = " ".join(lines[i:i+4])
            m2 = re.search(r"(?is)\b(\d\D?\d\D?\d\D?\d\D?\d\D?\d)\b", window)
            if m2:
                raw2 = m2.group(1)
                if not _looks_like_date_fragment(raw2):
                    digits2 = re.sub(r"\D", "", raw2)
                    if len(digits2) == 6 and re.fullmatch(r"\d{6}", digits2) and int(digits2) >= PO_MIN:
                        return digits2

    # 3) 'To' 规则（后一两行恰好6位）
    for i in range(len(lines) - 1):
        if lines[i].strip().lower() == "to":
            nxt = " ".join(lines[i+1:i+3])
            m3 = re.search(r"\b(\d{6})\b", nxt)
            if m3:
                cand = m3.group(1)
                if int(cand) >= PO_MIN:
                    return cand

    # 4) 全文兜底：距离 'Approved by' 最近的独立6位数字，且 >= PO_MIN
    approved_pos = re.search(r"(?i)approved\s*by", joined)
    pos = approved_pos.start() if approved_pos else 0
    candidates = [(m.start(), m.group(0))
                  for m in re.finditer(r"\b\d{6}\b", joined)]
    candidates = [c for c in candidates if int(c[1]) >= PO_MIN]
    if candidates:
        candidates.sort(key=lambda t: abs(t[0] - pos))
        return candidates[0][1]

    return ""


# ---------- PDF解析 ---------- #
def parse_pdf(pdf_path: Path):
    """
    适配 Savori 的“3行配对”版式：
      行1：纯数字货号
      行2：描述
      行3：数量+单位（如 '2 ctn' / '1 pkt' / '1 can'）
    门店：独立一行 'MOS Burger ... (nn)'（也支持关键字匹配）
    """
    pages = []
    with fitz.open(pdf_path) as doc:
        for page in doc:
            raw = page.get_text("text") or ""
            pages.append([ln.strip() for ln in raw.splitlines() if ln.strip()])
    return pages

# ---------- 单页信息抽取 ---------- #


def extract_orders_from_lines(lines: list[str]) -> dict:
    """从单页 lines 中抽取: store_line, po, items(list[{desc, qty, uom}])"""
    # 门店行（原样）
    store_line = None
    for ln in lines:
        if re.fullmatch(r"(?i)mos burger.+\(\d+\)", ln):
            store_line = ln
            break

    po = ""
    for i, ln in enumerate(lines):
        if "approved by:" in ln.lower():
            m = re.search(r"\b(\d{6})\b", ln)
            if m and int(m.group(1)) >= PO_MIN:
                po = m.group(1)
                break
            if i + 1 < len(lines) and re.fullmatch(r"\d{6}", lines[i + 1].strip()):
                if int(lines[i + 1].strip()) >= PO_MIN:
                    po = lines[i + 1].strip()
                    break

    if not po:
        for i in range(len(lines) - 1):
            if lines[i].strip().lower() == "to" and re.fullmatch(r"\d{6}", lines[i + 1].strip()):
                if int(lines[i + 1].strip()) >= PO_MIN:
                    po = lines[i + 1].strip()
                    break

    # 兜底（会再次套用 PO_MIN & 日期过滤）
    po2 = _extract_po_from_lines(lines)
    if po2:
        po = po2

    # 明细：三行配对
    items = []
    qty_re = re.compile(rf"^(\d+)\s+({UNIT_RE})$", re.I)
    i = 0
    while i < len(lines) - 2:
        if re.fullmatch(r"\d{5,}", lines[i]):   # 货号（5位以上纯数字）
            desc = lines[i + 1]
            m_qty = qty_re.fullmatch(lines[i + 2])
            if m_qty:
                items.append({
                    "desc": desc,
                    "qty": int(m_qty.group(1)),
                    "uom": m_qty.group(2).lower(),
                })
                i += 3
                continue
        i += 1

    return {"store_line": store_line, "po": po, "items": items}

# ---------- Excel表头映射 ---------- #


def build_mappings(ws):
    """生成 {店名: 行号}, {物品名: 列号}"""
    def _norm_store_for_excel(s: str) -> str:
        s = _norm_spaces(str(s))
        s = re.sub(r"\*+$", "", s).strip()
        return s

    store_row = {}
    for r in range(ROW_START, ROW_END + 1):
        v = ws[f"B{r}"].value
        if v is None:
            continue
        store_row[_norm_store_for_excel(v)] = r

    c1 = column_index_from_string(COL_START)
    c2 = column_index_from_string(COL_END)
    item_col = {}
    for c in range(c1, c2 + 1):
        v = ws.cell(38, c).value
        if v:
            item_col[_norm_spaces(str(v))] = c

    return store_row, item_col


def clear_target_cells(ws):
    # 清空目标范围内的历史值与高亮
    c1 = column_index_from_string(COL_START)
    c2 = column_index_from_string(COL_END)
    for r in range(ROW_START, ROW_END + 1):
        for c in range(c1, c2 + 1):
            cell = ws.cell(r, c)
            cell.value = None
            cell.fill = PatternFill()
        # 清空 PO 列
        ws[f"{PO_COL}{r}"].value = None

# ---------- 主流程 ---------- #


def main():
    wb = xl.load_workbook(EXCEL_PATH)
    ws = wb[SHEET_NAME]
    clear_target_cells(ws)
    store_row, item_col = build_mappings(ws)
    not_found_store, not_found_item, updated = [], [], []

    pdf_list = sorted(PDF_DIR.glob("*.pdf"))
    for pdf_file in pdf_list:
        pdf_had_update = False
        try:
            pages = parse_pdf(pdf_file)
            for lines in pages:
                page_info = extract_orders_from_lines(lines)
                # 检测是否为 amend 页（包含 amend/amended/amendment 等关键字）
                page_has_amend = any(
                    re.search(r"\bamend", ln, flags=re.I) for ln in lines)
                if not page_info["items"]:
                    continue  # 该页无明细

                # 门店解析：优先使用 store_line；支持关键字/部分匹配
                resolved_store_name = None
                if page_info["store_line"]:
                    resolved_store_name = resolve_store(
                        page_info["store_line"], list(store_row.keys()))
                else:
                    # 如果 PDF 没有清晰门店行，也尝试任何包含 MOS 的行
                    for ln in lines:
                        if "mos" in ln.lower():
                            resolved_store_name = resolve_store(
                                ln, list(store_row.keys()))
                            if resolved_store_name:
                                break

                if not resolved_store_name:
                    not_found_store.append(
                        {"store": page_info["store_line"] or "(missing)", "po": page_info["po"]})
                    continue

                r = store_row.get(resolved_store_name)
                if r is None:
                    not_found_store.append(
                        {"store": page_info["store_line"], "po": page_info["po"]})
                    continue

                # 写入每个明细
                for it in page_info["items"]:
                    item_std = _map_item(it["desc"])
                    c = item_col.get(item_std)
                    matched_header = item_std
                    if c is None:
                        # Fuzzy fallback by token overlap
                        c, matched_header = find_best_item_col(
                            it["desc"], item_col)
                    if c is None:
                        not_found_item.append({
                            "store": resolved_store_name,
                            "item": matched_header or item_std,
                            "qty": it.get("qty"),
                            "uom": it.get("uom", ""),
                            "po": page_info["po"]
                        })
                        continue

                    # 写入“数字 + UOM”，并根据数量进行单复数调整
                    if page_has_amend and int(it.get("qty", 0)) == 0:
                        cell = ws.cell(r, c)
                        cell.value = None
                        cell.fill = PatternFill()
                    else:
                        ws.cell(r, c).value = _format_qty_uom(
                            it["qty"], it.get("uom", ""))
                        ws.cell(r, c).fill = HIGHLIGHT
                    pdf_had_update = True

                # 写入 PO（保留原先的数字/文本写入策略）
                po_cell = ws[f"{PO_COL}{r}"]
                po_val = (page_info["po"] or "").strip()
                if re.fullmatch(r"\d{1,}\Z", po_val):
                    try:
                        po_cell.value = int(po_val)
                    except Exception:
                        po_cell.value = po_val
                else:
                    po_cell.value = po_val
                if po_val:
                    pdf_had_update = True

                updated.append((resolved_store_name, page_info["po"]))
        except Exception as e:
            print(f"[!] 解析 {pdf_file.name} 出错：{e}", file=sys.stderr)

        # 已写入则移动 PDF 到归档目录
        if pdf_had_update:
            _safe_move_file(pdf_file, ARCHIVE_DIR)

    # 覆盖保存（只保存一次）
    wb.save(EXCEL_PATH)
    wb.close()

    print(f"✅ 已更新 {len(updated)} 条记录，处理 {len(pdf_list)} 份 PDF → {EXCEL_PATH}")
    if not_found_store or not_found_item:
        print("\n⚠️ 未匹配成功：")
        if not_found_store:
            print("  [门店未匹配]")
            for od in not_found_store:
                print(f"   - store={od.get('store')} / PO {od.get('po')}")
        if not_found_item:
            print("  [物品未匹配]")
            for od in not_found_item:
                qty_uom = _format_qty_uom(od.get('qty') or 0, od.get(
                    'uom', '')) if od.get('qty') is not None else ''
                print(
                    f"   - {od['store']} / {od['item']} / {qty_uom} / PO {od['po']}")


if __name__ == "__main__":
    main()
