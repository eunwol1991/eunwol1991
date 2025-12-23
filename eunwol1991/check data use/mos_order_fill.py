"""
MOS / LTO Order → Excel 自动填表脚本（支持三行版 + 单行表格版）
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
PO_COL = "AI"
ROW_START = 39
ROW_END = 73
COL_START = "C"
COL_END = "AD"
HIGHLIGHT = PatternFill(fill_type="solid", fgColor="00FFFCD7")
ARCHIVE_DIR = Path(r"C:\Users\jhunj\Dropbox\for jj\mos order")

# —— PO 号下限（含）——
PO_MIN = 600000

# ======== 门店别名（键已全部小写） ======================================= #
STORE_ALIAS = {
    "mos burger 100 am (38)":                 "MOS - 100AM",
    "mos burger 18 tai seng (67)":            "MOS - 18 Tai Seng",
    "mos burger amk hub (24)":                "MOS - Ang Mo Kio Hub",
    "mos burger bedok mall (42)":             "MOS - Bedok Mall",
    "mos burger jewel changi airport (59)":   "MOS - Café Jewel Changi Airport*",
    "mos burger causeway point (33)":         "MOS - Causeway Point",
    "mos burger changi city point (80)":      "MOS - Changi City Point",
    "mos burger city square mall (69)":       "MOS - City Square Mall",
    "mos burger compass one (49)":           "MOS - Compass One",
    "mos burger eastpoint (64)":             "MOS - Eastpoint Mall",
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

# ======== 物品别名（键已全部小写） ======================================= #
ITEM_ALIAS = {
    "ikeda japanese chicken cutlet":                "Ikeda Japanese Chicken Cutlet (6 x 1.1kg)",
    "japanese chicken katsu 70g":                   "CS TAY Japanese Chicken Katsu 70G (6 x 1kg)",
    "tays golden crispy chicken wing":              "Golden Crispy Chicken Wing (5 pkts x 10 pcs x 1.15kg)",
    "cs tay crispy chicken patty":                  "CS Tay Crispy Chicken Patty (4 x 2.5kg)",
    "eb kranch alaska pollock fingers":             "Kranch Alaska Pollock Finger (5 x 900g)",
    "battered natural onion rings":                 "Salud Battered Natural Onion Ring (6 x 1kg)",
    "rosti hashbrown":                              "Rosti Hashbrown (10 x 1kg)",
    "tomato capsicum soup":                         "Tomato Capsicum Soup (10 x 1kg)",
    "par cooked beef patty 120g":                   "Cooked Hamburg Beef Patty 120G (10 x 720g)",
    "demi glace sauce":                             "Demi Glace Sauce (10 x 1kg)",
    "chilli crab sauce":                            "Chilli Crab Sauce (10 x 1kg)",
    "anchor minidish unsalted butter 7g":           "Anchor Minidish Unsalted Butter (1 x 144 x 7g)",
    "fried chicken tulip":                          "Fried Chicken Tulip (10 x 1kg)",
    "tsukune japanese minced chicken with potato":  "Tsukune Japanese Minced Chicken with Potato (20 x 480g)",
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
    "red sauce":                               "Red Sauce (10 x 1kg)",
    "j-lek sriracha sauce":                         "J-Lek Sriracha Hot Chilli Sauce (12 x 455ml)",
    "uht coconut water 1l":                         "UHT Coconut Water 1L (12 x 1L)",
}


# ========================================================================== #
UNIT_RE = r"(?:ctn|ctns|pkt|pkts|tin|tins|can|cans|box|boxes|btl|btls|pc|pcs)"

_PLURAL_MAP = {
    "ctn": "ctns",
    "pkt": "pkts",
    "tin": "tins",
    "can": "cans",
    "box": "boxes",
    "btl": "btls",
    "pc": "pcs",
}
_SINGULAR_MAP = {v: k for k, v in _PLURAL_MAP.items()}


def _normalize_uom(uom: str) -> str:
    u = (uom or "").strip().lower()
    if u in _SINGULAR_MAP:
        return _SINGULAR_MAP[u]
    return u


def _format_qty_uom(qty: int, uom: str) -> str:
    base = _normalize_uom(uom)
    if qty > 1:
        plural = _PLURAL_MAP.get(base, base + "s")
        return f"{qty} {plural}"
    return f"{qty} {base}"


def _norm_spaces(s: str) -> str:
    return re.sub(r"\s+", " ", s).strip()


def _looks_like_date_fragment(s: str) -> bool:
    if not s:
        return False
    return ("/" in s) or bool(re.search(r"20\d{2}", s))


# ---------- Item / Store normalization ---------- #

def _norm_item_key(desc: str) -> str:
    s = desc.lower()
    s = re.sub(r"\(.*?\)", "", s)
    s = re.sub(r"[^a-z0-9\s,.-]+", " ", s)
    return _norm_spaces(s)


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

    for e, orig in excel_norm_map.items():
        if pdf_norm in e or e in pdf_norm:
            return orig

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


# ---------- LTO 单行结构解析（新增） ---------- #

LINE_ITEM_RE = re.compile(
    rf"^\s*\d+\s+(.+?)\s+(\d+)\s+({UNIT_RE})\b", re.I
)


def parse_single_line_items(lines: list[str]) -> list:
    items = []
    for ln in lines:
        m = LINE_ITEM_RE.match(ln)
        if not m:
            continue

        desc = _norm_spaces(m.group(1))
        qty = int(m.group(2))
        uom = m.group(3).lower()

        items.append({"desc": desc, "qty": qty, "uom": uom})

    return items


# ---------- PDF basic parsing ---------- #

def parse_pdf(pdf_path: Path):
    pages = []
    with fitz.open(pdf_path) as doc:
        for page in doc:
            raw = page.get_text("text") or ""
            pages.append([ln.strip() for ln in raw.splitlines() if ln.strip()])
    return pages


# ---------- 解析 PO（保留你的 PO_MIN 逻辑） ---------- #

def _extract_po_from_lines(lines: list[str]) -> str:
    joined = " \n ".join(lines)

    m = re.search(r"(?is)approved\s*by[^\d]{0,80}?(\d{6})", joined)
    if m and int(m.group(1)) >= PO_MIN:
        return m.group(1)

    for m in re.finditer(r"\b\d{6}\b", joined):
        if int(m.group(0)) >= PO_MIN:
            return m.group(0)

    return ""


# ---------- 主解析：同时支持三行结构 + 单行结构 ---------- #

def extract_orders_from_lines(lines: list[str]) -> dict:
    store_line = None
    for ln in lines:
        if re.fullmatch(r"(?i)mos burger.+\(\d+\)", ln):
            store_line = ln
            break

    po = _extract_po_from_lines(lines)

    # 先试 LTO 单行结构
    items = parse_single_line_items(lines)

    # 若单行结构抓不到，再尝试三行结构（旧格式兼容）
    if not items:
        items = []
        qty_re = re.compile(rf"^(\d+)\s+({UNIT_RE})$", re.I)
        i = 0
        while i < len(lines) - 2:
            desc = lines[i]
            m_qty = qty_re.fullmatch(lines[i + 1])
            if m_qty:
                items.append({
                    "desc": desc,
                    "qty": int(m_qty.group(1)),
                    "uom": m_qty.group(2).lower()
                })
                i += 2
                continue
            i += 1

    return {"store_line": store_line, "po": po, "items": items}


# ---------- Excel 表头映射 ---------- #

def build_mappings(ws):
    store_row = {}
    for r in range(ROW_START, ROW_END + 1):
        v = ws[f"B{r}"].value
        if v:
            store_row[_norm_spaces(str(v))] = r

    c1 = column_index_from_string(COL_START)
    c2 = column_index_from_string(COL_END)

    item_col = {}
    for c in range(c1, c2 + 1):
        v = ws.cell(38, c).value
        if v:
            item_col[_norm_spaces(str(v))] = c

    return store_row, item_col


def clear_target_cells(ws):
    c1 = column_index_from_string(COL_START)
    c2 = column_index_from_string(COL_END)

    for r in range(ROW_START, ROW_END + 1):
        for c in range(c1, c2 + 1):
            ws.cell(r, c).value = None
            ws.cell(r, c).fill = PatternFill()
        ws[f"{PO_COL}{r}"].value = None


# ---------- 主流程 ---------- #

def main():
    wb = xl.load_workbook(EXCEL_PATH)
    ws = wb[SHEET_NAME]
    clear_target_cells(ws)

    store_row, item_col = build_mappings(ws)

    not_found_store = []
    # 记录找不到物品的详细信息：每一笔是一个 dict
    # [{"pdf": xxx, "store": xxx, "po": xxx, "desc": xxx, "qty": n, "uom": "ctn"}, ...]
    not_found_item = []
    updated = []

    pdf_list = sorted(PDF_DIR.glob("*.pdf"))

    for pdf_file in pdf_list:
        pdf_had_update = False

        try:
            pages = parse_pdf(pdf_file)
            for lines in pages:
                info = extract_orders_from_lines(lines)
                if not info["items"]:
                    continue

                resolved_store = None
                if info["store_line"]:
                    resolved_store = resolve_store(
                        info["store_line"], list(store_row.keys()))
                else:
                    for ln in lines:
                        if "mos" in ln.lower():
                            resolved_store = resolve_store(
                                ln, list(store_row.keys()))
                            if resolved_store:
                                break

                if not resolved_store:
                    not_found_store.append(info["store_line"])
                    continue

                r = store_row.get(resolved_store)
                if not r:
                    not_found_store.append(info["store_line"])
                    continue

                for it in info["items"]:
                    item_std = _map_item(it["desc"])
                    c = item_col.get(item_std)

                    if not c:
                        # 记录：是哪一个 PDF、门店、PO、原始品名、数量、单位
                        not_found_item.append({
                            "pdf": pdf_file.name,
                            "store": resolved_store,
                            "po": info["po"],
                            "desc": it["desc"],
                            "qty": it.get("qty"),
                            "uom": it.get("uom"),
                        })
                        continue

                    ws.cell(r, c).value = _format_qty_uom(
                        it["qty"], it["uom"])
                    ws.cell(r, c).fill = HIGHLIGHT
                    pdf_had_update = True

                # 写入 PO
                if info["po"]:
                    po_cell = ws[f"{PO_COL}{r}"]
                    po_cell.value = int(info["po"])
                    pdf_had_update = True

                updated.append((resolved_store, info["po"]))

        except Exception as e:
            print(f"[!] 解析出错 {pdf_file.name}: {e}")

        if pdf_had_update:
            _safe_move_file(pdf_file, ARCHIVE_DIR)

    wb.save(EXCEL_PATH)
    wb.close()

    print(f"已更新 {len(updated)} 条记录, 处理 {len(pdf_list)} 份 PDF → {EXCEL_PATH}")

    # 若有 PO 里的 items 没有在 Excel 找到对应列，提醒用户
    if not_found_item:
        print("\n[警告] 以下物品在 Excel 中找不到对应的列，请检查 ITEM_ALIAS 或表头：\n")
        seen = set()
        for rec in not_found_item:
            key = (rec["pdf"], rec["store"], rec["po"], rec["desc"])
            if key in seen:
                continue
            seen.add(key)
            print(
                f"- PDF: {rec['pdf']}, 门店: {rec['store']}, PO: {rec['po']}, "
                f"品名: {rec['desc']}, 数量: {rec['qty']} {rec['uom']}"
            )
        print("\n请根据以上资讯，\n1) 确认该物品是否应加入 ITEM_ALIAS；\n2) 或确认 Excel 第 38 行的品名是否与标准描述一致。")
    else:
        print("\n所有物品均已成功匹配到 Excel 列，没有遗漏的品项。")


if __name__ == "__main__":
    main()
