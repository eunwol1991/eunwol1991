import streamlit as st
import streamlit.components.v1 as components
from typing import Any, Optional, Tuple, List, Dict
import json
import datetime
import math
import pandas as pd
import re
import os
import sys
import io
import calendar
import os
import sys
from pathlib import Path

# 尝试检测当前是否已经在 streamlit 运行环境中
try:
    from streamlit.runtime.scriptrunner import get_script_run_ctx  # type: ignore
except Exception:
    def get_script_run_ctx():
        return None


def _short_path(p: str) -> str:
    """在 Windows 下把带空格的长路径转换成短路径，避免 cmd 解析问题。"""
    if os.name != "nt" or " " not in p:
        return p
    try:
        import ctypes
        from ctypes import wintypes
        GetShortPathNameW = ctypes.windll.kernel32.GetShortPathNameW
        GetShortPathNameW.argtypes = [
            wintypes.LPCWSTR, wintypes.LPWSTR, wintypes.DWORD
        ]
        GetShortPathNameW.restype = wintypes.DWORD
        buf = ctypes.create_unicode_buffer(260)
        res = GetShortPathNameW(p, buf, 260)
        return buf.value if res else p
    except Exception:
        return p


# 如果是直接 python xxx.py 启动，并且当前不在 streamlit 环境里，就自动帮你转成 streamlit run
if __name__ == "__main__" and os.environ.get("ST_REDIRECTED", "0") != "1" and (get_script_run_ctx() is None):
    import subprocess

    os.environ["ST_REDIRECTED"] = "1"

    script_path = str(Path(__file__).resolve())
    script_path = _short_path(script_path)

    # 把扩展名统一改成小写 .py，避免 Streamlit 讨厌 .PY
    if script_path.lower().endswith(".py"):
        script_path = script_path[:-3] + ".py"

    cmd = [sys.executable, "-m", "streamlit", "run", script_path]

    # 把原本 python xxx.py 后面的参数透传给 streamlit
    if len(sys.argv) > 1:
        cmd += ["--"] + sys.argv[1:]

    if os.environ.get("ST_DEBUG_REDIRECT") == "1":
        print("[streamlit-redirect] ", cmd)

    subprocess.run(cmd, check=False)
    sys.exit(0)


st.set_page_config(page_title='Warehouse Suite', layout='wide')

PRIMARY_WAREHOUSES = ["Savori Whse", "Lai Hock Whse"]

# ==== 新增：预编译正则 ====
HTML_TAG_RE = re.compile(r"<.*?>")
PAREN_CONTENT_RE = re.compile(r"\s*\([^)]*\)")


# ----------------------------- 基础工具函数 -----------------------------

def _normalize_warehouse_name(value) -> str:
    if pd.isna(value):
        return ""
    return str(value).strip()


def _normalize_pack_size_value(value) -> str:
    if pd.isna(value):
        return ""
    text = str(value).strip()
    if not text or text.lower() == "nan":
        return ""
    return text


def _format_pack_size_label(normalized: str) -> str:
    return normalized if normalized else "-"


def _canonical_unit(u: Optional[str]) -> str:
    if u is None:
        return ""
    s = str(u).strip().lower()
    if not s:
        return ""
    synonyms = {
        "ctn": {"ctn", "ctns", "carton", "cartons"},
        "pkt": {"pkt", "pkts", "pack", "packs", "package", "packages"},
        "box": {"box", "boxes"},
        "tin": {"tin", "tins"},
        "can": {"can", "cans"},
        "bag": {"bag", "bags"},
        "pc": {"pc", "pcs", "piece", "pieces"},
    }
    for key, vals in synonyms.items():
        if s in vals:
            return key
    return s


def _format_qty_number(value: Optional[float]) -> Optional[str]:
    if value is None or pd.isna(value):
        return None
    try:
        num = float(value)
    except (TypeError, ValueError):
        return None
    if math.isclose(num, 0.0, abs_tol=1e-9):
        return "0"
    if math.isclose(num, round(num), rel_tol=1e-9, abs_tol=1e-9):
        return f"{int(round(num))}"
    text_val = f"{num:,.2f}"
    return text_val.rstrip("0").rstrip(".")


def _format_quantity_pair(ctn: Optional[float], pkt: Optional[float]) -> str:
    parts: List[str] = []
    if ctn is not None and not pd.isna(ctn):
        ctn_val = float(ctn)
        if not math.isclose(ctn_val, 0.0, abs_tol=1e-9):
            qty_text = _format_qty_number(ctn_val)
            if qty_text:
                unit = "ctn" if math.isclose(
                    ctn_val, 1.0, abs_tol=1e-9) else "ctns"
                parts.append(f"{qty_text} {unit}")
    if pkt is not None and not pd.isna(pkt):
        pkt_val = float(pkt)
        if not math.isclose(pkt_val, 0.0, abs_tol=1e-9):
            qty_text = _format_qty_number(pkt_val)
            if qty_text:
                unit = "pkt" if math.isclose(
                    pkt_val, 1.0, abs_tol=1e-9) else "pkts"
                parts.append(f"{qty_text} {unit}")
    return " ".join(parts) if parts else "0"


def _strip_html_df(df: pd.DataFrame) -> pd.DataFrame:
    clean = df.copy()
    for col in clean.select_dtypes(include="object").columns:
        clean[col] = clean[col].astype(str).str.replace(
            HTML_TAG_RE, "", regex=True)
    return clean


MONTH_MAP: Dict[str, int] = {}
for index, name in enumerate(calendar.month_name):
    if name:
        MONTH_MAP[name.lower()] = index
for index, abbr in enumerate(calendar.month_abbr):
    if abbr:
        MONTH_MAP[abbr.lower()] = index


def _month_sort_key(value) -> int:
    if pd.isna(value):
        return 999
    text = str(value).strip()
    if not text:
        return 999
    if text.isdigit():
        num = int(text)
        if 1 <= num <= 12:
            return num
    lower = text.lower()
    if lower in MONTH_MAP:
        return MONTH_MAP[lower]
    if lower[:3] in MONTH_MAP:
        return MONTH_MAP[lower[:3]]
    return 999


def _format_month_label(value: Any) -> str:
    if pd.isna(value):
        return ""
    text = str(value).strip()
    if not text:
        return ""
    if text.isdigit():
        num = int(text)
        if 1 <= num <= 12:
            return calendar.month_abbr[num]
    lower = text.lower()
    if lower in MONTH_MAP:
        num = MONTH_MAP[lower]
        return calendar.month_abbr[num] if num else text
    if lower[:3] in MONTH_MAP:
        num = MONTH_MAP[lower[:3]]
        return calendar.month_abbr[num] if num else text
    return text


def _format_qty_display(value) -> str:
    if pd.isna(value):
        return ""
    try:
        num = float(value)
    except (TypeError, ValueError):
        return str(value)
    if math.isclose(num, round(num), rel_tol=1e-9, abs_tol=1e-9):
        return f"{int(round(num)):,}"
    formatted = f"{num:,.2f}".rstrip("0").rstrip(".")
    return formatted


def _format_price_display(value) -> str:
    if pd.isna(value):
        return ""
    try:
        num = float(value)
    except (TypeError, ValueError):
        return str(value)
    if math.isclose(num, 0.0, abs_tol=1e-9):
        return "0.00"
    return f"{num:,.2f}"


SALES_COLUMNS = [
    "Year",
    "Date",
    "Month",
    "Brand/Category",
    "Supplier",
    "Product Code",
    "Product Description",
    "Carton Packing",
    "Customer",
    "Outlet",
    "Qty in Pcs",
    "Qty in Ctns",
    "Total Qty in Pcs",
    "Total Qty in Ctns",
    "Invoice #",
    "Total Value",
    "GST",
    "Total Value Inclusive GST",
    "Account",
    "Customer PO#",
    "Remarks",
]

SALES_NUMERIC_COLUMNS = [
    "Qty in Pcs",
    "Qty in Ctns",
    "Total Value",
    "GST",
    "Total Value Inclusive GST",
]

SALES_DATE_COLUMNS = ["Date"]


def _safe_divide_series(numerator, denominator) -> pd.Series:
    numerator = pd.to_numeric(numerator, errors="coerce")
    denominator = pd.to_numeric(denominator, errors="coerce")
    result = numerator / denominator
    mask = denominator.isna() | (denominator == 0)
    return result.where(~mask, other=pd.NA)


def _highlight_missing_cell(value) -> str:
    if isinstance(value, str):
        if not value.strip():
            return "background-color: rgba(255, 235, 205, 0.6);"
    if pd.isna(value):
        return "background-color: rgba(255, 235, 205, 0.6);"
    return ""


_WEIGHT_UNIT_PATTERN = re.compile(
    r"\d+(?:\.\d+)?\s*(?:kg|kgs|g|grams?|gm|lt|ltr|l|ml|milliliter|millilitre|liter|litre|oz|lb|lbs)\b"
)


def _parse_carton_packing_numeric(value) -> Optional[float]:
    if pd.isna(value):
        return None
    text = str(value).strip()
    if not text:
        return None
    lowered = text.lower()

    match = re.search(r"(\d+)\s*[x×]", lowered)
    if not match:
        match = re.search(
            r"(\d+)(?=\s*(?:kg|g|ltr|lt|l|ml|litre|litres|liter|liters))",
            lowered,
        )
    if not match:
        match = re.search(r"(\d+)", lowered)
    if not match:
        return None
    try:
        return float(match.group(1))
    except ValueError:
        return None


def _infer_carton_pack_size(value) -> Optional[float]:
    if pd.isna(value):
        return None
    text = str(value).strip()
    if not text:
        return None
    lowered = text.lower()
    segments = [seg.strip()
                for seg in re.split(r"[x×*]", lowered) if seg.strip()]

    numeric_segments: List[Tuple[float, bool]] = []
    for seg in segments:
        match = re.search(r"(\d+(?:\.\d+)?)", seg)
        if not match:
            continue
        try:
            number_value = float(match.group(1))
        except ValueError:
            continue
        is_weight = bool(_WEIGHT_UNIT_PATTERN.search(seg))
        numeric_segments.append((number_value, is_weight))

    non_weight_numbers = [
        number for number, is_weight in numeric_segments
        if not is_weight and number > 0
    ]
    if non_weight_numbers:
        product = 1.0
        for number in non_weight_numbers:
            product *= number
        return product
    if numeric_segments:
        first_value, first_is_weight = numeric_segments[0]
        if first_is_weight:
            return 1.0
        if first_value > 0:
            return first_value
    return None


def _format_total_qty_text(total_pieces: float, carton_size: Optional[float]) -> str:
    if carton_size is None or carton_size <= 0 or pd.isna(total_pieces):
        return "-"
    try:
        pieces = int(round(total_pieces))
    except (TypeError, ValueError):
        return "-"
    if pieces == 0:
        return "0"
    cartons = pieces // int(carton_size)
    packets = pieces - cartons * int(carton_size)
    parts: List[str] = []
    if cartons > 0:
        parts.append(f"{cartons} ctn{'s' if cartons != 1 else ''}")
    if packets > 0:
        parts.append(f"{packets} pkt{'s' if packets != 1 else ''}")
    return " ".join(parts) if parts else "0"


def _coerce_qty_int(value) -> int:
    if value is None or pd.isna(value):
        return 0
    try:
        return int(round(float(value)))
    except (TypeError, ValueError):
        return 0


def _format_summary_total_qty(ctn_value, pkt_value, pack_size_value) -> str:
    cartons = _coerce_qty_int(ctn_value)
    packets = _coerce_qty_int(pkt_value)
    pack_size = _coerce_qty_int(pack_size_value)
    if pack_size > 0:
        total_packets = cartons * pack_size + packets
        formatted = _format_total_qty_text(total_packets, pack_size)
        if formatted and formatted != "-":
            return formatted
    return f"{cartons} ctns {packets} pkts"


def build_norm_desc(df: pd.DataFrame) -> pd.Series:
    if df is None or df.empty:
        return pd.Series(dtype="string", name="norm_desc")

    desc_series = df.get("description", pd.Series(
        dtype="string")).astype("string").fillna("")
    base = desc_series.str.replace(
        PAREN_CONTENT_RE, "", regex=True).str.strip()

    code_series = df.get("product_code", pd.Series(
        dtype="string")).astype("string").fillna("").str.strip()
    fallback_desc = desc_series.str.strip()

    norm = base.copy()
    mask_empty = norm.isna() | (norm.str.len() == 0)
    if mask_empty.any():
        norm = norm.mask(mask_empty, code_series)
        mask_empty = norm.isna() | (norm.str.len() == 0)
    if mask_empty.any():
        norm = norm.mask(mask_empty, fallback_desc)
        mask_empty = norm.isna() | (norm.str.len() == 0)
    if mask_empty.any():
        norm = norm.mask(mask_empty, "Unnamed product")

    norm = norm.fillna("").str.strip().replace("", "Unnamed product")
    norm.name = "norm_desc"
    return norm


def _extract_reorder_points(df_group: pd.DataFrame) -> Dict[str, Optional[float]]:
    result: Dict[str, Optional[float]] = {"ctn": None, "pkt": None}
    if df_group is None or df_group.empty:
        return result
    candidate_cols = [
        c for c in df_group.columns
        if isinstance(c, str) and ("reorder" in c.lower() or "rop" in c.lower())
    ]
    for col in candidate_cols:
        values = pd.to_numeric(df_group[col], errors="coerce")
        valid = values.dropna()
        if valid.empty:
            continue
        value = float(valid.iloc[0])
        lower = col.lower()
        if "pkt" in lower:
            result["pkt"] = value
        elif "ctn" in lower:
            result["ctn"] = value
        else:
            if result["ctn"] is None:
                result["ctn"] = value
            if result["pkt"] is None:
                result["pkt"] = value
    return result


# ----------------------------- 业务聚合 -----------------------------

def aggregate_summary(
    df: pd.DataFrame,
    warehouses: Optional[List[str]] = None,
) -> Tuple[pd.DataFrame, Dict[Tuple[str, str, str, str, str], pd.DataFrame]]:
    warehouses = warehouses or PRIMARY_WAREHOUSES
    columns = [
        "Supplier", "Brand", "Product", "Pack Size", "Product Code",
        "savori_ctn", "savori_pkt", "lai_hock_ctn", "lai_hock_pkt",
        "total_ctn", "total_pkt",
        "reorder_point_ctn", "reorder_point_pkt",
        "group_key",
    ]
    if df is None or df.empty:
        return pd.DataFrame(columns=columns), {}

    work = df.copy()
    if "warehouse" in work.columns:
        work["_warehouse_norm"] = work["warehouse"].map(
            _normalize_warehouse_name)
        work = work[work["_warehouse_norm"].isin(warehouses)]
    else:
        work["_warehouse_norm"] = ""
    if work.empty:
        return pd.DataFrame(columns=columns), {}

    work["norm_desc"] = build_norm_desc(work)
    work["_unit_norm"] = work.get("unit", pd.Series(
        dtype="string")).apply(_canonical_unit)
    work["_qty"] = pd.to_numeric(
        work.get("stock_qty", pd.Series(dtype=float)), errors="coerce").fillna(0.0)
    work["_product_code_norm"] = work.get("product_code", pd.Series(
        dtype="string")).astype("string").fillna("").str.strip()
    if "supplier" not in work.columns:
        work["supplier"] = pd.NA

    if "pack_size" in work.columns:
        work["_pack_size_norm"] = work["pack_size"].map(
            _normalize_pack_size_value)
    else:
        work["_pack_size_norm"] = ""
    if "brand" in work.columns:
        work["_brand_norm"] = (
            work["brand"].astype("string").fillna("").str.strip()
        )
    else:
        work["_brand_norm"] = ""

    rows = []
    detail_map: Dict[Tuple[str, str, str, str, str], pd.DataFrame] = {}

    for key, grp in work.groupby(
        ["supplier", "_brand_norm", "norm_desc", "_pack_size_norm", "_product_code_norm"],
        dropna=False,
    ):
        if not isinstance(key, tuple):
            key = (key,)
        supplier_val = key[0]
        brand_val = key[1] if len(key) > 1 else ""
        norm_desc_val = key[2] if len(key) > 2 else ""
        code_val = key[4] if len(key) > 4 else ""
        supplier_name = str(supplier_val).strip() if pd.notna(
            supplier_val) and str(supplier_val).strip() else "Unknown supplier"
        product_label = str(norm_desc_val).strip() if pd.notna(
            norm_desc_val) and str(norm_desc_val).strip() else "Unnamed product"
        raw_pack_val = key[3] if len(key) > 3 else ""
        pack_norm_val = _normalize_pack_size_value(raw_pack_val)
        pack_size_label = _format_pack_size_label(pack_norm_val)
        brand_text = str(brand_val).strip() if pd.notna(
            brand_val) and str(brand_val).strip() else ""
        brand_label = brand_text if brand_text else "Unknown brand"
        code_text = str(code_val).strip() if pd.notna(code_val) else ""
        product_code_label = code_text if code_text else "-"
        group_key = (supplier_name, brand_label,
                     product_label, pack_norm_val, code_text)
        detail_map[group_key] = grp.copy()

        wh_totals = {wh: {"ctn": 0.0, "pkt": 0.0} for wh in warehouses}
        wh_unit = grp.groupby(["_warehouse_norm", "_unit_norm"], dropna=False)[
            "_qty"].sum().reset_index()
        for _, r in wh_unit.iterrows():
            wh = r["_warehouse_norm"]
            unit = r["_unit_norm"]
            qty = float(r["_qty"])
            if wh in wh_totals and unit in ("ctn", "pkt"):
                wh_totals[wh][unit] += qty

        total_ctn = sum(wh_totals.get(wh, {}).get("ctn", 0.0)
                        for wh in warehouses)
        total_pkt = sum(wh_totals.get(wh, {}).get("pkt", 0.0)
                        for wh in warehouses)

        reorder_points = _extract_reorder_points(grp)

        rows.append({
            "Supplier": supplier_name,
            "Brand": brand_label,
            "Product": product_label,
            "Pack Size": pack_size_label,
            "Product Code": product_code_label,
            "savori_ctn": wh_totals.get("Savori Whse", {}).get("ctn", 0.0),
            "savori_pkt": wh_totals.get("Savori Whse", {}).get("pkt", 0.0),
            "lai_hock_ctn": wh_totals.get("Lai Hock Whse", {}).get("ctn", 0.0),
            "lai_hock_pkt": wh_totals.get("Lai Hock Whse", {}).get("pkt", 0.0),
            "total_ctn": total_ctn,
            "total_pkt": total_pkt,
            "reorder_point_ctn": reorder_points.get("ctn"),
            "reorder_point_pkt": reorder_points.get("pkt"),
            "group_key": group_key,
        })

    summary_df = pd.DataFrame(rows, columns=columns if rows else columns)
    return summary_df, detail_map


# ----------------------------- 判定规则（批次层/产品层） -----------------------------

def classify_batch_status(expiry_date: Optional[datetime.date], ctn: float, pkt: float, expiry_days: int):
    total = (ctn or 0.0) + (pkt or 0.0)
    if not isinstance(expiry_date, datetime.date):
        return ("Depleted" if total == 0 else "OK", None)
    days = (expiry_date - datetime.date.today()).days
    if total > 0:
        if days < 0:
            return ("Expired", days)
        if 0 <= days <= int(expiry_days):
            return ("Near-Expiry", days)
        return ("OK", days)
    else:
        return ("Depleted", days)


def classify_product_status(has_expired: bool, has_near: bool, is_low_stock: bool) -> str:
    if has_expired:
        return "Expired"
    if has_near:
        return "Near-Expiry"
    if is_low_stock:
        return "Low-Stock"
    return "OK"


def split_by_expiry(
    df_row_scope: pd.DataFrame,
    warehouses: Optional[List[str]] = None,
    *,
    expiry_days: int = 30,
    show_depleted: bool = True,
    batch_mode: str = "expiry",
) -> pd.DataFrame:
    warehouses = warehouses or PRIMARY_WAREHOUSES
    columns = [
        "Expiry", "Remark", "Savori Whse", "Lai Hock Whse", "Subtotal",
        "subtotal_ctn", "subtotal_pkt", "expiry_date",
        "status_batch", "days_to_expiry", "Info",
    ]
    if df_row_scope is None or df_row_scope.empty:
        return pd.DataFrame(columns=columns)

    work = df_row_scope.copy()
    if "warehouse" not in work.columns:
        return pd.DataFrame(columns=columns)

    work["_warehouse_norm"] = work["warehouse"].map(_normalize_warehouse_name)
    work = work[work["_warehouse_norm"].isin(warehouses)]
    if work.empty:
        return pd.DataFrame(columns=columns)

    work["_unit_norm"] = work.get("unit", pd.Series(
        dtype="string")).apply(_canonical_unit)
    work["_qty"] = pd.to_numeric(
        work.get("stock_qty", pd.Series(dtype=float)), errors="coerce").fillna(0.0)

    work["_expiry_norm"] = pd.to_datetime(
        work.get("expiry_date"), errors="coerce", format="mixed")
    work["_expiry_label"] = work["_expiry_norm"].apply(
        lambda x: x.date().isoformat() if pd.notna(x) else "No Expiry")
    work["_expiry_sort_key"] = work["_expiry_norm"].apply(
        lambda x: x.date() if pd.notna(x) else datetime.date.max)

    def _extract_remark(s: str) -> str:
        lst = re.findall(r"\(([^)]*)\)", str(s) or "")
        vals = [t.strip() for t in lst if t and t.strip()]
        return ", ".join(vals) if vals else "No Remark"

    work["_remark_label"] = work.get(
        "description", pd.Series(dtype="string")).apply(_extract_remark)

    if batch_mode == "remark":
        group_keys = ["_remark_label"]
    elif batch_mode == "both":
        group_keys = ["_expiry_label", "_remark_label"]
    else:
        group_keys = ["_expiry_label"]

    rows = []
    for key, grp in work.groupby(group_keys, dropna=False):
        if not isinstance(key, tuple):
            key = (key,)
        exp_label = None
        remark_label = None
        if batch_mode == "remark":
            remark_label = key[0]
        elif batch_mode == "both":
            exp_label, remark_label = key
        else:
            exp_label = key[0]

        sort_key = grp["_expiry_sort_key"].min()
        expiry_val = grp["_expiry_norm"].dropna().min()

        per_wh = {wh: {"ctn": 0.0, "pkt": 0.0} for wh in warehouses}
        wh_unit = grp.groupby(["_warehouse_norm", "_unit_norm"], dropna=False)[
            "_qty"].sum().reset_index()
        for _, r in wh_unit.iterrows():
            wh = r["_warehouse_norm"]
            unit = r["_unit_norm"]
            qty = float(r["_qty"])
            if wh in per_wh and unit in ("ctn", "pkt"):
                per_wh[wh][unit] += qty

        subtotal_ctn = sum(per_wh.get(wh, {}).get("ctn", 0.0)
                           for wh in warehouses)
        subtotal_pkt = sum(per_wh.get(wh, {}).get("pkt", 0.0)
                           for wh in warehouses)

        exp_date = expiry_val.date() if pd.notna(expiry_val) else None
        status_batch, d2e = classify_batch_status(
            exp_date, subtotal_ctn, subtotal_pkt, expiry_days)

        if (status_batch == "Depleted") and (not show_depleted):
            continue

        qty_text = _format_quantity_pair(subtotal_ctn, subtotal_pkt)
        info = ""
        if status_batch == "Expired":
            info = f"已过期 {abs(d2e)} 天，数量 {qty_text}"
        elif status_batch == "Near-Expiry":
            info = f"到期日 {exp_date}，距离到期 {d2e} 天，数量 {qty_text}"
        elif status_batch == "Depleted":
            info = "该批次已用完（0），等待仓库处理"

        rows.append({
            "Expiry": exp_label if exp_label is not None else "",
            "Remark": remark_label if remark_label is not None else "",
            "Savori Whse": _format_quantity_pair(per_wh.get('Savori Whse', {}).get('ctn', 0.0), per_wh.get('Savori Whse', {}).get('pkt', 0.0)),
            "Lai Hock Whse": _format_quantity_pair(per_wh.get('Lai Hock Whse', {}).get('ctn', 0.0), per_wh.get('Lai Hock Whse', {}).get('pkt', 0.0)),
            "Subtotal": qty_text,
            "subtotal_ctn": subtotal_ctn,
            "subtotal_pkt": subtotal_pkt,
            "expiry_date": exp_date,
            "status_batch": status_batch,
            "days_to_expiry": d2e,
            "Info": info,
            "_sort_key": sort_key,
        })

    if not rows:
        return pd.DataFrame(columns=columns)

    result = (
        pd.DataFrame(rows)
        .sort_values(by=["_sort_key", "Expiry", "Remark"], kind="stable")
        .drop(columns="_sort_key")
        .reset_index(drop=True)
    )
    return result


# ----------------------------- 侧边栏筛选 -----------------------------

def apply_filters_v2(df: pd.DataFrame):
    base = df.copy()

    def get_desc_base(series: pd.Series) -> pd.Series:
        return series.astype(str).str.replace(r"\s*\([^)]*\)", "", regex=True).str.strip()

    def extract_remarks(series: pd.Series) -> list:
        vals = series.astype(str).str.findall(r"\(([^)]*)\)").dropna().tolist()
        out = set()
        for lst in vals:
            for r in lst or []:
                s = str(r).strip()
                if s:
                    out.add(s)
        return sorted(out)

    ss = st.session_state

    def _set_focus_target(target: str) -> None:
        ss["__focus_target"] = target

    def _ensure_multiselect_key(key: str, options: list, init: list):
        if key not in ss:
            ss[key] = list(init)
        else:
            current = ss.get(key, [])
            if isinstance(current, (str, int, float)):
                current = [current]
            ss[key] = [v for v in current if v in options]

    sel_wh = list(ss.get("f_wh", []))
    sel_sup = list(ss.get("f_sup", []))
    sel_brand = list(ss.get("f_brand", []))
    sel_desc = list(ss.get("f_desc", []))
    sel_code = list(ss.get("f_code", []))
    sel_remark = list(ss.get("f_remark", []))

    def apply_all(df_in: pd.DataFrame, exclude: str = "") -> pd.DataFrame:
        d = df_in

        def _include(series: pd.Series, selected: list):
            if not selected:
                return pd.Series(True, index=series.index)
            return series.isin(selected)

        if exclude != "warehouse" and "warehouse" in d.columns:
            sel = list(ss.get("f_wh", []))
            if sel:
                d = d[_include(d["warehouse"], sel)]

        if exclude != "supplier" and "supplier" in d.columns:
            sel = list(ss.get("f_sup", []))
            exm = bool(ss.get("f_sup_ex", False))
            if sel:
                m = d["supplier"].isin(sel)
                d = d[~m] if exm else d[m]

        if exclude != "brand" and "brand" in d.columns:
            sel = list(ss.get("f_brand", []))
            if sel:
                d = d[_include(d["brand"], sel)]

        if exclude != "desc" and "description" in d.columns:
            base_ser = get_desc_base(d["description"])
            sel = list(ss.get("f_desc", []))
            if sel:
                m = base_ser.isin(sel)
                d = d[m]

        if exclude != "code" and "product_code" in d.columns:
            sel = list(ss.get("f_code", []))
            if sel:
                d = d[_include(d["product_code"], sel)]

        if exclude != "remark" and "description" in d.columns:
            sel = list(ss.get("f_remark", []))
            exm = bool(ss.get("f_remark_ex", False))
            if sel:
                matches = d["description"].astype(
                    str).str.findall(r"\(([^)]*)\)")
                has_any = matches.apply(lambda lst: any(
                    (str(x).strip() in sel) for x in (lst or [])))
                d = d[~has_any] if exm else d[has_any]

        return d

    with st.sidebar:
        st.header("筛选条件")

        sig_wo_wh = (
            tuple(ss.get("f_sup", [])),
            tuple(ss.get("f_brand", [])),
            tuple(ss.get("f_desc", [])),
            tuple(ss.get("f_code", [])),
            tuple(ss.get("f_remark", [])),
            bool(ss.get("use_date_filters", False)),
            ss.get("expiry_range"),
            ss.get("relabel_date_range"),
        )
        prev_sig = ss.get("__sig_wo_wh")
        sig_changed = (sig_wo_wh != prev_sig)

        if "warehouse" in base.columns:
            d = apply_all(base, exclude="warehouse")
            wh_options = [x for x in d["warehouse"].dropna().astype(
                str).unique().tolist()]
            ordered = [w for w in ["Savori Whse",
                                   "Lai Hock Whse"] if w in wh_options]
            ordered += [w for w in wh_options if w not in ordered]
            default_selection = [w for w in [
                "Savori Whse", "Lai Hock Whse"] if w in ordered] or list(ordered)

            cur = [w for w in ss.get("f_wh", []) if w in ordered]
            need_reset = sig_changed or (not cur and bool(ordered))
            ss["f_wh"] = list(default_selection) if need_reset else (
                cur or list(default_selection))

            st.multiselect(
                "仓库",
                options=ordered,
                key="f_wh",
                placeholder="选择一个或多个仓库",
            )

            ss["__sig_wo_wh"] = sig_wo_wh

        if "supplier" in base.columns:
            d = apply_all(base, exclude="supplier")
            sup_options = sorted(
                [x for x in d["supplier"].dropna().unique().tolist()])
            _ensure_multiselect_key("f_sup", sup_options, [])
            st.multiselect("Supplier", sup_options,
                           key="f_sup", placeholder="选择供应商")
            st.checkbox("Exclude selected (Supplier)",
                        key="f_sup_ex", value=False)
            sel_sup = list(ss.get("f_sup", []))

        if "brand" in base.columns:
            d = apply_all(base, exclude="brand")
            brand_options = sorted(
                [x for x in d["brand"].dropna().unique().tolist()])
            _ensure_multiselect_key("f_brand", brand_options, [])
            st.multiselect("Brand", brand_options,
                           key="f_brand", placeholder="选择品牌")
            sel_brand = list(ss.get("f_brand", []))

        if "description" in base.columns:
            d = apply_all(base, exclude="desc")
            base_ser = get_desc_base(
                d["description"]) if not d.empty else pd.Series(dtype=str)
            desc_options = sorted(
                [x for x in base_ser.dropna().unique().tolist() if x])
            _ensure_multiselect_key("f_desc", desc_options, [])
            st.multiselect("Description（去括号后）", desc_options,
                           key="f_desc", placeholder="选择描述",
                           on_change=_set_focus_target, args=("desc",))
            sel_desc = list(ss.get("f_desc", []))

        if "product_code" in base.columns:
            d = apply_all(base, exclude="code")
            code_options = sorted(
                [x for x in d["product_code"].dropna().unique().tolist()])
            _ensure_multiselect_key("f_code", code_options, [])
            st.multiselect("Product Code", code_options,
                           key="f_code", placeholder="选择产品编码")
            sel_code = list(ss.get("f_code", []))

        if "description" in base.columns:
            d = apply_all(base, exclude="remark")
            remark_options = extract_remarks(
                d["description"]) if not d.empty else []
            _ensure_multiselect_key("f_remark", remark_options, [])
            st.multiselect("Remark（来自描述括号）", remark_options,
                           key="f_remark", placeholder="选择 Remark",
                           on_change=_set_focus_target, args=("remark",))
            st.checkbox("Exclude selected (Remark)",
                        key="f_remark_ex", value=False)
            sel_remark = list(ss.get("f_remark", []))

        use_date_filters = st.checkbox("启用日期范围筛选", value=False)
        st.session_state["use_date_filters"] = use_date_filters
        start = end = None
        r_start = r_end = None

        def _clamp(cur_range, min_d, max_d):
            if pd.isna(min_d) or pd.isna(max_d):
                return None
            if not cur_range or len(cur_range) != 2:
                return (min_d.date(), max_d.date())
            a, b = cur_range
            a = max(min_d.date(), min(a, max_d.date()))
            b = max(min_d.date(), min(b, max_d.date()))
            if a > b:
                a, b = (min_d.date(), max_d.date())
            return (a, b)

        if use_date_filters:
            if "expiry_date" in base.columns:
                min_d, max_d = base["expiry_date"].min(
                ), base["expiry_date"].max()
                if pd.notna(min_d) and pd.notna(max_d):
                    st.session_state["expiry_range"] = _clamp(
                        st.session_state.get("expiry_range"), min_d, max_d)
                    st.date_input("有效期范围", key="expiry_range",
                                  min_value=min_d.date(), max_value=max_d.date())
                    start, end = st.session_state.get("expiry_range")

            if "relabel_to_date" in base.columns:
                min_r, max_r = base["relabel_to_date"].min(
                ), base["relabel_to_date"].max()
                if pd.notna(min_r) and pd.notna(max_r):
                    st.session_state["relabel_date_range"] = _clamp(
                        st.session_state.get("relabel_date_range"), min_r, max_r)
                    st.date_input("Relabel To 日期范围", key="relabel_date_range",
                                  min_value=min_r.date(), max_value=max_r.date())
                    r_start, r_end = st.session_state.get("relabel_date_range")

        focus_target = ss.get("__focus_target")
        if focus_target in {"desc", "remark"}:
            label_text = (
                "Description（去括号后）"
                if focus_target == "desc"
                else "Remark（来自描述括号）"
            )
            target_json = json.dumps(label_text)
            components.html(
                f"""
                <script>
                (function() {{
                    const targetLabel = {target_json};
                    const focusByLabel = () => {{
                        const widgets = window.parent.document.querySelectorAll('[data-testid="stWidget"]');
                        for (const w of widgets) {{
                            const label = w.querySelector('[data-testid="stWidgetLabel"]');
                            if (!label) continue;
                            const text = (label.textContent || '').trim();
                            if (!text.startsWith(targetLabel)) continue;
                            let input = w.querySelector('input[type="text"]');
                            if (!input) {{
                                input = w.querySelector('input');
                            }}
                            if (input) {{
                                input.focus();
                                if (input.setSelectionRange) {{
                                    const end = input.value ? input.value.length : 0;
                                    input.setSelectionRange(end, end);
                                }}
                                return true;
                            }}
                        }}
                        return false;
                    }};
                    let attempts = 0;
                    const handle = setInterval(() => {{
                        attempts += 1;
                        if (focusByLabel() || attempts >= 20) {{
                            clearInterval(handle);
                        }}
                    }}, 100);
                }})();
                </script>
                """,
                height=0,
                width=0,
            )
            ss["__focus_target"] = None

    work = apply_all(base)
    if use_date_filters and "expiry_date" in work.columns and start and end:
        mask_exp = work["expiry_date"].notna() & (
            work["expiry_date"].dt.date >= start) & (work["expiry_date"].dt.date <= end)
        work = work[mask_exp]
    if use_date_filters and "relabel_to_date" in work.columns and r_start and r_end:
        mask_rel = work["relabel_to_date"].notna() & (
            work["relabel_to_date"].dt.date >= r_start) & (work["relabel_to_date"].dt.date <= r_end)
        work = work[mask_rel]

    selections = {
        "warehouse": list(sel_wh),
        "supplier": list(sel_sup),
        "brand": list(sel_brand),
        "description": list(sel_desc),
        "product_code": list(sel_code),
        "remark": list(sel_remark),
        "use_date_filters": use_date_filters,
        "expiry_range": (start, end) if use_date_filters else None,
        "relabel_range": (r_start, r_end) if use_date_filters else None,
    }
    return work, sel_desc, st.session_state.get("summary_expiry_days", 30), selections


# ----------------------------- Excel 读入与规范化 -----------------------------

@st.cache_data(show_spinner=False)
def _find_sheet_name(file, desired: str) -> Optional[str]:
    try:
        xls = pd.ExcelFile(file, engine="openpyxl")
    except Exception:
        return None
    names = xls.sheet_names
    for n in names:
        if n == desired:
            return n
    for n in names:
        if n.lower() == desired.lower():
            return n
    d_norm = desired.replace(" ", "").lower()
    for n in names:
        if n.replace(" ", "").lower() == d_norm:
            return n
    return None


def _normalize_stocks_report(df: pd.DataFrame) -> Tuple[pd.DataFrame, Optional[str]]:
    expected_map = {
        0: "supplier",
        1: "brand",
        2: "product_code",
        3: "description",
        4: "pack_size",
        5: "unit",
        6: "expiry_date",
        7: "relabel_to_date",
        8: "daily_update_stock_i",
        9: "stock_qty",
    }
    warning = None
    pos_renames = {}
    for idx, name in expected_map.items():
        if idx < df.shape[1]:
            pos_renames[df.columns[idx]] = name

    norm = df.rename(columns=pos_renames).copy()
    norm["warehouse"] = "Savori Whse"

    required = ["supplier", "brand", "product_code", "description", "pack_size", "unit",
                "expiry_date", "relabel_to_date", "stock_qty"]
    missing = [c for c in required if c not in norm.columns]
    if missing:
        warning = (
            "根据预期位置缺少列，或表头不在第 3 行：" + ", ".join(missing) +
            "。请确保 Excel 的 'Stocks report' 工作表表头位于 A3:J3。"
        )

    for col in ["expiry_date", "relabel_to_date"]:
        if col in norm.columns:
            norm[col] = pd.to_datetime(
                norm[col], errors="coerce", format="mixed")

    if "stock_qty" in norm.columns:
        norm["stock_qty"] = pd.to_numeric(norm["stock_qty"].astype(
            str).str.replace(",", "", regex=False), errors="coerce")

    for col in ["supplier", "brand", "product_code", "description", "pack_size", "unit"]:
        if col in norm.columns:
            norm[col] = norm[col].astype("string").str.strip()

    return norm, warning


def _normalize_lai_hock_whse(df: pd.DataFrame) -> Tuple[pd.DataFrame, Optional[str]]:
    expected_map = {
        0: "supplier",
        1: "brand",
        2: "product_code",
        3: "description",
        5: "pack_size",
        6: "unit",
        7: "expiry_date",
        8: "relabel_to_date",
        9: "stocks_balance",
        10: "stock_qty",
    }

    warn_msgs = []
    pos_renames = {}
    for idx, name in expected_map.items():
        if idx < df.shape[1]:
            pos_renames[df.columns[idx]] = name

    norm = df.rename(columns=pos_renames).copy()
    norm["warehouse"] = "Lai Hock Whse"

    required = ["supplier", "brand", "product_code", "description", "pack_size", "unit",
                "expiry_date", "relabel_to_date", "stock_qty"]
    missing = [c for c in required if c not in norm.columns]
    if missing:
        warn_msgs.append(
            "根据预期位置缺少列，或表头不在第 3 行：" + ", ".join(missing) +
            "。请确保 Excel 的 'Lai Hock Whse' 工作表表头位于 A3:K3。"
        )

    for col in ["expiry_date", "relabel_to_date"]:
        if col in norm.columns:
            norm[col] = pd.to_datetime(
                norm[col], errors="coerce", format="mixed")

    if "stock_qty" in norm.columns:
        norm["stock_qty"] = pd.to_numeric(norm["stock_qty"].astype(
            str).str.replace(",", "", regex=False), errors="coerce")

    if "unit" not in norm.columns:
        norm["unit"] = "CTN"
        warn_msgs.append("缺少 Unit 列，已将缺失处默认填为 CTN")
    else:
        norm["unit"] = norm["unit"].astype("string").str.strip()
        norm.loc[norm["unit"].isna() | (norm["unit"] == ""), "unit"] = "CTN"

    for col in ["supplier", "brand", "product_code", "description", "pack_size"]:
        if col in norm.columns:
            norm[col] = norm[col].astype("string").str.strip()

    warning = "；".join(warn_msgs) if warn_msgs else None
    return norm, warning


def load_and_normalize(file) -> Tuple[pd.DataFrame, list]:
    warns = []
    name_sr = _find_sheet_name(file, "Stocks report")
    name_lh = _find_sheet_name(file, "Lai Hock Whse")

    df_sr = None
    df_lh = None
    try:
        if name_sr:
            df_sr = pd.read_excel(file, sheet_name=name_sr, header=2,
                                  dtype=str, engine="openpyxl").dropna(axis=0, how="all")
    except Exception as e:
        warns.append(f"读取工作表 '{name_sr}' 失败：{e}")
    try:
        if name_lh:
            df_lh = pd.read_excel(file, sheet_name=name_lh, header=2,
                                  dtype=str, engine="openpyxl").dropna(axis=0, how="all")
    except Exception as e:
        warns.append(f"读取工作表 '{name_lh}' 失败：{e}")

    frames = []
    if df_sr is not None:
        n_sr, w_sr = _normalize_stocks_report(df_sr)
        if w_sr:
            warns.append(f"Stocks report: {w_sr}")
        frames.append(n_sr)
    else:
        warns.append(
            f"未找到工作表：Stocks report（实际存在：{name_sr if name_sr else '无'}）")

    if df_lh is not None:
        n_lh, w_lh = _normalize_lai_hock_whse(df_lh)
        if w_lh:
            warns.append(f"Lai Hock Whse: {w_lh}")
        frames.append(n_lh)
    else:
        warns.append(
            f"未找到工作表：Lai Hock Whse（实际存在：{name_lh if name_lh else '无'}）")

    if frames:
        cols = [
            "supplier", "brand", "product_code", "description", "pack_size", "unit",
            "expiry_date", "relabel_to_date", "stock_qty", "warehouse",
        ]
        valid_frames = [f for f in frames if f is not None and not f.empty]
        combined = pd.concat(valid_frames, ignore_index=True,
                             sort=False) if valid_frames else pd.DataFrame(columns=cols)
        for c in cols:
            if c not in combined.columns:
                combined[c] = pd.NA
        combined["stock_qty"] = pd.to_numeric(
            combined["stock_qty"], errors="coerce")
        for col in ["expiry_date", "relabel_to_date"]:
            combined[col] = pd.to_datetime(
                combined[col], errors="coerce", format="mixed")
        if "unit" in combined.columns:
            combined["unit"] = combined["unit"].apply(_canonical_unit)
        combined = combined[cols]
        return combined, warns

    return pd.DataFrame(), warns


# ----------------------------- UI 状态回调 -----------------------------

def on_change_expiry_days() -> None:
    value = st.session_state.get("summary_expiry_days", 30)
    try:
        value_int = int(value)
    except (TypeError, ValueError):
        value_int = 30
    if value_int < 1:
        value_int = 1
    st.session_state["summary_expiry_days"] = value_int


def on_change_global_low_stock() -> None:
    for key in ("summary_global_low_ctn", "summary_global_low_pkt"):
        value = st.session_state.get(key, 0)
        try:
            value_int = int(value)
        except (TypeError, ValueError):
            value_int = 0
        if value_int < 0:
            value_int = 0
        st.session_state[key] = value_int


def on_toggle_near_expiry() -> None:
    st.session_state["toggle_only_near"] = bool(
        st.session_state.get("toggle_only_near", False))


def on_toggle_low_stock() -> None:
    st.session_state["toggle_only_low"] = bool(
        st.session_state.get("toggle_only_low", False))


# ----------------------------- 主程序：Stock 页 -----------------------------

def run_stock_page():
    st.title("Stock Dashboard (Stocks DataGrid)")
    st.caption(
        "Upload an Excel file containing the 'Stocks report' and 'Lai Hock Whse' sheets (data starts on row 3).")

    # 朴实无华的文件上传 + 缓存
    uploaded = st.file_uploader(
        "Upload Excel (.xlsx)", type=["xlsx"], key="stock_uploader"
    )
    if uploaded is not None:
        # 把最近一次上传的文件对象存进 session_state，方便页面切换后复用
        st.session_state["stock_file"] = uploaded

    stock_file = st.session_state.get("stock_file")

    if not stock_file:
        st.info("Upload a source workbook to begin.")
        return

    st.caption(f"Using file: {getattr(stock_file, 'name', 'cached workbook')}")

    try:
        df, warns = load_and_normalize(stock_file)
    except Exception as e:
        st.error(f"Failed to read Excel: {e}")
        return

    for w in warns:
        st.warning(w)

    display_cols = [
        "warehouse", "supplier", "brand", "product_code", "description",
        "pack_size", "unit", "expiry_date", "relabel_to_date", "stock_qty",
    ]
    display_cols = [c for c in display_cols if c in df.columns]
    df_display = df[display_cols].copy()

    filtered, selected_descs, _placeholder_due_days, filter_state = apply_filters_v2(
        df_display)
    total_rows = len(filtered)
    filter_summary_parts = []
    label_map = {
        "warehouse": "Warehouse",
        "supplier": "Supplier",
        "brand": "Brand",
        "description": "Description",
        "product_code": "Product Code",
        "remark": "Remark",
    }
    for key, label in label_map.items():
        vals = filter_state.get(key)
        if vals is None or vals == "":
            continue
        if isinstance(vals, (list, tuple, set)):
            formatted = ", ".join(str(v) for v in vals if v)
        else:
            formatted = str(vals).strip()
        if formatted:
            filter_summary_parts.append(f"{label}: {formatted}")

    def _fmt_range(value):
        if not value or not isinstance(value, (list, tuple)) or len(value) != 2:
            return ""
        start, end = value
        if start and end:
            return f"{start} ~ {end}"
        return ""

    if filter_state.get("use_date_filters"):
        exp_range = _fmt_range(filter_state.get("expiry_range"))
        relabel_range = _fmt_range(filter_state.get("relabel_range"))
        if exp_range:
            filter_summary_parts.append(f"Expiry: {exp_range}")
        if relabel_range:
            filter_summary_parts.append(f"Relabel: {relabel_range}")

    filter_summary_text = "; ".join(
        filter_summary_parts) if filter_summary_parts else "None (showing all rows)"
    st.caption(f"Active filters: {filter_summary_text}")

    summary_df, detail_map = aggregate_summary(filtered)
    summary_df = summary_df.copy()

    if "summary_expiry_days" not in st.session_state:
        st.session_state["summary_expiry_days"] = 30
    if "summary_global_low_ctn" not in st.session_state:
        st.session_state["summary_global_low_ctn"] = 0
    if "summary_global_low_pkt" not in st.session_state:
        st.session_state["summary_global_low_pkt"] = 0
    if "toggle_only_near" not in st.session_state:
        st.session_state["toggle_only_near"] = False
    if "toggle_only_low" not in st.session_state:
        st.session_state["toggle_only_low"] = False
    if "toggle_show_depleted" not in st.session_state:
        st.session_state["toggle_show_depleted"] = True
    if "product_quick_filter" not in st.session_state:
        st.session_state["product_quick_filter"] = ""

    summary_bar = st.container()
    with summary_bar:
        metric_holder = st.container()
        controls_holder = st.container()

    status_options = ["Expired", "Near-Expiry", "Low-Stock", "Depleted", "OK"]
    with controls_holder:
        toggle_cols = st.columns([1, 1, 1, 2])
        toggle_cols[0].toggle(
            "Show Near-Expiry Only", key="toggle_only_near", on_change=on_toggle_near_expiry)
        toggle_cols[1].toggle(
            "Show Low-Stock Only", key="toggle_only_low", on_change=on_toggle_low_stock)
        toggle_cols[2].toggle("Show Depleted", key="toggle_show_depleted")
        with toggle_cols[3]:
            threshold_cols = st.columns([1, 1, 1])
            threshold_cols[0].number_input(
                "expiry_days (days)", min_value=1, max_value=365, step=1,
                key="summary_expiry_days", on_change=on_change_expiry_days, format="%d",
                help="Day threshold applied when tagging Near-Expiry."
            )
            threshold_cols[1].number_input(
                "global_low_stock ctns", min_value=0, step=1,
                key="summary_global_low_ctn", on_change=on_change_global_low_stock, format="%d",
                help="Global carton threshold when a product ROP is missing."
            )
            threshold_cols[2].number_input(
                "global_low_stock pkts", min_value=0, step=1,
                key="summary_global_low_pkt", on_change=on_change_global_low_stock, format="%d",
                help="Global packet threshold when a product ROP is missing."
            )
            mode_label = st.radio(
                "批次分组显示",
                options=["Expiry only", "Remark only", "Expiry + Remark"],
                horizontal=True,
                key="batch_mode_radio",
                help="选择按到期、按备注，或两者同时分组展示批次"
            )
            mode_map = {
                "Expiry only": "expiry",
                "Remark only": "remark",
                "Expiry + Remark": "both",
            }
            batch_mode = mode_map.get(mode_label, "expiry")

        product_query_raw = st.text_input(
            "Product Quick Filter",
            key="product_quick_filter",
            placeholder="Filter by supplier / brand / product / code",
            help="Client-side filter that applies to Supplier, Brand, Product, and Product Code.",
        )
        status_selected = st.multiselect(
            "Status filter",
            options=status_options,
            default=status_options,
            key="status_filter_options",
            help="Limit the main table by status tags. Clear selection to show all statuses.",
        )

    expiry_days = int(st.session_state.get("summary_expiry_days", 30))
    global_low_ctn = int(st.session_state.get("summary_global_low_ctn", 0))
    global_low_pkt = int(st.session_state.get("summary_global_low_pkt", 0))
    near_only = bool(st.session_state.get("toggle_only_near", False))
    low_only = bool(st.session_state.get("toggle_only_low", False))
    show_depleted = bool(st.session_state.get("toggle_show_depleted", True))
    product_query = (product_query_raw or "").strip()
    status_selected = status_selected or status_options

    if summary_df.empty:
        with metric_holder:
            metric_cols = st.columns([1, 1, 1, 1])
            metric_cols[0].metric("Filtered Rows", f"{total_rows}")
            metric_cols[1].metric("Totals", "0")
            metric_cols[2].metric("Near-Expiry", "0")
            metric_cols[3].metric("Low-Stock", "0")
        st.info(
            "No data found for Savori Whse / Lai Hock Whse under the current filters.")
        return

    from functools import lru_cache
    normalized_detail_map = {
        tuple(k) if not isinstance(k, tuple) else tuple(k): v
        for k, v in (detail_map or {}).items()
        if v is not None
    }

    @lru_cache(maxsize=2048)
    def _cached_split(group_key_tuple, expiry_days, show_depleted, batch_mode):
        key = tuple(group_key_tuple) if not isinstance(
            group_key_tuple, tuple) else group_key_tuple
        frame = normalized_detail_map.get(key)
        if frame is None or frame.empty:
            return pd.DataFrame()
        return split_by_expiry(frame, expiry_days=expiry_days, show_depleted=show_depleted, batch_mode=batch_mode)

    def _get_expiry_table(group_key_tuple):
        return _cached_split(tuple(group_key_tuple) if not isinstance(group_key_tuple, tuple) else group_key_tuple, expiry_days, show_depleted, batch_mode)

    def _opt_float(val):
        try:
            return float(val)
        except Exception:
            return None

    summary_df["is_low_stock"] = False
    summary_df["low_stock_reason"] = None
    for i, r in summary_df.iterrows():
        total_ctn = _opt_float(r.get("total_ctn")) or 0.0
        total_pkt = _opt_float(r.get("total_pkt")) or 0.0
        rop_ctn = _opt_float(r.get("reorder_point_ctn"))
        rop_pkt = _opt_float(r.get("reorder_point_pkt"))
        if rop_ctn is not None and total_ctn <= rop_ctn + 1e-9:
            summary_df.at[i, "is_low_stock"] = True
            summary_df.at[i, "low_stock_reason"] = "ROP"
        elif rop_pkt is not None and total_pkt <= rop_pkt + 1e-9:
            summary_df.at[i, "is_low_stock"] = True
            summary_df.at[i, "low_stock_reason"] = "ROP"
        else:
            if (global_low_ctn and total_ctn <= float(global_low_ctn) + 1e-9) or (global_low_pkt and total_pkt <= float(global_low_pkt) + 1e-9):
                summary_df.at[i, "is_low_stock"] = True
                summary_df.at[i, "low_stock_reason"] = "Global"

    summary_df["has_expired_batch"] = False
    summary_df["has_near_batch"] = False
    for i, r in summary_df.iterrows():
        tuple_key = tuple(r["group_key"]) if not isinstance(
            r["group_key"], tuple) else r["group_key"]
        tbl = _get_expiry_table(tuple_key)
        if tbl is None or tbl.empty:
            continue
        pos_qty = (tbl["subtotal_ctn"].astype(float) +
                   tbl["subtotal_pkt"].astype(float)) > 0
        if not pos_qty.any():
            continue
        sub = tbl.loc[pos_qty]
        if (sub["status_batch"] == "Expired").any():
            summary_df.at[i, "has_expired_batch"] = True
        if (sub["status_batch"] == "Near-Expiry").any():
            summary_df.at[i, "has_near_batch"] = True

    summary_df["is_depleted"] = (
        (summary_df["total_ctn"].fillna(0) +
         summary_df["total_pkt"].fillna(0)) == 0
    )

    def _primary_status(has_expired: bool, has_near: bool, is_low: bool, is_depleted: bool) -> str:
        if is_depleted:
            return "Depleted"
        if has_expired:
            return "Expired"
        if has_near:
            return "Near-Expiry"
        if is_low:
            return "Low-Stock"
        return "OK"

    summary_df["Status Tags"] = [[] for _ in range(len(summary_df))]
    summary_df["status_product"] = ""

    for i, r in summary_df.iterrows():
        tags = []
        if bool(r["is_depleted"]):
            tags.append("Depleted")
        if bool(r["has_expired_batch"]):
            tags.append("Expired")
        if bool(r["has_near_batch"]):
            tags.append("Near-Expiry")
        if bool(r["is_low_stock"]):
            tags.append("Low-Stock")
        if not tags:
            tags = ["OK"]

        summary_df.at[i, "Status Tags"] = tags
        summary_df.at[i, "status_product"] = _primary_status(
            bool(r["has_expired_batch"]),
            bool(r["has_near_batch"]),
            bool(r["is_low_stock"]),
            bool(r["is_depleted"]),
        )

    if not show_depleted:
        summary_df = summary_df[summary_df["status_product"] != "Depleted"]

    summary_df["Savori Whse"] = [_format_quantity_pair(
        r.savori_ctn, r.savori_pkt) for r in summary_df.itertuples()]
    summary_df["Lai Hock Whse"] = [_format_quantity_pair(
        r.lai_hock_ctn, r.lai_hock_pkt) for r in summary_df.itertuples()]
    summary_df["Total"] = [_format_quantity_pair(
        r.total_ctn, r.total_pkt) for r in summary_df.itertuples()]

    def _format_status(tags, reason):
        if not isinstance(tags, (list, tuple)):
            return str(tags)
        formatted = []
        for tag in tags:
            if tag == "Low-Stock" and reason:
                formatted.append(f"{tag} ({reason})")
            else:
                formatted.append(str(tag))
        return " • ".join(formatted) if formatted else "OK"

    summary_df["Status"] = [
        _format_status(tags, reason)
        for tags, reason in zip(summary_df["Status Tags"], summary_df["low_stock_reason"])
    ]

    view_df = summary_df.copy()

    def _has_near(tags):
        return any(t in {"Expired", "Near-Expiry"} for t in (tags or []))

    def _has_low(tags):
        return "Low-Stock" in (tags or [])

    near_mask = view_df["Status Tags"].apply(_has_near)
    low_mask = view_df["Status Tags"].apply(_has_low)

    if near_only and not low_only:
        view_df = view_df[near_mask]
    elif low_only and not near_only:
        view_df = view_df[low_mask]
    elif near_only and low_only:
        view_df = view_df[near_mask | low_mask]

    if product_query:
        mask = (
            view_df["Product"].str.contains(
                product_query, case=False, na=False)
            | view_df["Supplier"].str.contains(product_query, case=False, na=False)
            | view_df["Brand"].str.contains(product_query, case=False, na=False)
            | view_df["Product Code"].str.contains(product_query, case=False, na=False)
        )
        view_df = view_df[mask]
    if status_selected and set(status_selected) != set(status_options):
        view_df = view_df[view_df["status_product"].isin(status_selected)]
    view_df = view_df.reset_index(drop=True)

    if view_df.empty:
        with metric_holder:
            metric_cols = st.columns([1, 1, 1, 1])
            metric_cols[0].metric("Filtered Rows", f"{total_rows}")
            metric_cols[1].metric("Totals", "0")
            metric_cols[2].metric("Near-Expiry", "0")
            metric_cols[3].metric("Low-Stock", "0")
        st.info(
            "No rows in the main table match the current view. Adjust filters or thresholds.")
        return

    total_ctn_all = float(view_df["total_ctn"].sum()
                          ) if not view_df.empty else 0.0
    total_pkt_all = float(view_df["total_pkt"].sum()
                          ) if not view_df.empty else 0.0
    totals_display = _format_quantity_pair(total_ctn_all, total_pkt_all)

    near_count = int(view_df["Status Tags"].apply(lambda tags: any(
        t in {"Expired", "Near-Expiry"} for t in (tags or []))).sum())
    low_count = int(view_df["Status Tags"].apply(
        lambda tags: "Low-Stock" in (tags or [])).sum())

    with metric_holder:
        metric_cols = st.columns([1, 1, 1, 1])
        metric_cols[0].metric("Filtered Rows", f"{total_rows}")
        metric_cols[1].metric(
            "Totals", totals_display if totals_display else "0")
        metric_cols[2].metric("Near-Expiry", f"{near_count}")
        metric_cols[3].metric("Low-Stock", f"{low_count}")

    nearest_days_map = {}
    for k in view_df["group_key"]:
        key = tuple(k) if not isinstance(k, tuple) else k
        tbl = _get_expiry_table(key)
        if tbl is None or tbl.empty:
            nearest_days_map[key] = 10**9
            continue
        pos = (tbl["subtotal_ctn"].astype(float) +
               tbl["subtotal_pkt"].astype(float)) > 0
        if not pos.any():
            nearest_days_map[key] = 10**9
            continue
        vals = tbl.loc[pos, "days_to_expiry"].dropna()
        try:
            nearest_days_map[key] = int(
                vals.min()) if not vals.empty else 10**9
        except Exception:
            nearest_days_map[key] = 10**9

    def _status_rank(s: str) -> int:
        return {"Depleted": 0, "Expired": 1, "Near-Expiry": 2, "Low-Stock": 3, "OK": 4}.get(s, 9)

    view_df["status_priority"] = view_df["status_product"].map(_status_rank)
    view_df["nearest_days"] = [nearest_days_map.get(tuple(k) if not isinstance(k, tuple) else k, 10**9)
                               for k in view_df["group_key"]]
    view_df = view_df.sort_values(
        by=["status_priority", "nearest_days",
            "total_ctn", "total_pkt", "Product"],
        kind="stable",
    ).reset_index(drop=True)

    display_columns = ["Supplier", "Brand", "Product", "Pack Size",
                       "Product Code", "Savori Whse", "Lai Hock Whse", "Total", "Status"]
    view_df_display = view_df[display_columns].copy()
    view_df_display_clean = _strip_html_df(view_df_display)
    total_idx = list(view_df_display_clean.columns).index("Total")

    def _tag_code(tags):
        s = set(tags or [])
        if "Expired" in s and "Low-Stock" in s:
            return 3
        if "Near-Expiry" in s and "Low-Stock" in s:
            return 2
        if "Expired" in s:
            return 1
        if "Near-Expiry" in s:
            return 4
        if "Low-Stock" in s:
            return 5
        if "Depleted" in s:
            return 6
        return 0
    view_df["__row_code"] = view_df["Status Tags"].apply(_tag_code)

    def _style_row(row):
        tags = view_df.loc[row.name, "Status Tags"]
        styles = ["" for _ in row]

        if "Expired" in tags and "Low-Stock" in tags:
            styles = ["background-color: rgba(102,51,153,0.25);"] * len(row)
            styles[total_idx] += "font-weight:600; color:#FFF;"
        elif "Near-Expiry" in tags and "Low-Stock" in tags:
            styles = ["background-color: rgba(255,140,0,0.22);"] * len(row)
            styles[total_idx] += "font-weight:600; color:#FFF;"
        elif "Expired" in tags:
            styles = ["background-color: rgba(178,34,34,0.18);"] * len(row)
        elif "Near-Expiry" in tags:
            styles = ["background-color: rgba(255,165,0,0.18);"] * len(row)
        elif "Low-Stock" in tags:
            styles[total_idx] = "background-color: rgba(255, 255, 0, 0.18); font-weight:600;"
        elif "Depleted" in tags:
            styles = ["background-color: rgba(128,128,128,0.12);"] * len(row)
        return styles

    styled_view = view_df_display_clean.style.apply(_style_row, axis=1)
    st.dataframe(styled_view, width="stretch", hide_index=True)

    export_summary_cols = [
        "Supplier", "Brand", "Product", "Pack Size", "Product Code",
        "Savori Whse", "Lai Hock Whse", "Total", "Status",
        "total_ctn", "total_pkt", "savori_ctn", "savori_pkt", "lai_hock_ctn", "lai_hock_pkt",
        "reorder_point_ctn", "reorder_point_pkt", "low_stock_reason", "status_product",
    ]
    export_summary = view_df[export_summary_cols].copy()

    expiry_export_frames = []
    for _, row in view_df.iterrows():
        tuple_key = tuple(row["group_key"]) if not isinstance(
            row["group_key"], tuple) else row["group_key"]
        expiry_table = _get_expiry_table(tuple_key)
        if expiry_table is None or expiry_table.empty:
            continue
        export_block = expiry_table.copy()
        export_block.insert(0, "Supplier", row["Supplier"])
        export_block.insert(1, "Product", row["Product"])
        export_block.insert(2, "Product Code", row["Product Code"])
        export_block.insert(3, "Product Status", row["status_product"])
        export_block.insert(4, "Parent Total", row["Total"])
        expiry_export_frames.append(export_block)

    valid_frames = [
        f for f in expiry_export_frames if f is not None and not f.empty]

    expected_cols = [
        "Supplier", "Product", "Product Code", "Product Status", "Parent Total",
        "Expiry", "Remark", "Savori Whse", "Lai Hock Whse", "Subtotal",
        "subtotal_ctn", "subtotal_pkt", "expiry_date", "status_batch", "days_to_expiry", "Info",
    ]

    normalized = []
    for f in valid_frames:
        g = f.copy()
        for c in expected_cols:
            if c not in g.columns:
                g[c] = pd.Series([pd.NA] * len(g), dtype="string")

        for c in ["Supplier", "Product", "Product Code", "Product Status",
                  "Parent Total", "Expiry", "Remark", "Savori Whse", "Lai Hock Whse",
                  "Subtotal", "status_batch", "Info"]:
            if c in g.columns:
                if g[c].dtype != "string":
                    g[c] = g[c].astype("string")

        for c in ["subtotal_ctn", "subtotal_pkt"]:
            if c in g.columns:
                g[c] = pd.to_numeric(g[c], errors="coerce").astype("float64")

        if "expiry_date" in g.columns:
            g["expiry_date"] = pd.to_datetime(
                g["expiry_date"], errors="coerce")

        if "days_to_expiry" in g.columns:
            g["days_to_expiry"] = pd.to_numeric(
                g["days_to_expiry"], errors="coerce").astype("Int64")

        g = g[expected_cols]
        normalized.append(g)

    if normalized:
        expiry_export_df = pd.concat(normalized, ignore_index=True)
    else:
        expiry_export_df = pd.DataFrame(columns=expected_cols)

    detail_frames = []
    if detail_map and not view_df.empty:
        key_order = [
            tuple(k) if not isinstance(k, tuple) else tuple(k)
            for k in view_df["group_key"]
        ]
        for key in dict.fromkeys(key_order):
            frame = detail_map.get(key)
            if frame is not None and not frame.empty:
                detail_frames.append(frame.copy())
    if detail_frames:
        detail_export_df = pd.concat(
            detail_frames, ignore_index=True, sort=False)
        helper_cols = [
            c for c in detail_export_df.columns if str(c).startswith("_")]
        detail_export_df = detail_export_df.drop(
            columns=helper_cols, errors="ignore")
        for col in ["expiry_date", "relabel_to_date"]:
            if col in detail_export_df.columns:
                detail_export_df[col] = pd.to_datetime(
                    detail_export_df[col], errors="coerce").dt.strftime("%Y-%m-%d")
    else:
        detail_export_df = pd.DataFrame()

    download_cols = st.columns([1, 1, 1, 1])
    csv_buffer = io.StringIO()
    export_summary_clean = _strip_html_df(export_summary)
    export_summary_clean.to_csv(csv_buffer, index=False)

    download_cols[0].download_button(
        "Export Summary CSV", data=csv_buffer.getvalue(),
        file_name="stock_summary.csv", mime="text/csv",
    )

    excel_buffer = io.BytesIO()
    with pd.ExcelWriter(excel_buffer) as writer:
        export_summary_clean = _strip_html_df(export_summary)
        expiry_export_clean = _strip_html_df(expiry_export_df)
        export_summary_clean.to_excel(
            writer, sheet_name="Summary", index=False)
        expiry_export_clean.to_excel(
            writer, sheet_name="Expiry Breakdown", index=False)

        meta_rows = []
        for k, v in (filter_state or {}).items():
            meta_rows.append({"key": str(k), "value": str(v)})
        meta_df = pd.DataFrame(meta_rows, columns=["key", "value"])
        meta_df.to_excel(writer, sheet_name="Filters", index=False)

    excel_buffer.seek(0)
    download_cols[1].download_button(
        "Export Excel", data=excel_buffer.getvalue(),
        file_name="stock_summary.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    detail_csv_buffer = io.StringIO()
    detail_export_clean = _strip_html_df(
        detail_export_df) if not detail_export_df.empty else detail_export_df
    detail_export_clean.to_csv(detail_csv_buffer, index=False)
    download_cols[2].download_button(
        "Export Filtered Detail CSV",
        data=detail_csv_buffer.getvalue(),
        file_name="stock_filtered_detail.csv",
        mime="text/csv",
        help="Download the raw rows that back the current table after all filters, quick search, and status filters.",
    )

    detail_excel_buffer = io.BytesIO()
    with pd.ExcelWriter(detail_excel_buffer) as writer:
        detail_export_clean.to_excel(
            writer, sheet_name="Stock Report", index=False)
        meta_rows = []
        for k, v in (filter_state or {}).items():
            meta_rows.append({"key": str(k), "value": str(v)})
        pd.DataFrame(meta_rows, columns=["key", "value"]).to_excel(
            writer, sheet_name="Filters", index=False)
    detail_excel_buffer.seek(0)
    download_cols[3].download_button(
        "Export Stock Report Excel",
        data=detail_excel_buffer.getvalue(),
        file_name="stock_report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        help="Export the filtered stock report as an Excel file.",
    )

    for _, row in view_df.iterrows():
        tuple_key = tuple(row["group_key"]) if not isinstance(
            row["group_key"], tuple) else row["group_key"]
        expiry_table = _get_expiry_table(tuple_key)
        title = f"{row['Supplier']} | {row['Product']}"
        with st.expander(title, expanded=False):
            st.caption(f"Product Code: {row['Product Code']}")
            st.caption(f"Status: {row['Status']}")
            if row.get("low_stock_reason"):
                st.caption(
                    f"Low-stock trigger: {'Product ROP' if row['low_stock_reason'] == 'ROP' else 'Global threshold'}")

            if expiry_table is None or expiry_table.empty:
                st.info("No expiry breakdown available.")
            else:
                if batch_mode == "expiry":
                    display_cols2 = [
                        "Expiry", "Savori Whse", "Lai Hock Whse", "Subtotal", "status_batch", "Info"]
                elif batch_mode == "remark":
                    display_cols2 = [
                        "Remark", "Savori Whse", "Lai Hock Whse", "Subtotal", "status_batch", "Info"]
                else:
                    display_cols2 = ["Expiry", "Remark", "Savori Whse",
                                     "Lai Hock Whse", "Subtotal", "status_batch", "Info"]

                table_display = expiry_table[display_cols2].rename(
                    columns={"status_batch": "Batch Status"})

                def _style_expiry(r):
                    status = table_display.loc[r.name, "Batch Status"]
                    if status == "Expired":
                        return ["background-color: rgba(220,20,60,0.16);"] * len(r)
                    if status == "Near-Expiry":
                        return ["background-color: rgba(255,165,0,0.18);"] * len(r)
                    if status == "Depleted":
                        return ["background-color: rgba(128,128,128,0.12);"] * len(r)
                    return ["" for _ in r]

                try:
                    styled_expiry = table_display.style.apply(
                        _style_expiry, axis=1)
                    st.dataframe(
                        styled_expiry, width="stretch", hide_index=True)
                except Exception:
                    st.dataframe(
                        table_display, width="stretch", hide_index=True)

                subtotal_ctn_sum = float(expiry_table["subtotal_ctn"].sum())
                subtotal_pkt_sum = float(expiry_table["subtotal_pkt"].sum())
                parent_ctn = float(row["total_ctn"])
                parent_pkt = float(row["total_pkt"])
                if (not math.isclose(subtotal_ctn_sum, parent_ctn, abs_tol=1e-6)
                        or not math.isclose(subtotal_pkt_sum, parent_pkt, abs_tol=1e-6)):
                    st.warning(
                        "Expiry breakdown totals do not match the parent totals; please verify the source data.")
                st.caption(f"Total: {row['Total']}")


# ----------------------------- Sales helpers -----------------------------

def _read_delivery_details_from_bytes(content: bytes, label: str) -> Tuple[pd.DataFrame, Optional[str]]:
    sheet_name = _find_sheet_name(io.BytesIO(content), "Delivery details")
    if not sheet_name:
        return pd.DataFrame(), f"{label}: 未找到 'Delivery details' 工作表。"
    try:
        frame = pd.read_excel(
            io.BytesIO(content),
            sheet_name=sheet_name,
            header=3,
            dtype=str,
            engine="openpyxl",
        ).dropna(axis=0, how="all")
        return frame, None
    except Exception as exc:
        return pd.DataFrame(), f"{label}: 读取 '{sheet_name}' 失败：{exc}"


def _normalize_sales_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    trimmed = df.copy().iloc[:, : len(SALES_COLUMNS)]
    trimmed.columns = SALES_COLUMNS[:trimmed.shape[1]]
    for col in SALES_COLUMNS:
        if col not in trimmed.columns:
            trimmed[col] = pd.NA
    trimmed = trimmed[SALES_COLUMNS].replace(r"^\s*$", pd.NA, regex=True)
    trimmed = _strip_html_df(trimmed)
    for col in trimmed.select_dtypes(include="object").columns:
        trimmed[col] = trimmed[col].astype("string").str.strip()
    for col in SALES_NUMERIC_COLUMNS:
        trimmed[col] = pd.to_numeric(trimmed[col], errors="coerce")
    for col in SALES_DATE_COLUMNS:
        if not pd.api.types.is_datetime64_any_dtype(trimmed[col]):
            trimmed[col] = pd.to_datetime(trimmed[col], errors="coerce")
    trimmed["carton_packing_numeric"] = trimmed["Carton Packing"].map(
        _infer_carton_pack_size)
    trimmed["pcs_to_ctn"] = trimmed["Qty in Pcs"] / \
        trimmed["carton_packing_numeric"]
    valid_pack = trimmed["carton_packing_numeric"].gt(0)
    trimmed["pcs_to_ctn"] = trimmed["pcs_to_ctn"].where(valid_pack, pd.NA)
    trimmed["Unit Price (per pcs)"] = _safe_divide_series(
        trimmed["Total Value"], trimmed["Qty in Pcs"])
    trimmed["Unit Price (per carton)"] = _safe_divide_series(
        trimmed["Total Value"], trimmed["Qty in Ctns"])
    return trimmed


@st.cache_data(show_spinner=False)
def _cached_sales_dataset(inputs: Tuple[Tuple[str, bytes], ...]) -> Tuple[pd.DataFrame, List[str]]:
    warns: List[str] = []
    frames: List[pd.DataFrame] = []
    for name, content in inputs:
        if not content:
            warns.append(f"{name}: 文件大小为 0，无法读取。")
            continue
        frame, err = _read_delivery_details_from_bytes(content, name)
        if err:
            warns.append(err)
            continue
        if frame.empty:
            warns.append(f"{name}: 'Delivery details' 表为空。")
            continue
        frames.append(frame)
    if not frames:
        return pd.DataFrame(columns=SALES_COLUMNS), warns
    combined = pd.concat(frames, ignore_index=True, sort=False)
    return _normalize_sales_dataframe(combined), warns


def load_sales_data(files) -> Tuple[pd.DataFrame, List[str]]:
    inputs: List[Tuple[str, bytes]] = []
    for uploaded in files or []:
        if uploaded is None:
            continue
        content = uploaded.getvalue()
        inputs.append((uploaded.name, content))
    if not inputs:
        return pd.DataFrame(columns=SALES_COLUMNS), []
    return _cached_sales_dataset(tuple(inputs))


def _ensure_multiselect_key_state(state_key: str, options: List[str], default: List[str]):
    # 第一次出现这个 key，用 default 初始化（注意要跟 options 交集一下）
    if state_key not in st.session_state:
        base = [v for v in default if v in options]
        st.session_state[state_key] = base
        return

    # 之后就只做「清理无效选项」，不再强制塞默认值
    cur = st.session_state.get(state_key, [])
    if isinstance(cur, (str, int, float)):
        cur = [str(cur)]
    # 只保留还在 options 里的选项
    cur = [v for v in cur if v in options]
    st.session_state[state_key] = cur


def apply_sales_filters(df: pd.DataFrame) -> Tuple[pd.DataFrame, Dict[str, str]]:
    if df.empty:
        return df, {}

    ss = st.session_state
    session_keys = {
        'Year': 'sales_filter_year',
        'Month': 'sales_filter_month',
        'Customer': 'sales_filter_customer',
        'Outlet': 'sales_filter_outlet',
        'Product Description': 'sales_filter_product_description',
        'Supplier': 'sales_filter_supplier',
        'Brand/Category': 'sales_filter_brand',
        'Product Code': 'sales_filter_product_code',
        'Account': 'sales_filter_account',
    }
    customer_exclude_key = "sales_filter_customer_exclude"

    def _series(name: str) -> pd.Series:
        if name not in df.columns:
            return pd.Series([''] * len(df), index=df.index)
        return df[name].astype('string').fillna('').str.strip()

    series_cache = {name: _series(name) for name in session_keys}
    invoice_series = _series('Invoice #')

    filter_columns = {
        'Year': 'Year',
        'Month': 'Month',
        'Customer': 'Customer',
        'Outlet': 'Outlet',
        'Product Description': 'Product Description',
        'Supplier': 'Supplier',
        'Brand/Category': 'Brand/Category',
        'Product Code': 'Product Code',
        'Account': 'Account',
    }

    date_series_raw = None
    default_date_from: Optional[datetime.date] = None
    default_date_to: Optional[datetime.date] = None
    if 'Date' in df.columns:
        date_series_raw = pd.to_datetime(df['Date'], errors='coerce')
        valid_dates = date_series_raw.dropna()
        if not valid_dates.empty:
            default_date_from = valid_dates.min().date()
            default_date_to = valid_dates.max().date()

    date_from: Optional[datetime.date] = None
    date_to: Optional[datetime.date] = None

    def _mask(
        exclude: Optional[str] = None,
        ignore_customer_exclude: bool = False,
    ) -> pd.Series:
        mask = pd.Series(True, index=df.index)
        for key, column in filter_columns.items():
            if key == exclude:
                continue
            sel = ss.get(session_keys[key], [])
            if not sel:
                continue
            mask &= series_cache[column].isin(sel)
        if not ignore_customer_exclude and exclude != "Exclude Customer":
            excluded_customers = ss.get(customer_exclude_key, [])
            if excluded_customers:
                mask &= ~series_cache["Customer"].isin(excluded_customers)
        if exclude != 'Invoice contains':
            invoice_query = ss.get('sales_filter_invoice', '')
            if invoice_query:
                mask &= invoice_series.str.contains(
                    invoice_query, case=False, na=False)
        if date_series_raw is not None:
            if date_from:
                mask &= date_series_raw >= pd.Timestamp(date_from)
            if date_to:
                mask &= date_series_raw <= pd.Timestamp(date_to)
        return mask

    def _options_for(key: str) -> List[str]:
        column = filter_columns[key]
        mask = _mask(
            exclude=key,
            ignore_customer_exclude=(key == "Customer"),
        )
        values = sorted(
            {val for val in series_cache[column][mask] if val},
            key=_month_sort_key if key == 'Month' else lambda x: x.upper(),
        )
        return values

    ss.setdefault(session_keys['Year'], [])
    ss.setdefault(session_keys['Month'], [])

    with st.form("sales_filters_form"):
        year_options = _options_for('Year')
        default_year: List[str] = []
        if year_options:
            target_year = '2025'
            if target_year in year_options:
                default_year = [target_year]
            else:
                default_year = [year_options[-1]]

        month_options = _options_for('Month')
        default_month: List[str] = []

        _ensure_multiselect_key_state(
            session_keys['Year'], year_options, default_year)
        _ensure_multiselect_key_state(
            session_keys['Month'], month_options, default_month)

        customer_options = _options_for('Customer')
        _ensure_multiselect_key_state(
            session_keys['Customer'], customer_options, default=[])
        _ensure_multiselect_key_state(
            customer_exclude_key, customer_options, default=[])
        outlet_options = _options_for('Outlet')
        _ensure_multiselect_key_state(
            session_keys['Outlet'], outlet_options, default=[])

        multi_cols = st.columns(5)
        with multi_cols[0]:
            st.multiselect(
                'Year',
                year_options,
                key=session_keys['Year'],
            )
        with multi_cols[1]:
            st.multiselect(
                'Month',
                month_options,
                key=session_keys['Month'],
            )
        with multi_cols[2]:
            st.multiselect(
                'Customer',
                customer_options,
                key=session_keys['Customer'],
            )
        with multi_cols[3]:
            st.multiselect(
                'Exclude Customer',
                customer_options,
                key=customer_exclude_key,
                help='Selected customers will be excluded from results.',
            )
        with multi_cols[4]:
            st.multiselect(
                'Outlet',
                outlet_options,
                key=session_keys['Outlet'],
            )

        product_desc_options = _options_for('Product Description')
        _ensure_multiselect_key_state(
            session_keys['Product Description'], product_desc_options, default=[])
        st.multiselect(
            'Product Description',
            product_desc_options,
            key=session_keys['Product Description'],
        )

        with st.expander('高级筛选', expanded=False):
            supplier_options = _options_for('Supplier')
            _ensure_multiselect_key_state(
                session_keys['Supplier'], supplier_options, default=[])
            brand_options = _options_for('Brand/Category')
            _ensure_multiselect_key_state(
                session_keys['Brand/Category'], brand_options, default=[])
            product_code_options = _options_for('Product Code')
            _ensure_multiselect_key_state(
                session_keys['Product Code'], product_code_options, default=[])
            account_options = _options_for('Account')
            _ensure_multiselect_key_state(
                session_keys['Account'], account_options, default=[])
            st.multiselect(
                'Supplier',
                supplier_options,
                key=session_keys['Supplier'],
            )
            st.multiselect(
                'Brand/Category',
                brand_options,
                key=session_keys['Brand/Category'],
            )
            st.multiselect(
                'Product Code',
                product_code_options,
                key=session_keys['Product Code'],
            )
            st.multiselect(
                'Account',
                account_options,
                key=session_keys['Account'],
            )
            st.text_input(
                'Invoice # contains',
                key='sales_filter_invoice',
                placeholder='',
            )
            date_preset_options = [
                'Default range',
                'This week',
                'Last week',
                'Next week',
                'Custom range',
            ]
            ss.setdefault('sales_date_preset', date_preset_options[0])
            preset = st.selectbox(
                'Date shortcut',
                date_preset_options,
                key='sales_date_preset',
            )
            today_date = datetime.date.today()

            def _week_bounds(offset: int) -> Tuple[datetime.date, datetime.date]:
                monday = today_date - \
                    datetime.timedelta(days=today_date.weekday())
                start = monday + datetime.timedelta(weeks=offset)
                end = start + datetime.timedelta(days=6)
                return start, end

            preset_from = default_date_from or today_date
            preset_to = default_date_to or today_date
            if preset == 'This week':
                preset_from, preset_to = _week_bounds(0)
            elif preset == 'Last week':
                preset_from, preset_to = _week_bounds(-1)
            elif preset == 'Next week':
                preset_from, preset_to = _week_bounds(1)
            elif preset == 'Default range':
                preset_from = default_date_from or today_date
                preset_to = default_date_to or today_date

            if preset != 'Custom range':
                st.caption(
                    f"Quick filter '{preset}': {preset_from} to {preset_to}")
                date_from = preset_from
                date_to = preset_to
                ss['sales_filter_date_from'] = date_from
                ss['sales_filter_date_to'] = date_to
            else:
                custom_from_default = ss.get(
                    'sales_filter_date_from',
                    default_date_from or today_date,
                )
                custom_to_default = ss.get(
                    'sales_filter_date_to',
                    default_date_to or today_date,
                )
                date_from = st.date_input(
                    'Date from',
                    value=custom_from_default,
                    key='sales_filter_date_from',
                )
                date_to = st.date_input(
                    'Date to',
                    value=custom_to_default,
                    key='sales_filter_date_to',
                )

        submitted = st.form_submit_button("Apply sales filters")

    filtered_df = df.loc[_mask()].reset_index(drop=True)
    selection_values = {
        'Year': ss.get(session_keys['Year'], []),
        'Month': ss.get(session_keys['Month'], []),
        'Customer': ss.get(session_keys['Customer'], []),
        'Exclude Customer': ss.get(customer_exclude_key, []),
        'Outlet': ss.get(session_keys['Outlet'], []),
        'Product Description': ss.get(session_keys['Product Description'], []),
        'Supplier': ss.get(session_keys['Supplier'], []),
        'Brand/Category': ss.get(session_keys['Brand/Category'], []),
        'Product Code': ss.get(session_keys['Product Code'], []),
        'Account': ss.get(session_keys['Account'], []),
    }
    filter_state: Dict[str, str] = {}
    for key, vals in selection_values.items():
        filter_state[key] = ', '.join(vals) if vals else 'All'
    invoice_query = ss.get('sales_filter_invoice', '')
    filter_state['Invoice contains'] = invoice_query if invoice_query else 'All'
    date_from_state = ss.get('sales_filter_date_from')
    date_to_state = ss.get('sales_filter_date_to')
    filter_state['Date from'] = (
        date_from_state.isoformat() if date_from_state else 'All'
    )
    filter_state['Date to'] = (
        date_to_state.isoformat() if date_to_state else 'All'
    )

    return filtered_df, filter_state


def build_sales_summary(
    filtered_df: pd.DataFrame,
    group_by_outlet: bool,
    include_date: bool,
    group_by_week: bool,
) -> pd.DataFrame:
    include_date = include_date and group_by_outlet and 'Date' in filtered_df.columns
    has_outlet = group_by_outlet and 'Outlet' in filtered_df.columns
    week_enabled = group_by_week and 'Date' in filtered_df.columns

    if filtered_df.empty:
        cols = ['Customer', 'Total Value', 'Total Qty']
        if week_enabled:
            cols.insert(0, 'Week')
        if not has_outlet:
            cols.append('Selling Price (per pkt)')
        return pd.DataFrame(columns=cols)

    working_df = filtered_df.copy()
    week_label_key = '_week_label'

    def _format_week_label(value: pd.Timestamp) -> str:
        if pd.isna(value):
            return ""
        dt = value.date()
        month_start = dt.replace(day=1)
        _, last_day = calendar.monthrange(dt.year, dt.month)
        month_end = dt.replace(day=last_day)
        week_start = dt - datetime.timedelta(days=dt.weekday())
        if week_start < month_start:
            week_start = month_start
        week_end = week_start + datetime.timedelta(days=4)
        if week_end > month_end:
            week_end = month_end
        return f"{week_start.strftime('%d/%m/%Y')} - {week_end.strftime('%d/%m/%Y')}"

    if week_enabled:
        date_series = pd.to_datetime(working_df['Date'], errors='coerce')
        working_df[week_label_key] = date_series.apply(_format_week_label)
    else:
        working_df[week_label_key] = pd.NA

    group_cols: List[str] = []
    if week_enabled:
        group_cols.append(week_label_key)
    group_cols.append('Customer')
    if include_date:
        group_cols.append('Date')
    if has_outlet:
        group_cols.append('Outlet')

    agg_map = {
        'Qty in Pcs': 'sum',
        'Qty in Ctns': 'sum',
        'Total Value': 'sum',
        'carton_packing_numeric': 'first',
    }
    summary = (
        working_df
        .groupby(group_cols, dropna=False, as_index=False)
        .agg(agg_map)
    )

    summary['Qty in Pcs'] = summary['Qty in Pcs'].fillna(0)
    summary['Qty in Ctns'] = summary['Qty in Ctns'].fillna(0)
    summary['Total Value'] = summary['Total Value'].fillna(0)

    summary['Total Qty'] = summary.apply(
        lambda row: _format_summary_total_qty(
            row.get('Qty in Ctns'),
            row.get('Qty in Pcs'),
            row.get('carton_packing_numeric'),
        ),
        axis=1,
    )

    if not has_outlet:
        pack_size = pd.to_numeric(
            summary['carton_packing_numeric'], errors='coerce')
        pack_size = pack_size.fillna(0)
        total_packets = summary['Qty in Ctns'] * \
            pack_size + summary['Qty in Pcs']
        summary['Selling Price (per pkt)'] = _safe_divide_series(
            summary['Total Value'], total_packets)

    if include_date and 'Date' in summary.columns:
        date_sort = pd.to_datetime(summary['Date'], errors='coerce')
        summary['_date_sort'] = date_sort
        summary['Date'] = date_sort.dt.strftime('%d/%m/%Y').fillna('')

    sort_keys = []
    if week_enabled:
        sort_keys.append(week_label_key)
    sort_keys.append('Customer')
    if include_date:
        sort_keys.append('_date_sort')
    if has_outlet:
        sort_keys.append('Outlet')
    summary = summary.sort_values(sort_keys, kind='stable')

    if week_label_key in summary.columns:
        summary['Week'] = summary[week_label_key]
    summary = summary.drop(
        columns=[col for col in sort_keys if col.startswith('_')], errors='ignore')
    summary = summary.reset_index(drop=True)
    summary = summary.drop(
        columns=['carton_packing_numeric', week_label_key], errors='ignore')
    return summary


def _customer_purchase_breakdown(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame()
    group_cols = [
        col for col in [
            "Customer",
            "Month",
            "Brand/Category",
            "Supplier",
            "Product Description",
            "Product Code",
        ]
        if col in df.columns
    ]
    agg_map = {
        col: "sum"
        for col in ["Qty in Ctns", "Qty in Pcs", "Total Value"]
        if col in df.columns
    }
    if "carton_packing_numeric" in df.columns:
        agg_map["carton_packing_numeric"] = "first"
    if not group_cols or not agg_map:
        return pd.DataFrame()
    summary = (
        df.groupby(group_cols, dropna=False, sort=False)
        .agg(agg_map)
        .reset_index()
    )
    for qty_col in ["Qty in Ctns", "Qty in Pcs"]:
        if qty_col in summary:
            summary[qty_col] = summary[qty_col].fillna(0)
    if "Total Value" in summary:
        summary["Total Value"] = summary["Total Value"].fillna(0)

    def _row_total_qty(row):
        return _format_summary_total_qty(
            row.get("Qty in Ctns"),
            row.get("Qty in Pcs"),
            row.get("carton_packing_numeric"),
        )

    summary["Total Qty"] = summary.apply(_row_total_qty, axis=1)

    pack_size_series = pd.to_numeric(
        summary.get(
            "carton_packing_numeric",
            pd.Series(0.0, index=summary.index),
        ),
        errors="coerce",
    ).fillna(0.0)
    total_ctns = pd.to_numeric(
        summary.get("Qty in Ctns", pd.Series(0.0, index=summary.index)),
        errors="coerce",
    ).fillna(0.0)
    total_pkts = pd.to_numeric(
        summary.get("Qty in Pcs", pd.Series(0.0, index=summary.index)),
        errors="coerce",
    ).fillna(0.0)
    total_packets = total_ctns * pack_size_series + total_pkts
    if "Total Value" in summary:
        summary["Selling Price (per pkt)"] = _safe_divide_series(
            summary["Total Value"],
            total_packets,
        )
    if "Selling Price (per pkt)" in summary:
        summary["Selling Price (per pkt)"] = summary[
            "Selling Price (per pkt)"].apply(_format_price_display)

    summary_cols = group_cols + ["Total Qty"]
    if "Selling Price (per pkt)" in summary:
        summary_cols.append("Selling Price (per pkt)")
    return summary[summary_cols]


def build_sales_usage_views(
    df: pd.DataFrame,
) -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    monthly_cols = ["Month", "Qty in Ctns", "Qty in Pcs", "Total Qty", "Total Value"]
    customer_cols = ["Customer", "Qty in Ctns", "Qty in Pcs", "Total Qty", "Total Value"]
    matrix_default = pd.DataFrame(columns=["Customer", "Overall"])

    if df.empty:
        return (
            pd.DataFrame(columns=monthly_cols),
            pd.DataFrame(columns=customer_cols),
            matrix_default,
        )

    working = df.copy()
    for numeric_col in ["Qty in Ctns", "Qty in Pcs", "Total Value"]:
        if numeric_col not in working.columns:
            working[numeric_col] = 0.0
        working[numeric_col] = pd.to_numeric(
            working[numeric_col], errors="coerce").fillna(0.0)

    if "Customer" in working.columns:
        customer_series = working["Customer"].astype(
            "string").fillna("").str.strip()
    else:
        customer_series = pd.Series([""] * len(working), index=working.index, dtype="string")
    working["__customer"] = customer_series.replace("", pd.NA).fillna("Unknown")

    if "Year" in working.columns:
        year_series = working["Year"].astype("string").fillna("").str.strip()
    else:
        year_series = pd.Series([""] * len(working), index=working.index, dtype="string")
    if "Month" in working.columns:
        month_series = working["Month"].astype("string").fillna("").str.strip()
    else:
        month_series = pd.Series([""] * len(working), index=working.index, dtype="string")

    month_label = month_series.apply(_format_month_label).replace("", pd.NA).fillna("Unknown")
    year_clean = year_series.replace("", pd.NA)
    working["__month_label"] = month_label.where(
        year_clean.isna(),
        year_clean.fillna("") + "-" + month_label,
    )

    month_sort = month_series.map(_month_sort_key)
    month_sort = pd.to_numeric(month_sort, errors="coerce").fillna(13)
    month_sort = month_sort.where(month_sort != 999, 13)
    year_sort = pd.to_numeric(year_clean, errors="coerce").fillna(9999)
    working["__month_order"] = year_sort * 100 + month_sort

    monthly_usage = (
        working
        .groupby(["__month_label", "__month_order"], as_index=False, dropna=False)[
            ["Qty in Ctns", "Qty in Pcs", "Total Value"]
        ]
        .sum()
        .sort_values(["__month_order", "__month_label"], kind="stable")
    )
    monthly_usage["Total Qty"] = monthly_usage.apply(
        lambda row: _format_quantity_pair(row.get("Qty in Ctns"), row.get("Qty in Pcs")),
        axis=1,
    )
    monthly_usage = monthly_usage.rename(columns={"__month_label": "Month"})
    monthly_usage = monthly_usage[monthly_cols].reset_index(drop=True)

    customer_usage = (
        working
        .groupby("__customer", as_index=False, dropna=False)[
            ["Qty in Ctns", "Qty in Pcs", "Total Value"]
        ]
        .sum()
        .rename(columns={"__customer": "Customer"})
        .sort_values(
            ["Qty in Ctns", "Qty in Pcs", "Total Value", "Customer"],
            ascending=[False, False, False, True],
            kind="stable",
        )
        .reset_index(drop=True)
    )
    customer_usage["Total Qty"] = customer_usage.apply(
        lambda row: _format_quantity_pair(row.get("Qty in Ctns"), row.get("Qty in Pcs")),
        axis=1,
    )
    customer_usage = customer_usage[customer_cols]

    customer_month = (
        working
        .groupby(["__customer", "__month_label", "__month_order"], as_index=False, dropna=False)[
            ["Qty in Ctns", "Qty in Pcs"]
        ]
        .sum()
    )
    customer_month["Total Qty"] = customer_month.apply(
        lambda row: _format_quantity_pair(row.get("Qty in Ctns"), row.get("Qty in Pcs")),
        axis=1,
    )
    if customer_month.empty:
        customer_month_matrix = matrix_default
    else:
        ordered_months = (
            customer_month[["__month_label", "__month_order"]]
            .drop_duplicates()
            .sort_values(["__month_order", "__month_label"], kind="stable")
        )
        customer_month_matrix = (
            customer_month.pivot(
                index="__customer",
                columns="__month_label",
                values="Total Qty",
            )
            .fillna("-")
            .reindex(columns=ordered_months["__month_label"].tolist())
            .reset_index()
            .rename(columns={"__customer": "Customer"})
        )
        overall_map = customer_usage.set_index("Customer")["Total Qty"]
        customer_month_matrix.insert(
            1,
            "Overall",
            customer_month_matrix["Customer"].map(overall_map).fillna("0"),
        )

    return monthly_usage, customer_usage, customer_month_matrix


# ----------------------------- Sales 页 -----------------------------

def run_sales_page():
    st.title("Sales Dashboard")
    st.caption(
        "Upload one or more workbooks containing a 'Delivery details' sheet (headers on row 4, columns A~U).")

    uploaded = st.file_uploader(
        "Upload Sales Excel (.xlsx)",
        type=["xlsx"],
        accept_multiple_files=True,
        key="sales_uploader",
    )

    # 朴实无华缓存：有新上传就覆盖缓存，没上传就用已有缓存
    if uploaded:
        st.session_state["sales_files"] = uploaded

    sales_files = st.session_state.get("sales_files", [])

    if not sales_files:
        st.info("Upload at least one Delivery details workbook to continue.")
        return

    try:
        sales_df, warns = load_sales_data(sales_files)
    except Exception as exc:
        st.error(f"Failed to read Sales workbooks: {exc}")
        return

    for warn in warns:
        st.warning(warn)

    if sales_df.empty:
        st.warning("无法从上传的工作簿读取任何 Sales 数据。")
        return

    filtered_df, filter_state = apply_sales_filters(sales_df)
    filter_order = [
        "Year", "Month", "Customer", "Exclude Customer", "Outlet", "Product Description",
        "Supplier", "Brand/Category", "Product Code", "Account", "Invoice contains",
        "Date from", "Date to",
    ]
    filter_summary_items = [
        f"{key} = {filter_state.get(key)}"
        for key in filter_order
        if filter_state.get(key) and filter_state.get(key) != "All"
    ]
    st.caption(
        "当前筛选：" + "；".join(filter_summary_items)
        if filter_summary_items else "当前筛选：全部数据"
    )

    selected_customers = st.session_state.get("sales_filter_customer") or []
    show_customer_brought = False
    if selected_customers:
        show_customer_brought = st.checkbox(
            "What they brought",
            key="sales_show_customer_purchases",
            help="Show the aggregated products and quantities for the filtered customer(s).",
        )
    else:
        st.session_state.pop("sales_show_customer_purchases", None)
    if show_customer_brought:
        customer_breakdown = _customer_purchase_breakdown(filtered_df)
        if customer_breakdown.empty:
            st.info("No purchase records remain after the current filters.")
        else:
            st.caption("What they brought")
            st.dataframe(
                customer_breakdown,
                width="stretch",
            )

    view_cols = st.columns([1, 1, 1])
    group_by_outlet = view_cols[0].checkbox(
        "Group by Outlet",
        key="sales_group_by_outlet",
    )
    include_date = view_cols[1].checkbox(
        "Include Date per Outlet",
        key="sales_include_date",
        disabled=not group_by_outlet,
    )
    include_date = include_date and group_by_outlet
    show_by_weeks = view_cols[2].checkbox(
        "Show by weeks",
        key="sales_show_by_weeks",
        disabled="Date" not in filtered_df.columns,
    )

    summary_df = build_sales_summary(
        filtered_df,
        group_by_outlet=group_by_outlet,
        include_date=include_date,
        group_by_week=show_by_weeks,
    )

    summary_display = summary_df.copy()
    if "Total Value" in summary_display.columns:
        summary_display["Total Value"] = summary_display["Total Value"].apply(
            _format_price_display
        )
    if "Selling Price (per pkt)" in summary_display.columns:
        summary_display["Selling Price (per pkt)"] = summary_display[
            "Selling Price (per pkt)"].apply(_format_price_display)

    def _int_total(series: Optional[pd.Series]) -> int:
        if series is None:
            return 0
        total = series.sum(min_count=1)
        if pd.isna(total):
            return 0
        try:
            return int(round(float(total)))
        except (TypeError, ValueError):
            return 0

    def _total_qty_text_with_pack(df: pd.DataFrame) -> str:
        if df is None or df.empty:
            return "0"

        total_ctns = _int_total(df.get("Qty in Ctns"))
        total_pkts = _int_total(df.get("Qty in Pcs"))

        pack_col = pd.to_numeric(
            df.get("carton_packing_numeric"), errors="coerce")
        valid = pack_col[pack_col > 0]
        unique_packs = sorted(valid.unique())

        if len(unique_packs) == 1:
            pack_size = int(round(unique_packs[0]))
            total_packets = total_ctns * pack_size + total_pkts
            return _format_total_qty_text(total_packets, pack_size)

        return f"{total_ctns} ctns {total_pkts} pkts"

    total_qty_text = _total_qty_text_with_pack(filtered_df)

    single_sku_mode = False
    single_sku_pack_size: Optional[int] = None
    if "Product Code" in filtered_df.columns:
        non_empty_codes = (
            filtered_df["Product Code"]
            .astype("string")
            .str.strip()
            .replace("", pd.NA)
            .dropna()
        )
        if not non_empty_codes.empty and non_empty_codes.nunique() == 1:
            pack_col = pd.to_numeric(
                filtered_df.get("carton_packing_numeric"), errors="coerce"
            )
            valid_packs = pack_col[pack_col > 0].unique()
            if len(valid_packs) == 1:
                single_sku_mode = True
                single_sku_pack_size = int(round(valid_packs[0]))

    selected_months = st.session_state.get("sales_filter_month") or []
    month_totals_df: Optional[pd.DataFrame] = None
    if "Month" in filtered_df.columns and not filtered_df.empty:
        month_totals_source = filtered_df.copy()
        for qty_col in ["Qty in Ctns", "Qty in Pcs"]:
            if qty_col not in month_totals_source.columns:
                month_totals_source[qty_col] = 0.0
        month_totals = (
            month_totals_source
            .groupby("Month", sort=False, dropna=False)[["Qty in Ctns", "Qty in Pcs"]]
            .sum()
            .reset_index()
        )
        month_totals["Qty in Ctns"] = month_totals["Qty in Ctns"].fillna(0)
        month_totals["Qty in Pcs"] = month_totals["Qty in Pcs"].fillna(0)

        if single_sku_mode and single_sku_pack_size:
            def _month_total_qty(row):
                try:
                    c = int(round(float(row["Qty in Ctns"])))
                except (TypeError, ValueError):
                    c = 0
                try:
                    p = int(round(float(row["Qty in Pcs"])))
                except (TypeError, ValueError):
                    p = 0
                total_packets = c * single_sku_pack_size + p
                return _format_total_qty_text(total_packets, single_sku_pack_size)

            month_totals["Total Qty"] = month_totals.apply(
                _month_total_qty, axis=1)
        else:
            month_totals["Total Qty"] = month_totals.apply(
                lambda row: _format_quantity_pair(
                    row["Qty in Ctns"], row["Qty in Pcs"]),
                axis=1,
            )

        month_totals_df = month_totals[["Month", "Total Qty"]].copy()
        if selected_months:
            order_map = {value: idx for idx,
                         value in enumerate(selected_months)}
            month_totals_df["__order"] = month_totals_df["Month"].map(
                lambda value: order_map.get(value, len(order_map)))
            month_totals_df = (
                month_totals_df.sort_values("__order").drop(
                    columns="__order").reset_index(drop=True)
            )

    monthly_usage_df, customer_usage_df, customer_month_matrix_df = build_sales_usage_views(
        filtered_df
    )
    monthly_usage_display = monthly_usage_df.copy()
    customer_usage_display = customer_usage_df.copy()
    for table in [monthly_usage_display, customer_usage_display]:
        for qty_col in ["Qty in Ctns", "Qty in Pcs"]:
            if qty_col in table.columns:
                table[qty_col] = table[qty_col].apply(_format_qty_display)
        if "Total Value" in table.columns:
            table["Total Value"] = table["Total Value"].apply(_format_price_display)

    display_cols = []
    if show_by_weeks and "Week" in summary_display.columns:
        display_cols.append("Week")
    display_cols.append("Customer")
    if include_date and "Date" in summary_display.columns:
        display_cols.append("Date")
    if group_by_outlet and "Outlet" in summary_display.columns:
        display_cols.append("Outlet")
    display_cols.append("Total Qty")
    if not group_by_outlet and "Selling Price (per pkt)" in summary_display.columns:
        display_cols.append("Selling Price (per pkt)")
    if "Total Value" in summary_display.columns:
        display_cols.append("Total Value")
    display_cols = list(dict.fromkeys(display_cols))

    detail_display = filtered_df.copy()
    if "Date" in detail_display.columns:
        if pd.api.types.is_datetime64_any_dtype(detail_display["Date"]):
            detail_display["Date"] = detail_display["Date"].dt.strftime(
                "%d/%m/%Y").fillna("")
        else:
            detail_display["Date"] = detail_display["Date"].astype(
                "string").str.strip()
    for qty_col in ["Qty in Pcs", "Qty in Ctns"]:
        if qty_col in detail_display.columns:
            detail_display[qty_col] = detail_display[qty_col].apply(
                _format_qty_display)
    for price_col in ["Total Value", "Unit Price (per pcs)", "Unit Price (per carton)"]:
        if price_col in detail_display.columns:
            detail_display[price_col] = detail_display[price_col].apply(
                _format_price_display)
    if "Month" in detail_display.columns:
        detail_display["Month"] = detail_display["Month"].apply(
            _format_month_label)

    summary_tab, raw_tab = st.tabs(["Summary", "Raw Records"])
    with summary_tab:
        view_label = "Customer"
        if show_by_weeks:
            view_label = "Week + " + view_label
        if group_by_outlet:
            view_label += " + Outlet"
        if include_date:
            view_label += " + Date"
        st.metric("TOTAL", total_qty_text)
        if month_totals_df is not None and len(selected_months) > 1 and not month_totals_df.empty:
            st.caption("Monthly totals for selected months")
            if single_sku_mode and single_sku_pack_size:
                st.caption(
                    f"(Single SKU mode, pack size = {single_sku_pack_size} per carton)")
            else:
                st.caption("Total Qty = ctns + pkts across mixed SKUs")
            st.table(month_totals_df)
        st.caption(
            f"View: {view_label}. Total Qty shows cartons + packets, selling price = Total Value / pack size.")
        st.caption(f"Total Qty (filtered): {total_qty_text}")
        if summary_display.empty:
            st.info("当前筛选未产出任何汇总")
        else:
            st.dataframe(
                summary_display[display_cols],
                width="stretch",
            )
        usage_tab_month, usage_tab_customer, usage_tab_matrix = st.tabs(
            ["Usage by Month", "Usage by Customer", "Customer x Month"]
        )
        with usage_tab_month:
            if monthly_usage_display.empty:
                st.info("当前筛选未产出月度用量数据。")
            else:
                st.caption("每个月的总用量（ctn/pkt）和销售额。")
                st.dataframe(monthly_usage_display, width="stretch")
                month_chart = monthly_usage_df[["Month", "Qty in Ctns", "Qty in Pcs"]].copy()
                month_chart = month_chart.set_index("Month")
                if not month_chart.empty:
                    st.bar_chart(month_chart)
        with usage_tab_customer:
            if customer_usage_display.empty:
                st.info("当前筛选未产出客户用量数据。")
            else:
                st.caption("每个客户的总用量（ctn/pkt）和销售额。")
                st.dataframe(customer_usage_display, width="stretch")
                customer_chart = customer_usage_df[[
                    "Customer", "Qty in Ctns", "Qty in Pcs"]].copy()
                customer_chart = customer_chart.head(15).set_index("Customer")
                if not customer_chart.empty:
                    st.bar_chart(customer_chart)
        with usage_tab_matrix:
            if customer_month_matrix_df.empty:
                st.info("当前筛选未产出客户-月份矩阵。")
            else:
                st.caption("每个客户在每个月的用量（Total Qty）。")
                st.dataframe(customer_month_matrix_df, width="stretch")
    with raw_tab:
        st.caption(
            f"Detailed records ({len(detail_display):,} rows) —— 当前筛选共享 Summary。")
        if detail_display.empty:
            st.info("当前没有满足筛选的明细。")
        else:
            try:
                styled_detail = detail_display.style.map(
                    _highlight_missing_cell)
                st.dataframe(
                    styled_detail,
                    width="stretch",
                )
            except Exception:
                st.dataframe(
                    detail_display,
                    width="stretch",
                )
        download_cols = st.columns([1, 1])
        detail_export = _strip_html_df(filtered_df)
        csv_buffer = io.StringIO()
        detail_export.to_csv(csv_buffer, index=False)
        download_cols[0].download_button(
            "Export Detail CSV",
            data=csv_buffer.getvalue(),
            file_name="sales_detail.csv",
            mime="text/csv",
        )
        excel_buffer = io.BytesIO()
        with pd.ExcelWriter(excel_buffer, engine="openpyxl") as writer:
            detail_export.to_excel(
                writer, sheet_name="Sales Detail", index=False)
            summary_export = _strip_html_df(summary_df)
            summary_export.to_excel(
                writer, sheet_name="Customer-Month Summary", index=False)
            monthly_usage_export = _strip_html_df(monthly_usage_display)
            monthly_usage_export.to_excel(
                writer, sheet_name="Usage by Month", index=False)
            customer_usage_export = _strip_html_df(customer_usage_display)
            customer_usage_export.to_excel(
                writer, sheet_name="Usage by Customer", index=False)
            customer_month_export = _strip_html_df(customer_month_matrix_df)
            customer_month_export.to_excel(
                writer, sheet_name="Customer x Month", index=False)
            meta_rows = [
                {"key": k, "value": str(v)} for k, v in filter_state.items()
            ]
            pd.DataFrame(meta_rows, columns=["key", "value"]).to_excel(
                writer, sheet_name="Filters", index=False)
        excel_buffer.seek(0)
        download_cols[1].download_button(
            "Export Excel",
            data=excel_buffer.getvalue(),
            file_name="sales_summary.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )


# ----------------------------- 主入口：使用 tabs 让两个界面独立运行 -----------------------------

def main():
    st.sidebar.title("导航")
    st.sidebar.write("选择上方标签切换 Sales 与 Stock 界面。")

    # 使用 tabs 同时渲染两个界面，互相不重置状态
    sales_tab, stock_tab = st.tabs(["Sales", "Stock"])

    with sales_tab:
        run_sales_page()

    with stock_tab:
        run_stock_page()

    st.sidebar.markdown("---")
    st.sidebar.caption("Powered by Streamlit")


if __name__ == "__main__":
    main()
