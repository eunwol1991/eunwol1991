import re
import os
import sys
import io

# 如果直接 python stock_datagrid.py 启动，则切换到 streamlit 运行（兼容 Windows 路径空格）
try:
    from streamlit.runtime.scriptrunner import get_script_run_ctx  # type: ignore
except Exception:  # Fallback when Streamlit internal API is unavailable
    def get_script_run_ctx():
        return None

if __name__ == "__main__" and os.environ.get("ST_REDIRECTED", "0") != "1" and (get_script_run_ctx() is None):
    import subprocess
    from pathlib import Path

    def _short_path(p: str) -> str:
        if os.name != "nt" or " " not in p:
            return p
        try:
            import ctypes
            from ctypes import wintypes
            GetShortPathNameW = ctypes.windll.kernel32.GetShortPathNameW
            GetShortPathNameW.argtypes = [wintypes.LPCWSTR, wintypes.LPWSTR, wintypes.DWORD]
            GetShortPathNameW.restype = wintypes.DWORD
            buf = ctypes.create_unicode_buffer(260)
            res = GetShortPathNameW(p, buf, 260)
            return buf.value if res else p
        except Exception:
            return p

    os.environ["ST_REDIRECTED"] = "1"
    script_path = str(Path(__file__).resolve())
    script_path = _short_path(script_path)
    cmd = [sys.executable, "-m", "streamlit", "run", script_path]
    if len(sys.argv) > 1:
        cmd += ["--"] + sys.argv[1:]
    if os.environ.get("ST_DEBUG_REDIRECT") == "1":
        print("[streamlit-redirect] ", cmd)
    subprocess.run(cmd, check=False)
    sys.exit(0)

import streamlit as st
import pandas as pd
import math
import datetime
from typing import Optional, Tuple, List, Dict

PRIMARY_WAREHOUSES = ["Savori Whse", "Lai Hock Whse"]


# ----------------------------- 基础工具函数 -----------------------------

def _normalize_warehouse_name(value) -> str:
    if pd.isna(value):
        return ""
    return str(value).strip()


def _canonical_unit(u: Optional[str]) -> str:
    """把单位同义词规整为统一小写键。"""
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
    """将数字格式化为紧凑字符串，不带无意义 0。"""
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
    """格式：'3 ctns 2 pkts'；为 0 的单位省略。"""
    parts: List[str] = []
    if ctn is not None and not pd.isna(ctn):
        ctn_val = float(ctn)
        if not math.isclose(ctn_val, 0.0, abs_tol=1e-9):
            qty_text = _format_qty_number(ctn_val)
            if qty_text:
                unit = "ctn" if math.isclose(ctn_val, 1.0, abs_tol=1e-9) else "ctns"
                parts.append(f"{qty_text} {unit}")
    if pkt is not None and not pd.isna(pkt):
        pkt_val = float(pkt)
        if not math.isclose(pkt_val, 0.0, abs_tol=1e-9):
            qty_text = _format_qty_number(pkt_val)
            if qty_text:
                unit = "pkt" if math.isclose(pkt_val, 1.0, abs_tol=1e-9) else "pkts"
                parts.append(f"{qty_text} {unit}")
    return " ".join(parts) if parts else "0"


def _strip_html_df(df: pd.DataFrame) -> pd.DataFrame:
    """
    移除 DataFrame 所有单元格里的 HTML 标签，确保显示与导出内容为纯文本。
    仅对 dtype==object 的列做清洗；不会影响数值/日期列。
    """
    clean = df.copy()
    for col in clean.columns:
        if clean[col].dtype == object:
            clean[col] = clean[col].astype(str).apply(lambda x: re.sub(r"<.*?>", "", x))
    return clean


def build_norm_desc(df: pd.DataFrame) -> pd.Series:
    """沿用既有 “去括号后” 规范化口径。"""
    if df is None or df.empty:
        return pd.Series(dtype="string", name="norm_desc")
    desc_series = df.get("description", pd.Series(dtype="string")).astype("string").fillna("")
    base = desc_series.str.replace(r"\s*\([^)]*\)", "", regex=True).str.strip()
    code_series = df.get("product_code", pd.Series(dtype="string")).astype("string").fillna("").str.strip()
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
    norm = norm.fillna("").str.strip()
    norm = norm.replace("", "Unnamed product")
    norm.name = "norm_desc"
    return norm


def _extract_reorder_points(df_group: pd.DataFrame) -> Dict[str, Optional[float]]:
    """从数据中抽取每商品的 ROP（如有），返回 {'ctn': float|None, 'pkt': float|None}。"""
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

# ----------------------------- 业务聚合 -----------------------------
def aggregate_summary(
    df: pd.DataFrame,
    warehouses: Optional[List[str]] = None,
) -> Tuple[pd.DataFrame, Dict[Tuple[str, str], pd.DataFrame]]:
    warehouses = warehouses or PRIMARY_WAREHOUSES
    columns = [
        "Supplier", "Product", "Pack Size", "Product Code",   # ← 新增 "Pack Size"
        "savori_ctn", "savori_pkt", "lai_hock_ctn", "lai_hock_pkt",
        "total_ctn", "total_pkt",
        "reorder_point_ctn", "reorder_point_pkt",
        "group_key",
    ]
    if df is None or df.empty:
        return pd.DataFrame(columns=columns), {}

    work = df.copy()
    if "warehouse" in work.columns:
        work["_warehouse_norm"] = work["warehouse"].map(_normalize_warehouse_name)
        work = work[work["_warehouse_norm"].isin(warehouses)]
    else:
        work["_warehouse_norm"] = ""
    if work.empty:
        return pd.DataFrame(columns=columns), {}

    work["norm_desc"] = build_norm_desc(work)
    work["_unit_norm"] = work.get("unit", pd.Series(dtype="string")).apply(_canonical_unit)
    work["_qty"] = pd.to_numeric(work.get("stock_qty", pd.Series(dtype=float)), errors="coerce").fillna(0.0)
    work["_product_code_norm"] = work.get("product_code", pd.Series(dtype="string")).astype("string").fillna("").str.strip()
    if "supplier" not in work.columns:
        work["supplier"] = pd.NA

    rows = []
    detail_map: Dict[Tuple[str, str], pd.DataFrame] = {}

    for key, grp in work.groupby(["supplier", "norm_desc"], dropna=False):
        if not isinstance(key, tuple):
            key = (key,)
        supplier_val = key[0]
        norm_desc_val = key[1] if len(key) > 1 else ""
        supplier_name = str(supplier_val).strip() if pd.notna(supplier_val) and str(supplier_val).strip() else "Unknown supplier"
        product_label = str(norm_desc_val).strip() if pd.notna(norm_desc_val) and str(norm_desc_val).strip() else "Unnamed product"
        group_key = (supplier_name, product_label)
        detail_map[group_key] = grp.copy()

        # 组合 Product Code
        codes = sorted({c for c in grp["_product_code_norm"].tolist() if c})
        product_code_label = " / ".join(codes) if codes else "-"

        # 组合 Pack Size（同组可能有多个，尽量并排展示）
        pack_sizes = grp.get("pack_size")
        if pack_sizes is not None:
            uniq_ps = sorted({str(x).strip() for x in pack_sizes.tolist() if str(x).strip() and str(x).strip().lower() != "nan"})
            pack_size_label = " / ".join(uniq_ps) if uniq_ps else "-"
        else:
            pack_size_label = "-"

        # 两仓分量
        wh_totals = {wh: {"ctn": 0.0, "pkt": 0.0} for wh in warehouses}
        wh_unit = grp.groupby(["_warehouse_norm", "_unit_norm"], dropna=False)["_qty"].sum().reset_index()
        for _, r in wh_unit.iterrows():
            wh = r["_warehouse_norm"]; unit = r["_unit_norm"]; qty = float(r["_qty"])
            if wh in wh_totals and unit in ("ctn", "pkt"):
                wh_totals[wh][unit] += qty

        total_ctn = sum(wh_totals.get(wh, {}).get("ctn", 0.0) for wh in warehouses)
        total_pkt = sum(wh_totals.get(wh, {}).get("pkt", 0.0) for wh in warehouses)

        reorder_points = _extract_reorder_points(grp)

        rows.append({
            "Supplier": supplier_name,
            "Product": product_label,
            "Pack Size": pack_size_label,                   # ← 写入
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
    """
    批次层唯一判定来源：
    - 数量>0 且 days<0 => Expired
    - 数量>0 且 0<=days<=expiry_days => Near-Expiry
    - 数量==0 => Depleted
    - 否则 OK
    返回: (status_batch:str, days_to_expiry:int|None)
    """
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
    """产品层状态优先级：Expired > Near-Expiry > Low-Stock > OK"""
    if has_expired:
        return "Expired"
    if has_near:
        return "Near-Expiry"
    if is_low_stock:
        return "Low-Stock"
    return "OK"


# ----------------------------- 批次层明细 -----------------------------

def split_by_expiry(
    df_row_scope: pd.DataFrame,
    warehouses: Optional[List[str]] = None,
    *,
    expiry_days: int = 30,
    show_depleted: bool = True,
) -> pd.DataFrame:
    """
    返回某产品的批次分布，并给出批次层判定
    """
    warehouses = warehouses or PRIMARY_WAREHOUSES
    columns = [
        "Expiry", "Savori Whse", "Lai Hock Whse", "Subtotal",
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

    work["_unit_norm"] = work.get("unit", pd.Series(dtype="string")).apply(_canonical_unit)
    work["_qty"] = pd.to_numeric(work.get("stock_qty", pd.Series(dtype=float)), errors="coerce").fillna(0.0)
    work["_expiry_norm"] = pd.to_datetime(work.get("expiry_date"), errors="coerce", format="mixed")
    work["_expiry_label"] = work["_expiry_norm"].apply(lambda x: x.date().isoformat() if pd.notna(x) else "No Expiry")
    work["_sort_key"] = work["_expiry_norm"].apply(lambda x: x.date() if pd.notna(x) else datetime.date.max)

    rows = []
    for label, grp in work.groupby("_expiry_label", dropna=False):
        sort_key = grp["_sort_key"].min()
        expiry_val = grp["_expiry_norm"].dropna().min()
        per_wh = {wh: {"ctn": 0.0, "pkt": 0.0} for wh in warehouses}
        wh_unit = grp.groupby(["_warehouse_norm", "_unit_norm"], dropna=False)["_qty"].sum().reset_index()
        for _, r in wh_unit.iterrows():
            wh = r["_warehouse_norm"]
            unit = r["_unit_norm"]
            qty = float(r["_qty"])
            if wh in per_wh and unit in ("ctn", "pkt"):
                per_wh[wh][unit] += qty

        subtotal_ctn = sum(per_wh.get(wh, {}).get("ctn", 0.0) for wh in warehouses)
        subtotal_pkt = sum(per_wh.get(wh, {}).get("pkt", 0.0) for wh in warehouses)

        exp_date = expiry_val.date() if pd.notna(expiry_val) else None
        status_batch, d2e = classify_batch_status(exp_date, subtotal_ctn, subtotal_pkt, expiry_days)

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
            "Expiry": label,
            "Savori Whse": _format_quantity_pair(per_wh.get('Savori Whse',{}).get('ctn',0.0), per_wh.get('Savori Whse',{}).get('pkt',0.0)),
            "Lai Hock Whse": _format_quantity_pair(per_wh.get('Lai Hock Whse',{}).get('ctn',0.0), per_wh.get('Lai Hock Whse',{}).get('pkt',0.0)),
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

    result = pd.DataFrame(rows).sort_values(by="_sort_key", kind="stable").drop(columns="_sort_key").reset_index(drop=True)
    return result


# ----------------------------- 侧边栏筛选 -----------------------------

def apply_filters_v2(df: pd.DataFrame):
    """
    Excel 式多维筛选，维度：Warehouse → Supplier → Brand → Description(去括号) → Product Code → Remark(括号内)。
    返回：筛选后的 DataFrame、所选 Description 列表、(占位)到期高亮天数、当前筛选状态。
    """
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

        # ---------- include-only 小工具 ----------
        def _include(series: pd.Series, selected: list):
            if not selected:
                return pd.Series(True, index=series.index)
            return series.isin(selected)

        # ---------- Warehouse（include-only） ----------
        if exclude != "warehouse" and "warehouse" in d.columns:
            sel = list(ss.get("f_wh", []))
            if sel:
                d = d[_include(d["warehouse"], sel)]

        # ---------- Supplier（支持 exclude） ----------
        if exclude != "supplier" and "supplier" in d.columns:
            sel = list(ss.get("f_sup", []))
            exm = bool(ss.get("f_sup_ex", False))
            if sel:
                m = d["supplier"].isin(sel)
                d = d[~m] if exm else d[m]

        # ---------- Brand（include-only） ----------
        if exclude != "brand" and "brand" in d.columns:
            sel = list(ss.get("f_brand", []))
            if sel:
                d = d[_include(d["brand"], sel)]

        # ---------- Description（include-only；用去括号基准） ----------
        if exclude != "desc" and "description" in d.columns:
            base_ser = get_desc_base(d["description"])
            sel = list(ss.get("f_desc", []))
            if sel:
                m = base_ser.isin(sel)
                d = d[m]

        # ---------- Product Code（include-only） ----------
        if exclude != "code" and "product_code" in d.columns:
            sel = list(ss.get("f_code", []))
            if sel:
                d = d[_include(d["product_code"], sel)]

        # ---------- Remark（支持 exclude） ----------
        if exclude != "remark" and "description" in d.columns:
            sel = list(ss.get("f_remark", []))
            exm = bool(ss.get("f_remark_ex", False))
            if sel:
                matches = d["description"].astype(str).str.findall(r"\(([^)]*)\)")
                has_any = matches.apply(lambda lst: any((str(x).strip() in sel) for x in (lst or [])))
                d = d[~has_any] if exm else d[has_any]

        return d



    with st.sidebar:
        st.header("筛选条件")

        # Warehouse
        wh_options = []
        if "warehouse" in base.columns:
            d = apply_all(base, exclude="warehouse")
            wh_options = [x for x in d["warehouse"].dropna().astype(str).unique().tolist()]
            ordered = [w for w in ["Savori Whse", "Lai Hock Whse"] if w in wh_options]
            ordered += [w for w in wh_options if w not in ordered]
            default_selection = [w for w in ["Savori Whse", "Lai Hock Whse"] if w in ordered] or list(ordered)
            _ensure_multiselect_key("f_wh", ordered, default_selection)
            st.multiselect("仓库", options=ordered, key="f_wh", placeholder="选择一个或多个仓库")
            sel_wh = list(ss.get("f_wh", []))

        # Supplier
        if "supplier" in base.columns:
            d = apply_all(base, exclude="supplier")
            sup_options = sorted([x for x in d["supplier"].dropna().unique().tolist()])
            _ensure_multiselect_key("f_sup", sup_options, [])
            st.multiselect("Supplier", sup_options, key="f_sup", placeholder="选择供应商")
            # ↓ 新增：Supplier 的排除模式开关
            st.checkbox("Exclude selected (Supplier)", key="f_sup_ex", value=False)
            sel_sup = list(ss.get("f_sup", []))

        # Brand
        if "brand" in base.columns:
            d = apply_all(base, exclude="brand")
            brand_options = sorted([x for x in d["brand"].dropna().unique().tolist()])
            _ensure_multiselect_key("f_brand", brand_options, [])
            st.multiselect("Brand", brand_options, key="f_brand", placeholder="选择品牌")
            sel_brand = list(ss.get("f_brand", []))

        # Description（去括号后）
        if "description" in base.columns:
            d = apply_all(base, exclude="desc")
            base_ser = get_desc_base(d["description"]) if not d.empty else pd.Series(dtype=str)
            if "f_desc_q" not in ss:
                ss["f_desc_q"] = ""
            q = st.text_input("Description 关键字", key="f_desc_q").strip()
            ser = base_ser.dropna()
            if q:
                ser = ser[ser.str.contains(q, case=False, na=False)]
            desc_options = sorted([x for x in ser.unique().tolist() if x])
            _ensure_multiselect_key("f_desc", desc_options, [])
            st.multiselect("Description（去括号后）", desc_options, key="f_desc", placeholder="选择描述")
            sel_desc = list(ss.get("f_desc", []))

        # Product Code
        if "product_code" in base.columns:
            d = apply_all(base, exclude="code")
            code_options = sorted([x for x in d["product_code"].dropna().unique().tolist()])
            _ensure_multiselect_key("f_code", code_options, [])
            st.multiselect("Product Code", code_options, key="f_code", placeholder="选择产品编码")
            sel_code = list(ss.get("f_code", []))

        # Remark（来自括号）
        if "description" in base.columns:
            d = apply_all(base, exclude="remark")
            remark_options = extract_remarks(d["description"]) if not d.empty else []
            _ensure_multiselect_key("f_remark", remark_options, [])
            st.multiselect("Remark（来自描述括号）", remark_options, key="f_remark", placeholder="选择 Remark")
            st.checkbox("Exclude selected (Remark)", key="f_remark_ex", value=False)
            sel_remark = list(ss.get("f_remark", []))

        # 日期范围筛选（可选）
        use_date_filters = st.checkbox("启用日期范围筛选", value=False)

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
                min_d, max_d = base["expiry_date"].min(), base["expiry_date"].max()
                if pd.notna(min_d) and pd.notna(max_d):
                    ss["expiry_range"] = _clamp(ss.get("expiry_range"), min_d, max_d)
                    st.date_input("有效期范围", key="expiry_range", min_value=min_d.date(), max_value=max_d.date())
                    start, end = ss.get("expiry_range")

            if "relabel_to_date" in base.columns:
                min_r, max_r = base["relabel_to_date"].min(), base["relabel_to_date"].max()
                if pd.notna(min_r) and pd.notna(max_r):
                    ss["relabel_date_range"] = _clamp(ss.get("relabel_date_range"), min_r, max_r)
                    st.date_input("Relabel To 日期范围", key="relabel_date_range", min_value=min_r.date(), max_value=max_r.date())
                    r_start, r_end = ss.get("relabel_date_range")

    # 应用筛选
    work = apply_all(base)
    if use_date_filters and "expiry_date" in work.columns and start and end:
        mask_exp = work["expiry_date"].notna() & (work["expiry_date"].dt.date >= start) & (work["expiry_date"].dt.date <= end)
        work = work[mask_exp]
    if use_date_filters and "relabel_to_date" in work.columns and r_start and r_end:
        mask_rel = work["relabel_to_date"].notna() & (work["relabel_to_date"].dt.date >= r_start) & (work["relabel_to_date"].dt.date <= r_end)
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
    # 这里不再从侧栏提供 due_days；返回一个占位值，主流程用 Summary Bar 的值
    return work, sel_desc, st.session_state.get("summary_expiry_days", 30), selections


# ----------------------------- Excel 读入与规范化 -----------------------------

@st.cache_data(show_spinner=False)
def _find_sheet_name(file, desired: str) -> Optional[str]:
    """在工作簿中查找最匹配的工作表名，优先精确匹配，其次忽略大小写/空格。"""
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
    """规范化 'Stocks report'（Savori Whse）。严格从 J 列获取库存。"""
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
        9: "stock_qty",  # 来自 J
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
            norm[col] = pd.to_datetime(norm[col], errors="coerce", format="mixed")

    if "stock_qty" in norm.columns:
        norm["stock_qty"] = pd.to_numeric(norm["stock_qty"].astype(str).str.replace(",", "", regex=False), errors="coerce")

    for col in ["supplier", "brand", "product_code", "description", "pack_size", "unit"]:
        if col in norm.columns:
            norm[col] = norm[col].astype("string").str.strip()

    return norm, warning


def _normalize_lai_hock_whse(df: pd.DataFrame) -> Tuple[pd.DataFrame, Optional[str]]:
    """规范化 'Lai Hock Whse'（第三方仓）。"""
    expected_map = {
        0: "supplier",
        1: "brand",
        2: "product_code",
        3: "description",
        # 4: 忽略
        5: "pack_size",
        6: "unit",
        7: "expiry_date",
        8: "relabel_to_date",
        9: "stocks_balance",
        10: "stock_qty",  # from updated_stocks
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
            norm[col] = pd.to_datetime(norm[col], errors="coerce", format="mixed")

    if "stock_qty" in norm.columns:
        norm["stock_qty"] = pd.to_numeric(norm["stock_qty"].astype(str).str.replace(",", "", regex=False), errors="coerce")

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
    """读取并整合两个工作表，返回统一列的合并 DataFrame 与告警列表。"""
    warns = []
    name_sr = _find_sheet_name(file, "Stocks report")
    name_lh = _find_sheet_name(file, "Lai Hock Whse")

    df_sr = None
    df_lh = None
    try:
        if name_sr:
            df_sr = pd.read_excel(file, sheet_name=name_sr, header=2, dtype=str, engine="openpyxl").dropna(axis=0, how="all")
    except Exception as e:
        warns.append(f"读取工作表 '{name_sr}' 失败：{e}")
    try:
        if name_lh:
            df_lh = pd.read_excel(file, sheet_name=name_lh, header=2, dtype=str, engine="openpyxl").dropna(axis=0, how="all")
    except Exception as e:
        warns.append(f"读取工作表 '{name_lh}' 失败：{e}")

    frames = []
    if df_sr is not None:
        n_sr, w_sr = _normalize_stocks_report(df_sr)
        if w_sr:
            warns.append(f"Stocks report: {w_sr}")
        frames.append(n_sr)
    else:
        warns.append(f"未找到工作表：Stocks report（实际存在：{name_sr if name_sr else '无'}）")

    if df_lh is not None:
        n_lh, w_lh = _normalize_lai_hock_whse(df_lh)
        if w_lh:
            warns.append(f"Lai Hock Whse: {w_lh}")
        frames.append(n_lh)
    else:
        warns.append(f"未找到工作表：Lai Hock Whse（实际存在：{name_lh if name_lh else '无'}）")

    if frames:
        cols = [
            "supplier", "brand", "product_code", "description", "pack_size", "unit",
            "expiry_date", "relabel_to_date", "stock_qty", "warehouse",
        ]
        valid_frames = [f for f in frames if f is not None and not f.empty]
        combined = pd.concat(valid_frames, ignore_index=True, sort=False) if valid_frames else pd.DataFrame(columns=cols)
        for c in cols:
            if c not in combined.columns:
                combined[c] = pd.NA
        combined["stock_qty"] = pd.to_numeric(combined["stock_qty"], errors="coerce")
        for col in ["expiry_date", "relabel_to_date"]:
            combined[col] = pd.to_datetime(combined[col], errors="coerce", format="mixed")
        if "unit" in combined.columns:
            combined["unit"] = combined["unit"].apply(_canonical_unit)
        combined = combined[cols]
        return combined, warns

    return pd.DataFrame(), warns


# ----------------------------- UI 状态回调 -----------------------------

def _touch_summary_refresh_token() -> None:
    st.session_state["__summary_refresh_token"] = st.session_state.get("__summary_refresh_token", 0) + 1

def on_change_expiry_days() -> None:
    value = st.session_state.get("summary_expiry_days", 30)
    try:
        value_int = int(value)
    except (TypeError, ValueError):
        value_int = 30
    if value_int < 1:
        value_int = 1
    st.session_state["summary_expiry_days"] = value_int
    _touch_summary_refresh_token()

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
    _touch_summary_refresh_token()

def on_toggle_near_expiry() -> None:
    st.session_state["toggle_only_near"] = bool(st.session_state.get("toggle_only_near", False))
    _touch_summary_refresh_token()

def on_toggle_low_stock() -> None:
    st.session_state["toggle_only_low"] = bool(st.session_state.get("toggle_only_low", False))
    _touch_summary_refresh_token()


# ----------------------------- 主程序 -----------------------------

def main():
    st.set_page_config(page_title="Stocks DataGrid", layout="wide")
    st.title("Stock Dashboard (Stocks DataGrid)")
    st.caption("Upload an Excel file containing the 'Stocks report' and 'Lai Hock Whse' sheets (data starts on row 3).")

    uploaded = st.file_uploader("Upload Excel (.xlsx)", type=["xlsx"])
    if not uploaded:
        st.info("Upload a source workbook to begin.")
        return

    try:
        df, warns = load_and_normalize(uploaded)
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

    # 侧栏筛选（阈值不在侧栏）
    filtered, selected_descs, _placeholder_due_days, filter_state = apply_filters_v2(df_display)
    total_rows = len(filtered)

    # 主聚合（产品层）
    summary_df, detail_map = aggregate_summary(filtered)
    summary_df = summary_df.copy()

    # 顶部合计
    total_ctn_all = float(summary_df["total_ctn"].sum()) if not summary_df.empty else 0.0
    total_pkt_all = float(summary_df["total_pkt"].sum()) if not summary_df.empty else 0.0
    totals_display = _format_quantity_pair(total_ctn_all, total_pkt_all)

    # 初始化 Summary Bar 状态
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

    # Summary Bar
    summary_bar = st.container()
    with summary_bar:
        metric_holder = st.container()
        controls_holder = st.container()

    with controls_holder:
        toggle_cols = st.columns([1, 1, 1, 2])
        toggle_cols[0].toggle("Show Near-Expiry Only", key="toggle_only_near", on_change=on_toggle_near_expiry)
        toggle_cols[1].toggle("Show Low-Stock Only", key="toggle_only_low", on_change=on_toggle_low_stock)
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
        product_query_raw = st.text_input(
            "Product Quick Filter",
            key="product_quick_filter",
            placeholder="Filter by supplier / product / code",
            help="Client-side filter that applies to Supplier, Product, and Product Code.",
        )

    expiry_days = int(st.session_state.get("summary_expiry_days", 30))
    global_low_ctn = int(st.session_state.get("summary_global_low_ctn", 0))
    global_low_pkt = int(st.session_state.get("summary_global_low_pkt", 0))
    near_only = bool(st.session_state.get("toggle_only_near", False))
    low_only = bool(st.session_state.get("toggle_only_low", False))
    show_depleted = bool(st.session_state.get("toggle_show_depleted", True))
    product_query = (product_query_raw or "").strip()

    if summary_df.empty:
        with metric_holder:
            metric_cols = st.columns([1, 1, 1, 1])
            metric_cols[0].metric("Filtered Rows", f"{total_rows}")
            metric_cols[1].metric("Totals", totals_display if totals_display else "0")
            metric_cols[2].metric("Near-Expiry", "0")
            metric_cols[3].metric("Low-Stock", "0")
        st.info("No data found for Savori Whse / Lai Hock Whse under the current filters.")
        return

    # 1) 为每个产品构建批次层表（带状态）
    expiry_tables_cache: Dict[Tuple[str, str], pd.DataFrame] = {}
    for key in summary_df["group_key"]:
        tuple_key = tuple(key) if not isinstance(key, tuple) else key
        if tuple_key not in expiry_tables_cache:
            source_df = detail_map.get(tuple_key, pd.DataFrame())
            expiry_tables_cache[tuple_key] = split_by_expiry(
                source_df, expiry_days=expiry_days, show_depleted=show_depleted
            )

    # 2) 计算 Low-Stock（产品层，优先 ROP 再全局）
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

    # 3) 由批次层结果得出产品层的 Expired / Near-Expiry
    summary_df["has_expired_batch"] = False
    summary_df["has_near_batch"] = False
    for i, r in summary_df.iterrows():
        tuple_key = tuple(r["group_key"]) if not isinstance(r["group_key"], tuple) else r["group_key"]
        tbl = expiry_tables_cache.get(tuple_key)
        if tbl is None or tbl.empty:
            continue
        pos_qty = (tbl["subtotal_ctn"].astype(float) + tbl["subtotal_pkt"].astype(float)) > 0
        if not pos_qty.any():
            continue
        sub = tbl.loc[pos_qty]
        if (sub["status_batch"] == "Expired").any():
            summary_df.at[i, "has_expired_batch"] = True
        if (sub["status_batch"] == "Near-Expiry").any():
            summary_df.at[i, "has_near_batch"] = True

    # 4) 计算“多标签”并产出主状态（用于排序/统计）
    def _primary_status(has_expired: bool, has_near: bool, is_low: bool) -> str:
        if has_expired:
            return "Expired"
        if has_near:
            return "Near-Expiry"
        if is_low:
            return "Low-Stock"
        return "OK"

    summary_df["Status Tags"] = [[] for _ in range(len(summary_df))]
    summary_df["status_product"] = ""  # 主状态（用于排序/统计）

    for i, r in summary_df.iterrows():
        tags = []
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
        )

    # 5) 顶部计数（Near-Expiry 包含 Expired）
    near_count = int(summary_df["status_product"].isin(["Expired", "Near-Expiry"]).sum())
    low_count = int(summary_df["status_product"].eq("Low-Stock").sum())

    with metric_holder:
        metric_cols = st.columns([1, 1, 1, 1])
        metric_cols[0].metric("Filtered Rows", f"{total_rows}")
        metric_cols[1].metric("Totals", totals_display if totals_display else "0")
        metric_cols[2].metric("Near-Expiry", f"{near_count}")
        metric_cols[3].metric("Low-Stock", f"{low_count}")

    # 6) 显示列 + Status 文本（把多标签渲染出来）
    summary_df["Savori Whse"] = [_format_quantity_pair(r.savori_ctn, r.savori_pkt) for r in summary_df.itertuples()]
    summary_df["Lai Hock Whse"] = [_format_quantity_pair(r.lai_hock_ctn, r.lai_hock_pkt) for r in summary_df.itertuples()]
    summary_df["Total"] = [_format_quantity_pair(r.total_ctn, r.total_pkt) for r in summary_df.itertuples()]

    def _format_status(tags, reason):
        # tags 现在是一个可能包含多个标签的 list，比如 ["Expired","Low-Stock"]
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


    # 7) 过滤（Near-Expiry Only / Low-Stock Only / 文本过滤）
    view_df = summary_df.copy()
    if near_only:
        view_df = view_df[view_df["status_product"].isin(["Expired", "Near-Expiry"])]
    if low_only:
        view_df = view_df[view_df["status_product"].eq("Low-Stock")]
    if product_query:
        mask = (
            view_df["Product"].str.contains(product_query, case=False, na=False)
            | view_df["Supplier"].str.contains(product_query, case=False, na=False)
            | view_df["Product Code"].str.contains(product_query, case=False, na=False)
        )
        view_df = view_df[mask]
    view_df = view_df.reset_index(drop=True)

    if view_df.empty:
        st.info("No rows in the main table match the current view. Adjust filters or thresholds.")
        return

    # 8) 排序：Expired → Near-Expiry → Low-Stock → OK；同组按最近有量批次天数升序，再按总量、产品名
    def _status_rank(s: str) -> int:
        return {"Expired": 0, "Near-Expiry": 1, "Low-Stock": 2, "OK": 3}.get(s, 9)

    def _nearest_pos_days(tuple_key):
        tbl = expiry_tables_cache.get(tuple_key)
        if tbl is None or tbl.empty:
            return 10**9
        pos = (tbl["subtotal_ctn"].astype(float) + tbl["subtotal_pkt"].astype(float)) > 0
        if not pos.any():
            return 10**9
        sub = tbl.loc[pos]
        vals = [v for v in sub["days_to_expiry"].tolist() if isinstance(v, int)]
        return min(vals) if vals else 10**9

    view_df["status_priority"] = view_df["status_product"].map(_status_rank)
    view_df["nearest_days"] = [_nearest_pos_days(tuple(k) if not isinstance(k, tuple) else k) for k in view_df["group_key"]]
    view_df = view_df.sort_values(
        by=["status_priority", "nearest_days", "total_ctn", "total_pkt", "Product"],
        kind="stable",
    ).reset_index(drop=True)

    # 9) 主表渲染与着色（先清理 HTML，再着色；只对命中状态着色）
    display_columns = ["Supplier", "Product", "Pack Size", "Product Code", "Savori Whse", "Lai Hock Whse", "Total", "Status"]
    view_df_display = view_df[display_columns].copy()

    # 去除任何 HTML 标签（修复 span 出现的问题）
    view_df_display_clean = _strip_html_df(view_df_display)

    # 在样式函数外部定义好 Total 的列索引
    total_idx = list(view_df_display_clean.columns).index("Total")


    def _style_row(row):
        tags = view_df.loc[row.name, "Status Tags"]
        styles = ["" for _ in row]

        if "Expired" in tags and "Low-Stock" in tags:
            # 双状态：整行紫色底，Total 列加粗
            styles = ["background-color: rgba(102,51,153,0.25);"] * len(row)  # 这里选中等亮度紫
            styles[total_idx] += "font-weight:600; color:#FFF;"
        elif "Near-Expiry" in tags and "Low-Stock" in tags:
            # 双状态：整行橙色底，Total 列加粗
            styles = ["background-color: rgba(255,140,0,0.22);"] * len(row)  # 深橙色
            styles[total_idx] += "font-weight:600; color:#FFF;"
        elif "Expired" in tags:
            styles = ["background-color: rgba(178,34,34,0.18);"] * len(row)
        elif "Near-Expiry" in tags:
            styles = ["background-color: rgba(255,165,0,0.18);"] * len(row)
        elif "Low-Stock" in tags:
            styles[total_idx] = "background-color: rgba(255, 255, 0, 0.18); font-weight:600;"
        return styles




    styled_view = view_df_display_clean.style.apply(_style_row, axis=1)
    st.dataframe(styled_view, use_container_width=True, hide_index=True)

    # 10) 导出：Summary + Expiry Breakdown（带批次状态/说明）
    export_summary_cols = [
    "Supplier", "Product", "Pack Size", "Product Code",
    "Savori Whse", "Lai Hock Whse", "Total", "Status",
    "total_ctn", "total_pkt", "savori_ctn", "savori_pkt", "lai_hock_ctn", "lai_hock_pkt",
    "reorder_point_ctn", "reorder_point_pkt", "low_stock_reason", "status_product",
    ]
    export_summary = view_df[export_summary_cols].copy()

    expiry_export_frames = []
    for _, row in view_df.iterrows():
        tuple_key = tuple(row["group_key"]) if not isinstance(row["group_key"], tuple) else row["group_key"]
        expiry_table = expiry_tables_cache.get(tuple_key)
        if expiry_table is None or expiry_table.empty:
            continue
        export_block = expiry_table.copy()
        export_block.insert(0, "Supplier", row["Supplier"])
        export_block.insert(1, "Product", row["Product"])
        export_block.insert(2, "Product Code", row["Product Code"])
        export_block.insert(3, "Product Status", row["status_product"])
        export_block.insert(4, "Parent Total", row["Total"])
        expiry_export_frames.append(export_block)
    # 过滤掉 None 或空的 DataFrame
    valid_frames = [f for f in expiry_export_frames if f is not None and not f.empty]

    # 统一列顺序与 dtype，避免 pandas 未来版本 concat 的 dtype 警告
    expected_cols = [
        "Supplier", "Product", "Product Code", "Product Status", "Parent Total",
        "Expiry", "Savori Whse", "Lai Hock Whse", "Subtotal",
        "subtotal_ctn", "subtotal_pkt", "expiry_date", "status_batch", "days_to_expiry", "Info",
    ]

    normalized = []
    for f in valid_frames:
        g = f.copy()

        # 确保所有期望列都存在；缺失的补 string-NA
        for c in expected_cols:
            if c not in g.columns:
                g[c] = pd.Series([pd.NA] * len(g), dtype="string")

        # 关键列统一 dtype
        # 文本列 → string
        for c in ["Supplier", "Product", "Product Code", "Product Status",
                "Parent Total", "Expiry", "Savori Whse", "Lai Hock Whse",
                "Subtotal", "status_batch", "Info"]:
            if c in g.columns:
                if g[c].dtype != "string":
                    g[c] = g[c].astype("string")

        # 数值列 → float64
        for c in ["subtotal_ctn", "subtotal_pkt"]:
            if c in g.columns:
                g[c] = pd.to_numeric(g[c], errors="coerce").astype("float64")

        # 日期列 → datetime64[ns]
        if "expiry_date" in g.columns:
            g["expiry_date"] = pd.to_datetime(g["expiry_date"], errors="coerce")

        # 天数字段 → 可空整型（避免 object）
        if "days_to_expiry" in g.columns:
            # 先转成数值，再转可空整型
            g["days_to_expiry"] = pd.to_numeric(g["days_to_expiry"], errors="coerce").astype("Int64")

        # 统一列顺序
        g = g[expected_cols]
        normalized.append(g)

    if normalized:
        expiry_export_df = pd.concat(normalized, ignore_index=True)
    else:
        expiry_export_df = pd.DataFrame(columns=expected_cols)



    download_cols = st.columns([1, 1])
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
        export_summary_clean.to_excel(writer, sheet_name="Summary", index=False)
        expiry_export_clean.to_excel(writer, sheet_name="Expiry Breakdown", index=False)

    excel_buffer.seek(0)
    download_cols[1].download_button(
        "Export Excel", data=excel_buffer.getvalue(),
        file_name="stock_summary.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    # 11) 展开明细：批次层高亮与 Info
    for _, row in view_df.iterrows():
        tuple_key = tuple(row["group_key"]) if not isinstance(row["group_key"], tuple) else row["group_key"]
        expiry_table = expiry_tables_cache.get(tuple_key)
        title = f"{row['Supplier']} | {row['Product']}"
        with st.expander(title, expanded=False):
            st.caption(f"Product Code: {row['Product Code']}")
            st.caption(f"Status: {row['Status']}")
            if row.get("low_stock_reason"):
                st.caption(f"Low-stock trigger: {'Product ROP' if row['low_stock_reason']=='ROP' else 'Global threshold'}")

            if expiry_table is None or expiry_table.empty:
                st.info("No expiry breakdown available.")
            else:
                display_cols2 = ["Expiry", "Savori Whse", "Lai Hock Whse", "Subtotal", "status_batch", "Info"]
                table_display = expiry_table[display_cols2].rename(columns={"status_batch": "Batch Status"})

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
                    styled_expiry = table_display.style.apply(_style_expiry, axis=1)
                    st.dataframe(styled_expiry, use_container_width=True, hide_index=True)
                except Exception:
                    st.dataframe(table_display, use_container_width=True, hide_index=True)

                # 校验小计与父合计一致（仅数量值）
                subtotal_ctn_sum = float(expiry_table["subtotal_ctn"].sum())
                subtotal_pkt_sum = float(expiry_table["subtotal_pkt"].sum())
                parent_ctn = float(row["total_ctn"])
                parent_pkt = float(row["total_pkt"])
                if (not math.isclose(subtotal_ctn_sum, parent_ctn, abs_tol=1e-6)
                        or not math.isclose(subtotal_pkt_sum, parent_pkt, abs_tol=1e-6)):
                    st.warning("Expiry breakdown totals do not match the parent totals; please verify the source data.")
                st.caption(f"Total: {row['Total']}")


if __name__ == "__main__" and os.environ.get("ST_REDIRECTED", "0") != "1" and (get_script_run_ctx() is None):
    # 直接 python 执行时，上面已重定向为 streamlit 运行，这里不再重复调用 main()
    pass
else:
    # 被 streamlit 执行时进入
    if __name__ == "__main__":
        main()
