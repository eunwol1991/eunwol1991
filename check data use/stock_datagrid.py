import os
import sys

# If executed directly via `python stock_datagrid.py`, re-launch with Streamlit.
# Uses a Windows-friendly subprocess approach to handle spaces in paths.
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
    # Optional debug of the redirect target
    if os.environ.get("ST_DEBUG_REDIRECT") == "1":
        print("[streamlit-redirect] ", cmd)
    subprocess.run(cmd, check=False)
    sys.exit(0)

import streamlit as st
import pandas as pd
import math
import datetime
from typing import Optional, Tuple, List

PRIMARY_WAREHOUSES = ["Savori Whse", "Lai Hock Whse"]


def _normalize_warehouse_name(value) -> str:
    if pd.isna(value):
        return ""
    return str(value).strip()


def _filter_primary_warehouses(df: pd.DataFrame) -> pd.DataFrame:
    if "warehouse" not in df.columns:
        return df
    mask = df["warehouse"].map(_normalize_warehouse_name).isin(PRIMARY_WAREHOUSES)
    return df.loc[mask].copy()


# Canonicalize unit strings to a normalized, lowercase key
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


@st.cache_data(show_spinner=False)
def _find_sheet_name(file, desired: str) -> Optional[str]:
    """在工作簿中查找最匹配的工作表名，优先精确匹配，其次忽略大小写/空格。"""
    try:
        xls = pd.ExcelFile(file, engine="openpyxl")
    except Exception:
        return None
    names = xls.sheet_names
    # Exact match first
    for n in names:
        if n == desired:
            return n
    # Case-insensitive
    for n in names:
        if n.lower() == desired.lower():
            return n
    # Ignore spaces and case
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
        9: "stock_qty",  # This must come from column J
    }

    warning = None
    pos_renames = {}
    for idx, name in expected_map.items():
        if idx < df.shape[1]:
            pos_renames[df.columns[idx]] = name

    norm = df.rename(columns=pos_renames).copy()
    norm["warehouse"] = "Savori Whse"

    # Ensure required columns exist
    required = [
        "supplier",
        "brand",
        "product_code",
        "description",
        "pack_size",
        "unit",
        "expiry_date",
        "relabel_to_date",
        "stock_qty",
    ]

    missing = [c for c in required if c not in norm.columns]
    if missing:
        warning = (
            "根据预期位置缺少列，或表头不在第 3 行：" + ", ".join(missing) +
            "。请确保 Excel 的 'Stocks report' 工作表表头位于 A3:J3。"
        )

    # Coerce dates
    for col in ["expiry_date", "relabel_to_date"]:
        if col in norm.columns:
            norm[col] = pd.to_datetime(norm[col], errors="coerce", format="mixed")

    # Coerce stock column strictly from J (index 9 mapped to 'stock_qty')
    if "stock_qty" in norm.columns:
        norm["stock_qty"] = pd.to_numeric(norm["stock_qty"].str.replace(
            ",", "", regex=False
        ), errors="coerce")

    # Clean strings (strip whitespace)
    for col in ["supplier", "brand", "product_code", "description", "pack_size", "unit"]:
        if col in norm.columns:
            norm[col] = norm[col].astype("string").str.strip()

    return norm, warning


def _normalize_lai_hock_whse(df: pd.DataFrame) -> Tuple[pd.DataFrame, Optional[str]]:
    """规范化 'Lai Hock Whse'（3rd party 仓）。

    列位映射：
    A supplier, B brand, C product_code, D description,
    E（忽略）, F pack_size, G unit, H expiry_date, I relabel_to_date,
    J stocks_balance, K updated_stocks -> 用于 stock_qty。
    G 列默认内容为 ctn（若缺失则补为 'ctn'）。
    """
    expected_map = {
        0: "supplier",
        1: "brand",
        2: "product_code",
        3: "description",
        # 4: 忽略 pallet 列
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

    # Validate required columns presence based on expected positions
    required = [
        "supplier",
        "brand",
        "product_code",
        "description",
        "pack_size",
        "unit",
        "expiry_date",
        "relabel_to_date",
        "stock_qty",
    ]
    missing = [c for c in required if c not in norm.columns]
    if missing:
        warn_msgs.append(
            "根据预期位置缺少列，或表头不在第 3 行：" + ", ".join(missing) +
            "。请确保 Excel 的 'Lai Hock Whse' 工作表表头位于 A3:K3。"
        )

    # Coerce dates
    for col in ["expiry_date", "relabel_to_date"]:
        if col in norm.columns:
            norm[col] = pd.to_datetime(norm[col], errors="coerce", format="mixed")

    # Stock comes from updated_stocks (mapped to stock_qty), ensure numeric
    if "stock_qty" in norm.columns:
        norm["stock_qty"] = pd.to_numeric(norm["stock_qty"].astype(str).str.replace(",", "", regex=False), errors="coerce")

    # Unit default to ctn if missing/blank
    if "unit" not in norm.columns:
        norm["unit"] = "CTN"
        warn_msgs.append("缺少 Unit 列，已将缺失处默认填为 CTN")
    else:
        norm["unit"] = norm["unit"].astype("string").str.strip()
        norm.loc[norm["unit"].isna() | (norm["unit"] == ""), "unit"] = "CTN"

    # Clean basic strings
    for col in ["supplier", "brand", "product_code", "description", "pack_size"]:
        if col in norm.columns:
            norm[col] = norm[col].astype("string").str.strip()

    warning = "；".join(warn_msgs) if warn_msgs else None
    return norm, warning


def load_and_normalize(file) -> Tuple[pd.DataFrame, list]:
    """读取并整合两个工作表，返回统一列的合并 DataFrame 与告警列表。"""
    warns = []
    # Find actual sheet names with some tolerance
    name_sr = _find_sheet_name(file, "Stocks report")
    name_lh = _find_sheet_name(file, "Lai Hock Whse")

    df_sr = None
    df_lh = None
    try:
        if name_sr:
            # 仅丢弃空行，不丢弃空列，避免因整列为空导致列位左移
            df_sr = pd.read_excel(
                file, sheet_name=name_sr, header=2, dtype=str, engine="openpyxl"
            ).dropna(axis=0, how="all")
    except Exception as e:
        warns.append(f"读取工作表 '{name_sr}' 失败：{e}")
    try:
        if name_lh:
            # 仅丢弃空行，不丢弃空列，确保 G 列（unit）即使全空也保留
            df_lh = pd.read_excel(
                file, sheet_name=name_lh, header=2, dtype=str, engine="openpyxl"
            ).dropna(axis=0, how="all")
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
        # unify expected columns
        cols = [
            "supplier", "brand", "product_code", "description", "pack_size", "unit",
            "expiry_date", "relabel_to_date", "stock_qty", "warehouse",
        ]
        valid_frames = [f for f in frames if f is not None and not f.empty]
        if valid_frames:
            combined = pd.concat(valid_frames, ignore_index=True, sort=False)
        else:
            combined = pd.DataFrame(columns=cols)
        for c in cols:
            if c not in combined.columns:
                combined[c] = pd.NA
        # Ensure types
        combined["stock_qty"] = pd.to_numeric(combined["stock_qty"], errors="coerce")
        for col in ["expiry_date", "relabel_to_date"]:
            combined[col] = pd.to_datetime(combined[col], errors="coerce", format="mixed")
        # Canonicalize unit values for consistency (e.g., carton -> ctn)
        if "unit" in combined.columns:
            combined["unit"] = combined["unit"].apply(_canonical_unit)
        # Order columns
        combined = combined[cols]
        return combined, warns

    return pd.DataFrame(), warns


def apply_filters(df: pd.DataFrame):
    """渲染侧边栏筛选并返回筛选后的 DataFrame、所选 Description 列表与“即将到期高亮”天数。

    分级筛选：Supplier → Brand → Description → Product Code。
    同时支持有效期与 Relabel To 的日期范围筛选。
    """
    work = df.copy()

    due_days = 30
    selected_descs = []

    with st.sidebar:
        st.header("筛选条件")

        # Warehouse filter (sheet selector)
        if "warehouse" in work.columns:
            warehouses = [x for x in work["warehouse"].dropna().astype(str).unique().tolist()]
            # Keep stable order: Savori first if present
            order = [w for w in ["Savori Whse", "Lai Hock Whse"] if w in warehouses]
            order += [w for w in warehouses if w not in order]
            sel_wh = st.multiselect(
                "仓库",
                options=order,
                default=order,
                placeholder="选择一个或多个仓库",
            )
            if sel_wh:
                work = work[work["warehouse"].isin(sel_wh)]

        # Supplier
        if "supplier" in work.columns:
            suppliers = sorted([x for x in work["supplier"].dropna().unique()])
            sel_suppliers = st.multiselect("Supplier", suppliers, placeholder="选择一个或多个供应商")
            if sel_suppliers:
                work = work[work["supplier"].isin(sel_suppliers)]

        # Brand (depends on Supplier)
        if "brand" in work.columns:
            brands = sorted([x for x in work["brand"].dropna().unique()])
            sel_brands = st.multiselect("Brand", brands, placeholder="选择一个或多个品牌")
            if sel_brands:
                work = work[work["brand"].isin(sel_brands)]

        # Description (depends on Supplier/Brand)
        if "description" in work.columns:
            # Optional keyword search to narrow options quickly
            q_desc = st.text_input("Description 关键字", value="").strip()
            # Base description: strip any parenthesis content, treat as same item
            base_series_all = work["description"].astype(str).str.replace(r"\s*\([^)]*\)", "", regex=True).str.strip()
            desc_series = base_series_all.dropna()
            if q_desc:
                desc_series = desc_series[desc_series.str.contains(q_desc, case=False, na=False)]
            descs = sorted(desc_series.unique().tolist())
            sel_descs = st.multiselect("Description（去括号后）", descs, placeholder="选择描述")
            if sel_descs:
                selected_descs = sel_descs
                mask = base_series_all.isin(sel_descs)
                work = work[mask]

            # Remark filter extracted from parentheses content in description
            # e.g., "Product Name (Frozen)" -> Remark = "Frozen"
            import re
            if not work["description"].dropna().empty:
                # collect all parenthesis contents
                remark_series = work["description"].astype(str).str.findall(r"\(([^)]*)\)")
                all_remarks = sorted({r.strip() for lst in remark_series.dropna().tolist() for r in (lst or []) if r.strip()})
                if all_remarks:
                    sel_remarks = st.multiselect("Remark（来自 Description 括号内）", all_remarks, placeholder="选择 Remark")
                    if sel_remarks:
                        pattern = "|".join([re.escape(x) for x in sel_remarks])
                        work = work[work["description"].astype(str).str.contains(r"\((?:" + pattern + r")\)", case=False, na=False)]

        # Product Code (depends on previous)
        if "product_code" in work.columns:
            codes = sorted([x for x in work["product_code"].dropna().unique()])
            sel_codes = st.multiselect("Product Code", codes, placeholder="选择产品编码")
            if sel_codes:
                work = work[work["product_code"].isin(sel_codes)]

        # Date ranges
        if "expiry_date" in work.columns:
            min_d, max_d = work["expiry_date"].min(), work["expiry_date"].max()
            if pd.notna(min_d) and pd.notna(max_d):
                start, end = st.date_input(
                    "有效期范围",
                    value=(min_d.date(), max_d.date()),
                    min_value=min_d.date(),
                    max_value=max_d.date(),
                )
                work = work[(work["expiry_date"].dt.date >= start) & (work["expiry_date"].dt.date <= end)]

        if "relabel_to_date" in work.columns:
            min_d, max_d = work["relabel_to_date"].min(), work["relabel_to_date"].max()
            if pd.notna(min_d) and pd.notna(max_d):
                start, end = st.date_input(
                    "Relabel To 日期范围",
                    value=(min_d.date(), max_d.date()),
                    min_value=min_d.date(),
                    max_value=max_d.date(),
                    key="relabel_date_range",
                )
                work = work[(work["relabel_to_date"].dt.date >= start) & (work["relabel_to_date"].dt.date <= end)]

        # Expiring soon highlighter (days)
        if "expiry_date" in work.columns:
            due_days = int(st.number_input(
                "即将到期高亮（天内）",
                min_value=1,
                max_value=365,
                value=30,
                step=1,
                help="将高亮所有已过期，以及从今天起未来指定天数内到期的行"
            ))

    return work, selected_descs, due_days


def apply_filters_v2(df: pd.DataFrame):
    """改进版筛选：实现 Excel 式逐步收缩，任意维度先选都能联动其它选项。
    维度顺序：Warehouse → Supplier → Brand → Description(去括号) → Product Code → Remark(来自括号)。
    返回：筛选后的 DataFrame、所选 Description 列表、到期高亮天数、当前筛选状态字典。
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

    # Helpers to keep selections stable when options shrink
    def _ensure_multiselect_key(key: str, options: list, init: list):
        # initialize if missing
        if key not in ss:
            ss[key] = list(init)
        else:
            # prune values not in current options (keep intersection)
            current = ss.get(key, [])
            if isinstance(current, (str, int, float)):
                current = [current]
            ss[key] = [v for v in current if v in options]

    # Prime local selections from session_state (may be updated after widgets)
    sel_wh = list(ss.get("f_wh", []))
    sel_sup = list(ss.get("f_sup", []))
    sel_brand = list(ss.get("f_brand", []))
    sel_desc = list(ss.get("f_desc", []))
    sel_code = list(ss.get("f_code", []))
    sel_remark = list(ss.get("f_remark", []))

    def apply_all(df_in: pd.DataFrame, exclude: str = "") -> pd.DataFrame:
        d = df_in
        if exclude != "warehouse" and sel_wh and "warehouse" in d.columns:
            d = d[d["warehouse"].isin(sel_wh)]
        if exclude != "supplier" and sel_sup and "supplier" in d.columns:
            d = d[d["supplier"].isin(sel_sup)]
        if exclude != "brand" and sel_brand and "brand" in d.columns:
            d = d[d["brand"].isin(sel_brand)]
        if exclude != "desc" and sel_desc and "description" in d.columns:
            base_ser = get_desc_base(d["description"]) if "description" in d.columns else pd.Series(dtype=str)
            d = d[base_ser.isin(sel_desc)]
        if exclude != "code" and sel_code and "product_code" in d.columns:
            d = d[d["product_code"].isin(sel_code)]
        if exclude != "remark" and sel_remark and "description" in d.columns:
            matches = d["description"].astype(str).str.findall(r"\(([^)]*)\)")
            mask = matches.apply(lambda lst: any((str(x).strip() in sel_remark) for x in (lst or [])))
            d = d[mask]
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
            default_selection = [w for w in ["Savori Whse", "Lai Hock Whse"] if w in ordered]
            if not default_selection:
                default_selection = list(ordered)
            _ensure_multiselect_key("f_wh", ordered, default_selection)
            st.multiselect("仓库", options=ordered, key="f_wh", placeholder="选择一个或多个仓库")
            sel_wh = list(ss.get("f_wh", []))

        # Supplier
        if "supplier" in base.columns:
            d = apply_all(base, exclude="supplier")
            sup_options = sorted([x for x in d["supplier"].dropna().unique().tolist()])
            _ensure_multiselect_key("f_sup", sup_options, [])
            st.multiselect("Supplier", sup_options, key="f_sup", placeholder="选择供应商")
            sel_sup = list(ss.get("f_sup", []))

        # Brand
        if "brand" in base.columns:
            d = apply_all(base, exclude="brand")
            brand_options = sorted([x for x in d["brand"].dropna().unique().tolist()])
            _ensure_multiselect_key("f_brand", brand_options, [])
            st.multiselect("Brand", brand_options, key="f_brand", placeholder="选择品牌")
            sel_brand = list(ss.get("f_brand", []))

        # Description (base, stripped)
        if "description" in base.columns:
            d = apply_all(base, exclude="desc")
            base_ser = get_desc_base(d["description"]) if not d.empty else pd.Series(dtype=str)
            # keyword input
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

        # Remark from parentheses
        if "description" in base.columns:
            d = apply_all(base, exclude="remark")
            remark_options = extract_remarks(d["description"]) if not d.empty else []
            _ensure_multiselect_key("f_remark", remark_options, [])
            st.multiselect("Remark（来自描述括号）", remark_options, key="f_remark", placeholder="选择 Remark")
            sel_remark = list(ss.get("f_remark", []))

        # --- 日期筛选开关（默认关闭） + 边界固定为全量 ---
        use_date_filters = st.checkbox("启用日期范围筛选", value=False)

        # 高亮天数（不依赖是否开启日期筛选）
        due_days = int(st.number_input(
            "即将到期高亮（天内）", min_value=1, max_value=365, value=30, step=1,
            help="将高亮所有已过期，以及从今天起未来指定天数内到期的行"
        ))

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
            # Expiry controls: bounds from full base
            if "expiry_date" in base.columns:
                min_d, max_d = base["expiry_date"].min(), base["expiry_date"].max()
                if pd.notna(min_d) and pd.notna(max_d):
                    ss["expiry_range"] = _clamp(ss.get("expiry_range"), min_d, max_d)
                    st.date_input("有效期范围", key="expiry_range", min_value=min_d.date(), max_value=max_d.date())
                    start, end = ss.get("expiry_range")

            # Relabel controls: bounds from full base
            if "relabel_to_date" in base.columns:
                min_r, max_r = base["relabel_to_date"].min(), base["relabel_to_date"].max()
                if pd.notna(min_r) and pd.notna(max_r):
                    ss["relabel_date_range"] = _clamp(ss.get("relabel_date_range"), min_r, max_r)
                    st.date_input("Relabel To 日期范围", key="relabel_date_range", min_value=min_r.date(), max_value=max_r.date())
                    r_start, r_end = ss.get("relabel_date_range")

    # Final filtered data
    work = apply_all(base)
    # 仅在开启时应用日期筛选，且对 NaT 安全
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
    return work, sel_desc, due_days, selections


def main():
    st.set_page_config(page_title="Stocks DataGrid", layout="wide")
    st.title("库存数据筛选 (Stocks DataGrid)")
    st.caption("整合 'Stocks report'（Savori Whse）与 'Lai Hock Whse' 两个工作表，表头均从 A3 开始。")

    uploaded = st.file_uploader("上传 Excel 文件 (.xlsx)", type=["xlsx"]) 
    if not uploaded:
        st.info("请上传包含 'Stocks report' 与/或 'Lai Hock Whse' 工作表的 Excel 文件以开始。")
        return

    try:
        df, warns = load_and_normalize(uploaded)
    except Exception as e:
        st.error(f"读取 Excel 失败：{e}")
        return

    for w in warns:
        st.warning(w)

    # Reorder columns for display if present
    display_cols = [
        "warehouse",
        "supplier",
        "brand",
        "product_code",
        "description",
        "pack_size",
        "unit",
        "expiry_date",
        "relabel_to_date",
        "stock_qty",
    ]
    display_cols = [c for c in display_cols if c in df.columns]
    df_display = df[display_cols].copy()

    # 使用改进后的联动筛选
    filtered, selected_descs, due_days, filter_state = apply_filters_v2(df_display)
    filter_state = filter_state or {}
    selected_suppliers = list(filter_state.get("supplier", []))

    # KPIs
    total_rows = len(filtered)

    def _plural(unit: str, n: float) -> str:
        u = (unit or "").strip()
        if not u:
            return u
        if abs(n) == 1:
            return u
        mapping = {
            "ctn": "ctns",
            "pkt": "pkts",
            "box": "boxes",
            "tin": "tins",
            "can": "cans",
            "bag": "bags",
            "piece": "pieces",
            "pc": "pcs",
            "pcs": "pcs",
        }
        return mapping.get(u.lower(), u + "s")

    def _normalize_unit(u: Optional[str]) -> str:
        return _canonical_unit(u)

    UNIT_PRIORITY_ORDER = {
        "ctn": 0,
        "box": 1,
        "tin": 2,
        "can": 3,
        "pkt": 4,
        "bag": 5,
        "pc": 6,
        "piece": 6,
    }
    _UNIT_PRIORITY_FALLBACK = len(UNIT_PRIORITY_ORDER) + 1

    def _unit_priority(u: Optional[str]) -> int:
        if u is None:
            return _UNIT_PRIORITY_FALLBACK
        if isinstance(u, float):
            try:
                if math.isnan(u):
                    return _UNIT_PRIORITY_FALLBACK
            except Exception:
                pass
        canon = _canonical_unit(u)
        if not canon:
            return _UNIT_PRIORITY_FALLBACK
        return UNIT_PRIORITY_ORDER.get(canon, _UNIT_PRIORITY_FALLBACK)

    def _format_quantity(value: float) -> str:
        if pd.isna(value):
            return ""
        val = float(value)
        if math.isclose(val, round(val), rel_tol=1e-9, abs_tol=1e-9):
            return f"{int(round(val))}"
        text_val = f"{val:,.2f}"
        return text_val.rstrip("0").rstrip(".")

    def _collect_unit_totals(df_subset: pd.DataFrame) -> List[Tuple[str, float]]:
        if df_subset.empty:
            return []
        working = df_subset.assign(
            stock_qty=pd.to_numeric(df_subset["stock_qty"], errors="coerce").fillna(0),
            unit_key=df_subset["unit"].apply(_normalize_unit),
        )
        grouped = (
            working.groupby("unit_key", dropna=False)["stock_qty"].sum().reset_index()
        )
        if grouped.empty:
            return []
        grouped["priority"] = grouped["unit_key"].apply(_unit_priority)
        grouped = grouped.sort_values(by=["priority", "unit_key"], kind="stable")
        results: List[Tuple[str, float]] = []
        for row in grouped.itertuples(index=False):
            unit_key = getattr(row, "unit_key")
            qty = getattr(row, "stock_qty")
            if pd.isna(unit_key):
                continue
            unit_text = str(unit_key).strip()
            if not unit_text:
                continue
            if pd.isna(qty):
                continue
            results.append((unit_text, float(qty)))
        return results

    def _format_unit_totals(unit_totals: List[Tuple[str, float]]) -> str:
        if not unit_totals:
            return ""
        parts = []
        for unit, qty in unit_totals:
            if pd.isna(qty) or math.isclose(float(qty), 0.0, abs_tol=1e-9):
                continue
            qty_text = _format_quantity(float(qty))
            parts.append(f"{qty_text} {_plural(unit, qty)}")
        return " ".join(parts)

    def unit_summary_text(df_subset: pd.DataFrame) -> str:
        if "unit" not in df_subset.columns or "stock_qty" not in df_subset.columns:
            return ""
        scoped = _filter_primary_warehouses(df_subset) if "warehouse" in df_subset.columns else df_subset
        totals = _collect_unit_totals(scoped)
        return _format_unit_totals(totals)

    def warehouse_summary_text(df_subset: pd.DataFrame) -> str:
        if "warehouse" not in df_subset.columns:
            return ""
        scoped = _filter_primary_warehouses(df_subset)
        if scoped.empty:
            return ""
        working = scoped.assign(_normalized_warehouse=scoped["warehouse"].map(_normalize_warehouse_name))
        texts = []
        for wh in PRIMARY_WAREHOUSES:
            subset = working[working["_normalized_warehouse"] == wh]
            if subset.empty:
                continue
            totals = _collect_unit_totals(subset)
            formatted = _format_unit_totals(totals)
            if formatted:
                texts.append(f"{wh}: {formatted}")
        return " | ".join(texts)

    def combined_unit_wh_text(df_subset: pd.DataFrame) -> str:
        """Return unit totals with optional warehouse breakdown in parentheses."""
        if "unit" not in df_subset.columns or "stock_qty" not in df_subset.columns:
            return ""
        scoped = _filter_primary_warehouses(df_subset) if "warehouse" in df_subset.columns else df_subset
        if scoped.empty:
            return ""
        totals = _collect_unit_totals(scoped)
        if "warehouse" not in scoped.columns:
            return _format_unit_totals(totals)
        working = scoped.assign(
            _normalized_warehouse=scoped["warehouse"].map(_normalize_warehouse_name),
            stock_qty=pd.to_numeric(scoped["stock_qty"], errors="coerce").fillna(0),
            unit_key=scoped["unit"].apply(_normalize_unit),
        )
        working = working[working["_normalized_warehouse"].isin(PRIMARY_WAREHOUSES)]
        if working.empty:
            return _format_unit_totals(totals)
        per_wh = (
            working.groupby(["unit_key", "_normalized_warehouse"], dropna=False)["stock_qty"]
            .sum()
            .reset_index()
        )
        wh_map = {
            unit: {row._normalized_warehouse: row.stock_qty for row in per_wh[per_wh["unit_key"] == unit].itertuples(index=False)}
            for unit, _ in totals
        }
        parts = []
        for unit, total in totals:
            if math.isclose(total, 0.0, abs_tol=1e-9):
                continue
            qty_text = _format_quantity(total)
            sub_parts = []
            for wh in PRIMARY_WAREHOUSES:
                wh_qty = wh_map.get(unit, {}).get(wh)
                if wh_qty is None or math.isclose(wh_qty, 0.0, abs_tol=1e-9):
                    continue
                sub_parts.append(f"{wh} {_format_quantity(wh_qty)}")
            detail = f" ({' + '.join(sub_parts)})" if sub_parts else ""
            parts.append(f"{qty_text} {_plural(unit, total)}{detail}")
        return " ".join(parts)

    def unit_totals_plain(df_subset: pd.DataFrame) -> str:
        """Return combined totals per unit without warehouse breakdown."""
        if "unit" not in df_subset.columns or "stock_qty" not in df_subset.columns:
            return ""
        scoped = _filter_primary_warehouses(df_subset) if "warehouse" in df_subset.columns else df_subset
        if scoped.empty:
            return ""
        totals = _collect_unit_totals(scoped)
        return _format_unit_totals(totals)

    def style_with_expiry(df_subset: pd.DataFrame):
        if "expiry_date" in df_subset.columns:
            cutoff = (pd.Timestamp.today().normalize() + pd.Timedelta(days=due_days)).date()

            def row_style(row):
                try:
                    d = pd.to_datetime(row.get("expiry_date"), errors="coerce", format="mixed")
                    if pd.notna(d):
                        d = d.date()
                        if d <= cutoff:
                            style = "background-color: rgba(255,140,0,0.22); border-left: 4px solid #ff8c00;"
                            return [style] * len(row)
                except Exception:
                    pass
                return [""] * len(row)

            try:
                return df_subset.style.apply(row_style, axis=1).format({"stock_qty": "{:,.0f}"})
            except Exception:
                try:
                    return df_subset.style.apply(row_style, axis=1)
                except Exception:
                    return df_subset

        try:
            return df_subset.style.format({"stock_qty": "{:,.0f}"})
        except Exception:
            return df_subset



    def supplier_product_summary(df_subset: pd.DataFrame, suppliers: list) -> pd.DataFrame:
        if not suppliers or "supplier" not in df_subset.columns:
            return pd.DataFrame()
        df_sup = df_subset[df_subset["supplier"].isin(suppliers)].copy()
        if df_sup.empty:
            return pd.DataFrame()

        df_sup["_base_description"] = (
            df_sup.get("description", pd.Series(dtype=str))
            .astype(str)
            .str.replace(r"\s*\([^)]*\)", "", regex=True)
            .str.strip()
        )
        df_sup["_product_code"] = (
            df_sup.get("product_code", pd.Series(dtype=str))
            .astype(str)
            .str.strip()
        )

        rows = []
        group_keys = ["_base_description", "_product_code"] if "product_code" in df_sup.columns else ["_base_description"]
        for sup_val, sup_group in df_sup.groupby("supplier", dropna=False):
            sup_name = str(sup_val).strip() if pd.notna(sup_val) and str(sup_val).strip() else "Unknown supplier"
            for key_vals, prod_group in sup_group.groupby(group_keys, dropna=False):
                if not isinstance(key_vals, tuple):
                    key_vals = (key_vals,)
                base_desc_value = key_vals[0]
                base_name = str(base_desc_value).strip() if pd.notna(base_desc_value) and str(base_desc_value).strip() else "Unnamed product"
                code_display = "-"
                if len(key_vals) > 1:
                    code_candidate = key_vals[1]
                    code_display = str(code_candidate).strip() if pd.notna(code_candidate) and str(code_candidate).strip() else "-"

                per_warehouse = {}
                if "warehouse" in prod_group.columns:
                    for wh_val, wh_group in prod_group.groupby("warehouse", dropna=False):
                        wh_name = _normalize_warehouse_name(wh_val)
                        if wh_name in PRIMARY_WAREHOUSES:
                            per_warehouse[wh_name] = unit_totals_plain(wh_group)

                rows.append({
                    "Supplier": sup_name,
                    "Product": base_name,
                    "Product Code": code_display,
                    "Savori Whse": per_warehouse.get("Savori Whse", ""),
                    "Lai Hock Whse": per_warehouse.get("Lai Hock Whse", ""),
                    "Total": unit_totals_plain(prod_group),
                })

        summary_df = pd.DataFrame(rows)
        if summary_df.empty:
            return summary_df
        summary_df = summary_df.sort_values(by=["Supplier", "Product", "Product Code"], kind="stable")
        return summary_df.reset_index(drop=True)

    def description_summary(df_subset: pd.DataFrame, descriptions: list) -> pd.DataFrame:
        if not descriptions or "description" not in df_subset.columns:
            return pd.DataFrame()
        base_series = (
            df_subset["description"].astype(str)
            .str.replace(r"\s*\([^)]*\)", "", regex=True)
            .str.strip()
        )
        working = df_subset.copy()
        working["_base_description"] = base_series
        target = working[working["_base_description"].isin(descriptions)]
        if target.empty:
            return pd.DataFrame()

        rows = []
        for desc_value, group in target.groupby("_base_description", dropna=False):
            desc_name = str(desc_value).strip() if pd.notna(desc_value) and str(desc_value).strip() else "Unnamed product"
            code_series = group.get("product_code", pd.Series(dtype=str))
            codes = sorted({str(c).strip() for c in code_series.dropna().tolist() if str(c).strip()})
            per_wh = {}
            if "warehouse" in group.columns:
                for wh_val, wh_group in group.groupby("warehouse", dropna=False):
                    wh_name = _normalize_warehouse_name(wh_val)
                    if wh_name in PRIMARY_WAREHOUSES:
                        per_wh[wh_name] = unit_totals_plain(wh_group)
            rows.append({
                "Description": desc_name,
                "Product Code(s)": " / ".join(codes) if codes else "-",
                "Savori Whse": per_wh.get("Savori Whse", ""),
                "Lai Hock Whse": per_wh.get("Lai Hock Whse", ""),
                "Total": unit_totals_plain(group),
            })
        summary_df = pd.DataFrame(rows)
        if summary_df.empty:
            return summary_df
        summary_df = summary_df.sort_values(by=["Description", "Product Code(s)"], kind="stable")
        return summary_df.reset_index(drop=True)

    def build_expiry_summary(df_subset: pd.DataFrame) -> pd.DataFrame:
        if "expiry_date" not in df_subset.columns:
            return pd.DataFrame()
        working = df_subset.copy()
        working["expiry_date"] = pd.to_datetime(working["expiry_date"], errors="coerce", format="mixed")
        if working["expiry_date"].isna().all():
            return pd.DataFrame()
        working["_expiry_date_only"] = working["expiry_date"].dt.date

        rows = []
        today = pd.Timestamp.today().normalize().date()
        for expiry_value, group in working.groupby("_expiry_date_only", dropna=False):
            if pd.isna(expiry_value):
                label = "无到期日"
                sort_key = datetime.date.max
                days_remaining = pd.NA
            else:
                label = expiry_value.isoformat()
                sort_key = expiry_value
                days_remaining = (expiry_value - today).days

            per_wh = {}
            if "warehouse" in group.columns:
                for wh_val, wh_group in group.groupby("warehouse", dropna=False):
                    wh_name = _normalize_warehouse_name(wh_val)
                    if wh_name in PRIMARY_WAREHOUSES:
                        per_wh[wh_name] = unit_totals_plain(wh_group)
            rows.append({
                "Expiry": label,
                "Savori Whse": per_wh.get("Savori Whse", ""),
                "Lai Hock Whse": per_wh.get("Lai Hock Whse", ""),
                "Total": unit_totals_plain(group),
                "Days Remaining": days_remaining,
                "_sort": sort_key,
            })

        summary_df = pd.DataFrame(rows)
        if summary_df.empty:
            return summary_df
        summary_df["Days Remaining"] = summary_df["Days Remaining"].astype("Int64")
        summary_df = summary_df.sort_values(by="_sort", kind="stable").drop(columns="_sort").reset_index(drop=True)
        return summary_df

    def style_expiry_summary(df_summary: pd.DataFrame, highlight_days: int):
        if df_summary.empty:
            return df_summary

        def _highlight(row):
            days = row.get("Days Remaining")
            if pd.isna(days):
                return [""] * len(row)
            try:
                if int(days) <= highlight_days:
                    style = "background-color: rgba(255,140,0,0.22); border-left: 4px solid #ff8c00;"
                    return [style] * len(row)
            except Exception:
                pass
            return [""] * len(row)

        try:
            return df_summary.style.apply(_highlight, axis=1)
        except Exception:
            return df_summary

    combined_txt = unit_totals_plain(filtered)
    supplier_summary_df = supplier_product_summary(filtered, selected_suppliers)
    warehouse_overview = warehouse_summary_text(filtered)

    def render_unit_summary(df_subset, title_prefix="单位汇总"):
        if "unit" in df_subset.columns and "stock_qty" in df_subset.columns:
            unit_sum = (
                df_subset.assign(stock_qty=pd.to_numeric(df_subset["stock_qty"], errors="coerce").fillna(0))
                .groupby("unit", dropna=False)["stock_qty"].sum()
                .sort_values(ascending=False)
                .rename("总量合计")
            )
            unit_cnt = (
                df_subset.groupby("unit", dropna=False)[
                    "product_code" if "product_code" in df_subset.columns else df_subset.columns[0]
                ].count()
                .rename("条目数")
            )
            summary_df = (
                pd.concat([unit_sum, unit_cnt], axis=1)
                .reset_index()
                .rename(columns={"unit": "单位"})
            )
            summary_df["_priority"] = summary_df["单位"].apply(_unit_priority)
            summary_df = summary_df.sort_values(by=["_priority", "单位"], kind="stable").drop(columns="_priority")
            st.subheader(f"{title_prefix}")
            try:
                styled_sum = summary_df.style.format({"总量合计": "{:,.0f}"})
                st.dataframe(styled_sum, use_container_width=True, hide_index=True)
            except Exception:
                st.dataframe(summary_df, use_container_width=True, hide_index=True)

    overview_tab, detail_tab = st.tabs(["Overview", "Details"])



    with overview_tab:
        col_kpi, col_totals = st.columns([1, 3])
        col_kpi.metric("Filtered Rows", f"{total_rows}")
        col_totals.markdown(
            f"<div style='font-size:1.05rem;font-weight:600;'>Totals: {combined_txt if combined_txt else 'No data'}</div>",
            unsafe_allow_html=True,
        )
        if warehouse_overview:
            st.caption(f"Warehouse split: {warehouse_overview}")

        if selected_suppliers:
            st.subheader("Supplier Summary")
            if supplier_summary_df.empty:
                st.info("No stock found for the selected supplier filter.")
            else:
                st.dataframe(supplier_summary_df, use_container_width=True, hide_index=True)

        render_unit_summary(filtered, title_prefix="单位汇总（当前筛选）")

    with detail_tab:
        def prepare_display(df_subset: pd.DataFrame):
            display_df = df_subset.copy()
            for col in ["expiry_date", "relabel_to_date"]:
                if col in display_df.columns:
                    display_df[col] = pd.to_datetime(display_df[col], errors="coerce", format="mixed").dt.date
            styled = style_with_expiry(display_df)
            return styled if hasattr(styled, "_repr_html_") else display_df

        if selected_descs:
            st.subheader("Descriptions")
            base_series_filtered = (
                filtered["description"].astype(str)
                .str.replace(r"\s*\([^)]*\)", "", regex=True)
                .str.strip()
            )
            summary_table = description_summary(filtered, selected_descs)
            if not summary_table.empty:
                st.dataframe(summary_table, use_container_width=True, hide_index=True)
            else:
                st.info("No summary available for the chosen descriptions.")
            for desc in selected_descs:
                subset_all = filtered[base_series_filtered == desc].copy()
                if subset_all.empty:
                    st.info(f"No rows for description: {desc}")
                    continue
                codes = sorted({
                    str(c).strip()
                    for c in subset_all.get("product_code", pd.Series(dtype=str)).dropna().unique().tolist()
                    if str(c).strip()
                })
                code_part = " / ".join(codes) if codes else ""
                title = (f"{code_part} - {desc}" if code_part else str(desc)).upper()
                st.markdown(f"**{title}**")
                combined_desc_txt = unit_totals_plain(subset_all)
                st.markdown(
                    f"<div style='font-size:1.05rem;font-weight:600;'>Totals: {combined_desc_txt if combined_desc_txt else 'No data'}</div>",
                    unsafe_allow_html=True,
                )
                wh_line = warehouse_summary_text(subset_all)
                if wh_line:
                    st.caption(f"Warehouse split: {wh_line}")
                expiry_summary = build_expiry_summary(subset_all)
                with st.expander("Expiry breakdown", expanded=False):
                    if expiry_summary.empty:
                        st.info("No expiry grouped data.")
                    else:
                        styled_expiry = style_expiry_summary(expiry_summary, due_days)
                        if hasattr(styled_expiry, "_repr_html_"):
                            st.dataframe(styled_expiry, use_container_width=True, hide_index=True)
                        else:
                            st.dataframe(expiry_summary, use_container_width=True, hide_index=True)
                if "warehouse" in subset_all.columns:
                    for wh, subset in subset_all.groupby("warehouse", dropna=False):
                        label = f"{wh if pd.notna(wh) else 'Unknown warehouse'}"
                        with st.expander(label, expanded=False):
                            render_unit_summary(subset, title_prefix="单位汇总（仓库）")
                            st.dataframe(prepare_display(subset), use_container_width=True, hide_index=True)
                else:
                    with st.expander("Details", expanded=False):
                        render_unit_summary(subset_all, title_prefix="单位汇总")
                        st.dataframe(prepare_display(subset_all), use_container_width=True, hide_index=True)

        else:
            render_unit_summary(filtered, title_prefix="单位汇总（全部）")
            st.divider()
            st.subheader("Warehouse breakdown")
            if "warehouse" in filtered.columns:
                for wh, subset in filtered.groupby("warehouse", dropna=False):
                    label = f"{wh if pd.notna(wh) else 'Unknown warehouse'}"
                    with st.expander(label, expanded=False):
                        render_unit_summary(subset, title_prefix="单位汇总（仓库）")
                        st.dataframe(prepare_display(subset), use_container_width=True, hide_index=True)
            else:
                st.dataframe(prepare_display(filtered), use_container_width=True, hide_index=True)
if __name__ == "__main__":
    main()
