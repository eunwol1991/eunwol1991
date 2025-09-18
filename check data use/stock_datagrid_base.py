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
from typing import Optional, Tuple


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
        combined = pd.concat(frames, ignore_index=True, sort=False)
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
    返回：筛选后的 DataFrame、所选 Description 列表、到期高亮天数。
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
            _ensure_multiselect_key("f_wh", ordered, ordered)
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

    return work, sel_desc, due_days


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
    filtered, selected_descs, due_days = apply_filters_v2(df_display)

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

    def unit_summary_text(df_subset: pd.DataFrame) -> str:
        if "unit" not in df_subset.columns or "stock_qty" not in df_subset.columns:
            return ""
        d = df_subset.assign(
            stock_qty=pd.to_numeric(df_subset["stock_qty"], errors="coerce").fillna(0),
            unit_key=df_subset["unit"].apply(_normalize_unit),
        )
        s = d.groupby("unit_key", dropna=False)["stock_qty"].sum()
        # Order by amount desc; show unit with value > 0 or all if all zeros
        s = s.sort_values(ascending=False)
        total_sum = s.sum()
        parts = [
            f"{int(v)} {_plural(u, v)}"
            for u, v in s.items()
            if pd.notna(u) and u != "" and (v != 0 or total_sum == 0)
        ]
        return " · ".join(parts)

    def warehouse_summary_text(df_subset: pd.DataFrame) -> str:
        if "warehouse" not in df_subset.columns:
            return ""
        order = ["Savori Whse", "Lai Hock Whse"]
        texts = []
        d = df_subset.assign(
            stock_qty=pd.to_numeric(df_subset["stock_qty"], errors="coerce").fillna(0),
            unit_key=df_subset["unit"].apply(_normalize_unit),
        )
        if d.empty:
            return ""
        grouped = d.groupby(["warehouse", "unit_key"], dropna=False)["stock_qty"].sum().reset_index()
        # Build per-warehouse text in order
        wh_names = list(grouped["warehouse"].dropna().unique())
        ordered_wh = [w for w in order if w in wh_names] + [w for w in wh_names if w not in order]
        for wh in ordered_wh:
            sub = grouped[grouped["warehouse"] == wh].sort_values("stock_qty", ascending=False)
            parts = [f"{int(v)} {_plural(u, v)}" for _, u, v in sub.itertuples(index=False) if u and v != 0]
            name = str(wh) if pd.notna(wh) else "未知仓库"
            if parts:
                texts.append(f"{name}: {' · '.join(parts)}")
        return " | ".join(texts)

    def combined_unit_wh_text(df_subset: pd.DataFrame) -> str:
        """按单位显示合计并在括号内给出分仓加总。"""
        if "unit" not in df_subset.columns or "stock_qty" not in df_subset.columns:
            return ""
        d = df_subset.assign(
            stock_qty=pd.to_numeric(df_subset["stock_qty"], errors="coerce").fillna(0),
            unit_key=df_subset["unit"].apply(_normalize_unit)
        )
        if "warehouse" in d.columns:
            piv = d.groupby(["unit_key", "warehouse"], dropna=False)["stock_qty"].sum().reset_index()
        else:
            piv = d.groupby(["unit_key"], dropna=False)["stock_qty"].sum().reset_index().assign(warehouse="")

        # total per unit
        totals = piv.groupby("unit_key", dropna=False)["stock_qty"].sum().sort_values(ascending=False)
        parts = []
        order_wh = ["Savori Whse", "Lai Hock Whse"]
        for unit, total in totals.items():
            if not unit:
                continue
            # breakdown per warehouse for this unit
            sub = piv[piv["unit_key"] == unit].set_index("warehouse")["stock_qty"]
            items = []
            for wh in order_wh:
                if wh in sub and sub[wh] != 0:
                    items.append(f"{wh} {int(sub[wh])}")
            for wh, val in sub.items():
                if wh not in order_wh and val != 0:
                    name = str(wh) if pd.notna(wh) else "未知仓库"
                    items.append(f"{name} {int(val)}")
            detail = f" ({' + '.join(items)})" if items else ""
            parts.append(f"{int(total)} {_plural(unit, total)}{detail}")
        return " · ".join(parts)

    def unit_totals_plain(df_subset: pd.DataFrame) -> str:
        """仅显示按单位的合计（不带分仓括号），如："448 ctns 2 pkts"。"""
        if "unit" not in df_subset.columns or "stock_qty" not in df_subset.columns:
            return ""
        d = df_subset.assign(
            stock_qty=pd.to_numeric(df_subset["stock_qty"], errors="coerce").fillna(0),
            unit_key=df_subset["unit"].apply(_normalize_unit)
        )
        s = d.groupby("unit_key", dropna=False)["stock_qty"].sum().sort_values(ascending=False)
        parts = [f"{int(v)} {_plural(u, v)}" for u, v in s.items() if pd.notna(u) and u != ""]
        return " ".join(parts)

    c1, c2, _c3 = st.columns([1, 5, 0.1])
    c1.metric("筛选后行数", f"{total_rows}")
    combined_txt = unit_totals_plain(filtered)
    c2.markdown(
        f"<div style='font-size:1.05rem;font-weight:600;'>加总：{combined_txt if combined_txt else '无数据'}</div>",
        unsafe_allow_html=True,
    )

    # Unit-wise summary per selected description (or overall when none selected)
    def render_unit_summary(df_subset, title_prefix="按单位汇总"):
        if "unit" in df_subset.columns and "stock_qty" in df_subset.columns:
            unit_sum = (
                df_subset.assign(stock_qty=pd.to_numeric(df_subset["stock_qty"], errors="coerce").fillna(0))
                .groupby("unit", dropna=False)["stock_qty"].sum()
                .sort_values(ascending=False)
                .rename("数量合计")
            )
            unit_cnt = (
                df_subset.groupby("unit", dropna=False)["product_code" if "product_code" in df_subset.columns else df_subset.columns[0]].count()
                .rename("条目数")
            )
            summary_df = (
                pd.concat([unit_sum, unit_cnt], axis=1)
                .reset_index()
                .rename(columns={"unit": "单位"})
            )
            st.subheader(f"{title_prefix}")
            # 数量合计显示为整数格式
            try:
                styled_sum = summary_df.style.format({"数量合计": "{:,.0f}"})
                st.dataframe(styled_sum, use_container_width=True, hide_index=True)
            except Exception:
                st.dataframe(summary_df, use_container_width=True, hide_index=True)

    if selected_descs:
        st.divider()
        st.subheader("按 Description 分组展示")
        # precompute styling helper
        def style_rows(df):
            if "expiry_date" in df.columns:
                today = pd.Timestamp.today().normalize().date()
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
                # 同时将 stock_qty 显示为整数
                try:
                    return df.style.apply(row_style, axis=1).format({"stock_qty": "{:,.0f}"})
                except Exception:
                    return df.style.apply(row_style, axis=1)
            return df

        for desc in selected_descs:
            base_series_filtered = filtered["description"].astype(str).str.replace(r"\s*\([^)]*\)", "", regex=True).str.strip()
            subset_all = filtered[base_series_filtered == desc].copy()
            # Uppercase title: CODE + DESCRIPTION
            codes = sorted({str(c) for c in subset_all.get("product_code", pd.Series(dtype=str)).dropna().unique().tolist()})
            code_part = " / ".join(codes) if codes else ""
            title = (f"{code_part} - {desc}" if code_part else str(desc)).upper()
            st.markdown(f"**{title}**")

            # Compact totals: combined math sum (plain) + warehouse breakdown on next line
            combined_txt = unit_totals_plain(subset_all)
            st.markdown(
                f"<div style='font-size:1.05rem;font-weight:600;'>加总：{combined_txt if combined_txt else '无数据'}</div>",
                unsafe_allow_html=True,
            )
            wh_line = warehouse_summary_text(subset_all)
            if wh_line:
                st.markdown(
                    f"<div style='font-size:1.05rem;font-weight:600;'>分仓：{wh_line}</div>",
                    unsafe_allow_html=True,
                )
            # Per-warehouse sections (collapsible)
            if "warehouse" in subset_all.columns:
                for wh, subset in subset_all.groupby("warehouse", dropna=False):
                    label = f"{wh if pd.notna(wh) else '未知仓库'}"
                    with st.expander(label, expanded=False):
                        render_unit_summary(subset, title_prefix="按单位汇总")

                        # Format dates
                        for col in ["expiry_date", "relabel_to_date"]:
                            if col in subset.columns:
                                subset[col] = pd.to_datetime(subset[col], errors="coerce", format="mixed").dt.date

                        styled = style_rows(subset)
                        st.dataframe(styled if hasattr(styled, "_repr_html_") else subset, use_container_width=True, hide_index=True)
            else:
                with st.expander("详情", expanded=False):
                    render_unit_summary(subset_all, title_prefix="按单位汇总")
                    for col in ["expiry_date", "relabel_to_date"]:
                        if col in subset_all.columns:
                            subset_all[col] = pd.to_datetime(subset_all[col], errors="coerce", format="mixed").dt.date
                    styled = style_rows(subset_all)
                    st.dataframe(styled if hasattr(styled, "_repr_html_") else subset_all, use_container_width=True, hide_index=True)
    else:
        # overall summary when no Description selection
        render_unit_summary(filtered, title_prefix="按单位汇总（全部）")

    # When descriptions were selected, per-section tables already shown above
    # Otherwise show the overall filtered table here
    if not selected_descs:
        # Per-warehouse overall sections
        st.divider()
        st.subheader("按仓库分组展示（全部）")

        def style_rows(df):
            if "expiry_date" in df.columns:
                today = pd.Timestamp.today().normalize().date()
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

                # 同时将 stock_qty 显示为整数
                try:
                    return df.style.apply(row_style, axis=1).format({"stock_qty": "{:,.0f}"})
                except Exception:
                    return df.style.apply(row_style, axis=1)
            return df

        if "warehouse" in filtered.columns:
            for wh, subset in filtered.groupby("warehouse", dropna=False):
                label = f"{wh if pd.notna(wh) else '未知仓库'}"
                with st.expander(label, expanded=False):
                    render_unit_summary(subset, title_prefix="按单位汇总")
                    display_df = subset.copy()
                    for col in ["expiry_date", "relabel_to_date"]:
                        if col in display_df.columns:
                            display_df[col] = pd.to_datetime(display_df[col], errors="coerce", format="mixed").dt.date
                    styled = style_rows(display_df)
                    st.dataframe(styled if hasattr(styled, "_repr_html_") else display_df, use_container_width=True, hide_index=True)
        else:
            # fallback single table
            display_df = filtered.copy()
            for col in ["expiry_date", "relabel_to_date"]:
                if col in display_df.columns:
                    display_df[col] = pd.to_datetime(display_df[col], errors="coerce", format="mixed").dt.date
            styled = style_rows(display_df)
            st.dataframe(styled if hasattr(styled, "_repr_html_") else display_df, use_container_width=True, hide_index=True)


if __name__ == "__main__":
    main()
