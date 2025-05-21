import streamlit as st
import pandas as pd
import re
def show_df(df, *, table=True, **st_kwargs):
    """把展示版复制并转成纯字符串，Arrow 永远不会再报错"""
    display = df.reset_index(drop=True).copy().astype(str)  # 统一索引避免类型冲突
    if table:
        st.table(display)
    else:
        st.dataframe(display, **st_kwargs)
# ──────────────────────────────
# 页面基础
st.set_page_config("C宝库存神器", layout="wide", page_icon="📦")

st.markdown("""
<style>
div[data-baseweb="select"] div { font-size:18px !important; }
.scroll-table { max-height: 420px; overflow-y: auto; }
</style>
""", unsafe_allow_html=True)

ALL = "全部（ALL）"
NO_BRACKET = "（无备注）"

# ──────────────────────────────
# 读文件函数
@st.cache_data(show_spinner="加载库存中…")
def load_stock(file):
    df = pd.read_excel(file, header=2)

    df["Product Code"] = (
        df["Product Code"].astype(str)
        .fillna("NA")
        .replace(["nan", "NaN", "NAN"], "NA")
    )
    if "Daily Updated stocks.1" in df.columns:
        df["Daily Updated stocks"] = df["Daily Updated stocks.1"]

    # ── 拆主描述 / 括号备注 ───────────────────
    def split_desc(txt: str):
        if not isinstance(txt, str):
            return pd.Series([txt, NO_BRACKET])
        s = txt.rstrip()
        m = re.search(r"\(([^()]*)\)\s*$", s)     # 捕最后一对(...)
        if m:
            main = s[: m.start()].rstrip(" -")
            bracket = m.group(1).strip()
        else:
            main, bracket = s, NO_BRACKET
        return pd.Series([main, bracket])

    # 根据真实列名选择
    desc_col = "Description" if "Description" in df.columns else "Desciption"
    df[["DescMain", "Bracket"]] = df[desc_col].apply(split_desc)
    df.drop(columns=["Relabel to date"], errors="ignore", inplace=True)

    df["Expiry_dt"] = pd.to_datetime(df["Expiry Date"], errors="coerce")

    df["Expiry Date"] = df["Expiry_dt"].dt.strftime("%d-%b-%Y")
    return df



@st.cache_data(show_spinner="加载销售汇总中…")
def load_sales(file):
    df = pd.read_excel(file, header=3)

    df["Year"] = (
        pd.to_numeric(df["Year"], errors="coerce")
          .astype("Int64").astype(str).str.replace("nan", "")
    )

    def clean_code(x):
        if pd.isna(x): return ""
        if isinstance(x, (int, float)) and float(x).is_integer():
            return str(int(x))
        return str(x).strip()

    df["Product Code"] = df["Product Code"].apply(clean_code)

    df["Month"] = df["Month"].astype(str)
    df["Date"]  = pd.to_datetime(df["Date"], errors="coerce").dt.strftime("%Y/%m/%d")
    return df

# ──────────────────────────────
# 侧边栏上传
st.sidebar.header("📁 上传数据")
stock_file = st.sidebar.file_uploader("Stock Report", type=["xlsx"])
sales_file = st.sidebar.file_uploader("Sales Summary", type=["xlsx"])

if stock_file: st.sidebar.success(f"库存已上传：{stock_file.name}")
if sales_file: st.sidebar.success(f"销售已上传：{sales_file.name}")

tab_stock, tab_sales = st.tabs(["库存查询", "销售分析"])

# ================= Tab 1 =================
with tab_stock:
    st.subheader("🎯 三级筛选查库存")

    if not stock_file:
        st.info("请先上传库存文件（Stock Report）")
    else:
        df_stock = load_stock(stock_file)

        # ── Step-1  产品名称（去重，大小写无关，跳过非字符串） ──
        clean_names = (
            df_stock["DescMain"]
            .dropna()
            .apply(lambda x: str(x).strip())
        )
        desc_opts = sorted(
            {name.lower(): name for name in clean_names}.values(),
            key=str.lower,
        )
        desc_sel = st.selectbox("Step 1：产品名称", desc_opts, key="stk_desc")

        # ── Step-2  产品代码 ──
        code_pool = (
            df_stock.loc[df_stock["DescMain"] == desc_sel, "Product Code"]
            .dropna()
            .unique()
        )
        code_sel = st.selectbox(
            "Step 2：产品代码", [ALL] + sorted(code_pool), key="stk_code"
        )

        # ── Step-3  备注 / 子标签（括号内容） ──
        cond = df_stock["DescMain"] == desc_sel
        if code_sel != ALL:
            cond &= df_stock["Product Code"] == code_sel

        bracket_pool = df_stock.loc[cond, "Bracket"].unique()
        bracket_display = ["（无备注）" if x == NO_BRACKET else x for x in bracket_pool]
        bracket_sel_disp = st.selectbox(
            "Step 3：备注/子标签", [ALL] + bracket_display, key="stk_bracket"
        )

        filt = cond.copy()
        if bracket_sel_disp != ALL:
            bracket_sel = NO_BRACKET if bracket_sel_disp == "（无备注）" else bracket_sel_disp
            filt &= df_stock["Bracket"] == bracket_sel

        result = df_stock[filt]

        # ── 表格 or 提示 ──
        if result.empty:
            st.warning("当前筛选无任何库存")
        else:
            grp = ["Expiry_dt", "Expiry Date", "Pack Size"]
            df_ctn = (
                result[result["Unit"] == "CTN"]
                .groupby(grp, as_index=False)["Daily Updated stocks"]
                .sum()
                .rename(columns={"Daily Updated stocks": "CTN"})
            )
            df_pkt = (
                result[result["Unit"] == "PKTS"]
                .groupby(grp, as_index=False)["Daily Updated stocks"]
                .sum()
                .rename(columns={"Daily Updated stocks": "PKTS"})
            )
            batch = (
                pd.merge(df_ctn, df_pkt, how="outer", on=grp)
                .fillna(0)
                .astype({"CTN": int, "PKTS": int})
            )

            total = pd.Series({
                "Expiry_dt": pd.NaT,
                "Expiry Date": "总计",
                "Pack Size": "",
                "CTN": batch["CTN"].sum(),
                "PKTS": batch["PKTS"].sum(),
            }, name="总计")
            batch = pd.concat([batch, total.to_frame().T])

            asc = st.radio(
                "Expiry Date 排序", ["升序", "降序"], key="stk_sort_dir", horizontal=True
            ) == "升序"

            data_rows = batch.drop(index="总计")
            sorted_rows = data_rows.sort_values("Expiry_dt", ascending=asc)
            summary_row = batch.loc[["总计"]]
            batch_sorted = pd.concat([sorted_rows, summary_row])

            show_df(batch_sorted.drop(columns="Expiry_dt"))


        # ── 全量表 ──
        st.write("## 📋 全量库存表")
        st.markdown('<div class="scroll-table">', unsafe_allow_html=True)
        show_df(df_stock, table=False, hide_index=True, use_container_width=True)
        st.markdown("</div>", unsafe_allow_html=True)


# ================= Tab 2 =================
with tab_sales:
    st.subheader("🔍 六级筛选销售（新顺序 v2）")

    if not sales_file:
        st.info("请先上传销售汇总文件（Sales Summary）")
    else:
        df = load_sales(sales_file)          # 已缓存
        ss = st.session_state               # 便捷

        # -------- 通用过滤函数 --------
        def filt(d, desc=ALL, code=ALL, year=ALL, month=ALL,
                 cust=ALL, outlet=ALL):
            m = pd.Series(True, index=d.index)
            if desc   != ALL: m &= d["Product Description"].str.lower() == desc.lower()
            if code   != ALL: m &= d["Product Code"] == code
            if year   != ALL: m &= d["Year"] == year
            if month  != ALL: m &= d["Month"] == month
            if cust   != ALL: m &= d["Customer"] == cust
            if outlet != ALL: m &= d["Outlet"] == outlet
            return d[m]

        # -------- Step-1 产品名称 --------
        def unique_ignore_case(series):
            seen, res = set(), []
            for x in series.dropna():
                low = x.lower()
                if low not in seen:
                    res.append(x)
                    seen.add(low)
            return res

        col_l, col_r = st.columns(2)

        with col_l:
            desc_opts = sorted(
                unique_ignore_case(df["Product Description"]),
                key=str.lower
            )
            old_desc = ss.get("desc", ALL)
            if old_desc.lower() not in [x.lower() for x in desc_opts] and old_desc != ALL:
                desc_opts.append(old_desc)
            desc_sel = st.selectbox(
                "Step 1：产品名称", [ALL] + desc_opts, key="desc"
            )

            df_d1 = filt(df, desc=desc_sel)

            code_opts = sorted([c for c in df_d1["Product Code"].unique() if c])
            old_code = ss.get("code", ALL)
            if old_code not in code_opts and old_code != ALL:
                code_opts.append(old_code)
            code_sel = st.selectbox(
                "Step 2：产品代码", [ALL] + code_opts, key="code"
            )

        df_d2 = filt(df, desc=desc_sel, code=code_sel)
        with col_r:
            year_opts = sorted([y for y in df_d2["Year"].unique() if y])
            old_year = ss.get("year", ALL)
            if old_year not in year_opts and old_year != ALL:
                year_opts.append(old_year)
            year_sel = st.selectbox(
                "Step 3：年份", [ALL] + year_opts, key="year"
            )

            df_d3 = filt(df_d2, year=year_sel)

            MONTH_ORDER = [
                "Jan", "Feb", "Mar", "Apr", "May", "Jun",
                "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"
            ]
            month_opts = [m for m in MONTH_ORDER if m in df_d3["Month"].unique()]
            old_month = ss.get("month", ALL)
            if old_month not in month_opts and old_month != ALL:
                month_opts.append(old_month)
            month_sel = st.selectbox(
                "Step 4：月份", [ALL] + month_opts, key="month"
            )

        # -------- Step-3 / 4  年份 & 月份 --------
        col_year, col_month = st.columns(2)

        with col_year:
            year_opts = sorted([y for y in df_d2["Year"].unique() if y])
            old_year  = ss.get("year", ALL)
            if old_year not in year_opts and old_year != ALL:
                year_opts.append(old_year)
            year_sel = st.selectbox("Step 3：年份",
                                    [ALL] + year_opts, key="year")

        df_d3 = filt(df_d2, year=year_sel)

        with col_month:
            MONTH_ORDER = ["Jan","Feb","Mar","Apr","May","Jun",
                           "Jul","Aug","Sep","Oct","Nov","Dec"]
            month_opts = [m for m in MONTH_ORDER if m in df_d3["Month"].unique()]
            old_month  = ss.get("month", ALL)
            if old_month not in month_opts and old_month != ALL:
                month_opts.append(old_month)
            month_sel = st.selectbox("Step 4：月份",
                                     [ALL] + month_opts, key="month")


        df_d4 = filt(df_d3, month=month_sel)

        # -------- Step-5 / 6  客户 & 门店 --------
        col_cust, col_out = st.columns(2)

        with col_cust:
            cust_opts = sorted(df_d4["Customer"].dropna().unique())
            old_cust = ss.get("cust", ALL)
            if old_cust not in cust_opts and old_cust != ALL:
                cust_opts.append(old_cust)
            cust_sel = st.selectbox(
                "Step 5：客户", [ALL] + cust_opts, key="cust"
            )

        df_d5 = filt(df_d4, cust=cust_sel)

        with col_out:
            outlet_opts = sorted(df_d5["Outlet"].dropna().unique())
            old_outlet = ss.get("outlet", ALL)
            if old_outlet not in outlet_opts and old_outlet != ALL:
                outlet_opts.append(old_outlet)
            outlet_sel = st.selectbox(
                "Step 6：门店", [ALL] + outlet_opts, key="outlet"
            )

        final_df = filt(df_d5, outlet=outlet_sel)

        # -------- 结果区域 --------
        if desc_sel == ALL:
            st.info("👉 先选“产品名称”再查看数据")
        elif final_df.empty:
            st.warning("当前筛选无数据")
        else:
            tbl = (
                final_df[["Customer", "Outlet", "Date", "Qty in Ctns", "Qty in Pcs"]]
                .groupby(["Customer", "Outlet", "Date"], as_index=False)
                .sum()
                .rename(columns={"Qty in Ctns": "CTN", "Qty in Pcs": "PCS"})
                .astype({"CTN": int, "PCS": int})
                .sort_values("Date")
            )
            tbl.loc["总计"] = ["", "", "总计", tbl["CTN"].sum(), tbl["PCS"].sum()]

            show_df(tbl)


        with col_cust:
            cust_opts = sorted(df_d4["Customer"].dropna().unique())
            old_cust  = ss.get("cust", ALL)
            if old_cust not in cust_opts and old_cust != ALL:
                cust_opts.append(old_cust)
            cust_sel = st.selectbox("Step 5：客户",
                                    [ALL] + cust_opts, key="cust")

        df_d5 = filt(df_d4, cust=cust_sel)

        with col_out:
            outlet_opts = sorted(df_d5["Outlet"].dropna().unique())
            old_outlet  = ss.get("outlet", ALL)
            if old_outlet not in outlet_opts and old_outlet != ALL:
                outlet_opts.append(old_outlet)
            outlet_sel = st.selectbox("Step 6：门店",
                                      [ALL] + outlet_opts, key="outlet")

        final_df = filt(df_d5, outlet=outlet_sel)

        # -------- 结果区域 --------
        if desc_sel == ALL:
            st.info("👉 先选“产品名称”再查看数据")
        elif final_df.empty:
            st.warning("当前筛选无数据")
        else:
            tbl = (final_df[["Customer","Outlet","Date",
                             "Qty in Ctns","Qty in Pcs"]]
                   .groupby(["Customer","Outlet","Date"], as_index=False).sum()
                   .rename(columns={"Qty in Ctns":"CTN","Qty in Pcs":"PCS"})
                   .astype({"CTN":int,"PCS":int}))
            tbl.loc["总计"] = ["", "", "总计",
                       tbl["CTN"].sum(), tbl["PCS"].sum()]

            # -------- 排序选项 --------
            sort_options = [c for c in tbl.columns if c != ""]
            default_idx = sort_options.index("Date") if "Date" in sort_options else 0
            sort_col2 = st.selectbox(
                "排序字段",
                sort_options,
                key="sale_sort_col",
                index=default_idx,
            )
            asc2 = st.radio(
                "排序方式", ["升序", "降序"], key="sale_sort_dir", horizontal=True
            ) == "升序"

            data_rows2 = tbl.drop(index="总计")
            summary_row2 = tbl.loc[["总计"]]
            sorted_rows2 = data_rows2.sort_values(sort_col2, ascending=asc2)
            tbl_sorted = pd.concat([sorted_rows2, summary_row2])

            display_tbl = tbl_sorted.astype(str)
            show_df(tbl_sorted)