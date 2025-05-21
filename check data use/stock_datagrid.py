import streamlit as st
import pandas as pd
import re
def show_df(df, *, table=True, **st_kwargs):
    """把展示版复制并转成纯字符串，Arrow 永远不会再报错"""
    display = df.copy().astype(str)          # ← 加 .copy() 避免改到原表
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
                    res.append(x.strip())
                    seen.add(low)
            return res

        desc_opts = sorted(unique_ignore_case(df["Product Description"]),
                           key=str.lower)
        old_desc = ss.get("desc", ALL)
        if old_desc.lower() not in [x.lower() for x in desc_opts] and old_desc != ALL:
            desc_opts.append(old_desc)
        desc_sel = st.selectbox("Step 1：产品名称",
                                [ALL] + desc_opts, key="desc")

        df_d1 = filt(df, desc=desc_sel)

        code_opts = sorted([c for c in df_d1["Product Code"].unique() if c])
        old_code = ss.get("code", ALL)
        if old_code not in code_opts and old_code != ALL:
            code_opts.append(old_code)
        code_sel = st.selectbox("Step 2：产品代码",
                                [ALL] + code_opts, key="code")

        df_d2 = filt(df, desc=desc_sel, code=code_sel)





