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
            .apply(lambda x: str(x).strip())           # 全转 str，去首尾空格
        )
        desc_opts = sorted(
            {name.lower(): name for name in clean_names}.values(),  # lower 去重
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
                .sort_values("Expiry_dt")
                .drop(columns="Expiry_dt")
            )
            batch.loc["总计"] = ["", "总计", batch["CTN"].sum(), batch["PKTS"].sum()]
            # -------- 排序选项 --------
            sort_col = st.selectbox(
                "排序字段",
                [c for c in batch.columns if c != ""],
                key="stk_sort_col",
            )
            asc = st.radio(
                "排序方式", ["升序", "降序"], key="stk_sort_dir", horizontal=True
            ) == "升序"

            data_rows = batch.drop(index="总计")
            summary_row = batch.loc[["总计"]]
            sorted_rows = data_rows.sort_values(sort_col, ascending=asc)
            batch_sorted = pd.concat([sorted_rows, summary_row])

            display_batch = batch_sorted.astype(str)
            show_df(batch_sorted)


        # ── 全量表 ──
        st.write("## 📋 全量库存表")
        display_stock = df_stock.astype(str)     # ← 新增
        st.markdown('<div class="scroll-table">', unsafe_allow_html=True)
        show_df(df_stock, table=False, hide_index=True, use_container_width=True)          # 用 display_batch 展示
        st.markdown("</div>", unsafe_allow_html=True)


# ================= Tab 2 =================
with tab_sales:
    st.subheader("🔍 六级筛选销售（新顺序 v2）")

    if not sales_file:
        st.info("请先上传销售汇总文件（Sales Summary）")
    else:



