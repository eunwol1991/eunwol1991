@@ -113,50 +113,55 @@ with tab_stock:
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

@@ -218,28 +223,108 @@ with tab_sales:
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

    # -------- Step-2 产品代码 --------
        code_opts = sorted([c for c in df_d1["Product Code"].unique() if c])
        old_code = ss.get("code", ALL)
        if old_code not in code_opts and old_code != ALL:
            code_opts.append(old_code)
        code_sel = st.selectbox("Step 2：产品代码",
                                [ALL] + code_opts, key="code")

        df_d2 = filt(df, desc=desc_sel, code=code_sel)

        # -------- Step-3 / 4  年份 & 月份 --------
        col_year, col_month = st.columns(2)

        with col_year:
            year_opts = sorted([y for y in df_d2["Year"].unique() if y])
            old_year = ss.get("sale_year", ALL)
            if old_year not in year_opts and old_year != ALL:
                year_opts.append(old_year)
            year_sel = st.selectbox("Step 3：年份",
                                     [ALL] + year_opts, key="sale_year")

        df_d3 = filt(df_d2, year=year_sel)

        with col_month:
            MONTH_ORDER = ["Jan","Feb","Mar","Apr","May","Jun",
                           "Jul","Aug","Sep","Oct","Nov","Dec"]
            month_opts = [m for m in MONTH_ORDER if m in df_d3["Month"].unique()]
            old_month = ss.get("sale_month", ALL)
            if old_month not in month_opts and old_month != ALL:
                month_opts.append(old_month)
            month_sel = st.selectbox("Step 4：月份",
                                     [ALL] + month_opts, key="sale_month")

        df_d4 = filt(df_d3, month=month_sel)

        # -------- Step-5 / 6  客户 & 门店 --------
        col_cust, col_out = st.columns(2)

        with col_cust:
            cust_opts = sorted(df_d4["Customer"].dropna().unique())
            old_cust = ss.get("sale_cust", ALL)
            if old_cust not in cust_opts and old_cust != ALL:
                cust_opts.append(old_cust)
            cust_sel = st.selectbox("Step 5：客户",
                                    [ALL] + cust_opts, key="sale_cust")

        df_d5 = filt(df_d4, cust=cust_sel)

        with col_out:
            outlet_opts = sorted(df_d5["Outlet"].dropna().unique())
            old_outlet = ss.get("sale_outlet", ALL)
            if old_outlet not in outlet_opts and old_outlet != ALL:
                outlet_opts.append(old_outlet)
            outlet_sel = st.selectbox("Step 6：门店",
                                      [ALL] + outlet_opts, key="sale_outlet")

        final_df = filt(df_d5, outlet=outlet_sel)

        # -------- 结果区域 --------
        if desc_sel == ALL:
            st.info("👉 先选“产品名称”再查看数据")
        elif final_df.empty:
            st.warning("当前筛选无数据")
        else:
            tbl = (
                final_df[["Customer", "Outlet", "Date",
                          "Qty in Ctns", "Qty in Pcs"]]
                .groupby(["Customer", "Outlet", "Date"], as_index=False)
                .sum()
                .rename(columns={"Qty in Ctns": "CTN", "Qty in Pcs": "PCS"})
                .astype({"CTN": int, "PCS": int})
            )
            tbl.loc["总计"] = ["", "", "总计",
                           tbl["CTN"].sum(), tbl["PCS"].sum()]

            # -------- 排序选项 --------
            sort_col2 = st.selectbox(
                "排序字段",
                [c for c in tbl.columns if c != ""],
                key="sale_sort_col",
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