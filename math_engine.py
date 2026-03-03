import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import matplotlib
from pathlib import Path
matplotlib.use("Agg")


def run_calculations(df, style_rows_with_ERR):
    metrics = ["Market Value","stock_net","bond_net","united_states_equity_exposure","real_estate_exposure","large_market_cap_exposure_generic",
            "medium_market_cap_exposure_generic","small_market_cap_exposure_generic","other_net","cash_net"]

    df[metrics] = df[metrics].apply(pd.to_numeric, errors="coerce")

      # --- NUMERIC CONVERSION ---
    exclude_cols = [
        "Q/NQ","Account","Symbol","broad_asset_class","detailed_security_type",
        "category_name","index_fund","equity_style","fixed_income_style",
        "Name","fund_family","share_class","hq_country"
    ]

    num_cols = df.columns.difference(exclude_cols)
    df[num_cols] = df[num_cols].apply(pd.to_numeric, errors="coerce")




    df["asset_class"] = "Allocation"

    # --- Country lists ---
    developed_countries = ["Canada","Australia","United Kingdom","Germany","France","Italy","Spain","Netherlands","Belgium","Austria","Switzerland","Sweden","Denmark",
                            "Norway","Finland","Ireland","Luxembourg","Iceland","Portugal","Greece","Japan","South Korea","Singapore","Hong Kong","New Zealand","Israel",
                            "Malta","Monaco","Liechtenstein","British Virgin Islands","Cayman Islands","Gibraltar","U.S. Virgin Islands","Puerto Rico","Guam"]

    emerging_countries = ["China","Indonesia","Brazil","Philippines","Bermuda","Turkey","Taiwan","South Africa","Mexico","Mauritius","Malaysia",
                        "Dominican Republic","Ukraine","Azerbaijan", "United Arab Emirates","Thailand","Bangladesh","Chile","Poland","India","Uruguay",
                        "Mongolia","Cyprus","Panama","Colombia","Peru","Argentina","Faroe Islands", "Kazakhstan","Lebanon","Oman","Papua New Guinea",
                        "Morocco","Qatar","Bahamas","Egypt", "Costa Rica","Romania","Ghana","Cambodia","Hungary","Nigeria","Liberia","Vietnam",
                        "Tanzania","Pakistan","Jordan","Anguilla","Latvia","Armenia","Montenegro","Belize","Venezuela","Barbados","Bulgaria","Jamaica",
                        "Namibia","Kenya","El Salvador","Zambia"]

    # --- International equities ---
    df.loc[(df["detailed_security_type"] == "stock") & (df["hq_country"].isin(emerging_countries)),"asset_class"] = "Emerging Markets"

    df.loc[(df["detailed_security_type"] == "stock") &(df["hq_country"].isin(developed_countries)),"asset_class"] = "International Developed"

    # --- U.S. equity size classification ---
    df.loc[(df["detailed_security_type"] == "stock") & (df["hq_country"] == "United States") &(df["equity_style"].str.contains("Large", na=False)), "asset_class"] = "Large Cap"

    df.loc[(df["detailed_security_type"] == "stock") & (df["hq_country"] == "United States") &(df["equity_style"].str.contains("Mid", na=False)), "asset_class"] = "Mid Cap"

    df.loc[(df["detailed_security_type"] == "stock") &(df["hq_country"] == "United States") & (df["equity_style"].str.contains("Small", na=False)), "asset_class"] = "Small Cap"


    # ============================================================
    # 1. BROAD ASSET CLASS — HIGH-LEVEL DIRECT MAPPINGS
    # ============================================================

    df.loc[df["broad_asset_class"] == "Money Market", "asset_class"] = "Cash"

    df.loc[df["broad_asset_class"].isin(["Taxable Bond", "Municipal Bond"]), "asset_class"] = "Fixed Income"

    df.loc[df["broad_asset_class"] == "Allocation", "asset_class"] = "Allocation"

    df.loc[df["broad_asset_class"] == "Commodities", "asset_class"] = "Commodities"

    df.loc[df["broad_asset_class"] == "Miscellaneous", "asset_class"] = "Allocation"

    df.loc[df["broad_asset_class"].isin(["Nontraditional Equity", "Other", "Alternative"]), "asset_class"] = "Other"


    # ============================================================
    # 2. U.S. EQUITY — STYLE-BASED SPLIT
    # ============================================================

    us_equity = df["broad_asset_class"].isin(["U.S. Equity", "U.S. Small/Mid Cap Equity","US Equity"])

    df.loc[us_equity & (df["equity_style"].str.contains("Large", na=False) | df["category_name"].str.contains("Large", na=False)), "asset_class"] = "Large Cap"
    df.loc[us_equity & (df["equity_style"].str.contains("Mid", na=False) | df["category_name"].str.contains("Mid", na=False)), "asset_class"] = "Mid Cap"
    df.loc[us_equity & (df["equity_style"].str.contains("Small", na=False) | df["category_name"].str.contains("Small", na=False)), "asset_class"] = "Small Cap"



    # ============================================================
    # 3. SECTOR EQUITY — CATEGORY-DRIVEN LOGIC
    # ============================================================

    sector = df["broad_asset_class"] == "Sector Equity"

    allocation_categories = ["Global Small/Mid Stock","Global Large-Stock Growth","Global Equity","Moderate Allocation","Global Large-Stock Blend","Miscellaneous Allocation",
        "Tactical Allocation","Foreign Large Blend","Miscellaneous Region",  "Trading--Leveraged Equity","Global Moderate Allocation", "Trading--Miscellaneous"]

    commodities_categories = ["Equity Precious Metals","Precious Metals Equity","Natural Resources", "Natural Resources Equity","Digital Assets",
                            "Equity Digital Assets","Commodities Broad Basket","Commodity","Commodities Focused"]

    em_categories = ["China Region","India Equity","Diversified Emerging Mkts"]

    other_categories = ["Long-Short Equity","Derivative Income","Private Multi-Asset","Venture Capital","Private Equity","Direct Infrastructure",
                        "Private Debt - General","Alternative Equity Focused","Alternative Private Equity"]

    real_estate_categories = ["Global Real Estate","Real Estate","Real Estate Equity","Direct Real Estate"]

    df.loc[sector & df["category_name"].isin(allocation_categories), "asset_class"] = "Allocation"
    df.loc[sector & df["category_name"].isin(commodities_categories), "asset_class"] = "Commodities"
    df.loc[sector & df["category_name"].isin(em_categories), "asset_class"] = "Emerging Markets"
    df.loc[sector & (df["category_name"] == "Europe Stock"), "asset_class"] = "International Developed"
    df.loc[sector & df["category_name"].isin(other_categories), "asset_class"] = "Other"
    df.loc[sector & df["category_name"].isin(real_estate_categories), "asset_class"] = "Real Estate"


    # Style fallback inside Sector Equity
    df.loc[sector & df["equity_style"].str.contains("Large", na=False), "asset_class"] = "Large Cap"
    df.loc[sector & df["equity_style"].str.contains("Mid",   na=False), "asset_class"] = "Mid Cap"
    df.loc[sector & df["equity_style"].str.contains("Small", na=False), "asset_class"] = "Small Cap"


    # ============================================================
    # 4. INTERNATIONAL EQUITY — CATEGORY-DRIVEN
    # ============================================================

    intl = df["broad_asset_class"] == "International Equity"

    intl_allocation = [
        "Foreign Large Growth","Foreign Large Value","Global Large-Stock Blend", "Equity Hedged","Foreign Large Blend","Foreign Small/Mid Growth",
        "Foreign Small/Mid Value","Global Large-Stock Value","Global Equity Balanced","Diversified Pacific/Asia","Global Small/Mid Stock","Global Large-Stock Growth",
        "Miscellaneous Region","Global Dividend & Income Equity","Global Equity","Tactical Allocation","Tactical Balanced","International Equity",
        "Global Small/Mid Cap Equity","Global Infrastructure Equity","Asia Pacific Equity","Global Neutral Balanced","Aggressive Allocation",
        "Global Moderately Aggressive Allocation","Global Aggressive Allocation","Global Moderate Allocation","Foreign Small/Mid Blend",
        "Global Moderately Conservative Allocation","Moderate Allocation"]

    intl_em = ["Diversified Emerging Mkts","China Region","Emerging Markets Equity","Greater China Equity","India Equity","Asia Pacific ex-Japan Equity",
                "Latin America Stock","Emerging Markets Bond"]

    intl_dev = ["Pacific/Asia ex-Japan Stk","Canadian Dividend & Income Equity","Canadian Focused Equity","Canadian Equity","European Equity",
                "Canadian Small/Mid Cap Equity","Canadian Equity Balanced","Europe Stock","Japan Stock","North American Equity",
                "Canadian Focused Small/Mid Cap Equity"]

    df.loc[intl & df["category_name"].isin(intl_allocation), "asset_class"] = "Allocation"
    df.loc[intl & df["category_name"].isin(intl_em),         "asset_class"] = "Emerging Markets"
    df.loc[intl & df["category_name"].isin(intl_dev),        "asset_class"] = "International Developed"

    df.loc[intl & (df["category_name"] == "Global Real Estate"), "asset_class"] = "Real Estate"
    df.loc[intl & df["category_name"].isin(["Venture Capital","Alternative Equity Focused","Private Equity"]), "asset_class"] = "Other"


    # ============================================================
    # 5. FINAL OVERWRITES — STRING-BASED, HIGHEST PRIORITY
    # ============================================================

    df.loc[df["category_name"].str.contains("Real Estate", na=False), "asset_class"] = "Real Estate"

    #df.loc[df["category_name"].str.contains("Canada|France", na=False), "asset_class"] = "International Developed"

    df.loc[df["category_name"].str.contains("China|India", na=False), "asset_class"] = "Emerging Markets"
    df.loc[df["category_name"].str.contains("Allocation", na=False), "asset_class"] = "Allocation"
    df.loc[df["category_name"].str.contains("Private Equity|Venture Capital|Equity Hedged|Derivative Income|Alternative|Private Debt", na=False), "asset_class"] = "Other"
    df.loc[df["category_name"].str.contains("Target", na=False) &~df["category_name"].str.contains("Muni", na=False),"asset_class"] = "Allocation"

    df.loc[df["category_name"].str.contains("Fixed Income|Bond", na=False), "asset_class"] = "Fixed Income"

    df.loc[df["category_name"].str.contains("Digital", na=False), "asset_class"] = "Commodities"


    # 1/28
    df.loc[df["category_name"].str.contains("ERR:", na=False) & (df["detailed_security_type"] != "stock"), "asset_class"] = "Other"

    df=df.fillna(0)
    df["LC"] = df["Market Value"] * df["stock_net"] * (df["united_states_equity_exposure"] - df["real_estate_exposure"]) * df["large_market_cap_exposure_generic"]
    df["MC"] = df["Market Value"] * df["stock_net"] * (df["united_states_equity_exposure"] - df["real_estate_exposure"]) * df["medium_market_cap_exposure_generic"]
    df["SC"] = df["Market Value"] * df["stock_net"] * (df["united_states_equity_exposure"] - df["real_estate_exposure"]) * df["small_market_cap_exposure_generic"]
    df["RE"] = df["Market Value"] * df["stock_net"] * df["real_estate_exposure"]
    df["Dev"] = df["Market Value"] * df["stock_net"] * (df["developed_market_exposure"] - df["united_states_equity_exposure"])
    df["EM"] = df["Market Value"] * df["stock_net"] * df["emerging_market_exposure"]
    df["Comm"] = 0
    df["FI"] = df["Market Value"] * df["bond_net"] + df["Market Value"] * df["convertible_net"] +df["Market Value"] * df["preferred_net"]
    df["cash"] = df["Market Value"] * df["cash_net"]
    df["Other"] = df["Market Value"] - (df["LC"] + df["MC"] + df["SC"] + df["RE"] +df["Dev"] + df["EM"] + df["Comm"] +df["FI"] + df["cash"])
    df["Error"] = df["Market Value"] - (df["LC"]+df["MC"]+df["SC"]+df["RE"]+df["Dev"]+df["EM"]+df["Comm"]+df["FI"]+df["Other"]+df["cash"])

    # =========================
    # OLD-FASHIONED GAA (LC/MC/... already dollar values)
    # =========================

    # --- Split data ---
    base = df[df["asset_class"] != "Allocation"]
    alloc = df[df["asset_class"] == "Allocation"]

    base_q = base[base["Q/NQ"] == "Q"]
    base_nq = base[base["Q/NQ"] == "NQ"]

    alloc_q = alloc[alloc["Q/NQ"] == "Q"]
    alloc_nq = alloc[alloc["Q/NQ"] == "NQ"]

    general_gaa = pd.DataFrame({"asset_class": ["Large Cap", "Mid Cap", "Small Cap","International Developed", "Emerging Markets",
            "Real Estate", "Commodities","Other","Fixed Income", "Cash"],

        # -------- ALL --------
        "Market Value": [
            base.loc[base["asset_class"] == "Large Cap", "Market Value"].sum() + alloc["LC"].sum(),
            base.loc[base["asset_class"] == "Mid Cap", "Market Value"].sum() + alloc["MC"].sum(),
            base.loc[base["asset_class"] == "Small Cap", "Market Value"].sum() + alloc["SC"].sum(),
            base.loc[base["asset_class"] == "International Developed", "Market Value"].sum() + alloc["Dev"].sum(),
            base.loc[base["asset_class"] == "Emerging Markets", "Market Value"].sum() + alloc["EM"].sum(),
            base.loc[base["asset_class"] == "Real Estate", "Market Value"].sum() + alloc["RE"].sum(),
            base.loc[base["asset_class"] == "Commodities", "Market Value"].sum() + alloc["Comm"].sum(),
            base.loc[base["asset_class"] == "Other", "Market Value"].sum() + alloc["Other"].sum(),
            base.loc[base["asset_class"] == "Fixed Income", "Market Value"].sum() + alloc["FI"].sum(),
            base.loc[base["asset_class"] == "Cash", "Market Value"].sum() + alloc["cash"].sum(),
        ],

        # -------- Q --------
        "Q_MV": [
            base_q.loc[base_q["asset_class"] == "Large Cap", "Market Value"].sum() + alloc_q["LC"].sum(),
            base_q.loc[base_q["asset_class"] == "Mid Cap", "Market Value"].sum() + alloc_q["MC"].sum(),
            base_q.loc[base_q["asset_class"] == "Small Cap", "Market Value"].sum() + alloc_q["SC"].sum(),
            base_q.loc[base_q["asset_class"] == "International Developed", "Market Value"].sum() + alloc_q["Dev"].sum(),
            base_q.loc[base_q["asset_class"] == "Emerging Markets", "Market Value"].sum() + alloc_q["EM"].sum(),
            base_q.loc[base_q["asset_class"] == "Real Estate", "Market Value"].sum() + alloc_q["RE"].sum(),
            base_q.loc[base_q["asset_class"] == "Commodities", "Market Value"].sum() + alloc_q["Comm"].sum(),
            base_q.loc[base_q["asset_class"] == "Other", "Market Value"].sum() + alloc_q["Other"].sum(),
            base_q.loc[base_q["asset_class"] == "Fixed Income", "Market Value"].sum() + alloc_q["FI"].sum(),
            base_q.loc[base_q["asset_class"] == "Cash", "Market Value"].sum() + alloc_q["cash"].sum(),
        ],

        # -------- NQ --------
        "NQ_MV": [
            base_nq.loc[base_nq["asset_class"] == "Large Cap", "Market Value"].sum() + alloc_nq["LC"].sum(),
            base_nq.loc[base_nq["asset_class"] == "Mid Cap", "Market Value"].sum() + alloc_nq["MC"].sum(),
            base_nq.loc[base_nq["asset_class"] == "Small Cap", "Market Value"].sum() + alloc_nq["SC"].sum(),
            base_nq.loc[base_nq["asset_class"] == "International Developed", "Market Value"].sum() + alloc_nq["Dev"].sum(),
            base_nq.loc[base_nq["asset_class"] == "Emerging Markets", "Market Value"].sum() + alloc_nq["EM"].sum(),
            base_nq.loc[base_nq["asset_class"] == "Real Estate", "Market Value"].sum() + alloc_nq["RE"].sum(),
            base_nq.loc[base_nq["asset_class"] == "Commodities", "Market Value"].sum() + alloc_nq["Comm"].sum(),
            base_nq.loc[base_nq["asset_class"] == "Other", "Market Value"].sum() + alloc_nq["Other"].sum(),
            base_nq.loc[base_nq["asset_class"] == "Fixed Income", "Market Value"].sum() + alloc_nq["FI"].sum(),
            base_nq.loc[base_nq["asset_class"] == "Cash", "Market Value"].sum() + alloc_nq["cash"].sum(),
        ]
    })

    # --- Convert to allocations ---
    general_gaa["Target Allocation"] = ""
    general_gaa["Current Allocation"] = general_gaa["Market Value"] / df["Market Value"].sum()
    general_gaa["Q Allocation"] = general_gaa["Q_MV"] / df.loc[df["Q/NQ"] == "Q", "Market Value"].sum()
    general_gaa["NQ Allocation"] = general_gaa["NQ_MV"] / df.loc[df["Q/NQ"] == "NQ", "Market Value"].sum()

    general_gaa_total_row = general_gaa[general_gaa.select_dtypes(include="number").columns].sum()
    general_gaa_total_row["asset_class"] = "Total:"
    general_gaa.loc[len(general_gaa)] = general_gaa_total_row


    general_gaa = general_gaa[    ["asset_class", "Target Allocation", "Current Allocation", "Q Allocation", "NQ Allocation"]]
    general_gaa= general_gaa.fillna("")

    # =========================
    # TABLE 2: NQ GAA BY ACCOUNT (old-fashioned, explicit)
    # =========================

    # --- Keep only NQ rows ---
    df_nq = df[df["Q/NQ"] == "NQ"]

    # --- Split Allocation vs non-Allocation ---
    base_nq = df_nq[df_nq["asset_class"] != "Allocation"]
    alloc_nq = df_nq[df_nq["asset_class"] == "Allocation"]

    rows = []

    for acct in df_nq["Account"].unique():

        b = base_nq[base_nq["Account"] == acct]
        a = alloc_nq[alloc_nq["Account"] == acct]


        rows.extend([
            (acct, "Large Cap", b.loc[b["asset_class"] == "Large Cap", "Market Value"].sum() + a["LC"].sum()),
            (acct, "Mid Cap", b.loc[b["asset_class"] == "Mid Cap", "Market Value"].sum() + a["MC"].sum()),
            (acct, "Small Cap", b.loc[b["asset_class"] == "Small Cap", "Market Value"].sum() + a["SC"].sum()),
            (acct, "International Developed", b.loc[b["asset_class"] == "International Developed", "Market Value"].sum() + a["Dev"].sum()),
            (acct, "Emerging Markets", b.loc[b["asset_class"] == "Emerging Markets", "Market Value"].sum() + a["EM"].sum()),
            (acct, "Real Estate", b.loc[b["asset_class"] == "Real Estate", "Market Value"].sum() + a["RE"].sum()),
            (acct, "Commodities", b.loc[b["asset_class"] == "Commodities", "Market Value"].sum() + a["Comm"].sum()),
            (acct, "Other", b.loc[b["asset_class"] == "Other", "Market Value"].sum() + a["Other"].sum()),
            (acct, "Fixed Income", b.loc[b["asset_class"] == "Fixed Income", "Market Value"].sum() + a["FI"].sum()),
            (acct, "Cash", b.loc[b["asset_class"] == "Cash", "Market Value"].sum() + a["cash"].sum()),
        ])

    # --- Build dataframe ---
    nq_gaa = pd.DataFrame(rows, columns=["Account", "asset_class", "Market Value"])

    # --- Convert to allocation ---
    nq_totals = df_nq.groupby("Account")["Market Value"].sum()
    nq_gaa["NQ Allocation"] = nq_gaa["Market Value"] / nq_gaa["Account"].map(nq_totals)

    # --- Final shape ---
    nq_gaa = nq_gaa[["Account", "asset_class", "NQ Allocation"]]
    nq_gaa = nq_gaa.pivot(index="asset_class",columns="Account",values="NQ Allocation").fillna(0)

    nq_gaa = nq_gaa.reset_index()

    # =========================
    # TABLE 3: Q GAA BY ACCOUNT (old-fashioned, explicit)
    # =========================

    # --- Keep only Q rows ---
    df_q = df[df["Q/NQ"] == "Q"]

    # --- Split Allocation vs non-Allocation ---
    base_q = df_q[df_q["asset_class"] != "Allocation"]
    alloc_q = df_q[df_q["asset_class"] == "Allocation"]

    rows = []

    for acct in df_q["Account"].unique():

        b = base_q[base_q["Account"] == acct]
        a = alloc_q[alloc_q["Account"] == acct]

        rows.extend([
            (acct, "Large Cap", b.loc[b["asset_class"] == "Large Cap", "Market Value"].sum() + a["LC"].sum()),
            (acct, "Mid Cap", b.loc[b["asset_class"] == "Mid Cap", "Market Value"].sum() + a["MC"].sum()),
            (acct, "Small Cap", b.loc[b["asset_class"] == "Small Cap", "Market Value"].sum() + a["SC"].sum()),
            (acct, "International Developed", b.loc[b["asset_class"] == "International Developed", "Market Value"].sum() + a["Dev"].sum()),
            (acct, "Emerging Markets", b.loc[b["asset_class"] == "Emerging Markets", "Market Value"].sum() + a["EM"].sum()),
            (acct, "Real Estate", b.loc[b["asset_class"] == "Real Estate", "Market Value"].sum() + a["RE"].sum()),
            (acct, "Commodities", b.loc[b["asset_class"] == "Commodities", "Market Value"].sum() + a["Comm"].sum()),
            (acct, "Fixed Income", b.loc[b["asset_class"] == "Fixed Income", "Market Value"].sum() + a["FI"].sum()),
            (acct, "Cash", b.loc[b["asset_class"] == "Cash", "Market Value"].sum() + a["cash"].sum()),
            (acct, "Other", b.loc[b["asset_class"] == "Other", "Market Value"].sum() + a["Other"].sum()),
        ])

    # --- Build dataframe ---
    q_gaa = pd.DataFrame(rows, columns=["Account", "asset_class", "Market Value"])

    # --- Convert to allocation ---
    q_totals = df_q.groupby("Account")["Market Value"].sum()
    q_gaa["Q Allocation"] = q_gaa["Market Value"] / q_gaa["Account"].map(q_totals)

    q_gaa = q_gaa[["Account", "asset_class", "Q Allocation"]]

    q_gaa = q_gaa.pivot(index="asset_class",columns="Account",values="Q Allocation").fillna(0)
    q_gaa = q_gaa.reset_index()


    q_gaa  = q_gaa.rename(columns={"asset_class": "Asset Class"})
    nq_gaa = nq_gaa.rename(columns={"asset_class": "Asset Class"})

    row_order = ["Large Cap","Mid Cap","Small Cap","International Developed","Emerging Markets",
                "Real Estate","Commodities","Other","Fixed Income","Cash","Total:"]

    for gaa_table in [q_gaa, nq_gaa]:
        num_cols_gaa = gaa_table.select_dtypes(include="number").columns
        total_row = gaa_table[num_cols_gaa].sum()
        total_row["Asset Class"] = "Total:"
        gaa_table.loc[len(gaa_table)] = total_row

        gaa_table["Asset Class"] = pd.Categorical(gaa_table["Asset Class"], categories=row_order, ordered=True)
        gaa_table.sort_values("Asset Class", inplace=True)
        gaa_table.reset_index(drop=True, inplace=True)

        gaa_table[num_cols_gaa] = gaa_table[num_cols_gaa].applymap(lambda x: f"{x*100:.2f}%" if pd.notna(x) else "")

    # =====================================================
    # STYLE COLUMNS
    # =====================================================

    style_cols = ["equity_stylebox_large_cap_blend_exposure", "equity_stylebox_large_cap_growth_exposure", "equity_stylebox_large_cap_value_exposure",
                "equity_stylebox_mid_cap_blend_exposure","equity_stylebox_mid_cap_growth_exposure", "equity_stylebox_mid_cap_value_exposure",
                "equity_stylebox_small_cap_blend_exposure", "equity_stylebox_small_cap_growth_exposure", "equity_stylebox_small_cap_value_exposure"]

    style_map = {"Large Cap/Blend":  "equity_stylebox_large_cap_blend_exposure","Large Cap/Growth": "equity_stylebox_large_cap_growth_exposure",
                "Large Cap/Value":  "equity_stylebox_large_cap_value_exposure","Mid Cap/Blend":    "equity_stylebox_mid_cap_blend_exposure",
                "Mid Cap/Growth":   "equity_stylebox_mid_cap_growth_exposure","Mid Cap/Value":    "equity_stylebox_mid_cap_value_exposure",
                "Small Cap/Blend":  "equity_stylebox_small_cap_blend_exposure","Small Cap/Growth": "equity_stylebox_small_cap_growth_exposure",
                "Small Cap/Value":  "equity_stylebox_small_cap_value_exposure",}

    valid_rows = (
        # keep rows where equity_style is NOT an error
        ~style_rows_with_ERR["equity_style"].astype(str).str.contains("ERR:", na=False)

        # AND remove rows that are:
        #   (not stock) AND (all style columns are ERR)
        & ~((style_rows_with_ERR["detailed_security_type"] != "stock") &style_rows_with_ERR[style_cols].astype(str).apply(lambda x: x.str.contains("ERR:")).all(axis=1)))

    style_df = style_rows_with_ERR.loc[valid_rows].copy()

    # zero all style exposures for stocks first
    style_df.loc[style_df["detailed_security_type"] == "stock", style_cols] = 0.0

    # then assign 100% to the correct bucket based on equity_style
    style_df.loc[(style_df["detailed_security_type"]=="stock")&(style_df["equity_style"]=="Large Cap/Blend"),  "equity_stylebox_large_cap_blend_exposure"]=1.0
    style_df.loc[(style_df["detailed_security_type"]=="stock")&(style_df["equity_style"]=="Large Cap/Growth"), "equity_stylebox_large_cap_growth_exposure"]=1.0
    style_df.loc[(style_df["detailed_security_type"]=="stock")&(style_df["equity_style"]=="Large Cap/Value"),  "equity_stylebox_large_cap_value_exposure"]=1.0
    style_df.loc[(style_df["detailed_security_type"]=="stock")&(style_df["equity_style"]=="Mid Cap/Blend"),    "equity_stylebox_mid_cap_blend_exposure"]=1.0
    style_df.loc[(style_df["detailed_security_type"]=="stock")&(style_df["equity_style"]=="Mid Cap/Growth"),   "equity_stylebox_mid_cap_growth_exposure"]=1.0
    style_df.loc[(style_df["detailed_security_type"]=="stock")&(style_df["equity_style"]=="Mid Cap/Value"),    "equity_stylebox_mid_cap_value_exposure"]=1.0
    style_df.loc[(style_df["detailed_security_type"]=="stock")&(style_df["equity_style"]=="Small Cap/Blend"),  "equity_stylebox_small_cap_blend_exposure"]=1.0
    style_df.loc[(style_df["detailed_security_type"]=="stock")&(style_df["equity_style"]=="Small Cap/Growth"), "equity_stylebox_small_cap_growth_exposure"]=1.0
    style_df.loc[(style_df["detailed_security_type"]=="stock")&(style_df["equity_style"]=="Small Cap/Value"),  "equity_stylebox_small_cap_value_exposure"]=1.0

    # sanity check: exposures should sum to 1 for stocks (0 for non-stocks)
    #style_df["style_exposure_sanity"] = style_df[style_cols].sum(axis=1)

    # dollar exposure = exposure * Market Value

    style_df["large_cap_blend_market_value"]  = style_df["equity_stylebox_large_cap_blend_exposure"]  * style_df["Market Value"]
    style_df["large_cap_growth_market_value"] = style_df["equity_stylebox_large_cap_growth_exposure"] * style_df["Market Value"]
    style_df["large_cap_value_market_value"]  = style_df["equity_stylebox_large_cap_value_exposure"]  * style_df["Market Value"]

    style_df["mid_cap_blend_market_value"]    = style_df["equity_stylebox_mid_cap_blend_exposure"]    * style_df["Market Value"]
    style_df["mid_cap_growth_market_value"]   = style_df["equity_stylebox_mid_cap_growth_exposure"]   * style_df["Market Value"]
    style_df["mid_cap_value_market_value"]    = style_df["equity_stylebox_mid_cap_value_exposure"]    * style_df["Market Value"]

    style_df["small_cap_blend_market_value"]  = style_df["equity_stylebox_small_cap_blend_exposure"]  * style_df["Market Value"]
    style_df["small_cap_growth_market_value"] = style_df["equity_stylebox_small_cap_growth_exposure"] * style_df["Market Value"]
    style_df["small_cap_value_market_value"]  = style_df["equity_stylebox_small_cap_value_exposure"]  * style_df["Market Value"]

    #style_df["style_market_value_sanity_sum"] = style_df[["large_cap_blend_market_value","large_cap_growth_market_value","large_cap_value_market_value","mid_cap_blend_market_value","mid_cap_growth_market_value","mid_cap_value_market_value","small_cap_blend_market_value","small_cap_growth_market_value","small_cap_value_market_value"]].sum(axis=1)
    #style_df["style_market_value_sanity"] = (style_df["style_market_value_sanity_sum"] - style_df["Market Value"]).round()



    large_cap_blend_percentage  = style_df["large_cap_blend_market_value"].sum()  / style_df["Market Value"].sum()
    large_cap_growth_percentage = style_df["large_cap_growth_market_value"].sum() / style_df["Market Value"].sum()
    large_cap_value_percentage  = style_df["large_cap_value_market_value"].sum()  / style_df["Market Value"].sum()

    mid_cap_blend_percentage    = style_df["mid_cap_blend_market_value"].sum()    / style_df["Market Value"].sum()
    mid_cap_growth_percentage   = style_df["mid_cap_growth_market_value"].sum()   / style_df["Market Value"].sum()
    mid_cap_value_percentage    = style_df["mid_cap_value_market_value"].sum()    / style_df["Market Value"].sum()

    small_cap_blend_percentage  = style_df["small_cap_blend_market_value"].sum()  / style_df["Market Value"].sum()
    small_cap_growth_percentage = style_df["small_cap_growth_market_value"].sum() / style_df["Market Value"].sum()
    small_cap_value_percentage  = style_df["small_cap_value_market_value"].sum()  / style_df["Market Value"].sum()


    # sanity check → should be ~1.0
    #sanity_check = (large_cap_blend_percentage +large_cap_growth_percentage +large_cap_value_percentage +mid_cap_blend_percentage +mid_cap_growth_percentage +mid_cap_value_percentage +small_cap_blend_percentage +small_cap_growth_percentage +small_cap_value_percentage)

    #print("Style weight total:", sanity_check)

    df.loc[len(df), df.columns[0]] = "Aggregate"
    for col in ["expense_ratio", "market_beta_36_month",]:
        df.loc[df[df.columns[0]] == "Aggregate", col] = (df["Market Value"] * df[col]).sum() / df["Market Value"].sum()



    fi_total_mv = df.loc[df["asset_class"] == "Fixed Income", "Market Value"].sum()
    fi_cols = ["securitized_fixed_income_exposure_generic", "current_yield","yield_to_maturity_generic","effective_duration","effective_maturity",
            "corporate_fixed_income_exposure_generic", "government_fixed_income_exposure_generic","municipal_fixed_income_exposure_generic",
            "other_fixed_income_exposure_generic", "aaa_bond_exposure","aa_bond_exposure","a_bond_exposure","bbb_bond_exposure", "bb_bond_exposure",
            "b_bond_exposure","below_b_bond_exposure", "not_rated_bond_exposure","duration_generic"]

    if fi_total_mv > 0:
        for col in fi_cols:
            df.loc[df[df.columns[0]] == "Aggregate", col] = ((df.loc[df["asset_class"] == "Fixed Income", "Market Value"] * df.loc[df["asset_class"] == "Fixed Income", col]).sum() / fi_total_mv)
    else:
        df.loc[df[df.columns[0]] == "Aggregate", fi_cols] = " "

    plot_data = general_gaa[general_gaa["asset_class"] != "Total:"]

    fig, ax = plt.subplots()

    wedges, texts, autotexts = ax.pie(
        plot_data["Current Allocation"],
        startangle=90,
        pctdistance=0.7,
        autopct=lambda p: f"{p:.1f}%" if p >= 2 else ""
    )

    ax.legend(wedges, plot_data["asset_class"], title="Asset Class",
            loc="center left", bbox_to_anchor=(1, 0.5))

    ax.set_title("Asset Allocation")
    ax.axis("equal")
    plt.tight_layout()



    BASE_DIR = Path(__file__).resolve().parent
    OBS_PATH = BASE_DIR / "Observations"
    OBS_PATH.mkdir(exist_ok=True)

    plt.savefig(OBS_PATH / "Title Pie chart.png", dpi=300, bbox_inches="tight")

    # In[ ]:


    summary_gaa = (pd.concat([df.loc[(df["Q/NQ"] == "Q") & (df["Market Value"] != 0)].groupby(["Q/NQ", "Account"], as_index=False)["Market Value"].sum(),

            pd.DataFrame([{"Q/NQ": "Total Q:","Account": "","Market Value": df.loc[(df["Q/NQ"] == "Q") & (df["Market Value"] != 0),"Market Value"].sum()}]),
            df.loc[(df["Q/NQ"] == "NQ") & (df["Market Value"] != 0)].groupby(["Q/NQ", "Account"], as_index=False)["Market Value"].sum(),
            pd.DataFrame([{"Q/NQ": "Total NQ:","Account": "","Market Value": df.loc[(df["Q/NQ"] == "NQ") & (df["Market Value"] != 0), "Market Value"].sum()}]) ],ignore_index=True))
    summary_gaa.loc["Total:"] = ["Total:", "", summary_gaa["Market Value"].sum() / 2]
    summary_gaa["Date"] = summary_gaa.iloc[:, 0].astype(str).str.contains("Total", case=False).map(lambda x: "" if x else "Values as of ")


    # In[ ]:


    #Group by ticker, drop Q/NQ and Account because you dont need them
    grouped_by_ticker = df.groupby("Symbol", as_index=False).agg({"Market Value": "sum", **{col: "first" for col in df.columns if col not in ["Symbol", "Market Value"]}})
    grouped_by_ticker= grouped_by_ticker.drop(columns=["Q/NQ", "Account"])
    grouped_by_ticker['Expense Paid'] = grouped_by_ticker['Market Value']* grouped_by_ticker['expense_ratio']

    #grouped by Q and NQ
    grouped_by_ticker_account_nq = (df[df["Q/NQ"] == "NQ"].groupby(["Symbol", "Account"], as_index=False).agg({"Market Value": "sum",**{col: "first" for col in df.columns if col not in ["Symbol", "Account", "Market Value"]}}))
    grouped_by_ticker_account_qualified = (df[df["Q/NQ"] == "Q"].groupby(["Symbol", "Account"], as_index=False).agg({"Market Value": "sum",**{col: "first" for col in df.columns if col not in ["Symbol", "Account", "Market Value"]}}))


    #Bring to numeric
    grouped_by_ticker[num_cols] = grouped_by_ticker[num_cols].apply(pd.to_numeric, errors="coerce")
    grouped_by_ticker_account_nq[num_cols] = grouped_by_ticker_account_nq[num_cols].apply(pd.to_numeric, errors="coerce")
    grouped_by_ticker_account_qualified[num_cols] = grouped_by_ticker_account_qualified[num_cols].apply(pd.to_numeric, errors="coerce")



    high_expense_ratio = grouped_by_ticker.loc[grouped_by_ticker["broad_asset_class"] != "Money Market", ["Name","Symbol","Market Value","expense_ratio","expense_ratio_rank", "Expense Paid"]]
    high_expense_ratio = high_expense_ratio[high_expense_ratio["expense_ratio"] > 0.007]
    high_expense_ratio = high_expense_ratio.sort_values(by="Expense Paid", ascending=False)

    tax_cost_ratio = grouped_by_ticker.loc[grouped_by_ticker["broad_asset_class"] != "Money Market", ["Name","Symbol","Market Value","one_year_tax_cost_ratio","three_year_tax_cost_ratio","five_year_tax_cost_ratio","ten_year_tax_cost_ratio"]]
    tax_cost_ratio = tax_cost_ratio[tax_cost_ratio["one_year_tax_cost_ratio"] > 0.02]
    tax_cost_ratio = tax_cost_ratio.sort_values(by="three_year_tax_cost_ratio", ascending=False)

    taxable_bonds_in_nq = grouped_by_ticker_account_nq[grouped_by_ticker_account_nq["broad_asset_class"] == "Taxable Bond"][['Account', 'Name', 'Symbol', 'Market Value', 'current_yield', 'broad_asset_class']]
    taxable_bonds_in_nq['Estimated Interest'] = taxable_bonds_in_nq['current_yield'] * taxable_bonds_in_nq['Market Value']
    taxable_bonds_in_nq = taxable_bonds_in_nq[taxable_bonds_in_nq["current_yield"] > 0.01]
    taxable_bonds_in_nq = taxable_bonds_in_nq.sort_values(by="Estimated Interest", ascending=False)


    taxable_bonds_in_nq

    #muni_bond_qualified = grouped_by_ticker_account_qualified[grouped_by_ticker_account_qualified["broad_asset_class"] == "Municipal Bond"][['Account', 'Name', 'Symbol', 'Market Value', 'current_yield', 'broad_asset_class','annualized_daily_one_year_return','annualized_three_year_return','annualized_daily_five_year_return']]

    # muni bonds
    muni_bond_qualified = grouped_by_ticker_account_qualified[grouped_by_ticker_account_qualified["broad_asset_class"] == "Municipal Bond"]
    muni_bond_qualified= muni_bond_qualified.sort_values(by="yield_to_maturity_generic", ascending=True)

    index_bnd_row = df[(df["Q/NQ"] == "Index") & (df["Symbol"] == "BND")]  #index

    # combine + select columns
    muni_bond_qualified = pd.concat([index_bnd_row, muni_bond_qualified], ignore_index=True)[["Account","Name","Symbol","Market Value","yield_to_maturity_generic","broad_asset_class"]]


    muni_bond_qualified

    tracking_error = grouped_by_ticker.loc[grouped_by_ticker["broad_asset_class"] != "Money Market", ["Name","Symbol","Market Value","tracking_error_1y_vs_category","tracking_error_3y_vs_category","tracking_error_5y_vs_category","tracking_error_10y_vs_category"]]
    tracking_error["tracking_error_3y_vs_category"] = pd.to_numeric(tracking_error["tracking_error_3y_vs_category"], errors="coerce")
    tracking_error = tracking_error[tracking_error["tracking_error_3y_vs_category"] > 6]

    tracking_error = tracking_error.sort_values(by="tracking_error_3y_vs_category", ascending=False)
    #Note is tracking_error_3y_vs_category the best metric??

    tracking_error

    active_passive = grouped_by_ticker.loc[grouped_by_ticker["broad_asset_class"] != "Money Market",["Name", "Symbol", "Market Value", "category_name", "index_fund", "detailed_security_type"]]
    active_passive = active_passive[active_passive["Symbol"].astype(str).str.len() <= 7] # 2. Remove tickers longer than 7 characters
    active_passive = (active_passive.groupby(["category_name", "index_fund"], as_index=False).agg({"Market Value": "sum"}))
    active_passive = active_passive[~active_passive.astype(str).apply(lambda col: col.str.contains("ERR:", na=False)).any(axis=1)] # drop any row where any column contains "ERR:"
    active_passive = (active_passive.pivot(index="category_name", columns="index_fund", values="Market Value").fillna(0).reset_index())      # bring category_name back as a normal column
    active_passive = active_passive.rename(columns={"False": "Active $","True": "Passive $"}) # 5. Rename the pivoted columns to something readable
    active_passive['Market Value'] = active_passive['Active $']  + active_passive['Passive $']
    active_passive = active_passive[(active_passive["Active $"] + active_passive["Passive $"]) >= 0.01*(df['Market Value'].sum())]
    active_passive


    individual_stocks_df =  grouped_by_ticker  [['Name','Symbol','Market Value','detailed_security_type','annualized_daily_one_year_return','annualized_three_year_return', 'annualized_daily_five_year_return','max_drawdown_1y','max_drawdown_3y','max_drawdown_5y','daily_standard_deviation_1y','daily_standard_deviation_3y','daily_standard_deviation_5y']]
    individual_stocks_df = individual_stocks_df[individual_stocks_df["detailed_security_type"] == "stock"]
    individual_stocks_df = individual_stocks_df[['Name','Symbol','Market Value','annualized_daily_one_year_return','annualized_three_year_return','max_drawdown_1y','max_drawdown_3y','daily_standard_deviation_1y','daily_standard_deviation_3y']]
    #Refiltering for formatting

    retail_class = grouped_by_ticker.loc[grouped_by_ticker["broad_asset_class"] != "Money Market", ["Name","Symbol","Market Value","share_class","expense_ratio","actual_12b1", "Expense Paid"]]
    retail_class = retail_class.loc[retail_class["share_class"].isin(["A","B","C","D","F","Inv","K","M","N","P","Premier","Retail","Service","T","F3","I2"])]

    fund_families = grouped_by_ticker.groupby('fund_family', as_index=False).agg({"Market Value": "sum", "Expense Paid": "sum"})
    fund_families = fund_families[~fund_families["fund_family"].str.contains("ERR:", case=False, na=False)]
    fund_families = fund_families[['fund_family','Market Value','Expense Paid']]
    fund_families['Percentage Paid']= fund_families['Expense Paid']/fund_families['Expense Paid'].sum()

    target_date_fund = grouped_by_ticker[['Name','Symbol','Market Value','expense_ratio','category_name']]
    target_date_fund = target_date_fund[target_date_fund["category_name"].str.contains("target", case=False, na=False)]


    #See maybe I can add more to it
    countries_sector_values = ["Japan Stock", "India Equity", "Latin America Stock", "Europe Stock","Pacific/Asia ex-Japan Stock", "China Region",
            "Diversified Pacific/Asia","Miscellaneous Region", "Technology", "Utilities", "Health", "Industrials","Communications","Consumer Cyclical",
                "Consumer Defensive", "Financial","Natural Resources", "Infrastructure", "Equity Energy","Equity Precious Metals","Miscellaneous Sector"]

    countries_sector_rows = grouped_by_ticker[['Name','Symbol','Market Value','expense_ratio','category_name','max_drawdown_3y', 'annualized_three_year_return']]
    countries_sector_rows = countries_sector_rows[countries_sector_rows["category_name"].isin(countries_sector_values)]

    leverage_inverse_df = grouped_by_ticker[['Name','Symbol','Market Value','expense_ratio','category_name','alpha_3y_vs_category','max_drawdown_3y','annualized_three_year_return']]
    leverage_inverse_df = leverage_inverse_df[leverage_inverse_df["category_name"].str.contains("Trading--", case=False, na=False)]

    digital_df = grouped_by_ticker[['Name','Symbol','Market Value','expense_ratio','category_name','max_drawdown_1y',#'max_drawdown_3y','max_drawdown_5y',
                                    'annualized_daily_one_year_return']]#,'annualized_three_year_return', 'annualized_daily_five_year_return']]
    digital_df = digital_df[digital_df["category_name"].str.contains("Digital", case=False, na=False)]

    dilution = df.loc[df["broad_asset_class"] != "Money Market", ["Q/NQ","Symbol","Market Value","category_name","annualized_daily_one_year_return","annualized_three_year_return","annualized_daily_five_year_return"]]

    # Base extract
    dilution = df.loc[
        df["broad_asset_class"] != "Money Market",
        ["Q/NQ","Symbol","Market Value","category_name","annualized_daily_one_year_return","annualized_three_year_return","annualized_daily_five_year_return"]]

    # 1. Keep only valid Q/NQ values
    dilution = dilution[dilution["Q/NQ"].isin(["Q", "NQ", "Index","SP500"])]

    # 2. Keep only categories that have BOTH Index and non-Index (Q or NQ)
    index_cats = set(dilution.loc[dilution["Q/NQ"].isin(["Index", "SP500"]), "category_name"])
    non_index_cats = set(dilution.loc[dilution["Q/NQ"].isin(["Q", "NQ"]), "category_name"])

    valid_categories = index_cats & non_index_cats
    dilution = dilution[dilution["category_name"].isin(valid_categories)]

    # 3. Drop non-Index rows that do not underperform their Index by the tolerance
    tolerance = 0.02

    index_returns = (dilution.loc[dilution["Q/NQ"] == "Index", ["category_name", "annualized_three_year_return"]].rename(columns={"annualized_three_year_return": "index_3y_return"}))

    dilution = pd.merge(dilution,index_returns,on="category_name",how="left",)

    dilution = dilution[(dilution["Q/NQ"] == "Index")| (dilution["annualized_three_year_return"] <= dilution["index_3y_return"] - tolerance)]
    dilution = dilution.drop(columns="index_3y_return")

    # 4. Final guardrail: re-enforce Index + non-Index existence after deletions
    index_cats = set(dilution.loc[dilution["Q/NQ"] == "Index", "category_name"])
    non_index_cats = set(dilution.loc[dilution["Q/NQ"].isin(["Q", "NQ"]), "category_name"])
    valid_categories = index_cats & non_index_cats
    dilution = dilution[dilution["category_name"].isin(valid_categories)]

    # 5. Sort: Index first within each category
    order_map = {"Index": 0, "SP500": 0.5,  "Q": 1, "NQ": 2}
    dilution = (dilution.assign(_order=dilution["Q/NQ"].map(order_map)).sort_values(["category_name", "_order"]).drop(columns="_order"))

    # ============================================================
    # CONFIG: sector → exposure column mapping (single source of truth)
    # ============================================================
    sector_map = { "Information Technology": "technology_exposure", "Consumer Staples": "consumer_defensive_exposure", "Consumer Discretionary": "consumer_cyclical_exposure", "Communication Services": "communication_services_exposure", "Financials": "financial_services_exposure", "Health Care": "healthcare_exposure", "Energy": "energy_exposure", "Utilities": "utilities_exposure", "Industrials": "industrials_exposure", "Real Estate": "real_estate_exposure", "Materials": "basic_materials_exposure" }

    sector_cols = [f"{s}_value" for s in sector_map]
    df[sector_cols] = 0.0

    # STEP 2: Masks
    # ============================================================
    stock_mask = df["detailed_security_type"] == "stock"
    equity_like_mask = (df["detailed_security_type"] != "stock") & (df["asset_class"].isin(["Emerging Markets", "International Developed", "Large Cap", "Mid Cap", "Small Cap", "Allocation"]))
    commodities_mask = df["asset_class"] == "Commodities"

    # STEP 3: Stocks → 100% allocated to their sector
    # ============================================================
    for sector in sector_map:
        df.loc[stock_mask & (df["Sector"] == sector), f"{sector}_value"] = df["Market Value"]

    # STEP 4: Equity-like funds → allocate using exposures
    # ============================================================
    for sector, exposure_col in sector_map.items():
        df.loc[equity_like_mask, f"{sector}_value"] = df["Market Value"] * df[exposure_col].fillna(0)

    # STEP 5: Commodities → 100% to Materials
    # ============================================================
    df.loc[commodities_mask, "Materials_value"] = df["Market Value"]

    # STEP 6: Denominator
    # ============================================================
    include_mask = equity_like_mask | commodities_mask
    total_value = df.loc[include_mask, "Market Value"].sum()

    # STEP 7: Final output
    # ============================================================
    sectors = df[sector_cols].sum().div(total_value).rename(lambda c: c.replace("_value", ""), axis=0).reset_index().rename(columns={"index": "Sector", 0: "Weight"})


    # STEP 8: Add aggregate exposures + enforce order
    # ============================================================

    aggregates = pd.DataFrame({
        "Sector": ["Cyclical Exposure", "Sensitive Exposure", "Defensive Exposure"],
        "Weight": [sectors.loc[sectors["Sector"] == "Materials", "Weight"].values[0]+ sectors.loc[sectors["Sector"] == "Consumer Discretionary", "Weight"].values[0]
                + sectors.loc[sectors["Sector"] == "Financials", "Weight"].values[0]+ sectors.loc[sectors["Sector"] == "Real Estate", "Weight"].values[0],

            sectors.loc[sectors["Sector"] == "Communication Services", "Weight"].values[0]+ sectors.loc[sectors["Sector"] == "Information Technology", "Weight"].values[0]
            + sectors.loc[sectors["Sector"] == "Energy", "Weight"].values[0]+ sectors.loc[sectors["Sector"] == "Industrials", "Weight"].values[0],

            sectors.loc[sectors["Sector"] == "Consumer Staples", "Weight"].values[0]+ sectors.loc[sectors["Sector"] == "Health Care", "Weight"].values[0]
            + sectors.loc[sectors["Sector"] == "Utilities", "Weight"].values[0]]})

    sectors = pd.concat([aggregates, sectors], ignore_index=True)

    final_order = ["Cyclical Exposure", "Materials", "Consumer Discretionary", "Financials", "Real Estate",
        "Sensitive Exposure", "Communication Services", "Information Technology", "Energy", "Industrials",
        "Defensive Exposure", "Consumer Staples", "Health Care", "Utilities"]

    sectors = sectors.set_index("Sector").reindex(final_order).reset_index()


    # Map display labels (what you want in sectors["Sector"]) to the exact column names in df
    col_map = {"Cyclical Exposure":"cyclical_exposure","Materials":"basic_materials_exposure","Consumer Discretionary":"consumer_cyclical_exposure","Financials":"financial_services_exposure","Real Estate":"real_estate_exposure",
            "Sensitive Exposure":"sensitive_exposure","Communication Services":"communication_services_exposure","Information Technology":"technology_exposure","Energy":"energy_exposure","Industrials":"industrials_exposure",
            "Defensive Exposure":"defensive_exposure","Consumer Staples":"consumer_defensive_exposure","Health Care":"healthcare_exposure","Utilities":"utilities_exposure"}

    # Step 1: explicitly list the columns we want to pull from df (prevents accidental column mismatch)
    exposure_cols = list(col_map.values())

    # Step 2: select the single row where the first column equals "S&P 500"
    sp500_row = df.loc[df.iloc[:,0] == "SP500", exposure_cols]

    # Step 3: rename columns from internal names (cyclical_exposure) to display names (Cyclical Exposure)
    sp500_row = sp500_row.rename(columns={v:k for k,v in col_map.items()})

    # Step 4: transpose so each sector becomes its own row instead of a column
    sp500 = sp500_row.T.reset_index()

    # Step 5: force final column names to match sectors dataframe structure exactly
    sp500.columns = ["Sector","S&P 500"]

    # Step 6: merge safely by Sector name (guarantees correct alignment, never positional)
    sectors = pd.merge(sectors, sp500, on= 'Sector', how = 'left')

    sectors['Difference'] = sectors['Weight'] - sectors['S&P 500']
    sectors = sectors.rename(columns={"Weight": "Percentage"})


    stylebox_cols = [  # ordered Growth → Blend → Value (Large, Mid, Small)
        large_cap_growth_percentage,large_cap_blend_percentage,large_cap_value_percentage,
        mid_cap_growth_percentage,mid_cap_blend_percentage,mid_cap_value_percentage,
        small_cap_growth_percentage,small_cap_blend_percentage,small_cap_value_percentage]


    table = np.array(stylebox_cols).reshape(3, 3)
    aggregate_stylebox = pd.DataFrame(table, columns=["Growth", "Blend", "Value"], index=["Large", "Mid", "Small"])



    credit_yield = (
        df.loc[df.iloc[:, 0] == "Aggregate",
            ["current_yield","distribution_yield","yield_to_maturity_generic","yield_to_worst_generic"]]
        .rename(columns={
            "current_yield": "Current Yield","distribution_yield": "Distribution Yield",
            "yield_to_maturity_generic": "Yield To Maturity","yield_to_worst_generic": "Yield To Worst",}).T.reset_index(names="Metric").set_axis(["Metric", "Value"], axis=1))


    credit_duration = (
        df.loc[df.iloc[:, 0] == "Aggregate",
            ["duration_generic","effective_duration","effective_maturity"]]
        .rename(columns={
            "duration_generic": "Duration","effective_duration": "Effective Duration","effective_maturity": "Effective Maturity"}).T.reset_index(names="Metric").set_axis(["Metric", "Value"], axis=1))
    credit_duration["Value"] = credit_duration["Value"].round(2).astype(str) + " yr"





    credit_issuer_exposure = (
        df.loc[df.iloc[:, 0] == "Aggregate",
            ["corporate_fixed_income_exposure_generic","government_fixed_income_exposure_generic","municipal_fixed_income_exposure_generic","securitized_fixed_income_exposure_generic","other_fixed_income_exposure_generic"]]
        .rename(columns={
            "corporate_fixed_income_exposure_generic": "Corporate Fixed Income","government_fixed_income_exposure_generic": "Government Fixed Income","municipal_fixed_income_exposure_generic": "Municipal Fixed Income", "securitized_fixed_income_exposure_generic":"Securitized Exposure","other_fixed_income_exposure_generic": "Other Fixed Income"}).T.reset_index(names="Metric").set_axis(["Metric", "Value"], axis=1))


    credit_quality_exposure = (
        df.loc[df.iloc[:, 0] == "Aggregate",["aaa_bond_exposure","aa_bond_exposure","a_bond_exposure","bbb_bond_exposure","bb_bond_exposure","b_bond_exposure","below_b_bond_exposure","not_rated_bond_exposure"]])
    credit_quality_exposure['Investment Grade']= credit_quality_exposure['aaa_bond_exposure']+ credit_quality_exposure['aa_bond_exposure'] + credit_quality_exposure['a_bond_exposure'] + credit_quality_exposure['bbb_bond_exposure']
    credit_quality_exposure = credit_quality_exposure[['Investment Grade', 'bb_bond_exposure','b_bond_exposure', 'below_b_bond_exposure', 'not_rated_bond_exposure']].rename(columns={'bb_bond_exposure':' BB Exposure', 'b_bond_exposure':'B Exposure','below_b_bond_exposure':'Below B Exposure','not_rated_bond_exposure': 'Not Rated Exposure'}).T.reset_index(names="Metric").set_axis(["Rating", "Exposure"], axis=1)



    total_mv = df["Market Value"].sum()

    mutual_fund      = df.loc[df["detailed_security_type"] == "mutual_fund", "Market Value"].sum()
    individual_stocks = df.loc[df["detailed_security_type"] == "stock", "Market Value"].sum()
    individual_bonds  = df.loc[df["Symbol"].astype(str).str.len() >= 8, "Market Value"].sum()
    etf              = df.loc[df["detailed_security_type"] == "etf", "Market Value"].sum() - individual_bonds
    cef              = df.loc[df["detailed_security_type"].str.contains("cef", case=False, na=False), "Market Value"].sum()
    other            = total_mv - (mutual_fund + individual_stocks + individual_bonds + etf + cef)


    security_type_table = pd.DataFrame({
        "Security Type": ["ETF","Mutual Fund","Individual Stocks","Individual Bonds","CEF","Other"],
        "Market Value": [etf,mutual_fund,individual_stocks,individual_bonds,cef,other]})
    security_type_table['Percentage'] = security_type_table['Market Value']/total_mv

    alpha_summary = df.loc[(df["market_alpha_36_month"] < -0.5) & (~df["broad_asset_class"].isin(["Municipal Bond", "Taxable Bond", "Money Market"])),"Market Value"].sum()
    beta_summary = df.loc[df.iloc[:, 0] == "Aggregate", "market_beta_36_month"].iloc[0]

    #performance_summary = pd.DataFrame({ "Metric": [" Market Alpha < -0.5%)", "Beta (36 months)"], "Value": [f"${alpha_summary:,.2f}", f"{beta_summary:.2f}"]})


    performance_summary = pd.DataFrame({"Metric": ["Market Alpha < -0.5%)", "Beta (36 months)"], "Value": [f"${alpha_summary:,.2f}", f"{beta_summary:.2f}"],
        "Percentage": [f"{alpha_summary / grouped_by_ticker['Market Value'].sum():.2%}", ""]})



    performance_summary

    current_yield   = df.loc[df.iloc[:, 0] == "Aggregate", "current_yield"].iloc[0]
    credit_quality  = df.loc[df.iloc[:, 0] == "Aggregate", "average_credit_quality_score"].iloc[0]
    duration        = df.loc[df.iloc[:, 0] == "Aggregate", "duration_generic"].iloc[0]

    mini_fixed_income = pd.DataFrame({"Metric": ["Current Yield", "Credit Quality", "Duration"], "Value": [f"{current_yield*100:.2f}%", f"{credit_quality:.2f}", f"{duration:.2f} yr"]})

    mini_fixed_income

    high_tcr = grouped_by_ticker_account_nq.loc[(grouped_by_ticker_account_nq["three_year_tax_cost_ratio"] > 0.0005) & (grouped_by_ticker_account_nq["broad_asset_class"] != "Money Market"), "Market Value"].sum()
    inefficient_share_class = retail_class["Market Value"].sum()
    taxable_bonds_nq = grouped_by_ticker_account_nq.loc[grouped_by_ticker_account_nq["broad_asset_class"] == "Taxable Bond","Market Value"].sum()
    expense_ratio = df.loc[df.iloc[:, 0] == "Aggregate", "expense_ratio"].iloc[0]


    expense_summary = pd.DataFrame({
        "Metric": ["Tax Cost Ratio > 20 bps","Inefficient Share Class","Taxable Bonds (NQ)",],
        "Value": [high_tcr,inefficient_share_class,taxable_bonds_nq]})
    expense_summary['Percentage']= expense_summary['Value']/total_mv
    expense_summary["Percentage"] = expense_summary["Percentage"].map(lambda x: f"{x * 100:.2f}%")
    expense_summary["Value"] = expense_summary["Value"].map(lambda x: f"${x:,.2f}")
    expense_summary.loc[len(expense_summary)] = ["Expense Ratio (Aggregate)", "", f"{expense_ratio*10000:.2f} bps"]

    def check_equal(name, value, expected, tol=0.01):
        value_r = round(value, 2)
        expected_r = round(expected, 2)
        if abs(value_r - expected_r) <= tol:
            print(f"✅ {name}: CHECKED ({value_r})")
        else:
            print(f"❌ {name}: Expected {expected_r}, got {value_r}")


    def check_sum_equals_one(name, series):
        total = round(series.sum(), 2)
        if total == 1:
            print(f"✅ {name}: CHECKED (sum = 1.00)")
        else:
            print(f"❌ {name}: Sum should be 1.00, got {total}")


    def check_sum_equals_two(name, series):
        total = round(series.sum(), 2)
        if total == 2:
            print(f"✅ {name}: CHECKED (sum = 1.00)")
        else:
            print(f"❌ {name}: Sum should be 1.00, got {total}")

    check_equal("Summary GAA vs DF Market Value ratio", summary_gaa["Market Value"].sum() / df["Market Value"].sum(), 3)
    check_sum_equals_two("General GAA Q Allocation", general_gaa.loc[general_gaa["Q Allocation"] != "Total", "Q Allocation"])
    check_sum_equals_two("General GAA NQ Allocation", general_gaa.loc[general_gaa["NQ Allocation"] != "Total", "NQ Allocation"])


    check_equal("Equity Style Exposure Sum", large_cap_blend_percentage + large_cap_growth_percentage + large_cap_value_percentage + mid_cap_blend_percentage + mid_cap_growth_percentage + mid_cap_value_percentage + small_cap_blend_percentage + small_cap_growth_percentage + small_cap_value_percentage, 1)
    check_sum_equals_one("Credit Quality Exposure", credit_quality_exposure["Exposure"])
    check_sum_equals_one("Credit Issuer Exposure", credit_issuer_exposure["Value"])

    check_sum_equals_two("Sector Exposure", sectors["Percentage"])
    check_sum_equals_two("S&P 500 Sector Exposure", sectors["S&P 500"])
    check_equal("Security Type Market Value Coverage", security_type_table["Market Value"].sum() / df["Market Value"].sum(), 1)
    check_equal("Grouped by Ticker Coverage", grouped_by_ticker["Market Value"].sum() / df["Market Value"].sum(), 1)
    check_equal("Qualified + Non-Qualified Reconciliation", grouped_by_ticker_account_nq["Market Value"].sum() + grouped_by_ticker_account_qualified["Market Value"].sum(), df["Market Value"].sum())



    # DataFrames to format
    dfs = [high_expense_ratio, tax_cost_ratio, taxable_bonds_in_nq, muni_bond_qualified, active_passive, fund_families, aggregate_stylebox, credit_yield,credit_quality_exposure, individual_stocks_df,
        credit_issuer_exposure, tracking_error, countries_sector_rows, leverage_inverse_df, digital_df, dilution,retail_class, target_date_fund, sectors, summary_gaa, general_gaa, security_type_table]

    # Short display names (easy to edit later)
    rename_map = {
        "Market Value": "Market Value", "Estimated Interest": "Est. Interest", "Expense Paid": "Expense Paid",
        "expense_ratio": "Expense Ratio", "one_year_tax_cost_ratio": "1Y Tax Cost",
        "three_year_tax_cost_ratio": "3Y Tax Cost", "five_year_tax_cost_ratio": "5Y Tax Cost",
        "ten_year_tax_cost_ratio": "10Y Tax Cost", "current_yield": "Current Yield", "yield_to_maturity_generic": "Yield to Maturity",
        "annualized_daily_one_year_return": "1Y Return", "annualized_three_year_return": "3Y Return",
        "annualized_daily_five_year_return": "5Y Return",
        "max_drawdown_1y": "Max DD 1Y", "max_drawdown_3y": "Max DD 3Y", "max_drawdown_5y": "Max DD 5Y",
        "daily_standard_deviation_1y": "Vol 1Y", "daily_standard_deviation_3y": "Vol 3Y",
        "daily_standard_deviation_5y": "Vol 5Y", "actual_12b1": "12b-1 Fee",
        "alpha_1y_vs_category": "Alpha 1Y", "alpha_3y_vs_category": "Alpha 3Y",
        "alpha_5y_vs_category": "Alpha 5Y",
        "expense_ratio_rank": "Expense Ratio Rank", "broad_asset_class": "Broad Asset Class",
        "tracking_error_1y_vs_category": "TE 1Y", "tracking_error_3y_vs_category": "TE 3Y",
        "tracking_error_5y_vs_category": "TE 5Y", "tracking_error_10y_vs_category": "TE 10Y",
        "fund_family": "Fund Family", "category_name": "Category", "share_class": "Share Class", "asset_class":"Asset Class"}

    # Formatting groups
    DOLLAR_COLS = {"Market Value", "Est. Interest", "Expense Paid", "Active $", "Passive $"}
    PERCENT_COLS = {"Expense Ratio", "1Y Tax Cost", "3Y Tax Cost", "5Y Tax Cost", "10Y Tax Cost",
                    "Current Yield", "1Y Return", "3Y Return", "5Y Return", "Yield to Maturity",
                    "Max DD 1Y", "Max DD 3Y", "Max DD 5Y",
                    "Vol 1Y", "Vol 3Y", "Vol 5Y", "12b-1 Fee", "Current Allocation", "Q Allocation", "NQ Allocation",
                    "Growth", "Blend", "Value", "Percentage Paid", "Weight",
                    "Alpha 1Y", "Alpha 3Y", "Alpha 5Y", "Aggregate", "S&P 500","Difference", "Percentage","Exposure"}


    for df in dfs:
        df.rename(columns=rename_map, inplace=True)

        if df is not dilution and df is not muni_bond_qualified and "Market Value" in df.columns:# only apply filter to tables that actually have Market Value
            df.drop(df[df["Market Value"].astype(float) < 1000].index, inplace=True)  # delete rows below $1,000

        for c in DOLLAR_COLS & set(df.columns):
            df[c] = df[c].astype(float).round(0).map(lambda x: f"${x:,.0f}" if pd.notna(x) else "")
        for c in PERCENT_COLS & set(df.columns):
            df[c] = df[c].astype(float).mul(100).round(2).map(lambda x: f"{x:.2f}%" if pd.notna(x) else "")


    ABBREV_MAP = {
        "Russell Investments": "Russell", "John Hancock": "JH","Canadian Inflation Protected Fixed Income":"Inflation Protected FI Canada",
        "Tax-Managed": "Tax Mng", "Managed Account": "Mng Acct", "Trading--":"",
        "Managed": "Mng", "Account": "Acct", "Shares": "Shs", "GQG Partners US Select Quality Eq Fund Inv":"GQG US Quality Eq Fund Inv",
        "Share Class": "Shr Cl", "Class": "Cl", "Miscellaneous":"Misc",
        "Equity": "Eq", "Fixed Income": "FI", "Bond": "Bond",
        "Dual Directional": "DualDir", "Directional": "Dir",
        "Buffered": "Buff", "Buffer": "Buff", "Institutional":"Inst",
        "Defined Outcome": "DefOut", "Income": "Inc", "Emerging Markets":"EM", "Emerging-Markets":"EM",
        "United States": "US", "International": "Intl","Infrastructure":"Infra",
        "Emerging Markets": "EM", "Allocation": "Alloc","FdTM Cn Sr F":"","F/m":"","ETF":"",
        "Moderately": "Mod", "Aggressive": "Aggr","High Dividend ":"High Dvdnd", "Inv Grd":"IG",
        "Conservative": "Cons", "Global": "Glob", "Investment Grade":"IG", "Artificial Intelligence":"AI",
        "Series EF":""," Srs F USD":""," Srs E USD":"", "AT8":"","September":"Sept", "December":"Dec","New York":"NY",
        "Canadian Focused Small/Mid Cap Eq":"Canadian Small/Mid Cap", "Government Mortgage-Backed Bond":"Gov Mortgage-Backed Bond",
        "Short Term Target Date Portfolio":"Short Term Target Date","Target Date Portfolio":"Target Date",
        "Short-Term Inflation-Protected Bond":"ST Inflation-Protected Bond", "Securitized Bond - Diversified": "Securitized Bond - Dvrsfd",
        
    }
    def map_text(x, max_len):
        if pd.isna(x): return x          # Preserve NaNs
        out = x                         # Working copy
        #print(out)                      # Debug: original text
        #print("")
        for k, v in ABBREV_MAP.items(): # Iteratively apply abbreviations
            out = out.replace(k, v)     # Replace long → short
            if len(out) <= max_len:     # Stop once length constraint is met
                #print(len(out))         # Debug: final length
                break
        return out.strip()              # Clean whitespace before returning


    for df in dfs:
        if "Name" in df.columns:
            df["Name"] = df["Name"].apply(lambda x: map_text(x, 40))          # Shorten Name if present
        if "Category" in df.columns:
            df["Category"] = df["Category"].apply(lambda x: map_text(x, 22))  # Shorten category_name if present


    return {
    "high_expense_ratio": high_expense_ratio,
    "tax_cost_ratio": tax_cost_ratio,
    "taxable_bonds_in_nq": taxable_bonds_in_nq,
    "muni_bond_qualified": muni_bond_qualified,
    "active_passive": active_passive,
    "tracking_error": tracking_error,
    "countries_sector_rows": countries_sector_rows,
    "leverage_inverse_df": leverage_inverse_df,
    "digital_df": digital_df,
    "retail_class": retail_class,
    "target_date_fund": target_date_fund,
    "sectors": sectors,
    "summary_gaa": summary_gaa,
    "general_gaa": general_gaa,
    "q_gaa": q_gaa,
    "nq_gaa": nq_gaa,
    "aggregate_stylebox": aggregate_stylebox,
    "credit_yield": credit_yield,
    "credit_duration": credit_duration,
    "credit_issuer_exposure": credit_issuer_exposure,
    "credit_quality_exposure": credit_quality_exposure,
    "security_type_table": security_type_table,
    "performance_summary": performance_summary,
    "expense_summary": expense_summary,
    "mini_fixed_income": mini_fixed_income,
    "individual_stocks_df": individual_stocks_df,
    "fund_families": fund_families,
    "dilution": dilution}
