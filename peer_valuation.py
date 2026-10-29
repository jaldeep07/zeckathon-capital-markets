"""
peer_valuation.py
Computes Relative Valuation Multiples (P/E, EV/EBITDA, PEG, RoIC, Cash Conversion)
and ₹500 Crore Institutional Deployment Liquidity Analysis across all 4 candidates.
"""

import pandas as pd

def compute_peer_valuation_and_liquidity(csv_path="data/peer_financials_fy26_fy29.csv", fund_mandate_cr=500.0):
    df = pd.read_csv(csv_path)

    # Valuation Multiples
    df["FY27E_PE"] = (df["Market_Cap_Cr"] / df["FY27E_PAT_Cr"]).round(1)
    df["FY28E_PE"] = (df["Market_Cap_Cr"] / df["FY28E_PAT_Cr"]).round(1)
    df["FY27E_EV_EBITDA"] = (df["Enterprise_Value_Cr"] / df["FY27E_EBITDA_Cr"]).round(1)
    df["FY28E_EV_EBITDA"] = (df["Enterprise_Value_Cr"] / df["FY28E_EBITDA_Cr"]).round(1)
    df["FY27E_PEG"] = (df["FY27E_PE"] / df["EPS_CAGR_FY26_29E_Pct"]).round(2)

    # Cash Conversion Due Diligence
    df["OCF_to_EBITDA_Pct"] = ((df["OCF_FY26A_Cr"] / df["FY26A_EBITDA_Cr"]) * 100).round(1)
    df["FCF_to_PAT_Pct"] = ((df["FCF_FY26A_Cr"] / df["FY26A_PAT_Cr"]) * 100).round(1)
    df["Software_Cap_Pct_Rev"] = ((df["Capex_Intangible_Software_Cr"] / df["FY26A_Rev_Cr"]) * 100).round(2)

    # ₹500 Crore Deployment & Liquidity Analysis (at 15% of 3M ADV)
    df["Daily_Buy_15Pct_ADV_Cr"] = (df["ADV_3M_Cr"] * 0.15).round(2)
    df["Days_to_Deploy_500Cr"] = (fund_mandate_cr / df["Daily_Buy_15Pct_ADV_Cr"]).round(1)
    df["Fund_Ownership_Stake_Pct"] = ((fund_mandate_cr / df["Market_Cap_Cr"]) * 100).round(2)

    return df

if __name__ == "__main__":
    df_val = compute_peer_valuation_and_liquidity()
    print(df_val[["Company", "FY27E_PE", "FY27E_EV_EBITDA", "RoIC_Ex_Cash_Pct", "FCF_to_PAT_Pct", "Days_to_Deploy_500Cr"]])
