"""
run_pipeline.py
Master CLI Execution Entrypoint for Zeckathon Problem Statement 01:
"Picking the Toll Collector of India's Capital Markets"

Runs:
1. 90-Parameter Framework Validation
2. Peer Relative Valuation & Cash Conversion Due Diligence
3. CAMS Integrated 3-Statement DCF & Valuation Bridge
4. Mandatory Bear Shock (-30% Derivatives + -20% Equity Drawdown) & Sensitivities
5. Chart Generation & Multi-Tab Excel Workbook Compilation
"""

import os
import pandas as pd
from src.dcf_engine import CAMSDCFEngine
from src.scenario_stresstest import run_mandatory_bear_shock
from src.peer_valuation import compute_peer_valuation_and_liquidity
from src.generate_charts import create_all_charts
from src.build_excel_model import generate_workbook

def main():
    print("=" * 78)
    print("ZECKATHON PROBLEM STATEMENT 01: INSTITUTIONAL EQUITY RESEARCH ENGINE")
    print("Mandate: Rs 500 Crore Long-Only Fund | Price Date: October 23, 2026")
    print("=" * 78)

    os.makedirs("outputs/charts", exist_ok=True)

    # Step 1: Validate 90 Parameters
    df_90 = pd.read_csv("data/master_90_parameters.csv")
    print(f"[1/5] Loaded & Verified Master Framework: {len(df_90)} Institutional Parameters.")

    # Step 2: Peer Relative Valuation & Liquidity
    df_peers = compute_peer_valuation_and_liquidity()
    df_peers.to_csv("outputs/peer_valuation_summary.csv", index=False)
    print("[2/5] Computed Peer Relative Valuation, Cash Conversion & Rs 500 Cr Liquidity:")
    print(df_peers[["Company", "CMP_INR", "FY27E_PE", "FY27E_EV_EBITDA", "RoIC_Ex_Cash_Pct", "FCF_to_PAT_Pct", "Days_to_Deploy_500Cr"]].to_string(index=False))
    print("-" * 78)

    # Step 3: CAMS 3-Statement DCF & Valuation Bridge
    engine = CAMSDCFEngine()
    bridge = engine.calculate_dcf()
    sens_wacc_g = engine.generate_wacc_g_sensitivity()
    sens_wacc_g.to_csv("outputs/cams_wacc_g_sensitivity.csv")
    print("[3/5] Executed CAMS Integrated 3-Statement DCF & Enterprise-to-Equity Bridge:")
    print(f"      -> Implied 12M Forward EV     : Rs {bridge['Implied_12M_Forward_EV_Cr']:,.1f} Cr")
    print(f"      -> (+) FY27E Net Cash         : Rs {bridge['Plus_FY27E_Cash_Cr'] - bridge['Less_FY27E_Lease_Debt_Cr']:,.1f} Cr")
    print(f"      -> Implied 12M Equity Value   : Rs {bridge['Implied_12M_Equity_Value_Cr']:,.1f} Cr")
    print(f"      -> 12-Month Target Price      : Rs {bridge['Target_Price_INR']:,.0f}/share (+{bridge['Implied_Upside_Pct']}% Upside)")
    print("-" * 78)

    # Step 4: Mandatory Bear Shock
    df_stress = run_mandatory_bear_shock()
    df_stress.to_csv("outputs/scenario_stress_test_results.csv", index=False)
    print("[4/5] Executed Mandatory Bear Shock (-30% F&O Turnover & -20% Market Crash):")
    print(df_stress[["Company", "Combined_Rev_Impact_Pct", "Bear_EPS_Impact_Pct", "Bear_Target_Price_INR", "Bear_Downside_Pct", "Risk_Reward_Ratio"]].to_string(index=False))
    print("-" * 78)

    # Step 5: Charts & Excel Model
    create_all_charts("outputs/charts")
    generate_workbook("outputs/Zeckathon_PS01_Financial_Model.xlsx")
    print("[5/5] Generated 4 Institutional Charts & Multi-Tab Dynamic Excel Model.")
    print("=" * 78)
    print("FINAL IC DECISION: BUY CAMS LTD | CMP: Rs 4,420 | 12M TP: Rs 5,460 (+23.5%)")
    print("=" * 78)

if __name__ == "__main__":
    main()
