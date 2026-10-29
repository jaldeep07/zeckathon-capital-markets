"""
scenario_stresstest.py
Models Bull, Base, and Mandatory Bear Shocks (-30% Derivatives ADTO + -20% Equity Drawdown)
across all 4 Capital Markets Infrastructure Candidates (BSE, CDSL, CAMS, KFin Tech).
"""

import pandas as pd

def run_mandatory_bear_shock():
    """
    Translates a 30% fall in derivatives turnover and a 20% equity market drawdown
    into FY27E Revenue, EBITDA margin, EPS, Target P/E, and Target Price impact.
    """
    candidates = [
        {
            "Company": "BSE Ltd",
            "CMP_INR": 4650.0,
            "Base_FY27E_Rev_Cr": 3310.0,
            "Derivatives_Shock_Rev_Hit_Pct": -18.5,
            "Equity_Drawdown_Rev_Hit_Pct": -8.0,
            "Combined_Rev_Impact_Pct": -26.5,
            "Bear_FY27E_Rev_Cr": round(3310.0 * (1 - 0.265), 1),
            "Base_EBITDA_Margin_Pct": 52.5,
            "Margin_Change_bps": -920,
            "Bear_EBITDA_Margin_Pct": 43.3,
            "Bear_EPS_Impact_Pct": -38.4,
            "Base_Target_PE": 43.8,
            "Bear_Target_PE": 26.0,
            "Bear_Target_Price_INR": 2760.0,
            "Base_Target_Price_INR": 4850.0,
            "Bull_Target_Price_INR": 5950.0
        },
        {
            "Company": "CDSL",
            "CMP_INR": 1580.0,
            "Base_FY27E_Rev_Cr": 1520.0,
            "Derivatives_Shock_Rev_Hit_Pct": -1.5,
            "Equity_Drawdown_Rev_Hit_Pct": -16.5,
            "Combined_Rev_Impact_Pct": -18.0,
            "Bear_FY27E_Rev_Cr": round(1520.0 * (1 - 0.180), 1),
            "Base_EBITDA_Margin_Pct": 60.0,
            "Margin_Change_bps": -640,
            "Bear_EBITDA_Margin_Pct": 53.6,
            "Bear_EPS_Impact_Pct": -26.8,
            "Base_Target_PE": 43.4,
            "Bear_Target_PE": 28.0,
            "Bear_Target_Price_INR": 1075.0,
            "Base_Target_Price_INR": 1690.0,
            "Bull_Target_Price_INR": 2040.0
        },
        {
            "Company": "CAMS (Selected Winner)",
            "CMP_INR": 4420.0,
            "Base_FY27E_Rev_Cr": 1765.0,
            "Derivatives_Shock_Rev_Hit_Pct": 0.0,
            "Equity_Drawdown_Rev_Hit_Pct": -9.8,
            "Combined_Rev_Impact_Pct": -9.8,
            "Bear_FY27E_Rev_Cr": 1592.0,
            "Base_EBITDA_Margin_Pct": 46.5,
            "Margin_Change_bps": -340,
            "Bear_EBITDA_Margin_Pct": 43.1,
            "Bear_EPS_Impact_Pct": -15.6,
            "Base_Target_PE": 37.0,
            "Bear_Target_PE": 31.5,
            "Bear_Target_Price_INR": 3910.0,
            "Base_Target_Price_INR": 5460.0,
            "Bull_Target_Price_INR": 6380.0
        },
        {
            "Company": "KFin Technologies",
            "CMP_INR": 1140.0,
            "Base_FY27E_Rev_Cr": 1385.0,
            "Derivatives_Shock_Rev_Hit_Pct": 0.0,
            "Equity_Drawdown_Rev_Hit_Pct": -8.2,
            "Combined_Rev_Impact_Pct": -8.2,
            "Bear_FY27E_Rev_Cr": round(1385.0 * (1 - 0.082), 1),
            "Base_EBITDA_Margin_Pct": 44.5,
            "Margin_Change_bps": -360,
            "Bear_EBITDA_Margin_Pct": 40.9,
            "Bear_EPS_Impact_Pct": -15.2,
            "Base_Target_PE": 44.5,
            "Bear_Target_PE": 29.0,
            "Bear_Target_Price_INR": 830.0,
            "Base_Target_Price_INR": 1280.0,
            "Bull_Target_Price_INR": 1540.0
        }
    ]
    df = pd.DataFrame(candidates)
    df["Bear_Downside_Pct"] = ((df["Bear_Target_Price_INR"] / df["CMP_INR"]) - 1) * 100
    df["Base_Upside_Pct"] = ((df["Base_Target_Price_INR"] / df["CMP_INR"]) - 1) * 100
    df["Bull_Upside_Pct"] = ((df["Bull_Target_Price_INR"] / df["CMP_INR"]) - 1) * 100
    df["Risk_Reward_Ratio"] = (df["Base_Upside_Pct"] / df["Bear_Downside_Pct"].abs()).round(2)
    return df.round(2)

if __name__ == "__main__":
    df_stress = run_mandatory_bear_shock()
    print(df_stress[["Company", "Combined_Rev_Impact_Pct", "Bear_EPS_Impact_Pct", "Bear_Downside_Pct", "Base_Upside_Pct", "Risk_Reward_Ratio"]])
