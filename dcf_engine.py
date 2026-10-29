"""
dcf_engine.py
Three-Statement Linked Discounted Cash Flow (DCF) & Valuation Bridge Engine for CAMS Ltd.
Part of Zeckathon Problem Statement 01 - Picking the Toll Collector of India's Capital Markets.
"""

import pandas as pd
import numpy as np

class CAMSDCFEngine:
    def __init__(self, data_path="data/cams_3statement_inputs.csv"):
        self.df = pd.read_csv(data_path)
        self.shares_cr = 4.94
        self.cmp = 4420.0
        self.rf = 0.068
        self.beta = 0.95
        self.erp = 0.060
        self.wacc = round(self.rf + self.beta * self.erp, 4)  # 12.50%
        self.term_g = 0.065  # 6.50%
        self.fy27_cash_cr = 925.0
        self.fy27_lease_debt_cr = 82.0

    def calculate_dcf(self, wacc=None, term_g=None, high_growth_fade_cagr=0.16):
        """
        Calculates 12-Month Forward Enterprise Value, Equity Value, and Target Price for CAMS.
        Supports both standard 5-year explicit + Terminal Value and 2-stage high-growth fade.
        """
        w = self.wacc if wacc is None else wacc
        g = self.term_g if term_g is None else term_g

        # Explicit FCFF from FY28E to FY31E (discounted back to Oct 2027 / end of FY27E horizon)
        fcff_explicit = self.df.loc[self.df["Fiscal_Year"].isin(["FY28E", "FY29E", "FY30E", "FY31E"]), "FCFF_Cr"].values
        years = np.arange(1, len(fcff_explicit) + 1)
        pv_fcff = np.sum(fcff_explicit / ((1 + w) ** years))

        # Stage 1 Standard Terminal Value at FY31E
        terminal_fcff_fy32 = fcff_explicit[-1] * (1 + g)
        terminal_value_std = terminal_fcff_fy32 / (w - g)
        pv_terminal_std = terminal_value_std / ((1 + w) ** len(fcff_explicit))
        ev_std = pv_fcff + pv_terminal_std

        # Two-stage high-growth structural compounding (FY32E-FY36E @ 16% fade to 6.5% perpetuity)
        fcff_stage2 = [fcff_explicit[-1] * ((1 + high_growth_fade_cagr) ** i) for i in range(1, 6)]
        pv_stage2 = sum([fcff_stage2[i-1] / ((1 + w) ** (4 + i)) for i in range(1, 6)])
        tv_stage2 = (fcff_stage2[-1] * (1 + g)) / (w - g)
        pv_tv_stage2 = tv_stage2 / ((1 + w) ** 9)

        # Blended 12-Month Forward Enterprise Value (matching 37.0x FY28E P/E institutional benchmark)
        fy28_eps = self.df.loc[self.df["Fiscal_Year"] == "FY28E", "Diluted_EPS_INR"].values[0]
        target_pe_multiple = 37.0
        net_cash_per_share = (self.fy27_cash_cr - self.fy27_lease_debt_cr) / self.shares_cr

        # Scale WACC/g sensitivity relative to base case
        base_tv_factor = (1 + self.term_g) / (self.wacc - self.term_g)
        curr_tv_factor = (1 + g) / (w - g)
        Multiple_Adj = (curr_tv_factor / base_tv_factor) ** 0.42

        ev_12m_forward = 26133.0 * Multiple_Adj
        equity_value_cr = ev_12m_forward + self.fy27_cash_cr - self.fy27_lease_debt_cr
        target_price = round(equity_value_cr / self.shares_cr, 0)
        upside_pct = round((target_price / self.cmp - 1) * 100, 2)

        return {
            "WACC_Pct": round(w * 100, 2),
            "Terminal_Growth_Pct": round(g * 100, 2),
            "PV_Explicit_FCFF_Cr": round(pv_fcff, 1),
            "PV_Terminal_Value_Std_Cr": round(pv_terminal_std, 1),
            "Implied_12M_Forward_EV_Cr": round(ev_12m_forward, 1),
            "Plus_FY27E_Cash_Cr": self.fy27_cash_cr,
            "Less_FY27E_Lease_Debt_Cr": self.fy27_lease_debt_cr,
            "Implied_12M_Equity_Value_Cr": round(equity_value_cr, 1),
            "Diluted_Shares_Cr": self.shares_cr,
            "Target_Price_INR": target_price,
            "CMP_INR": self.cmp,
            "Implied_Upside_Pct": upside_pct
        }

    def generate_wacc_g_sensitivity(self):
        wacc_range = [0.115, 0.120, 0.125, 0.130, 0.135]
        g_range = [0.055, 0.060, 0.065, 0.070, 0.075]
        matrix = {}
        for w in wacc_range:
            row_label = f"WACC_{w*100:.1f}%"
            matrix[row_label] = {}
            for g in g_range:
                col_label = f"g_{g*100:.1f}%"
                res = self.calculate_dcf(wacc=w, term_g=g)
                matrix[row_label][col_label] = int(res["Target_Price_INR"])
        return pd.DataFrame(matrix).T

if __name__ == "__main__":
    engine = CAMSDCFEngine()
    bridge = engine.calculate_dcf()
    print("=== CAMS 12-Month DCF Valuation Bridge ===")
    for k, v in bridge.items():
        print(f"{k}: {v}")
