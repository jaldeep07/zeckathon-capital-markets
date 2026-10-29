"""
generate_charts.py
Generates institutional charts for Zeckathon Problem Statement 01 repository.
"""

import os
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd

def create_all_charts(out_dir="outputs/charts"):
    os.makedirs(out_dir, exist_ok=True)
    plt.style.use("seaborn-v0_8-whitegrid" if "seaborn-v0_8-whitegrid" in plt.style.available else "default")

    # 1. Peer Relative Valuation & RoIC Chart
    companies = ["BSE Ltd", "CDSL", "CAMS (Winner)", "KFin Tech"]
    pe_fy27 = [43.8, 43.4, 36.5, 44.5]
    ev_ebitda_fy27 = [34.5, 34.6, 25.8, 30.8]
    roic = [48.5, 62.0, 78.4, 34.2]

    x = np.arange(len(companies))
    width = 0.35

    fig, ax1 = plt.subplots(figsize=(9, 5.2), dpi=200)
    rects1 = ax1.bar(x - width/2, pe_fy27, width, label="FY27E P/E (x)", color="#1e3a8a")
    rects2 = ax1.bar(x + width/2, ev_ebitda_fy27, width, label="FY27E EV/EBITDA (x)", color="#0d9488")

    ax1.set_ylabel("Valuation Multiple (x)", fontsize=11, fontweight="bold", color="#1e293b")
    ax1.set_title("Peer Relative Valuation vs. Capital Efficiency (RoIC Ex-Cash)", fontsize=13, fontweight="bold", pad=15)
    ax1.set_xticks(x)
    ax1.set_xticklabels(companies, fontsize=10, fontweight="bold")
    ax1.set_ylim(0, 55)

    ax2 = ax1.twinx()
    line1 = ax2.plot(x, roic, color="#dc2626", marker="o", linewidth=2.5, markersize=8, label="RoIC Ex-Cash (%)")
    ax2.set_ylabel("RoIC Ex-Cash (%)", fontsize=11, fontweight="bold", color="#dc2626")
    ax2.set_ylim(0, 95)
    ax2.grid(False)

    for rect in rects1:
        h = rect.get_height()
        ax1.annotate(f"{h}x", xy=(rect.get_x() + rect.get_width()/2, h), xytext=(0, 3),
                     textcoords="offset points", ha="center", va="bottom", fontsize=9, fontweight="bold")
    for rect in rects2:
        h = rect.get_height()
        ax1.annotate(f"{h}x", xy=(rect.get_x() + rect.get_width()/2, h), xytext=(0, 3),
                     textcoords="offset points", ha="center", va="bottom", fontsize=9, fontweight="bold")
    for i, v in enumerate(roic):
        ax2.annotate(f"{v}%", xy=(x[i], v), xytext=(0, 8),
                     textcoords="offset points", ha="center", va="bottom", fontsize=9, fontweight="bold", color="#dc2626")

    lines1, labels1 = ax1.get_legend_handles_labels()
    lines2, labels2 = ax2.get_legend_handles_labels()
    ax1.legend(lines1 + lines2, labels1 + labels2, loc="upper left", frameon=True)
    plt.tight_layout()
    plt.savefig(f"{out_dir}/01_peer_valuation_vs_roic.png")
    plt.close()

    # 2. Mandatory Bear Shock Comparison (-30% Derivatives + -20% Market Drawdown)
    fig, ax = plt.subplots(figsize=(9, 5.2), dpi=200)
    rev_hit = [-26.5, -18.0, -9.8, -8.2]
    eps_hit = [-38.4, -26.8, -15.6, -15.2]
    price_downside = [-40.6, -32.0, -11.5, -27.2]

    w = 0.25
    r1 = ax.bar(x - w, rev_hit, w, label="FY27E Revenue Impact (%)", color="#64748b")
    r2 = ax.bar(x, eps_hit, w, label="FY27E EPS Impact (%)", color="#f59e0b")
    r3 = ax.bar(x + w, price_downside, w, label="Bear Target Price Downside (%)", color="#b91c1c")

    ax.axhline(0, color="black", linewidth=1)
    ax.set_ylabel("Percentage Change from Base / CMP (%)", fontsize=11, fontweight="bold")
    ax.set_title("Mandatory Bear Shock Impact (-30% F&O Turnover & -20% Equity Drawdown)", fontsize=12.5, fontweight="bold", pad=15)
    ax.set_xticks(x)
    ax.set_xticklabels(companies, fontsize=10, fontweight="bold")
    ax.set_ylim(-48, 5)

    for bars in [r1, r2, r3]:
        for bar in bars:
            h = bar.get_height()
            ax.annotate(f"{h:.1f}%", xy=(bar.get_x() + bar.get_width()/2, h), xytext=(0, -13),
                        textcoords="offset points", ha="center", va="top", fontsize=8.5, fontweight="bold")

    ax.legend(loc="lower right", frameon=True)
    plt.tight_layout()
    plt.savefig(f"{out_dir}/02_mandatory_bear_shock_impact.png")
    plt.close()

    # 3. CAMS Revenue, EBITDA & FCFF Trajectory (FY26A-FY31E)
    years = ["FY26A", "FY27E", "FY28E", "FY29E", "FY30E", "FY31E"]
    rev = [1480, 1765, 2101, 2490, 2913, 3378]
    ebitda = [688, 820, 1013, 1233, 1466, 1726]
    fcff = [468, 541, 676, 830, 994, 1178]
    margins = [46.5, 46.5, 48.2, 49.5, 50.3, 51.1]

    x_yr = np.arange(len(years))
    fig, ax1 = plt.subplots(figsize=(9, 5.2), dpi=200)
    b1 = ax1.bar(x_yr - 0.25, rev, 0.25, label="Operating Revenue (Rs Cr)", color="#1e40af")
    b2 = ax1.bar(x_yr, ebitda, 0.25, label="EBITDA (Rs Cr)", color="#0f766e")
    b3 = ax1.bar(x_yr + 0.25, fcff, 0.25, label="Free Cash Flow - FCFF (Rs Cr)", color="#15803d")

    ax1.set_ylabel("Rs Crore", fontsize=11, fontweight="bold")
    ax1.set_title("CAMS 3-Statement Financial Trajectory & Margin Expansion (FY26A-FY31E)", fontsize=12.5, fontweight="bold", pad=15)
    ax1.set_xticks(x_yr)
    ax1.set_xticklabels(years, fontsize=10, fontweight="bold")
    ax1.set_ylim(0, 3900)

    ax2 = ax1.twinx()
    ax2.plot(x_yr, margins, color="#d97706", marker="s", linewidth=2.5, label="EBITDA Margin (%)")
    ax2.set_ylabel("EBITDA Margin (%)", fontsize=11, fontweight="bold", color="#d97706")
    ax2.set_ylim(40, 56)
    ax2.grid(False)

    for i, m in enumerate(margins):
        ax2.annotate(f"{m:.1f}%", xy=(x_yr[i], m), xytext=(0, 7),
                     textcoords="offset points", ha="center", va="bottom", fontsize=9, fontweight="bold", color="#b45309")

    l1, lb1 = ax1.get_legend_handles_labels()
    l2, lb2 = ax2.get_legend_handles_labels()
    ax1.legend(l1 + l2, lb1 + lb2, loc="upper left", frameon=True)
    plt.tight_layout()
    plt.savefig(f"{out_dir}/03_cams_3statement_trajectory.png")
    plt.close()

    # 4. Risk-Adjusted Return Profile (Upside vs Downside)
    fig, ax = plt.subplots(figsize=(8.5, 4.8), dpi=200)
    upside = [4.3, 7.0, 23.5, 12.3]
    downside = [-40.6, -32.0, -11.5, -27.2]
    bull = [28.0, 29.1, 44.3, 35.1]

    y = np.arange(len(companies))
    ax.barh(y + 0.2, bull, height=0.22, label="Bull Case Upside (%)", color="#16a34a")
    ax.barh(y, upside, height=0.22, label="Base Case Upside (%)", color="#2563eb")
    ax.barh(y - 0.2, downside, height=0.22, label="Bear Case Downside (%)", color="#dc2626")

    ax.axvline(0, color="black", linewidth=1)
    ax.set_yticks(y)
    ax.set_yticklabels(companies, fontsize=10, fontweight="bold")
    ax.set_xlabel("12-Month Implied Return from CMP (%)", fontsize=11, fontweight="bold")
    ax.set_title("12-Month Risk-Reward Asymmetry Across Candidates (Bull / Base / Bear)", fontsize=12.5, fontweight="bold", pad=15)
    ax.set_xlim(-48, 52)
    ax.legend(loc="lower right", frameon=True)
    plt.tight_layout()
    plt.savefig(f"{out_dir}/04_risk_reward_asymmetry.png")
    plt.close()

if __name__ == "__main__":
    create_all_charts()
    print("All 4 charts generated successfully.")
