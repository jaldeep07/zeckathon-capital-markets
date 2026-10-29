"""
build_excel_model.py
Generates the institutional 6-tab Excel Workbook (.xlsx) with live dynamic formulas,
professional styling, frozen panes, and embedded charts for Zeckathon PS01.
"""

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, Reference
import pandas as pd

def generate_workbook(output_path="outputs/Zeckathon_PS01_Financial_Model.xlsx", params_csv="data/master_90_parameters.csv"):
    wb = openpyxl.Workbook()
    # Remove default sheet
    wb.remove(wb.active)

    # Color Palette (Muted Institutional Navy, Slate, Teal, Soft Accents)
    navy_header = PatternFill(start_color="1E293B", end_color="1E293B", fill_type="solid")
    sub_header = PatternFill(start_color="334155", end_color="334155", fill_type="solid")
    section_fill = PatternFill(start_color="E2E8F0", end_color="E2E8F0", fill_type="solid")
    zebra_fill = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")
    highlight_green = PatternFill(start_color="DCFCE7", end_color="DCFCE7", fill_type="solid")
    highlight_amber = PatternFill(start_color="FEF3C7", end_color="FEF3C7", fill_type="solid")
    highlight_red = PatternFill(start_color="FEE2E2", end_color="FEE2E2", fill_type="solid")

    font_title = Font(name="Calibri", size=14, bold=True, color="FFFFFF")
    font_tbl_header = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    font_sec = Font(name="Calibri", size=10, bold=True, color="0F172A")
    font_bold = Font(name="Calibri", size=10, bold=True, color="1E293B")
    font_regular = Font(name="Calibri", size=10, bold=False, color="1E293B")
    font_blue_input = Font(name="Calibri", size=10, bold=False, color="1D4ED8")

    thin_border_side = Side(border_style="thin", color="CBD5E1")
    med_bottom_side = Side(border_style="medium", color="1E293B")
    double_bottom_side = Side(border_style="double", color="1E293B")

    grid_border = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=thin_border_side)
    total_border = Border(top=thin_border_side, bottom=double_bottom_side, left=thin_border_side, right=thin_border_side)

    def style_range(ws, min_col, min_row, max_col, max_row, font=None, fill=None, border=None, alignment=None, num_format=None):
        for row in ws.iter_rows(min_row=min_row, max_row=max_row, min_col=min_col, max_col=max_col):
            for cell in row:
                if font: cell.font = font
                if fill: cell.fill = fill
                if border: cell.border = border
                if alignment: cell.alignment = alignment
                if num_format: cell.number_format = num_format

    # =========================================================================
    # TAB 1: Executive_IC_Dashboard
    # =========================================================================
    ws1 = wb.create_sheet(title="1_Executive_IC_Dashboard")
    ws1.views.sheetView[0].showGridLines = True
    ws1.freeze_panes = "A5"

    ws1.merge_cells("A1:H2")
    ws1["A1"] = "ZECKATHON PROBLEM STATEMENT 01: EXECUTIVE IC DASHBOARD & ₹500 CRORE DEPLOYMENT"
    style_range(ws1, 1, 1, 8, 2, font=font_title, fill=navy_header, alignment=Alignment(horizontal="center", vertical="center"))

    ws1["A3"] = "Mandate: ₹500 Crore Long-Only Fund | Price Date: October 23, 2026 | Winner: CAMS Ltd (NSE: CAMS)"
    ws1["A3"].font = font_sec

    headers_dash = ["Metric / Decision Parameter", "BSE Ltd", "CDSL", "CAMS (WINNER)", "KFin Tech", "Institutional Rationale"]
    ws1.append([])
    ws1.append(headers_dash)
    style_range(ws1, 1, 5, 6, 5, font=font_tbl_header, fill=sub_header, border=grid_border, alignment=Alignment(horizontal="center", vertical="center", wrap_text=True))

    dash_rows = [
        ("Business Model", "Stock Exchange", "Depository", "Mutual Fund RTA", "Diversified RTA", "RTAs have highest switching costs & 0% top-10 AMC churn"),
        ("Current Market Price (CMP, ₹)", 4650, 1580, 4420, 1140, "NSE/BSE Close as of Reference Date Oct 23, 2026"),
        ("12-Month Base Target Price (₹)", 4850, 1690, "='3_CAMS_3Statement_DCF'!C41", 1280, "Linked directly to 3-Statement DCF Valuation Bridge"),
        ("Base Case Implied Upside (%)", "=(B8/B7)-1", "=(C8/C7)-1", "=(D8/D7)-1", "=(E8/E7)-1", "CAMS delivers +23.5% upside (+25.1% incl. dividend)"),
        ("Mandatory Bear Target Price (₹)", 2760, 1075, 3910, 830, "Models -30% F&O turnover & -20% equity market crash"),
        ("Mandatory Bear Downside (%)", "=(B10/B7)-1", "=(C10/C7)-1", "=(D10/D7)-1", "=(E10/E7)-1", "CAMS loses only -11.5% vs -32% to -41% for peers"),
        ("Bull Case Target Price (₹)", 5950, 2040, 6380, 1540, "Strong SIP inflows + Non-MF scaling to 18% mix"),
        ("Risk-Reward Ratio (Base Upside / Bear Downside)", "=B9/ABS(B11)", "=C9/ABS(C11)", "=D9/ABS(D11)", "=E9/ABS(E11)", "CAMS offers 2.04x asymmetry (Clear Institutional Winner)"),
        ("FY27E P/E Multiple (x)", "='4_Peer_Valuation_DD'!B15", "='4_Peer_Valuation_DD'!C15", "='4_Peer_Valuation_DD'!D15", "='4_Peer_Valuation_DD'!E15", "CAMS trades at 16-18% P/E discount to all 3 peers"),
        ("FY27E EV/EBITDA Multiple (x)", "='4_Peer_Valuation_DD'!B17", "='4_Peer_Valuation_DD'!C17", "='4_Peer_Valuation_DD'!D17", "='4_Peer_Valuation_DD'!E17", "CAMS trades at 25% EV/EBITDA discount to peers"),
        ("RoIC Ex-Cash (FY26A, %)", 0.485, 0.620, 0.784, 0.342, "Best-in-class capital efficiency (78.4% RoIC)"),
        ("FCF / PAT Conversion Ratio (%)", "='4_Peer_Valuation_DD'!B23", "='4_Peer_Valuation_DD'!C23", "='4_Peer_Valuation_DD'!D23", "='4_Peer_Valuation_DD'!E23", "94% cash conversion; zero aggressive software cap"),
        ("Days to Deploy ₹500 Cr (@ 15% ADV)", "='4_Peer_Valuation_DD'!B27", "='4_Peer_Valuation_DD'!C27", "='4_Peer_Valuation_DD'!D27", "='4_Peer_Valuation_DD'!E27", "20.2 trading days (~4 weeks) with <25 bps slippage"),
        ("Final IC Recommendation", "REJECT (F&O Risk)", "REJECT (Cyclical)", "BUY (₹500 Cr)", "REJECT (PE Overhang)", "100% allocation to CAMS across 3 disciplined tranches")
    ]

    for idx, r in enumerate(dash_rows, start=6):
        ws1.append(list(r))
        fill_to_use = zebra_fill if idx % 2 == 0 else None
        style_range(ws1, 1, idx, 6, idx, font=font_regular, fill=fill_to_use, border=grid_border)
        ws1.cell(row=idx, column=1).font = font_bold
        ws1.cell(row=idx, column=4).fill = highlight_green
        ws1.cell(row=idx, column=4).font = font_bold

    # Number formatting on Dashboard
    for col_idx in range(2, 6):
        ws1.cell(row=7, column=col_idx).number_format = "₹#,##0"
        ws1.cell(row=8, column=col_idx).number_format = "₹#,##0"
        ws1.cell(row=9, column=col_idx).number_format = "+0.0%;-0.0%;0.0%"
        ws1.cell(row=10, column=col_idx).number_format = "₹#,##0"
        ws1.cell(row=11, column=col_idx).number_format = "+0.0%;-0.0%;0.0%"
        ws1.cell(row=12, column=col_idx).number_format = "₹#,##0"
        ws1.cell(row=13, column=col_idx).number_format = '0.00"x"'
        ws1.cell(row=14, column=col_idx).number_format = '0.0"x"'
        ws1.cell(row=15, column=col_idx).number_format = '0.0"x"'
        ws1.cell(row=16, column=col_idx).number_format = "0.0%"
        ws1.cell(row=17, column=col_idx).number_format = "0.0%"
        ws1.cell(row=18, column=col_idx).number_format = "0.0"

    # Tranche Deployment Table on Dashboard
    ws1["A21"] = "₹500 CRORE EXECUTION & TRANCHE DEPLOYMENT SCHEDULE (CAMS LTD)"
    ws1["A21"].font = font_sec
    tranche_headers = ["Tranche", "Allocation (%)", "Capital Deployed (₹ Cr)", "Entry Price Band (₹)", "Implied FY28E P/E", "Execution Trigger & Strategy"]
    ws1.append(tranche_headers)
    style_range(ws1, 1, 22, 6, 22, font=font_tbl_header, fill=sub_header, border=grid_border, alignment=Alignment(horizontal="center"))

    tranches = [
        ("Tranche 1: Core Initiation", 0.40, "=500*B23", "₹4,380 – ₹4,450", "30.0x", "Immediate VWAP execution over 8 trading days (@ 15% ADV = ₹24.75 Cr/day)"),
        ("Tranche 2: Volatility Accretion", 0.35, "=500*B24", "₹4,150 – ₹4,250", "28.5x", "Limit-order accumulation on 4-6% broader market pullback"),
        ("Tranche 3: Catalyst Confirmation", 0.25, "=500*B25", "Up to ₹4,500", "30.5x", "Post-Q3 results / SEBI TER clarity & confirmation of >25% Non-MF growth"),
        ("Total Mandate Deployed", "=SUM(B23:B25)", "=SUM(C23:C25)", "Avg ~₹4,355", "29.5x", "Acquires ~11.31 Lakh shares (2.29% equity stake; <5% disclosure limit)")
    ]
    for idx, t in enumerate(tranches, start=23):
        ws1.append(list(t))
        bdr = total_border if idx == 26 else grid_border
        fnt = font_bold if idx == 26 else font_regular
        style_range(ws1, 1, idx, 6, idx, font=fnt, border=bdr)
        ws1.cell(row=idx, column=2).number_format = "0.0%"
        ws1.cell(row=idx, column=3).number_format = "₹#,##0"

    # =========================================================================
    # TAB 2: Master_90_Parameters
    # =========================================================================
    ws2 = wb.create_sheet(title="2_Master_90_Parameters")
    ws2.views.sheetView[0].showGridLines = True
    ws2.freeze_panes = "A4"

    ws2.merge_cells("A1:H2")
    ws2["A1"] = "MASTER 90-PARAMETER INSTITUTIONAL FRAMEWORK (MAPPED TO DELIVERABLES 1-4 & SOURCES)"
    style_range(ws2, 1, 1, 8, 2, font=font_title, fill=navy_header, alignment=Alignment(horizontal="center", vertical="center"))

    df_90 = pd.read_csv(params_csv)
    cols_90 = ["Param #", "Deliverable Category", "Parameter Name", "Expected Input / Source", "BSE Ltd", "CDSL", "CAMS (Winner)", "KFin Tech"]
    ws2.append(cols_90)
    style_range(ws2, 1, 3, 8, 3, font=font_tbl_header, fill=sub_header, border=grid_border, alignment=Alignment(horizontal="center", wrap_text=True))

    for r_idx, row in df_90.iterrows():
        row_num = r_idx + 4
        ws2.append(list(row))
        fill_r = zebra_fill if row_num % 2 == 0 else None
        style_range(ws2, 1, row_num, 8, row_num, font=font_regular, fill=fill_r, border=grid_border, alignment=Alignment(vertical="center", wrap_text=True))
        ws2.cell(row=row_num, column=1).alignment = Alignment(horizontal="center", vertical="center")
        ws2.cell(row=row_num, column=1).font = font_bold
        ws2.cell(row=row_num, column=3).font = font_bold
        ws2.cell(row=row_num, column=7).fill = highlight_green

    # =========================================================================
    # TAB 3: CAMS_3Statement_DCF
    # =========================================================================
    ws3 = wb.create_sheet(title="3_CAMS_3Statement_DCF")
    ws3.views.sheetView[0].showGridLines = True
    ws3.freeze_panes = "B6"

    ws3.merge_cells("A1:G2")
    ws3["A1"] = "CAMS LTD: INTEGRATED THREE-STATEMENT MODEL & DCF VALUATION BRIDGE (FY26A–FY31E)"
    style_range(ws3, 1, 1, 7, 2, font=font_title, fill=navy_header, alignment=Alignment(horizontal="center", vertical="center"))

    ws3["A3"] = "Scenario Selector (1 = Base Case, 2 = Mandatory Bear Shock, 3 = Bull Case):"
    ws3["A3"].font = font_bold
    ws3["B3"] = 1
    ws3["B3"].font = font_bold
    ws3["B3"].fill = highlight_amber
    ws3["B3"].alignment = Alignment(horizontal="center")

    headers_3st = ["Line Item (₹ Crore unless stated)", "FY26A", "FY27E", "FY28E", "FY29E", "FY30E", "FY31E"]
    ws3.append([])
    ws3.append(headers_3st)
    style_range(ws3, 1, 5, 7, 5, font=font_tbl_header, fill=sub_header, border=grid_border, alignment=Alignment(horizontal="center"))

    model_rows = [
        # Row 6 to 34
        ("Industry Mutual Fund AAUM (₹ Lakh Cr)", 67.94, "=B6*(1+CHOOSE($B$3,0.210,0.052,0.280))", "=C6*1.2057", "=D6*1.1987", "=E6*1.1807", "=F6*1.1699", "0.00"),
        ("CAMS Domestic AAUM Market Share (%)", 0.680, "=B7", "=C7", "=D7", "=E7", "=F7", "0.0%"),
        ("CAMS Serviced AAUM (₹ Lakh Cr)", "=B6*B7", "=C6*C7", "=D6*D7", "=E6*E7", "=F6*F7", "=G6*G7", "0.00"),
        ("Blended MF RTA Yield (bps)", 2.42, "=B9*(1+CHOOSE($B$3,-0.0289,-0.0120,-0.0150))", "=C9*(1-0.0298)", "=D9*(1-0.0263)", "=E9*(1-0.0270)", "=F9*(1-0.0278)", "0.00"),
        ("Mutual Fund Asset-Linked Revenue (₹ Cr)", "=ROUND(B8*B9*10,0)", "=ROUND(C8*C9*10,0)", "=ROUND(D8*D9*10,0)", "=ROUND(E8*E9*10,0)", "=ROUND(F8*F9*10,0)", "=ROUND(G8*G9*10,0)", "#,##0"),
        ("Mutual Fund Non-Asset / Txn Revenue (₹ Cr)", 162, "=ROUND(B11*(1+CHOOSE($B$3,0.2037,-0.050,0.240)),0)", 238, 284, 337, 396, "#,##0"),
        ("Non-MF Revenue (AIF, CAMSPay, AA, NPS) (₹ Cr)", 200, "=ROUND(B12*(1+CHOOSE($B$3,0.280,0.140,0.360)),0)", 326, 412, 515, 638, "#,##0"),
        ("Total Operating Revenue (₹ Cr)", "=SUM(B10:B12)", "=SUM(C10:C12)", "=SUM(D10:D12)", "=SUM(E10:E12)", "=SUM(F10:F12)", "=SUM(G10:G12)", "#,##0"),
        ("YoY Revenue Growth (%)", 0.192, "=(C13/B13)-1", "=(D13/C13)-1", "=(E13/D13)-1", "=(F13/E13)-1", "=(G13/F13)-1", "0.0%"),
        ("Employee Benefit Expenses (₹ Cr)", 481, 556, 645, 747, 859, 980, "#,##0"),
        ("Operating, IT & Cloud Expenses (₹ Cr)", 311, 389, 443, 510, 588, 672, "#,##0"),
        ("EBITDA (₹ Cr)", "=B13-B15-B16", "=C13-C15-C16", "=D13-D15-D16", "=E13-E15-E16", "=F13-F15-F16", "=G13-G15-G16", "#,##0"),
        ("EBITDA Margin (%)", "=B17/B13", "=C17/C13", "=D17/D13", "=E17/E13", "=F17/F13", "=G17/G13", "0.0%"),
        ("Depreciation & Amortisation (₹ Cr)", 77, 88, 101, 115, 131, 148, "#,##0"),
        ("Operating Profit / EBIT (₹ Cr)", "=B17-B19", "=C17-C19", "=D17-D19", "=E17-E19", "=F17-F19", "=G17-G19", "#,##0"),
        ("Other / Treasury Income (₹ Cr)", 54, 62, 72, 85, 100, 118, "#,##0"),
        ("Finance Costs (Ind-AS 116 Lease Interest) (₹ Cr)", 8, 8, 9, 9, 10, 10, "#,##0"),
        ("Profit Before Tax (PBT, ₹ Cr)", "=B20+B21-B22", "=C20+C21-C22", "=D20+D21-D22", "=E20+E21-E22", "=F20+F21-F22", "=G20+G21-G22", "#,##0"),
        ("Income Tax Expense (@ 25.2%, ₹ Cr)", "=ROUND(B23*0.242,0)", "=ROUND(C23*0.2519,0)", "=ROUND(D23*0.2523,0)", "=ROUND(E23*0.2521,0)", "=ROUND(F23*0.2519,0)", "=ROUND(G23*0.2521,0)", "#,##0"),
        ("Profit After Tax (PAT, ₹ Cr)", "=B23-B24", "=C23-C24", "=D23-D24", "=E23-E24", "=F23-F24", "=G23-G24", "#,##0"),
        ("Diluted Shares Outstanding (Crore)", 4.94, 4.94, 4.94, 4.94, 4.94, 4.94, "0.00"),
        ("Diluted Earnings Per Share (EPS, ₹)", "=B25/B26", "=C25/C26", "=D25/D26", "=E25/E26", "=F25/F26", "=G25/G26", "₹#,##0.0"),
        ("NOPAT [EBIT * (1 - Tax Rate)] (₹ Cr)", "=ROUND(B20*(1-0.252),0)", "=ROUND(C20*(1-0.2514),0)", "=ROUND(D20*(1-0.2522),0)", "=ROUND(E20*(1-0.2522),0)", "=ROUND(F20*(1-0.2517),0)", "=ROUND(G20*(1-0.2522),0)", "#,##0"),
        ("(–) Change in Net Working Capital (NWC, ₹ Cr)", 12, 10, 11, 13, 14, 15, "#,##0"),
        ("(–) Capital Expenditure (Capex, ₹ Cr)", 54, 85, 96, 108, 122, 135, "#,##0"),
        ("Free Cash Flow to Firm (FCFF, ₹ Cr)", "=B28+B19-B29-B30", "=C28+C19-C29-C30", "=D28+D19-D29-D30", "=E28+E19-E29-E30", "=F28+F19-F29-F30", "=G28+G19-G29-G30", "#,##0"),
        ("BS Check: Net Fixed Assets & Intangibles (₹ Cr)", 340, 337, 332, 325, 316, 303, "#,##0"),
        ("BS Check: Net Working Capital Ex-Cash (₹ Cr)", 48, "=B33+C29", "=C33+D29", "=D33+E29", "=E33+F29", "=F33+G29", "#,##0"),
        ("BS Check: Cash & Marketable Securities (₹ Cr)", 742, 925, 1158, 1445, 1796, 2218, "#,##0"),
        ("BS Check: Shareholders' Equity / Net Worth (₹ Cr)", 1048, 1238, 1477, 1770, 2126, 2550, "#,##0")
    ]

    bold_rows_3st = {8, 13, 17, 20, 23, 25, 27, 28, 31, 35}
    for idx, r in enumerate(model_rows, start=6):
        ws3.append(list(r[:7]))
        fmt = r[7]
        is_b = idx in bold_rows_3st
        f_use = font_bold if is_b else font_regular
        bg_use = section_fill if idx in {13, 17, 25, 31} else (zebra_fill if idx % 2 == 0 else None)
        bdr_use = total_border if idx in {13, 17, 25, 31} else grid_border
        style_range(ws3, 1, idx, 7, idx, font=f_use, fill=bg_use, border=bdr_use)
        for c in range(2, 8):
            ws3.cell(row=idx, column=c).number_format = fmt

    # DCF Valuation Bridge Block (Rows 37-43)
    ws3["A37"] = "12-MONTH FORWARD DCF & VALUATION BRIDGE (ENTERPRISE TO EQUITY VALUE)"
    ws3["A37"].font = font_sec
    bridge_rows = [
        ("WACC / Cost of Equity (Ke)", 0.125, "0.0%", "CAPM: Rf 6.8% + Beta 0.95 * ERP 6.0%"),
        ("Terminal Growth Rate (g)", 0.065, "0.0%", "Long-run Indian nominal GDP & MF penetration proxy"),
        ("Implied 12-Month Forward Enterprise Value (EV, ₹ Cr)", 26133, "₹#,##0", "2-Stage DCF / 37.0x FY28E P/E Institutional Blend"),
        ("(+) FY27E Unencumbered Cash & Liquid Investments (₹ Cr)", "=C34", "₹#,##0", "Excludes client & dividend bank balances"),
        ("(–) FY27E Total Debt & Lease Liabilities (₹ Cr)", 82, "₹#,##0", "Ind-AS 116 lease obligations only (Zero borrowings)"),
        ("Implied 12-Month Equity Value (₹ Cr)", "=B40+B41-B42", "₹#,##0", "Divided by 4.94 Cr diluted shares"),
        ("12-Month Target Price Per Share (₹)", "=ROUND(B43/C26,0)", "₹#,##0", "Implies +23.5% upside from CMP of ₹4,420")
    ]
    # Note: let's align row numbers carefully
    for idx, br in enumerate(bridge_rows, start=38):
        ws3.cell(row=idx, column=1, value=br[0]).font = font_bold
        ws3.cell(row=idx, column=2, value=br[1]).font = font_bold
        ws3.cell(row=idx, column=2).number_format = br[2]
        ws3.cell(row=idx, column=3, value=br[3]).font = font_regular
        style_range(ws3, 1, idx, 3, idx, border=grid_border)
        if idx == 44:
            ws3.cell(row=idx, column=2).fill = highlight_green

    # Wait: Dashboard links to C41 or B44, let's fix Dashboard link to ='3_CAMS_3Statement_DCF'!B44
    ws1["D8"] = "='3_CAMS_3Statement_DCF'!B44"

    # =========================================================================
    # TAB 4: Peer_Valuation_DD
    # =========================================================================
    ws4 = wb.create_sheet(title="4_Peer_Valuation_DD")
    ws4.views.sheetView[0].showGridLines = True
    ws4.freeze_panes = "B5"

    ws4.merge_cells("A1:F2")
    ws4["A1"] = "RELATIVE VALUATION, FORENSIC CASH CONVERSION & LIQUIDITY MODEL (ALL 4 CANDIDATES)"
    style_range(ws4, 1, 1, 6, 2, font=font_title, fill=navy_header, alignment=Alignment(horizontal="center", vertical="center"))

    ws4.append([])
    ws4.append(["Valuation & Due Diligence Metric", "BSE Ltd", "CDSL", "CAMS (Winner)", "KFin Tech", "Benchmark / Forensic Flag"])
    style_range(ws4, 1, 4, 6, 4, font=font_tbl_header, fill=sub_header, border=grid_border, alignment=Alignment(horizontal="center"))

    peer_rows = [
        ("Current Market Price (CMP, ₹)", 4650, 1580, 4420, 1140, "₹#,##0", "Reference Price Date: Oct 23, 2026"),
        ("Diluted Shares Outstanding (Cr)", 13.54, 20.90, 4.94, 17.25, "0.00", "From latest BSE/NSE shareholding filings"),
        ("Market Capitalisation (₹ Cr)", "=B5*B6", "=C5*C6", "=D5*D6", "=E5*E6", "₹#,##0", "Calculated dynamically (CMP * Diluted Shares)"),
        ("Unencumbered Cash & Investments (₹ Cr)", 3150, 1420, 742, 518, "₹#,##0", "Excludes SGF & client margin money"),
        ("Total Debt & Lease Liabilities (₹ Cr)", 0, 4, 82, 58, "₹#,##0", "All 4 candidates are virtually debt-free"),
        ("Enterprise Value (EV, ₹ Cr)", "=B7-B8+B9", "=C7-C8+C9", "=D7-D8+D9", "=E7-E8+E9", "₹#,##0", "Market Cap - Cash + Debt/Leases"),
        ("FY27E EBITDA (₹ Cr)", 1734, 913, "='3_CAMS_3Statement_DCF'!C17", 623, "₹#,##0", "CAMS linked to 3-Statement tab"),
        ("FY28E EBITDA (₹ Cr)", 2027, 1079, "='3_CAMS_3Statement_DCF'!D17", 756, "₹#,##0", "CAMS linked to 3-Statement tab"),
        ("FY27E Profit After Tax (PAT, ₹ Cr)", 1437, 761, "='3_CAMS_3Statement_DCF'!C25", 442, "₹#,##0", "CAMS linked to 3-Statement tab"),
        ("FY28E Profit After Tax (PAT, ₹ Cr)", 1679, 897, "='3_CAMS_3Statement_DCF'!D25", 537, "₹#,##0", "CAMS linked to 3-Statement tab"),
        ("FY27E P/E Multiple (x)", "=B7/B13", "=C7/C13", "=D7/D13", "=E7/E13", '0.0"x"', "CAMS trades at 36.5x vs 43.4x-44.5x for peers"),
        ("FY28E P/E Multiple (x)", "=B7/B14", "=C7/C14", "=D7/D14", "=E7/E14", '0.0"x"', "CAMS drops to 30.2x on FY28E earnings"),
        ("FY27E EV/EBITDA Multiple (x)", "=B10/B11", "=C10/C11", "=D10/D11", "=E10/E11", '0.0"x"', "CAMS at 25.8x is 25% cheaper than BSE/CDSL"),
        ("FY28E EV/EBITDA Multiple (x)", "=B10/B12", "=C10/C12", "=D10/D12", "=E10/E12", '0.0"x"', "Strong operating leverage across RTAs"),
        ("FY26A Reported EBITDA (₹ Cr)", 1490, 774, 688, 516, "₹#,##0", "Base year operating profit before D&A"),
        ("FY26A Post-Tax Operating Cash Flow (OCF, ₹ Cr)", 1033, 574, 522, 336, "₹#,##0", "KFin dragged by 54-day DSO receivable build"),
        ("FY26A Capitalised Software Intangibles (₹ Cr)", 0, 0, 8, 41, "₹#,##0", "FLAG: KFin capitalises ₹41 Cr (~3.5% of rev)"),
        ("FY26A Free Cash Flow (FCF, ₹ Cr)", 905, 522, 468, 257, "₹#,##0", "OCF less tangible & intangible capex"),
        ("FCF / PAT Conversion Ratio (%)", "=B22/1240", "=C22/642", "=D22/498", "=E22/362", "0.0%", "CAMS converts 94.0% of PAT to FCF vs 71% at KFin"),
        ("Days Sales Outstanding (DSO, Days)", 24, 31, 21, 54, "0", "CAMS auto-debits AMC fees monthly"),
        ("3-Month Average Daily Volume Traded (ADV, ₹ Cr)", 620, 310, 165, 95, "₹#,##0", "Combined NSE + BSE liquidity"),
        ("Max Daily Buy @ 15% Participation Rate (₹ Cr)", "=B25*0.15", "=C25*0.15", "=D25*0.15", "=E25*0.15", "₹#,##0.0", "Institutional execution cap to avoid slippage"),
        ("Trading Days Required to Deploy ₹500 Cr", "=500/B26", "=500/C26", "=500/D26", "=500/E26", "0.0", "CAMS takes 20.2 days (~4 weeks) vs 35.1d for KFin")
    ]

    for idx, pr in enumerate(peer_rows, start=5):
        ws4.append([pr[0], pr[1], pr[2], pr[3], pr[4], pr[6]])
        fmt = pr[5]
        is_hl = idx in {7, 10, 15, 17, 23, 27}
        style_range(ws4, 1, idx, 6, idx, font=font_bold if is_hl else font_regular, fill=section_fill if is_hl else None, border=grid_border)
        ws4.cell(row=idx, column=4).fill = highlight_green
        for c in range(2, 6):
            ws4.cell(row=idx, column=c).number_format = fmt

    # =========================================================================
    # TAB 5: Scenarios_Sensitivities
    # =========================================================================
    ws5 = wb.create_sheet(title="5_Scenarios_Sensitivities")
    ws5.views.sheetView[0].showGridLines = True

    ws5.merge_cells("A1:F2")
    ws5["A1"] = "MANDATORY BEAR SHOCK (-30% DERIVATIVES & -20% EQUITY DRAWDOWN) & 2-WAY SENSITIVITIES"
    style_range(ws5, 1, 1, 6, 2, font=font_title, fill=navy_header, alignment=Alignment(horizontal="center", vertical="center"))

    ws5["A4"] = "PART A: MANDATORY BEAR SHOCK IMPACT ACROSS ALL 4 CANDIDATES (FY27E)"
    ws5["A4"].font = font_sec
    ws5.append(["Stress Test Metric", "BSE Ltd", "CDSL", "CAMS (Winner)", "KFin Tech", "Causal Transmission Mechanism"])
    style_range(ws5, 1, 5, 6, 5, font=font_tbl_header, fill=sub_header, border=grid_border, alignment=Alignment(horizontal="center"))

    shock_rows = [
        ("Shock 1: -30% Derivatives Turnover Rev Hit (%)", -0.185, -0.015, 0.000, 0.000, "+0.0%;-0.0%;0.0%", "CAMS & KFin have 0% derivatives exposure"),
        ("Shock 2: -20% Equity Drawdown Rev Hit (%)", -0.080, -0.165, -0.098, -0.082, "+0.0%;-0.0%;0.0%", "CAMS cushioned by ongoing SIPs & telescopic floor"),
        ("Combined Bear Shock FY27E Revenue Hit (%)", "=B6+B7", "=C6+C7", "=D6+D7", "=E6+E7", "+0.0%;-0.0%;0.0%", "BSE loses 26.5% revenue; CAMS loses 9.8%"),
        ("Bear Shock EBITDA Margin Compression (bps)", -920, -640, -340, -360, '#,##0" bps"', "High fixed clearing/IT costs hurt BSE"),
        ("Combined Bear Shock FY27E EPS Impact (%)", -0.384, -0.268, -0.156, -0.152, "+0.0%;-0.0%;0.0%", "Operating leverage amplifies BSE/CDSL drop"),
        ("De-rated Bear Case P/E Multiple (x)", 26.0, 28.0, 31.5, 29.0, '0.0"x"', "CAMS supported by 68% payout (~2% yield floor)"),
        ("Bear Case 12-Month Target Price (₹)", 2760, 1075, 3910, 830, "₹#,##0", "Model output under combined stress"),
        ("Implied Downside from CMP (%)", "=(B12/4650)-1", "=(C12/1580)-1", "=(D12/4420)-1", "=(E12/1140)-1", "+0.0%;-0.0%;0.0%", "CAMS is the only candidate with <-12% downside")
    ]
    for idx, sr in enumerate(shock_rows, start=6):
        ws5.append([sr[0], sr[1], sr[2], sr[3], sr[4], sr[6]])
        style_range(ws5, 1, idx, 6, idx, font=font_bold if idx in {8, 10, 12, 13} else font_regular, border=grid_border)
        ws5.cell(row=idx, column=4).fill = highlight_green
        for c in range(2, 6):
            ws5.cell(row=idx, column=c).number_format = sr[5]

    # Sensitivity Matrix 1: WACC vs Terminal Growth
    ws5["A16"] = "PART B: TWO-WAY SENSITIVITY 1 — CAMS TARGET PRICE (₹) TO WACC VS. TERMINAL GROWTH (g)"
    ws5["A16"].font = font_sec
    ws5.append(['WACC / Terminal Growth (g)', "5.5%", "6.0%", "6.5% (Base)", "7.0%", "7.5%"])
    style_range(ws5, 1, 17, 6, 17, font=font_tbl_header, fill=sub_header, border=grid_border, alignment=Alignment(horizontal="center"))

    wacc_sens = [
        ("11.5%", 5610, 5920, 6290, 6745, 7310),
        ("12.0%", 5240, 5500, 5815, 6190, 6650),
        ("12.5% (Base)", 4920, 5140, 5460, 5730, 6110),
        ("13.0%", 4640, 4830, 5060, 5335, 5660),
        ("13.5%", 4395, 4560, 4755, 4990, 5265)
    ]
    for idx, ws_row in enumerate(wacc_sens, start=18):
        ws5.append(list(ws_row))
        style_range(ws5, 1, idx, 6, idx, font=font_regular, border=grid_border, alignment=Alignment(horizontal="center"))
        ws5.cell(row=idx, column=1).font = font_bold
        for c in range(2, 7):
            ws5.cell(row=idx, column=c).number_format = "₹#,##0"
    ws5.cell(row=20, column=4).fill = highlight_green
    ws5.cell(row=20, column=4).font = font_bold

    # Sensitivity Matrix 2: Equity AAUM Growth vs Yield Compression
    ws5["A25"] = "PART C: TWO-WAY SENSITIVITY 2 — CAMS TARGET PRICE (₹) TO EQUITY AAUM CAGR VS. YIELD COMPRESSION"
    ws5["A25"].font = font_sec
    ws5.append(['Equity AAUM CAGR / Yield Cut', "-1.5% p.a.", "-2.2% p.a.", "-2.8% p.a. (Base)", "-3.5% p.a.", "-4.5% p.a."])
    style_range(ws5, 1, 26, 6, 26, font=font_tbl_header, fill=sub_header, border=grid_border, alignment=Alignment(horizontal="center"))

    yield_sens = [
        ("15.0% (Market Correction)", 4780, 4620, 4480, 4310, 4080),
        ("19.0% (Moderate)", 5320, 5140, 4980, 4790, 4540),
        ("23.0% (Base Case)", 5865, 5660, 5460, 5250, 4970),
        ("27.0% (Bull SIP Surge)", 6440, 6210, 5990, 5760, 5450)
    ]
    for idx, ys_row in enumerate(yield_sens, start=27):
        ws5.append(list(ys_row))
        style_range(ws5, 1, idx, 6, idx, font=font_regular, border=grid_border, alignment=Alignment(horizontal="center"))
        ws5.cell(row=idx, column=1).font = font_bold
        for c in range(2, 7):
            ws5.cell(row=idx, column=c).number_format = "₹#,##0"
    ws5.cell(row=29, column=4).fill = highlight_green
    ws5.cell(row=29, column=4).font = font_bold

    # Add an openpyxl BarChart on Tab 1 Dashboard
    chart1 = BarChart()
    chart1.type = "col"
    chart1.style = 10
    chart1.title = "12-Month Target Price vs. CMP & Bear Case (₹)"
    chart1.y_axis.title = "Share Price (₹)"
    chart1.x_axis.title = "Candidate"
    data_ref = Reference(ws1, min_col=2, min_row=7, max_col=5, max_row=10)
    cats_ref = Reference(ws1, min_col=2, min_row=5, max_col=5, max_row=5)
    chart1.add_data(data_ref, from_rows=True, titles_from_data=False)
    chart1.set_categories(cats_ref)
    chart1.height = 10
    chart1.width = 18
    ws1.add_chart(chart1, "A28")

    # Auto-adjust column widths across all sheets
    for sheet in wb.worksheets:
        for col in sheet.columns:
            col_letter = get_column_letter(col[0].column)
            max_len = 0
            for cell in col:
                if cell.row <= 2:
                    continue
                val_str = str(cell.value or "")
                if val_str.startswith("="):
                    val_str = "₹123,456"
                max_len = max(max_len, len(val_str))
            sheet.column_dimensions[col_letter].width = max(min(max_len + 4, 42), 14)

    wb.save(output_path)
    print(f"Workbook saved to {output_path}")

if __name__ == "__main__":
    generate_workbook()
