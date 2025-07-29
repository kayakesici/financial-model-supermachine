import pandas as pd
import re

def extract_numeric_values(row):
    values = []
    for v in row[1:]:
        try:
            values.append(float(v))
        except:
            values.append(None)
    return values

def get_inputs_from_excel(file):
    xls = pd.ExcelFile(file, engine="openpyxl")

    assumptions = {}

    # ---------- 1) Read First View ----------
    if "First View" in xls.sheet_names:
        df = pd.read_excel(file, sheet_name="First View", header=None)

        for i, row in df.iterrows():
            row_text = str(row[0]).strip().lower()

            if "turnover" in row_text or "revenue" in row_text:
                vals = extract_numeric_values(row)
                assumptions["historical_revenue"] = vals
            if "cost of sales" in row_text:
                vals = extract_numeric_values(row)
                assumptions["historical_costs"] = vals
            if "ebitda" in row_text:
                vals = extract_numeric_values(row)
                assumptions["historical_ebitda"] = vals
            if "depreciation" in row_text:
                vals = extract_numeric_values(row)
                assumptions["historical_depreciation"] = vals

    # ---------- 2) Read Fixed Assets ----------
    if "Fixed Assets" in xls.sheet_names:
        df = pd.read_excel(file, sheet_name="Fixed Assets", header=None)

        for i, row in df.iterrows():
            if "additions" in str(row[0]).lower() or "capex" in str(row[0]).lower():
                assumptions["historical_capex"] = extract_numeric_values(row)

    # ---------- 3) Read Balance Sheet ----------
    if "Balance Sheet" in xls.sheet_names:
        df = pd.read_excel(file, sheet_name="Balance Sheet", header=None)

        for i, row in df.iterrows():
            row_text = str(row[0]).strip().lower()
            if "debt" in row_text or "loan" in row_text:
                assumptions["historical_debt"] = extract_numeric_values(row)
            if "cash" in row_text:
                assumptions["historical_cash"] = extract_numeric_values(row)
            if "equity" in row_text:
                assumptions["historical_equity"] = extract_numeric_values(row)

    # ---------- 4) Read Cashflow ----------
    if "Cashflow" in xls.sheet_names:
        df = pd.read_excel(file, sheet_name="Cashflow", header=None)
        for i, row in df.iterrows():
            if "net cash flow" in str(row[0]).lower():
                assumptions["historical_cashflow"] = extract_numeric_values(row)

    # ---------- Auto-generate basic assumptions ----------
    if "historical_revenue" in assumptions and assumptions["historical_revenue"]:
        revs = [v for v in assumptions["historical_revenue"] if v is not None]
        if len(revs) >= 2:
            growth = (revs[-1] / revs[-2]) - 1
        else:
            growth = 0.1
        assumptions["revenue_growth"] = growth
        assumptions["starting_revenue"] = revs[-1]
    else:
        assumptions["revenue_growth"] = 0.1
        assumptions["starting_revenue"] = 1_000_000

    if "historical_costs" in assumptions and assumptions["historical_costs"]:
        last_rev = assumptions["starting_revenue"]
        last_cost = [c for c in assumptions["historical_costs"] if c is not None][-1]
        assumptions["margin"] = last_cost / last_rev
    else:
        assumptions["margin"] = 0.4

    assumptions["debt"] = 500_000
    assumptions["interest_rate"] = 0.05
    assumptions["discount_rate"] = 0.1
    assumptions["exit_multiple"] = 5

    return assumptions
