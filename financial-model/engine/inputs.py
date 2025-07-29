import pandas as pd

def extract_numeric_values(row):
    vals = []
    for v in row[1:]:
        try:
            vals.append(float(v))
        except:
            vals.append(None)
    return vals

def get_inputs_from_excel(file):
    xls = pd.ExcelFile(file, engine="openpyxl")
    assumptions = {}

    # — First View (P&L) —
    if "First View" in xls.sheet_names:
        df = pd.read_excel(file, sheet_name="First View", header=None)
        for _, row in df.iterrows():
            key = str(row[0]).lower()
            if "turnover" in key or "revenue" in key:
                assumptions["historical_revenue"] = extract_numeric_values(row)
            if "cost of sales" in key:
                assumptions["historical_costs"] = extract_numeric_values(row)
            if "ebitda" in key:
                assumptions["historical_ebitda"] = extract_numeric_values(row)
            if "depreciation" in key:
                assumptions["historical_depreciation"] = extract_numeric_values(row)

    # — Fixed Assets (Capex) —
    if "Fixed Assets" in xls.sheet_names:
        df = pd.read_excel(file, sheet_name="Fixed Assets", header=None)
        for _, row in df.iterrows():
            key = str(row[0]).lower()
            if "capex" in key or "additions" in key:
                assumptions["historical_capex"] = extract_numeric_values(row)

    # — Balance Sheet —
    if "Balance Sheet" in xls.sheet_names:
        df = pd.read_excel(file, sheet_name="Balance Sheet", header=None)
        for _, row in df.iterrows():
            key = str(row[0]).lower()
            if "debt" in key or "loan" in key:
                assumptions["historical_debt"] = extract_numeric_values(row)
            if key == "cash" or "cash" in key:
                assumptions["historical_cash"] = extract_numeric_values(row)
            if "equity" in key:
                assumptions["historical_equity"] = extract_numeric_values(row)

    # — Cashflow —
    if "Cashflow" in xls.sheet_names:
        df = pd.read_excel(file, sheet_name="Cashflow", header=None)
        for _, row in df.iterrows():
            key = str(row[0]).lower()
            if "net cash flow" in key:
                assumptions["historical_cashflow"] = extract_numeric_values(row)

    # — Derive basic assumptions —
    # Revenue growth & starting revenue
    revs = [v for v in assumptions.get("historical_revenue", []) if v is not None]
    if len(revs) >= 2:
        assumptions["revenue_growth"] = revs[-1]/revs[-2] - 1
        assumptions["starting_revenue"] = revs[-1]
    else:
        assumptions["revenue_growth"] = 0.1
        assumptions["starting_revenue"] = revs[-1] if revs else 1_000_000

    # Cost margin
    costs = [v for v in assumptions.get("historical_costs", []) if v is not None]
    if costs:
        assumptions["margin"] = costs[-1] / assumptions["starting_revenue"]
    else:
        assumptions["margin"] = 0.4

    # Defaults for anything missing
    assumptions.setdefault("historical_debt", [])
    assumptions.setdefault("historical_cash", [])
    assumptions.setdefault("historical_capex", [])
    assumptions.setdefault("interest_rate", 0.05)
    assumptions.setdefault("discount_rate", 0.1)
    assumptions.setdefault("exit_multiple", 5)

    return assumptions
