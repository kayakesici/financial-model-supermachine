import pandas as pd

def extract_numeric_values(row):
    """Turn a row of mixed data into a list of floats or None."""
    vals = []
    for v in row[1:]:
        try:
            vals.append(float(v))
        except:
            vals.append(None)
    return vals

def get_inputs_from_excel(file):
    """
    Reads your four sheets—First View, Fixed Assets, Balance Sheet, Cashflow—
    pulls out revenue, costs, capex, debt, cash, etc., and then auto-derives
    growth rates, margins and sets sensible defaults for everything else.
    """
    xls = pd.ExcelFile(file, engine="openpyxl")
    assumptions = {}

    # — First View (P&L) —
    if "First View" in xls.sheet_names:
        df = pd.read_excel(file, sheet_name="First View", header=None)
        for _, row in df.iterrows():
            key = str(row[0]).strip().lower()
            if "turnover" in key or "revenue" in key:
                assumptions["historical_revenue"] = extract_numeric_values(row)
            elif "cost of sales" in key:
                assumptions["historical_costs"] = extract_numeric_values(row)
            elif "ebitda" in key:
                assumptions["historical_ebitda"] = extract_numeric_values(row)
            elif "depreciation" in key:
                assumptions["historical_depreciation"] = extract_numeric_values(row)

    # — Fixed Assets (Capex) —
    if "Fixed Assets" in xls.sheet_names:
        df = pd.read_excel(file, sheet_name="Fixed Assets", header=None)
        for _, row in df.iterrows():
            key = str(row[0]).strip().lower()
            if "capex" in key or "additions" in key:
                assumptions["historical_capex"] = extract_numeric_values(row)

    # — Balance Sheet —
    if "Balance Sheet" in xls.sheet_names:
        df = pd.read_excel(file, sheet_name="Balance Sheet", header=None)
        for _, row in df.iterrows():
            key = str(row[0]).strip().lower()
            if "debt" in key or "loan" in key:
                assumptions["historical_debt"] = extract_numeric_values(row)
            elif key == "cash" or "cash" in key:
                assumptions["historical_cash"] = extract_numeric_values(row)
            elif "equity" in key:
                assumptions["historical_equity"] = extract_numeric_values(row)

    # — Cashflow —
    if "Cashflow" in xls.sheet_names:
        df = pd.read_excel(file, sheet_name="Cashflow", header=None)
        for _, row in df.iterrows():
            key = str(row[0]).strip().lower()
            if "net cash flow" in key or "cash flow" == key:
                assumptions["historical_cashflow"] = extract_numeric_values(row)

    # — DERIVE CLEAN ASSUMPTIONS — #
    # Historical revenue and growth
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

    # Debt: pull last historical, else default
    debts = [v for v in assumptions.get("historical_debt", []) if v is not None]
    if debts:
        assumptions["debt"] = debts[-1]
    else:
        assumptions["debt"] = 0.0  # or choose a default like 500_000

    # Cash and capex (just store arrays; you can use these later if desired)
    assumptions.setdefault("historical_cash", [])
    assumptions.setdefault("historical_capex", [])
    assumptions.setdefault("historical_cashflow", [])

    # Fixed defaults for rates
    assumptions.setdefault("interest_rate", 0.05)
    assumptions.setdefault("discount_rate", 0.1)
    assumptions.setdefault("exit_multiple", 5)

    return assumptions
