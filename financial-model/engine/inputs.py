import pandas as pd

def get_inputs_from_excel(file):
    # Reads assumptions from "IC Financial Projections"
    df = pd.read_excel(file, sheet_name="IC Financial Projections", engine="openpyxl")
    assumptions = {}

    # Example mapping - adjust if column names differ
    for _, row in df.iterrows():
        key = str(row[0]).strip().lower()
        if "growth" in key:
            assumptions["revenue_growth"] = float(row[1])
        elif "margin" in key:
            assumptions["margin"] = float(row[1])
        elif "discount" in key:
            assumptions["discount_rate"] = float(row[1])
        elif "exit" in key:
            assumptions["exit_multiple"] = float(row[1])

    return assumptions
