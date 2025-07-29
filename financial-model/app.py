import streamlit as st
import pandas as pd
import numpy as np
from engine.inputs import get_inputs_from_excel
from engine.statements import create_3_statements
from engine.valuation import dcf_valuation
from engine.reporting import create_excel_report, create_powerpoint_report

st.set_page_config(page_title="M&A Financial Model", layout="wide")
st.title("💼 M&A Financial Model Super Machine")

DEFAULT_FILE = "Kayas NEW Model.xlsx"  # Default file in project folder

uploaded_file = st.file_uploader("Upload Excel file", type=["xlsx"])

# Choose file: uploaded or default
if uploaded_file:
    file_to_use = uploaded_file
else:
    try:
        file_to_use = DEFAULT_FILE
        st.info(f"📄 Using default model: {DEFAULT_FILE}")
    except:
        st.error("❌ Default file not found. Please upload an Excel file.")
        st.stop()

def run_model(assumptions):
    model = create_3_statements(assumptions, years=5)
    cash_flows = model["cash_flow"]["Cash Flow"]
    ev = dcf_valuation(cash_flows, assumptions["discount_rate"], assumptions["exit_multiple"])
    return model, ev

try:
    # Extract assumptions
    assumptions = get_inputs_from_excel(file_to_use)
    st.subheader("📄 Extracted Assumptions (Auto)")
    st.json(assumptions)

    # Base model
    base_model, base_ev = run_model(assumptions)
    income_df = pd.DataFrame(base_model["income_statement"])
    cash_flows = base_model["cash_flow"]["Cash Flow"]

    st.subheader("📊 Base Case Income Statement")
    st.dataframe(income_df)
    st.metric("Enterprise Value (Base)", f"${base_ev:,.0f}")

    # Scenario Analysis
    st.subheader("📈 Scenario Analysis")
    scenarios = {
        "Base": assumptions.copy(),
        "Upside": {**assumptions, "revenue_growth": assumptions["revenue_growth"] + 0.05, "margin": assumptions["margin"] + 0.05},
        "Downside": {**assumptions, "revenue_growth": max(0, assumptions["revenue_growth"] - 0.05), "margin": max(0, assumptions["margin"] - 0.05)},
    }

    results = {}
    for name, scenario_inputs in scenarios.items():
        _, ev = run_model(scenario_inputs)
        results[name] = ev

    scenario_df = pd.DataFrame(results.items(), columns=["Scenario", "Enterprise Value ($)"])
    st.dataframe(scenario_df)

    # Sensitivity Table
    st.subheader("📊 Valuation Sensitivity (EV)")
    discount_rates = np.arange(0.05, 0.21, 0.02)
    exit_multiples = range(3, 9)

    table = []
    for dr in discount_rates:
        row = []
        for em in exit_multiples:
            _, ev = run_model({**assumptions, "discount_rate": dr, "exit_multiple": em})
            row.append(ev)
        table.append(row)

    df_sens = pd.DataFrame(table, index=[f"{dr:.0%}" for dr in discount_rates], columns=[f"x{em}" for em in exit_multiples])
    st.dataframe(df_sens.style.format("${:,.0f}"))

    # Reports
    excel_report = create_excel_report(income_df, cash_flows, base_ev, scenario_df, df_sens, base_model)
    ppt_report = create_powerpoint_report(income_df, base_ev, scenario_df)

    st.download_button("📊 Download Full Excel Report", data=excel_report,
                       file_name="Full_Financial_Model.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    st.download_button("📈 Download PowerPoint Report", data=ppt_report,
                       file_name="Financial_Model_Report.pptx",
                       mime="application/vnd.openxmlformats-officedocument.presentationml.presentation")

except Exception as e:
    st.error(f"Error: {e}")
