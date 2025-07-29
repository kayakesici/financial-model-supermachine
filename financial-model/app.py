import streamlit as st
import pandas as pd
import numpy as np
from engine.inputs import get_inputs_from_excel
from engine.statements import create_3_statements
from engine.valuation import dcf_valuation
from engine.reporting import create_excel_report, create_powerpoint_report
from io import BytesIO

st.set_page_config(page_title="M&A Financial Model", layout="wide")
st.title("💼 M&A Financial Model Super Machine")

uploaded_file = st.file_uploader("Upload Excel/CSV file", type=["xlsx", "csv"])

def run_model(assumptions):
    model = create_3_statements(assumptions, years=5)
    cash_flows = model["cash_flow"]["cash_flow"]
    ev = dcf_valuation(cash_flows, assumptions["discount_rate"], assumptions["exit_multiple"])
    return model, ev

if uploaded_file:
    try:
        # Load file
        if uploaded_file.name.endswith(".csv"):
            df_uploaded = pd.read_csv(uploaded_file)
        else:
            df_uploaded = pd.read_excel(uploaded_file, sheet_name=None, engine="openpyxl")

        st.success("✅ File Loaded Successfully")

        # Get assumptions (or default if not found)
        assumptions = get_inputs_from_excel(uploaded_file)
        if not assumptions:
            assumptions = {
                "starting_revenue": 1_000_000,
                "revenue_growth": 0.1,
                "margin": 0.4,
                "debt": 500_000,
                "interest_rate": 0.05,
                "discount_rate": 0.1,
                "exit_multiple": 5
            }

        # Manual Adjustments
        st.sidebar.header("Manual Adjustments")
        assumptions["starting_revenue"] = st.sidebar.number_input("Starting Revenue", value=assumptions["starting_revenue"])
        assumptions["revenue_growth"] = st.sidebar.slider("Revenue Growth", 0.0, 0.5, assumptions["revenue_growth"])
        assumptions["margin"] = st.sidebar.slider("Cost Margin", 0.0, 1.0, assumptions["margin"])
        assumptions["debt"] = st.sidebar.number_input("Debt", value=assumptions["debt"])
        assumptions["interest_rate"] = st.sidebar.slider("Interest Rate", 0.0, 0.2, assumptions["interest_rate"])
        assumptions["discount_rate"] = st.sidebar.slider("Discount Rate", 0.0, 0.3, assumptions["discount_rate"])
        assumptions["exit_multiple"] = st.sidebar.slider("Exit Multiple", 1, 15, assumptions["exit_multiple"])

        # Base case
        base_model, base_ev = run_model(assumptions)
        income_df = pd.DataFrame(base_model["income_statement"])
        cash_flows = base_model["cash_flow"]["cash_flow"]

        st.subheader("📊 Base Case Income Statement")
        st.dataframe(income_df)
        st.metric("Enterprise Value (Base)", f"${base_ev:,.0f}")
        st.line_chart(income_df[["revenue", "profit"]])

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

        # Generate Reports
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

else:
    st.info("Upload an Excel or CSV file to get started.")
