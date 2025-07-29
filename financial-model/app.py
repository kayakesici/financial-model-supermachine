import streamlit as st
import pandas as pd
import numpy as np
from engine.inputs import get_inputs_from_excel
from engine.statements import create_3_statements
from engine.valuation import dcf_valuation
from engine.reporting import create_excel_report, create_powerpoint_report

st.set_page_config(page_title="M&A Financial Model", layout="wide")
st.title("💼 M&A Financial Model Super Machine")

# If you drop your Excel here it overrides this default
DEFAULT_FILE = "Kayas NEW Model.xlsx"
uploaded = st.file_uploader("Upload Excel file", type=["xlsx"])

# Choose which file to use
if uploaded:
    file_to_use = uploaded
else:
    file_to_use = DEFAULT_FILE
    st.info(f"📄 Using default model: {DEFAULT_FILE}")

# MAIN
try:
    # 1) extract everything
    assumptions = get_inputs_from_excel(file_to_use)
    st.subheader("📄 Auto‑Extracted Assumptions")
    st.json(assumptions)

    # 2) run 3‑statement + DCF
    def run(assumps):
        m, ev = create_3_statements(assumps, years=5), None
        cf = m["cash_flow"]["Cash Flow"]
        ev = dcf_valuation(cf, assumps["discount_rate"], assumps["exit_multiple"])
        return m, ev

    model, ev = run(assumptions)

    # 3) display Income Statement & EV
    inc = pd.DataFrame(model["income_statement"])
    st.subheader("📊 Income Statement")
    st.dataframe(inc)
    st.metric("Enterprise Value", f"${ev:,.0f}")

    # 4) Scenarios
    st.subheader("📈 Scenario Analysis")
    scenarios = {
        "Base": assumptions,
        "Upside": {**assumptions, "revenue_growth": assumptions["revenue_growth"] + 0.05, "margin": assumptions["margin"] + 0.05},
        "Downside": {**assumptions, "revenue_growth": max(0, assumptions["revenue_growth"] - 0.05), "margin": max(0, assumptions["margin"] - 0.05)},
    }
    results = {k: run(v)[1] for k, v in scenarios.items()}
    df_scen = pd.DataFrame(results.items(), columns=["Scenario", "EV"])
    st.dataframe(df_scen.style.format({"EV":"${:,.0f}"}))

    # 5) Sensitivity
    st.subheader("📊 Sensitivity Table")
    drs = np.arange(0.05, 0.21, 0.02)
    ems = range(3,9)
    sens = []
    for dr in drs:
        row = [run({**assumptions, "discount_rate":dr,"exit_multiple":em})[1] for em in ems]
        sens.append(row)
    df_sens = pd.DataFrame(sens, index=[f"{dr:.0%}" for dr in drs], columns=[f"x{em}" for em in ems])
    st.dataframe(df_sens.style.format("${:,.0f}"))

    # 6) Full‑report downloads
    st.subheader("📥 Download Reports")
    excel_bytes = create_excel_report(
        pd.DataFrame(model["income_statement"]),
        model["cash_flow"]["Cash Flow"],
        ev,
        df_scen,
        df_sens,
        model
    )
    ppt_bytes = create_powerpoint_report(
        pd.DataFrame(model["income_statement"]),
        ev,
        df_scen
    )

    st.download_button("Download Excel", excel_bytes, "Full_Report.xlsx", 
                       "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    st.download_button("Download PowerPoint", ppt_bytes, "Full_Report.pptx",
                       "application/vnd.openxmlformats-officedocument.presentationml.presentation")

except Exception as e:
    st.error(f"⚠️ Something went wrong: {e}")
