import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from pptx import Presentation
from pptx.util import Inches
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table as PDFTable
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.pagesizes import letter

from engine.inputs import get_inputs_from_excel
from engine.statements import create_3_statements
from engine.valuation import dcf_valuation
from engine.reporting import create_excel_report, create_powerpoint_report

# ─── CONFIG ──────────────────────────────────────────────────────────
st.set_page_config(page_title="M&A Super Machine", layout="wide")
DEFAULT_FILE = "Kayas NEW Model.xlsx"

# ─── LOAD HISTORICAL ASSUMPTIONS ────────────────────────────────────
hist_assumps = get_inputs_from_excel(DEFAULT_FILE)

# ─── USER INPUTS PANEL ──────────────────────────────────────────────
st.sidebar.header("🔧 Model Inputs (override)")

starting_rev   = st.sidebar.number_input(
    "Starting Revenue",
    value=int(hist_assumps["starting_revenue"]),
    step=10000
)

revenue_growth = st.sidebar.slider(
    "Revenue Growth Rate",
    0.0, 1.0,
    float(hist_assumps["revenue_growth"])
)

cost_margin    = st.sidebar.slider(
    "Cost of Sales Margin",
    0.0, 1.0,
    float(hist_assumps["margin"])
)

depr_rate      = st.sidebar.slider(
    "Depreciation Rate (%)",
    0.0, 0.5,
    0.1
)

capex_pct      = st.sidebar.slider(
    "CapEx % of Revenue",
    0.0, 0.5,
    0.1
)

debt_amt       = st.sidebar.number_input(
    "Debt (Year 0)",
    value=int(hist_assumps.get("debt", 0)),
    step=10000
)

interest_rate  = st.sidebar.slider(
    "Interest Rate",
    0.0, 0.3,
    float(hist_assumps["interest_rate"])
)

discount_rate  = st.sidebar.slider(
    "DCF Discount Rate",
    0.0, 0.3,
    float(hist_assumps["discount_rate"])
)

exit_multiple  = st.sidebar.slider(
    "Exit Multiple",
    1, 15,
    int(hist_assumps["exit_multiple"])
)

# Assemble the final assumptions dict
assumptions = {
    "starting_revenue": starting_rev,
    "revenue_growth":   revenue_growth,
    "margin":           cost_margin,
    "depreciation_rate": depr_rate,
    "capex_pct":        capex_pct,
    "debt":             debt_amt,
    "interest_rate":    interest_rate,
    "discount_rate":    discount_rate,
    "exit_multiple":    exit_multiple
}

# ─── RUN MODEL ──────────────────────────────────────────────────────
model = create_3_statements(assumptions, years=5)
income_df   = pd.DataFrame(model["income_statement"]).set_index("Year")
cashflow_df = pd.DataFrame(model["cash_flow"]).set_index("Year")
bs_df       = pd.DataFrame(model["balance_sheet"]).set_index("Year")

# Calculate DCF EV
ev = dcf_valuation(cashflow_df["Cash Flow"].tolist(), discount_rate, exit_multiple)

# ─── NAVBAR (TABS) ─────────────────────────────────────────────────
tabs = st.tabs(["📈 Summary", "📄 P&L", "🏗 Fixed Assets", "🏦 Balance Sheet", "💵 Cash Flow"])

with tabs[0]:
    st.header("📊 Executive Summary")
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Starting Revenue", f"${starting_rev:,.0f}")
    c2.metric("EV (DCF)", f"${ev:,.0f}")
    c3.metric("Revenue Growth", f"{revenue_growth:.1%}")
    c4.metric("Cost Margin", f"{cost_margin:.1%}")
    st.line_chart(income_df[["Revenue","Profit"]], height=300)
    st.bar_chart(cashflow_df["Cash Flow"], height=300)

with tabs[1]:
    st.header("📄 Profit & Loss (Modeled)")
    st.dataframe(income_df, use_container_width=True)

with tabs[2]:
    st.header("🏗 Fixed Assets & CapEx")
    st.markdown(f"- **Depreciation Rate:** {depr_rate:.1%}")
    st.markdown(f"- **CapEx Assumed:** {capex_pct:.1%} of revenue")
    # If you have a separate Fixed Assets DF, you can show it here

with tabs[3]:
    st.header("🏦 Balance Sheet")
    st.dataframe(bs_df, use_container_width=True)
    st.area_chart(bs_df[["Assets","Debt","Equity"]], height=300)

with tabs[4]:
    st.header("💵 Cash Flow Forecast")
    st.dataframe(cashflow_df, use_container_width=True)
    st.metric("Enterprise Value (DCF)", f"${ev:,.0f}")
    st.line_chart(cashflow_df["Cash Flow"], height=300)

# ─── DOWNLOAD REPORTS ───────────────────────────────────────────────
st.markdown("---")
st.subheader("📥 Download Reports")

excel_bytes = create_excel_report(
    income_df.reset_index(),
    cashflow_df["Cash Flow"].tolist(),
    ev,
    pd.DataFrame([("Base",ev)],columns=["Scenario","EV"]),
    pd.DataFrame(),  # sensitivity placeholder
    model
)
ppt_bytes = create_powerpoint_report(
    income_df.reset_index(),
    ev,
    pd.DataFrame([("Base",ev)],columns=["Scenario","EV"])
)

col1, col2 = st.columns(2)
with col1:
    st.download_button("Download Excel", excel_bytes, "M&A_Model.xlsx",
                       "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
with col2:
    st.download_button("Download PPT",   ppt_bytes,   "M&A_Model.pptx",
                       "application/vnd.openxmlformats-officedocument.presentationml.presentation")
