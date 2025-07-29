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
st.set_page_config(page_title="M&A Super Machine", layout="wide", initial_sidebar_state="collapsed")
DEFAULT_FILE = "Kayas NEW Model.xlsx"

@st.cache_data(show_spinner=False)
def load_excel(path):
    return pd.ExcelFile(path, engine="openpyxl")

xls = None
try:
    xls = load_excel(DEFAULT_FILE)
except FileNotFoundError:
    st.error(f"🚨 Default file not found: {DEFAULT_FILE}")
    st.stop()

# ─── EXTRACT & MODEL ─────────────────────────────────────────────────
assumptions = get_inputs_from_excel(DEFAULT_FILE)
model, _ = None, None
try:
    model = create_3_statements(assumptions, years=5)
    cash_flows = model["cash_flow"]["Cash Flow"]
    base_ev = dcf_valuation(cash_flows, assumptions["discount_rate"], assumptions["exit_multiple"])
except Exception as e:
    st.error(f"Error building model: {e}")
    st.stop()

income_df = pd.DataFrame(model["income_statement"])
cashflow_df = pd.DataFrame(model["cash_flow"])
bs_df = pd.DataFrame(model["balance_sheet"])
fixed_assets_df = pd.read_excel(DEFAULT_FILE, sheet_name="Fixed Assets", engine="openpyxl", header=None)
first_view_df = pd.read_excel(DEFAULT_FILE, sheet_name="First View", engine="openpyxl", header=None)

# ─── NAVBAR / TABS ───────────────────────────────────────────────────
tabs = st.tabs(["📈 Summary", "📄 P&L", "🏗️ Fixed Assets", "🏦 Balance Sheet", "💵 Cash Flow"])

# ─── TAB 1: SUMMARY ───────────────────────────────────────────────────
with tabs[0]:
    st.header("📊 Executive Summary")
    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Starting Revenue", f"${assumptions['starting_revenue']:,.0f}")
    k2.metric("EV (DCF)", f"${base_ev:,.0f}")
    k3.metric("Revenue Growth", f"{assumptions['revenue_growth']:.1%}")
    k4.metric("Cost Margin", f"{assumptions['margin']:.1%}")

    col1, col2 = st.columns(2)
    with col1:
        st.subheader("Revenue vs Profit")
        st.line_chart(income_df.set_index("Year")[["Revenue", "Profit"]])
    with col2:
        st.subheader("Cash Flow Trend")
        st.bar_chart(cashflow_df.set_index("Year")["Cash Flow"])

    st.subheader("🥇 Top Ratios")
    ratios = {
        "EBITDA Margin": f"{((assumptions.get('historical_ebitda') or [0])[-1]/assumptions['starting_revenue']):.1%}" if assumptions.get('historical_ebitda') else "n/a",
        "Debt / EBITDA": f"{(assumptions['historical_debt'][-1]/(assumptions.get('historical_ebitda')[-1])):.1f}" if assumptions.get('historical_debt') and assumptions.get('historical_ebitda') else "n/a"
    }
    for name, val in ratios.items():
        st.write(f"**{name}:** {val}")

# ─── TAB 2: P&L ───────────────────────────────────────────────────────
with tabs[1]:
    st.header("📄 Profit & Loss (First View)")
    st.dataframe(first_view_df, use_container_width=True)
    st.subheader("Modeled Income Statement")
    st.dataframe(income_df.set_index("Year"), use_container_width=True)

# ─── TAB 3: FIXED ASSETS ──────────────────────────────────────────────
with tabs[2]:
    st.header("🏗️ Fixed Assets & CapEx")
    st.dataframe(fixed_assets_df, use_container_width=True)
    st.subheader("Modeled CapEx Trend")
    hist_capex = assumptions.get("historical_capex", [])
    st.line_chart(pd.DataFrame({"CapEx": hist_capex}, index=range(len(hist_capex))))

# ─── TAB 4: BALANCE SHEET ────────────────────────────────────────────
with tabs[3]:
    st.header("🏦 Balance Sheet")
    st.dataframe(bs_df.set_index("Year"), use_container_width=True)
    st.subheader("Balance Sheet Breakdown")
    st.area_chart(bs_df.set_index("Year")[["Assets", "Debt", "Equity"]])

# ─── TAB 5: CASH FLOW ────────────────────────────────────────────────
with tabs[4]:
    st.header("💵 Cash Flow Forecast")
    st.dataframe(cashflow_df.set_index("Year"), use_container_width=True)
    st.subheader("DCF Enterprise Value")
    st.metric("", f"${base_ev:,.0f}")

    # Sensitivity pivot
    st.subheader("🔀 Sensitivity (EV)")
    drs = np.arange(0.05, 0.21, 0.02)
    ems = range(3, 9)
    table = [[
        dcf_valuation(cash_flows, dr, em) for em in ems
    ] for dr in drs]
    sens_df = pd.DataFrame(table, index=[f"{d:.0%}" for d in drs], columns=[f"x{e}" for e in ems])
    st.dataframe(sens_df.style.format("${:,.0f}"))

# ─── MONTE CARLO SIMULATION ───────────────────────────────────────────
st.markdown("---")
st.header("🎲 Monte Carlo Scenario Simulation")
n = st.slider("Number of runs", 50, 500, 100)
mc = []
for i in range(n):
    g = np.random.normal(assumptions["revenue_growth"], 0.02)
    m = np.random.normal(assumptions["margin"], 0.03)
    tmp_assump = {**assumptions, "revenue_growth": g, "margin": m}
    cf = create_3_statements(tmp_assump, years=5)["cash_flow"]["Cash Flow"]
    mc.append(dcf_valuation(cf, assumptions["discount_rate"], assumptions["exit_multiple"]))
st.line_chart(pd.Series(mc, name="Simulated EV"))

# ─── DOWNLOADS ────────────────────────────────────────────────────────
st.markdown("---")
st.header("📥 Download Reports")
excel_data = create_excel_report(
    income_df, cash_flows, base_ev,
    pd.DataFrame([(k,v) for k,v in [("Base", base_ev)]], columns=["Scenario","EV"]),
    sens_df, model
)
ppt_data = create_powerpoint_report(income_df, base_ev, pd.DataFrame([("Base", base_ev)], columns=["Scenario","EV"]))

col1, col2, col3 = st.columns([1,1,2])
with col1:
    st.download_button("Download Excel", excel_data, "M&A_Model.xlsx",
                       "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
with col2:
    st.download_button("Download PPT", ppt_data, "M&A_Model.pptx",
                       "application/vnd.openxmlformats-officedocument.presentationml.presentation")
with col3:
    # PDF export
    buffer = BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=letter)
    styles = getSampleStyleSheet()
    elems = [Paragraph("M&A Financial Model Report", styles["Title"]), Spacer(1,12)]
    # include summary table
    data = [["Metric","Value"], ["EV", f"${base_ev:,.0f}"], ["Revenue Growth", f"{assumptions['revenue_growth']:.1%}"]]
    elems.append(PDFTable(data))
    doc.build(elems)
    pdf = buffer.getvalue()
    st.download_button("Download PDF", pdf, "M&A_Report.pdf", "application/pdf")
