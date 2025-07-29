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

@st.cache_resource  # <— changed from cache_data to cache_resource
def load_excel(path):
    return pd.ExcelFile(path, engine="openpyxl")

# Attempt to load
try:
    xls = load_excel(DEFAULT_FILE)
except FileNotFoundError:
    st.error(f"🚨 Default file not found: {DEFAULT_FILE}")
    st.stop()

# ─── EXTRACT & MODEL ─────────────────────────────────────────────────
assumptions = get_inputs_from_excel(DEFAULT_FILE)
try:
    model = create_3_statements(assumptions, years=5)
    cash_flows = model["cash_flow"]["Cash Flow"]
    base_ev = dcf_valuation(cash_flows, assumptions["discount_rate"], assumptions["exit_multiple"])
except Exception as e:
    st.error(f"Error building model: {e}")
    st.stop()

income_df     = pd.DataFrame(model["income_statement"])
cashflow_df   = pd.DataFrame(model["cash_flow"])
bs_df         = pd.DataFrame(model["balance_sheet"])
fixed_assets_df = pd.read_excel(DEFAULT_FILE, sheet_name="Fixed Assets", engine="openpyxl", header=None)
first_view_df   = pd.read_excel(DEFAULT_FILE, sheet_name="First View",   engine="openpyxl", header=None)

# ─── NAVBAR / TABS ───────────────────────────────────────────────────
tabs = st.tabs(["📈 Summary", "📄 P&L", "🏗️ Fixed Assets", "🏦 Balance Sheet", "💵 Cash Flow"])

# ─── TAB 1: SUMMARY ───────────────────────────────────────────────────
with tabs[0]:
    st.header("📊 Executive Summary")
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Starting Revenue", f"${assumptions['starting_revenue']:,.0f}")
    c2.metric("EV (DCF)",          f"${base_ev:,.0f}")
    c3.metric("Revenue Growth",    f"{assumptions['revenue_growth']:.1%}")
    c4.metric("Cost Margin",       f"{assumptions['margin']:.1%}")
    # … rest of summary …

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
    # … rest …

# ─── TAB 4: BALANCE SHEET ────────────────────────────────────────────
with tabs[3]:
    st.header("🏦 Balance Sheet")
    st.dataframe(bs_df.set_index("Year"), use_container_width=True)
    # … rest …

# ─── TAB 5: CASH FLOW ────────────────────────────────────────────────
with tabs[4]:
    st.header("💵 Cash Flow Forecast")
    st.dataframe(cashflow_df.set_index("Year"), use_container_width=True)
    # … rest …

# ─── DOWNLOADS ────────────────────────────────────────────────────────
st.markdown("---")
st.header("📥 Download Reports")
excel_data = create_excel_report(income_df, cash_flows, base_ev,
                                 pd.DataFrame([("Base", base_ev)], columns=["Scenario","EV"]),
                                 pd.DataFrame(), model)
ppt_data   = create_powerpoint_report(income_df, base_ev,
                                      pd.DataFrame([("Base", base_ev)], columns=["Scenario","EV"]))

col1, col2 = st.columns(2)
with col1:
    st.download_button("Download Excel", excel_data, "M&A_Model.xlsx",
                       "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
with col2:
    st.download_button("Download PPT",   ppt_data,   "M&A_Model.pptx",
                       "application/vnd.openxmlformats-officedocument.presentationml.presentation")
