"""
================================================================================
 NSE INTRADAY MOMENTUM SCANNER -- STREAMLIT APP
================================================================================
UI wrapper around scanner_core.py (your original scanning engine, unchanged).
Run locally with:  streamlit run app.py
Deploy free on Streamlit Community Cloud by pushing this folder to GitHub.
================================================================================
"""

import os
import io
import pandas as pd
import streamlit as st

from scanner_core import Config, run_scan, export_to_excel
from datetime import datetime

st.set_page_config(page_title="NSE Intraday Scanner", page_icon="📈", layout="wide")

st.title("📈 NSE Intraday Momentum Scanner")
st.caption("Ranks liquid NSE stocks and builds an Entry/SL/Target trade plan. For research/education — not investment advice.")

# --------------------------------------------------------------------------
# Sidebar: inputs (replaces CLI args)
# --------------------------------------------------------------------------
with st.sidebar:
    st.header("Scan Settings")
    capital = st.number_input("Capital (Rs)", min_value=1000.0, value=100000.0, step=5000.0)
    leverage = st.number_input("Leverage (MIS)", min_value=1.0, max_value=20.0, value=5.0, step=0.5)
    risk_pct = st.number_input("Risk per trade (%)", min_value=0.1, max_value=5.0, value=0.5, step=0.1)
    top_n = st.slider("Number of stocks to show", min_value=5, max_value=150, value=20, step=5)
    signal_time = st.text_input("Signal cutoff time (HH:MM, IST)", value="09:45")
    allow_shorts = st.checkbox("Allow SELL (short) setups", value=True)

    st.divider()
    run_button = st.button("🔍 Run Scan", type="primary", use_container_width=True)

# --------------------------------------------------------------------------
# Session state to hold results across reruns
# --------------------------------------------------------------------------
if "trades" not in st.session_state:
    st.session_state.trades = None
if "cfg" not in st.session_state:
    st.session_state.cfg = None
if "label" not in st.session_state:
    st.session_state.label = None

# --------------------------------------------------------------------------
# Run scan
# --------------------------------------------------------------------------
if run_button:
    cfg = Config(
        capital=capital,
        leverage=leverage,
        risk_pct=risk_pct,
        top_n=top_n,
        signal_time=signal_time,
        allow_shorts=allow_shorts,
        outdir="scan_output",
        no_save=True,   # we handle Excel export in-memory for the download button instead
    )

    label = datetime.now().strftime("%Y-%m-%d")

    with st.spinner("Downloading data and ranking stocks... this can take 1-3 minutes for the full universe."):
        try:
            trades = run_scan(cfg)
        except Exception as e:
            st.error(f"Scan failed: {e}")
            trades = []

    st.session_state.trades = trades
    st.session_state.cfg = cfg
    st.session_state.label = label

    if not trades:
        st.warning("No tradable setups found right now. Try loosening filters (lower min_atr_pct / min_turnover) or check back closer to market open.")

# --------------------------------------------------------------------------
# Display results
# --------------------------------------------------------------------------
trades = st.session_state.trades
cfg = st.session_state.cfg
label = st.session_state.label

if trades:
    st.success(f"Found {len(trades)} ranked setups — {label}")

    rows = []
    for t in trades:
        rows.append({
            "Rank": t.rank, "Symbol": t.symbol, "Sector": t.sector,
            "Direction": t.direction, "Score": t.score, "Confidence": t.confidence,
            "Entry": t.entry, "Stop": t.stop, "Target1": t.target1, "Target2": t.target2,
            "Qty": t.qty, "Margin (Rs)": t.margin, "Risk (Rs)": t.risk_rs,
            "Reward (Rs)": t.reward_rs, "R:R": t.rr, "RVOL": t.rvol, "Gap %": t.gap_pct,
        })
    df = pd.DataFrame(rows)

    # Color BUY/SELL
    def highlight_dir(row):
        color = "background-color: #c6efce; color: #006100" if row["Direction"] == "BUY" else "background-color: #ffc7ce; color: #9c0006"
        return [color if col == "Direction" else "" for col in row.index]

    st.dataframe(df.style.apply(highlight_dir, axis=1), use_container_width=True, height=500)

    # Summary metrics row
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Total Trades", len(trades))
    c2.metric("Buy / Sell", f"{sum(1 for t in trades if t.direction=='BUY')} / {sum(1 for t in trades if t.direction=='SELL')}")
    c3.metric("Avg Score", f"{sum(t.score for t in trades)/len(trades):.1f}")
    c4.metric("Avg R:R", f"{sum(t.rr for t in trades)/len(trades):.2f}")

    # Expandable reasons per stock
    with st.expander("📋 Reasons & Warnings per stock"):
        for t in trades:
            st.markdown(f"**#{t.rank} {t.symbol} ({t.direction}, {t.confidence})**")
            st.write("- " + "\n- ".join(t.reasons))
            if t.warnings:
                st.warning(" | ".join(t.warnings))
            st.divider()

    # --------------------------------------------------------------------
    # Excel download (built in-memory, no local file needed)
    # --------------------------------------------------------------------
    os.makedirs("scan_output", exist_ok=True)
    xlsx_path = export_to_excel(trades, cfg, label)
    with open(xlsx_path, "rb") as f:
        excel_bytes = f.read()

    st.download_button(
        label="⬇️ Download Excel (5 sheets: Rankings, Trade Plan, Reasons, Summary, Full Data)",
        data=excel_bytes,
        file_name=f"picks_{label}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )

    csv_bytes = df.to_csv(index=False).encode("utf-8")
    st.download_button(
        label="⬇️ Download CSV",
        data=csv_bytes,
        file_name=f"picks_{label}.csv",
        mime="text/csv",
        use_container_width=True,
    )

elif not run_button:
    st.info("Set your capital/leverage/risk in the sidebar, then tap **Run Scan**.")
