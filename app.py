"""
Stock Backtesting Dashboard  —  Streamlit UI
=============================================
Run: streamlit run app.py

Requires: pip install streamlit yfinance pandas numpy openpyxl requests
"""

import os
import time
import tempfile
import streamlit as st
import pandas as pd
from backtest import run_backtest, classify_signal, classify_win_loss

st.set_page_config(
    page_title="NSE Stock Backtester",
    page_icon="📊",
    layout="wide",
)

# ─── Custom CSS ────────────────────────────────────────────────────────────────
st.markdown("""
<style>
    .metric-card {
        background: #1F3864; color: white;
        padding: 18px 22px; border-radius: 10px;
        text-align: center; margin: 4px;
    }
    .metric-card .value { font-size: 28px; font-weight: bold; }
    .metric-card .label { font-size: 13px; opacity: 0.85; margin-top: 4px; }
    .win-badge  { background:#C6EFCE; color:#375623; padding:3px 9px; border-radius:5px; font-weight:bold; }
    .loss-badge { background:#FFC7CE; color:#9C0006; padding:3px 9px; border-radius:5px; font-weight:bold; }
    .momentum-badge { background:#C6EFCE; color:#375623; padding:3px 8px; border-radius:5px; }
    .weak-badge     { background:#FFC7CE; color:#9C0006; padding:3px 8px; border-radius:5px; }
    .neutral-badge  { background:#FFEB9C; color:#7D6608; padding:3px 8px; border-radius:5px; }
</style>
""", unsafe_allow_html=True)


# ─── Header ────────────────────────────────────────────────────────────────────
st.title("📊 NSE Stock Backtesting System")
st.caption(
    "Upload your screener results · Fetch live prices & RSI · Measure forward returns")

st.divider()

# ─── Sidebar ───────────────────────────────────────────────────────────────────
with st.sidebar:
    st.header("⚙️ Settings")
    output_filename = st.text_input(
        "Output Filename", value="final_output.xlsx")
    st.markdown("---")
    st.markdown("""
**Input File Requirements**

| Column        | Example          |
|---------------|------------------|
| Symbol        | RELIANCE, TCS    |
| Sector        | Technology       |
| Market Cap    | Largecap         |
| Date of Entry | 2025-08-22       |

Accepts `.xlsx` or `.csv`
    """)

# ─── File Upload ───────────────────────────────────────────────────────────────
uploaded_file = st.file_uploader(
    "📁 Upload Screener File (Excel or CSV)",
    type=["xlsx", "csv"],
    help="File must contain: Symbol, Sector, Market Cap, Date of Entry"
)

if uploaded_file is not None:
    try:
        if uploaded_file.name.endswith(".csv"):
            df_preview = pd.read_csv(uploaded_file)
        else:
            df_preview = pd.read_excel(uploaded_file)
        uploaded_file.seek(0)

        st.success(
            f"✅ File loaded: **{uploaded_file.name}** — {len(df_preview)} rows")

        with st.expander("📋 Preview Input Data", expanded=False):
            st.dataframe(df_preview.head(10), use_container_width=True)

        st.divider()

        col1, col2, col3 = st.columns([1, 2, 1])
        with col2:
            run_clicked = st.button(
                "🚀 Run Full Backtest Analysis",
                use_container_width=True,
                type="primary",
            )

        if run_clicked:
            # Save input to temp file
            suffix = ".xlsx" if uploaded_file.name.endswith(
                ".xlsx") else ".csv"
            with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp_in:
                tmp_in.write(uploaded_file.getbuffer())
                input_path = tmp_in.name

            output_path = os.path.join(tempfile.gettempdir(), output_filename)

            progress_bar = st.progress(0, text="Initialising…")
            status_text = st.empty()

            t_start = time.time()

            with st.spinner(""):
                try:
                    df_result = run_backtest(input_path, output_path)
                except Exception as e:
                    st.error(f"❌ Analysis failed: {e}")
                    st.stop()

            progress_bar.progress(100, text="Done!")
            elapsed = time.time() - t_start
            st.success(f"✅ Analysis completed in {elapsed:.0f}s")

            # ── Summary Metrics ───────────────────────────────────────────────
            st.divider()
            st.subheader("📈 Summary Metrics")

            valid = df_result[df_result["Entry Price"].notna()]
            wins = (valid["Win/Loss"] == "Win").sum()
            win_rate = wins / len(valid) * 100 if len(valid) else 0
            avg_1w = valid["1W Return %"].mean()
            avg_2w = valid["2W Return %"].mean()
            avg_3w = valid["3W Return %"].mean()

            m1, m2, m3, m4, m5 = st.columns(5)
            m1.metric("Total Stocks",    len(df_result))
            m2.metric("Fetched OK",       len(valid))
            m3.metric("Win Rate (1W)",    f"{win_rate:.1f}%")
            m4.metric("Avg 1W Return",
                      f"{avg_1w:+.2f}%" if not pd.isna(avg_1w) else "N/A")
            m5.metric("Avg 3W Return",
                      f"{avg_3w:+.2f}%" if not pd.isna(avg_3w) else "N/A")

            # ── Signal Breakdown ──────────────────────────────────────────────
            sig_counts = valid["Signal"].value_counts()
            sc1, sc2, sc3 = st.columns(3)
            sc1.metric("🟢 Momentum", sig_counts.get("Momentum", 0))
            sc2.metric("🟡 Neutral",  sig_counts.get("Neutral", 0))
            sc3.metric("🔴 Weak",     sig_counts.get("Weak", 0))

            # ── Results Table ─────────────────────────────────────────────────
            st.divider()
            st.subheader("📊 Backtest Results")

            # Filter controls
            f1, f2, f3 = st.columns(3)
            with f1:
                sig_filter = st.multiselect(
                    "Filter by Signal",
                    options=["Momentum", "Neutral", "Weak", "N/A"],
                    default=["Momentum", "Neutral", "Weak", "N/A"]
                )
            with f2:
                wl_filter = st.multiselect(
                    "Filter by Win/Loss",
                    options=["Win", "Loss", "N/A"],
                    default=["Win", "Loss", "N/A"]
                )
            with f3:
                sectors = df_result["Sector"].dropna().unique().tolist()
                sec_filter = st.multiselect(
                    "Filter by Sector",
                    options=sorted(sectors),
                    default=sectors
                )

            mask = (
                df_result["Signal"].isin(sig_filter) &
                df_result["Win/Loss"].isin(wl_filter) &
                df_result["Sector"].isin(sec_filter)
            )
            df_display = df_result[mask].copy()
            df_display["Original Entry Date"] = df_display["Original Entry Date"].astype(
                str)
            df_display["Entry Date"] = df_display["Entry Date"].astype(str)
            df_display["Date Shift (Days)"] = df_display["Date Shift (Days)"].apply(
                lambda x: f"{x} days" if pd.notna(x) else "—"
            )

            # Format for display
            for ret_col in ("1W Return %", "2W Return %", "3W Return %"):
                df_display[ret_col] = df_display[ret_col].apply(
                    lambda x: f"{x:+.2f}%" if pd.notna(x) else "—"
                )
            df_display["Entry Price"] = df_display["Entry Price"].apply(
                lambda x: f"₹{x:,.2f}" if pd.notna(x) else "—"
            )
            df_display["RSI"] = df_display["RSI"].apply(
                lambda x: f"{x:.2f}" if pd.notna(x) else "—"
            )
            df_display["Delivery %"] = df_display["Delivery %"].apply(
                lambda x: f"{x:.2f}%" if pd.notna(x) else "—"
            )

            st.dataframe(
                df_display,
                use_container_width=True,
                height=480,
            )

            # ── Download ──────────────────────────────────────────────────────
            st.divider()
            with open(output_path, "rb") as f:
                st.download_button(
                    label="📥 Download Excel Report",
                    data=f,
                    file_name=output_filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    type="primary",
                )

            os.unlink(input_path)

    except Exception as e:
        st.error(f"❌ Error reading file: {e}")

else:
    st.info("👆 Upload a screener file above to get started.")
    with st.expander("ℹ️ How it works"):
        st.markdown("""
1. **Upload** your screener Excel/CSV file with columns: Symbol, Sector, Market Cap, Date of Entry
2. **Click Run** — the system fetches historical prices from Yahoo Finance (NSE)
3. **Results** include: Entry Price, RSI(14), 1W/2W/3W returns, Delivery %, Signal, Win/Loss
4. **Download** the formatted Excel report

**Signal Logic:**
- 🟢 **Momentum**: RSI between 55 and 75
- 🔴 **Weak**: RSI below 40
- 🟡 **Neutral**: everything else

**Win/Loss:** Based on 1-week (5 trading days) forward return.
        """)
