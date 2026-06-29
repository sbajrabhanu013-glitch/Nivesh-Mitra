
import streamlit as st
import pandas as pd
import numpy as np
import yfinance as yf
from datetime import datetime, timedelta
import plotly.graph_objects as go
import plotly.express as px
import urllib.parse

st.set_page_config(
    page_title="No-API Market Assistant",
    page_icon="📈",
    layout="wide"
)

DISCLAIMER = """
This app is for learning and market screening only. It does not connect to Groww,
does not place orders, and does not provide guaranteed profit calls. Always verify
prices in your broker app before trading and use a stop-loss.
"""

DEFAULT_TICKERS = [
    "YESBANK.NS", "IDFCFIRSTB.NS", "SUZLON.NS", "SJVN.NS", "UCOBANK.NS",
    "CENTRALBK.NS", "IOB.NS", "NMDC.NS", "BANKMAHARAS.NS", "IDEA.NS",
    "IRFC.NS", "PNB.NS", "SAIL.NS", "HUDCO.NS", "GAIL.NS"
]

def clean_symbol(symbol: str) -> str:
    symbol = symbol.strip().upper()
    if not symbol:
        return ""
    if "." not in symbol:
        symbol = symbol + ".NS"
    return symbol

@st.cache_data(ttl=60)
def fetch_intraday(symbol: str, period="1d", interval="5m"):
    data = yf.download(
        symbol,
        period=period,
        interval=interval,
        progress=False,
        auto_adjust=False,
        threads=False
    )
    if data is None or data.empty:
        return pd.DataFrame()
    if isinstance(data.columns, pd.MultiIndex):
        data.columns = data.columns.get_level_values(0)
    data = data.dropna().copy()
    return data

@st.cache_data(ttl=300)
def fetch_daily(symbol: str, period="3mo"):
    data = yf.download(
        symbol,
        period=period,
        interval="1d",
        progress=False,
        auto_adjust=False,
        threads=False
    )
    if data is None or data.empty:
        return pd.DataFrame()
    if isinstance(data.columns, pd.MultiIndex):
        data.columns = data.columns.get_level_values(0)
    return data.dropna().copy()

def calculate_vwap(df: pd.DataFrame) -> pd.Series:
    typical = (df["High"] + df["Low"] + df["Close"]) / 3
    volume = df["Volume"].replace(0, np.nan)
    return (typical * volume).cumsum() / volume.cumsum()

def analyze_symbol(symbol: str, max_price: float, min_volume: int):
    intraday = fetch_intraday(symbol)
    daily = fetch_daily(symbol)

    if intraday.empty or len(intraday) < 5:
        return None, intraday

    last = intraday.iloc[-1]
    current_price = float(last["Close"])
    day_high = float(intraday["High"].max())
    day_low = float(intraday["Low"].min())
    day_volume = int(intraday["Volume"].sum())

    if current_price > max_price or day_volume < min_volume:
        return None, intraday

    intraday["VWAP"] = calculate_vwap(intraday)
    vwap = float(intraday["VWAP"].iloc[-1])

    avg_vol_20 = None
    if not daily.empty and "Volume" in daily.columns:
        avg_vol_20 = float(daily["Volume"].tail(20).mean())

    above_vwap = current_price > vwap
    near_day_high = current_price >= day_high * 0.995
    above_open = current_price > float(intraday["Open"].iloc[0])

    volume_ok = True
    if avg_vol_20 and avg_vol_20 > 0:
        # Intraday volume is partial-day, so this is only a rough liquidity check
        volume_ok = day_volume >= avg_vol_20 * 0.10

    risk_per_share = max(current_price - vwap, current_price * 0.005)
    stop_loss = current_price - risk_per_share
    target_1 = current_price + (risk_per_share * 1.5)
    target_2 = current_price + (risk_per_share * 2)

    score = 0
    reasons = []
    if above_vwap:
        score += 30
        reasons.append("Price above VWAP")
    else:
        reasons.append("Price below VWAP")

    if near_day_high:
        score += 25
        reasons.append("Near day high")
    else:
        reasons.append("Not near day high")

    if above_open:
        score += 20
        reasons.append("Above opening price")
    else:
        reasons.append("Below opening price")

    if volume_ok:
        score += 25
        reasons.append("Liquidity/volume acceptable")
    else:
        reasons.append("Weak volume")

    if score >= 75:
        signal = "BUY WATCH"
    elif score >= 50:
        signal = "WAIT / WATCH"
    else:
        signal = "AVOID"

    return {
        "Symbol": symbol,
        "Price": round(current_price, 2),
        "VWAP": round(vwap, 2),
        "Day High": round(day_high, 2),
        "Day Low": round(day_low, 2),
        "Volume": day_volume,
        "Score": score,
        "Signal": signal,
        "Stop-loss idea": round(stop_loss, 2),
        "Target 1 idea": round(target_1, 2),
        "Target 2 idea": round(target_2, 2),
        "Reasons": "; ".join(reasons),
        "Groww Search": f"https://groww.in/search?query={urllib.parse.quote(symbol.replace('.NS',''))}"
    }, intraday

def quantity_from_risk(capital, max_risk_pct, entry, stop):
    risk_amount = capital * (max_risk_pct / 100)
    risk_per_share = max(entry - stop, 0.01)
    qty = int(risk_amount // risk_per_share)
    affordable_qty = int(capital // entry)
    return max(0, min(qty, affordable_qty)), risk_amount

st.title("📈 No-API Market Assistant")
st.caption("Public-data screening dashboard. No broker login. No order placement.")

st.warning(DISCLAIMER)

with st.sidebar:
    st.header("Settings")
    capital = st.number_input("Capital available (₹)", min_value=100.0, value=100.0, step=50.0)
    max_price = st.number_input("Maximum stock price (₹)", min_value=1.0, value=100.0, step=1.0)
    min_volume = st.number_input("Minimum intraday volume", min_value=0, value=100000, step=50000)
    max_risk_pct = st.slider("Max risk per trade (%)", 0.5, 5.0, 1.0, 0.5)

    st.subheader("Symbols")
    raw = st.text_area(
        "NSE symbols, one per line",
        value="\n".join(DEFAULT_TICKERS),
        height=220,
        help="Use NSE symbols. Example: YESBANK or YESBANK.NS"
    )
    run_scan = st.button("Run scan", type="primary")

symbols = [clean_symbol(x) for x in raw.splitlines() if clean_symbol(x)]
symbols = list(dict.fromkeys(symbols))

tab1, tab2, tab3, tab4 = st.tabs(["Scanner", "Chart", "Risk Calculator", "How to use"])

if run_scan or "scan_results" not in st.session_state:
    results = []
    data_cache = {}
    progress = st.progress(0)
    for i, symbol in enumerate(symbols):
        result, df = analyze_symbol(symbol, max_price=max_price, min_volume=min_volume)
        if result:
            results.append(result)
            data_cache[symbol] = df
        progress.progress((i + 1) / max(len(symbols), 1))
    progress.empty()
    st.session_state["scan_results"] = pd.DataFrame(results)
    st.session_state["data_cache"] = data_cache

df_results = st.session_state.get("scan_results", pd.DataFrame())
data_cache = st.session_state.get("data_cache", {})

with tab1:
    st.subheader("Stocks under your price limit")
    if df_results.empty:
        st.info("No matching stocks found. Lower volume filter or add more symbols.")
    else:
        df_show = df_results.sort_values(["Score", "Volume"], ascending=[False, False]).reset_index(drop=True)
        st.dataframe(
            df_show.drop(columns=["Groww Search"]),
            use_container_width=True,
            hide_index=True
        )
        st.caption("Signals are rule-based watch labels, not guaranteed buy/sell calls.")

        csv = df_show.to_csv(index=False).encode("utf-8")
        st.download_button("Download scan as CSV", data=csv, file_name="market_scan.csv", mime="text/csv")

with tab2:
    st.subheader("Chart view")
    if df_results.empty:
        st.info("Run the scanner first.")
    else:
        selected = st.selectbox("Choose symbol", df_results["Symbol"].tolist())
        chart_df = data_cache.get(selected)
        if chart_df is None or chart_df.empty:
            chart_df = fetch_intraday(selected)
            if not chart_df.empty:
                chart_df["VWAP"] = calculate_vwap(chart_df)

        if chart_df is not None and not chart_df.empty:
            fig = go.Figure()
            fig.add_trace(go.Candlestick(
                x=chart_df.index,
                open=chart_df["Open"],
                high=chart_df["High"],
                low=chart_df["Low"],
                close=chart_df["Close"],
                name="Price"
            ))
            if "VWAP" in chart_df.columns:
                fig.add_trace(go.Scatter(
                    x=chart_df.index,
                    y=chart_df["VWAP"],
                    mode="lines",
                    name="VWAP"
                ))
            fig.update_layout(height=550, xaxis_rangeslider_visible=False)
            st.plotly_chart(fig, use_container_width=True)

            row = df_results[df_results["Symbol"] == selected].iloc[0]
            col1, col2, col3, col4 = st.columns(4)
            col1.metric("Signal", row["Signal"])
            col2.metric("Score", int(row["Score"]))
            col3.metric("Stop-loss idea", f"₹{row['Stop-loss idea']}")
            col4.metric("Target 1 idea", f"₹{row['Target 1 idea']}")

            st.link_button("Open Groww search manually", row["Groww Search"])
        else:
            st.error("Could not load chart data for this symbol.")

with tab3:
    st.subheader("Position size calculator")
    entry = st.number_input("Entry price (₹)", min_value=0.01, value=50.0, step=0.05)
    stop = st.number_input("Stop-loss price (₹)", min_value=0.01, value=49.5, step=0.05)
    qty, risk_amount = quantity_from_risk(capital, max_risk_pct, entry, stop)
    st.write(f"Maximum risk amount: **₹{risk_amount:.2f}**")
    st.write(f"Suggested quantity by risk and affordability: **{qty} share(s)**")
    if qty > 0:
        st.write(f"Approx order value: **₹{qty * entry:.2f}**")
        st.write(f"Approx loss if stop-loss hits: **₹{qty * max(entry - stop, 0):.2f}**")
    else:
        st.warning("Quantity is 0. Increase capital or reduce entry price/risk gap.")

with tab4:
    st.subheader("How this app works")
    st.markdown("""
1. Add NSE symbols in the sidebar.
2. Set maximum price, for example ₹100.
3. Click **Run scan**.
4. Use **BUY WATCH** only as a shortlist, then verify in Groww manually.
5. Never trade without stop-loss.
6. Do not use margin or F&O while learning.
    """)
    st.info("The app uses public market data through yfinance. Data may be delayed or unavailable. Verify every price in your broker app.")
