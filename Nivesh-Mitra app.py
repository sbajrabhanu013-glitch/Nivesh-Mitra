import streamlit as st
import pandas as pd
import numpy as np
import yfinance as yf
import plotly.graph_objects as go
from datetime import datetime
from pathlib import Path
import urllib.parse
import time
import requests
from io import StringIO

st.set_page_config(
    page_title="Full Market Assistant - No API",
    page_icon="📊",
    layout="wide"
)

DATA_DIR = Path("data")
DATA_DIR.mkdir(exist_ok=True)

NSE_SEC_LIST_URL = "https://nsearchives.nseindia.com/content/equities/sec_list.csv"
CACHE_FILE = DATA_DIR / "nse_equity_universe_cache.csv"

DISCLAIMER = """
This app is for learning and screening only. It does not connect to Groww, does not ask
for broker login, does not place orders, and does not provide guaranteed profit calls.
Public market data may be delayed or unavailable. Always verify live prices in your broker
app before any trade and use a stop-loss.
"""

STARTER_SYMBOLS = [
    {"symbol": "YESBANK.NS", "name": "Yes Bank", "series": "EQ"},
    {"symbol": "IDFCFIRSTB.NS", "name": "IDFC First Bank", "series": "EQ"},
    {"symbol": "SUZLON.NS", "name": "Suzlon Energy", "series": "EQ"},
    {"symbol": "SJVN.NS", "name": "SJVN", "series": "EQ"},
    {"symbol": "UCOBANK.NS", "name": "UCO Bank", "series": "EQ"},
    {"symbol": "CENTRALBK.NS", "name": "Central Bank of India", "series": "EQ"},
    {"symbol": "IOB.NS", "name": "Indian Overseas Bank", "series": "EQ"},
    {"symbol": "NMDC.NS", "name": "NMDC", "series": "EQ"},
    {"symbol": "BANKMAHARAS.NS", "name": "Bank of Maharashtra", "series": "EQ"},
    {"symbol": "IDEA.NS", "name": "Vodafone Idea", "series": "EQ"},
    {"symbol": "IRFC.NS", "name": "Indian Railway Finance Corporation", "series": "EQ"},
    {"symbol": "PNB.NS", "name": "Punjab National Bank", "series": "EQ"},
    {"symbol": "SAIL.NS", "name": "Steel Authority of India", "series": "EQ"},
    {"symbol": "HUDCO.NS", "name": "HUDCO", "series": "EQ"},
    {"symbol": "GAIL.NS", "name": "GAIL", "series": "EQ"},
]

def normalize_col(col: str) -> str:
    return col.strip().upper().replace(" ", "_").replace("-", "_")

def safe_symbol(symbol: str) -> str:
    symbol = str(symbol).strip().upper()
    symbol = symbol.replace("&", "%26")
    if not symbol:
        return ""
    if "." not in symbol:
        return f"{symbol}.NS"
    return symbol

def groww_url(symbol: str) -> str:
    raw = symbol.replace(".NS", "").replace("%26", "&")
    return f"https://groww.in/search?query={urllib.parse.quote(raw)}"

@st.cache_data(ttl=60 * 60)
def download_nse_universe():
    """
    Download NSE securities list. This is a symbol universe, not a live-price feed.
    If NSE blocks or internet fails, fallback to local cache/starter list.
    """
    headers = {
        "User-Agent": "Mozilla/5.0",
        "Accept": "text/csv,application/csv,text/plain,*/*",
        "Referer": "https://www.nseindia.com/"
    }
    try:
        response = requests.get(NSE_SEC_LIST_URL, headers=headers, timeout=15)
        response.raise_for_status()
        text = response.text
        df_raw = pd.read_csv(StringIO(text))
        df_raw.columns = [normalize_col(c) for c in df_raw.columns]

        # NSE sec_list commonly contains SYMBOL, SERIES, NAME_OF_COMPANY.
        symbol_col = "SYMBOL"
        name_col = "NAME_OF_COMPANY" if "NAME_OF_COMPANY" in df_raw.columns else None
        series_col = "SERIES" if "SERIES" in df_raw.columns else None

        if symbol_col not in df_raw.columns:
            raise ValueError("NSE CSV did not contain SYMBOL column")

        df = pd.DataFrame()
        df["symbol_raw"] = df_raw[symbol_col].astype(str).str.strip().str.upper()
        df["symbol"] = df["symbol_raw"].apply(safe_symbol)
        df["name"] = df_raw[name_col].astype(str) if name_col else df["symbol_raw"]
        df["series"] = df_raw[series_col].astype(str).str.strip().str.upper() if series_col else "EQ"

        df = df[df["symbol_raw"].str.len() > 0].drop_duplicates("symbol")
        df = df.sort_values(["series", "symbol"]).reset_index(drop=True)

        df.to_csv(CACHE_FILE, index=False)
        return df, f"Downloaded from NSE at {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
    except Exception as e:
        if CACHE_FILE.exists():
            df = pd.read_csv(CACHE_FILE)
            return df, f"Loaded cached NSE universe because download failed: {e}"
        df = pd.DataFrame(STARTER_SYMBOLS)
        df["symbol_raw"] = df["symbol"].str.replace(".NS", "", regex=False)
        return df[["symbol_raw", "symbol", "name", "series"]], f"Loaded starter list because NSE download failed: {e}"

@st.cache_data(ttl=90)
def fetch_intraday(symbol: str, period="1d", interval="5m"):
    try:
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
        required = {"Open", "High", "Low", "Close", "Volume"}
        if not required.issubset(set(data.columns)):
            return pd.DataFrame()
        return data.dropna().copy()
    except Exception:
        return pd.DataFrame()

@st.cache_data(ttl=300)
def fetch_batch_daily(tickers, period="5d"):
    if not tickers:
        return pd.DataFrame()
    try:
        data = yf.download(
            tickers=tickers,
            period=period,
            interval="1d",
            group_by="ticker",
            progress=False,
            auto_adjust=False,
            threads=True
        )
        return data
    except Exception:
        return pd.DataFrame()

def calculate_vwap(df):
    typical = (df["High"] + df["Low"] + df["Close"]) / 3
    volume = df["Volume"].replace(0, np.nan)
    return (typical * volume).cumsum() / volume.cumsum()

def analyze_intraday(symbol, name="", max_price=100.0, min_volume=100000):
    df = fetch_intraday(symbol)
    if df.empty or len(df) < 4:
        return None, df

    df["VWAP"] = calculate_vwap(df)
    last = df.iloc[-1]
    current_price = float(last["Close"])
    day_high = float(df["High"].max())
    day_low = float(df["Low"].min())
    day_open = float(df["Open"].iloc[0])
    day_volume = int(df["Volume"].sum())
    vwap = float(df["VWAP"].iloc[-1]) if not pd.isna(df["VWAP"].iloc[-1]) else current_price

    if current_price > max_price or day_volume < min_volume:
        return None, df

    above_vwap = current_price > vwap
    near_day_high = current_price >= day_high * 0.995
    above_open = current_price > day_open
    green_last_candle = float(last["Close"]) > float(last["Open"])
    not_extended = current_price <= vwap * 1.04  # avoids very extended entries

    score = 0
    reasons = []

    if above_vwap:
        score += 30
        reasons.append("Above VWAP")
    else:
        reasons.append("Below VWAP")

    if near_day_high:
        score += 25
        reasons.append("Near day high")
    else:
        reasons.append("Not near day high")

    if above_open:
        score += 15
        reasons.append("Above open")
    else:
        reasons.append("Below open")

    if green_last_candle:
        score += 10
        reasons.append("Last candle green")
    else:
        reasons.append("Last candle weak")

    if not_extended:
        score += 20
        reasons.append("Not too extended")
    else:
        reasons.append("Extended from VWAP")

    risk_per_share = max(current_price - min(vwap, current_price * 0.995), current_price * 0.005)
    stop_loss = current_price - risk_per_share
    target_1 = current_price + risk_per_share * 1.5
    target_2 = current_price + risk_per_share * 2.0

    if score >= 80:
        signal = "BUY WATCH"
    elif score >= 55:
        signal = "WAIT / WATCH"
    else:
        signal = "AVOID"

    return {
        "Symbol": symbol,
        "Name": name,
        "Price": round(current_price, 2),
        "VWAP": round(vwap, 2),
        "Day Open": round(day_open, 2),
        "Day High": round(day_high, 2),
        "Day Low": round(day_low, 2),
        "Volume": day_volume,
        "Score": int(score),
        "Signal": signal,
        "Stop-loss idea": round(stop_loss, 2),
        "Target 1 idea": round(target_1, 2),
        "Target 2 idea": round(target_2, 2),
        "Reasons": "; ".join(reasons),
        "Groww Search": groww_url(symbol),
    }, df

def quick_daily_scan(universe, max_price=100.0, limit=500):
    """
    A faster scan based on daily candles. Useful for big universe pre-filtering.
    Then run intraday scan on the shortlist.
    """
    universe = universe.head(limit).copy()
    tickers = universe["symbol"].tolist()
    data = fetch_batch_daily(tickers, period="5d")
    rows = []
    if data.empty:
        return pd.DataFrame()

    for _, urow in universe.iterrows():
        t = urow["symbol"]
        name = urow.get("name", "")
        try:
            if isinstance(data.columns, pd.MultiIndex):
                if t not in data.columns.get_level_values(0):
                    continue
                df = data[t].dropna()
            else:
                df = data.dropna()
            if df.empty or len(df) < 2:
                continue
            last = df.iloc[-1]
            prev = df.iloc[-2]
            price = float(last["Close"])
            if price > max_price or price <= 0:
                continue
            pct_change = ((price - float(prev["Close"])) / float(prev["Close"])) * 100
            volume = int(last["Volume"])
            rows.append({
                "Symbol": t,
                "Name": name,
                "Price": round(price, 2),
                "% Change": round(pct_change, 2),
                "Volume": volume,
                "Groww Search": groww_url(t),
            })
        except Exception:
            continue

    out = pd.DataFrame(rows)
    if out.empty:
        return out
    return out.sort_values(["% Change", "Volume"], ascending=[False, False]).reset_index(drop=True)

def quantity_from_risk(capital, max_risk_pct, entry, stop):
    risk_amount = capital * (max_risk_pct / 100)
    risk_per_share = max(entry - stop, 0.01)
    qty_by_risk = int(risk_amount // risk_per_share)
    qty_by_cash = int(capital // entry)
    return max(0, min(qty_by_risk, qty_by_cash)), risk_amount

st.title("📊 Full Market Assistant - No API")
st.caption("Large NSE stock-universe screener using public data. No broker login. No order placement.")

st.warning(DISCLAIMER)

universe, universe_status = download_nse_universe()

with st.sidebar:
    st.header("Market universe")
    st.caption(universe_status)

    series_options = sorted(universe["series"].dropna().unique().tolist())
    default_series = ["EQ"] if "EQ" in series_options else series_options[:1]
    selected_series = st.multiselect(
        "NSE series to include",
        options=series_options,
        default=default_series,
        help="EQ is normal equity. BE/SME can be less liquid/riskier."
    )

    search_text = st.text_input("Search symbol/name", "")

    filtered_universe = universe[universe["series"].isin(selected_series)].copy()
    if search_text.strip():
        q = search_text.strip().upper()
        filtered_universe = filtered_universe[
            filtered_universe["symbol_raw"].str.contains(q, na=False) |
            filtered_universe["name"].str.upper().str.contains(q, na=False)
        ]

    st.write(f"Loaded symbols: **{len(universe):,}**")
    st.write(f"Selected universe: **{len(filtered_universe):,}**")

    st.header("Scan settings")
    capital = st.number_input("Capital available (₹)", min_value=100.0, value=100.0, step=50.0)
    max_price = st.number_input("Maximum stock price (₹)", min_value=1.0, value=100.0, step=1.0)
    min_volume = st.number_input("Minimum intraday volume", min_value=0, value=100000, step=50000)
    max_risk_pct = st.slider("Max risk per trade (%)", 0.5, 5.0, 1.0, 0.5)

    st.header("Batch control")
    batch_size = st.slider("Symbols to scan now", 20, 500, 100, 20)
    start_at = st.number_input("Start from row number", min_value=0, max_value=max(len(filtered_universe)-1, 0), value=0, step=100)
    auto_refresh = st.checkbox("Auto-refresh every 90 seconds", value=False)

    run_intraday_scan = st.button("Run intraday batch scan", type="primary")
    run_daily_scan = st.button("Fast daily pre-scan")

    if st.button("Clear cached market data"):
        st.cache_data.clear()
        st.success("Cache cleared. Run scan again.")

if auto_refresh:
    st.markdown("<script>setTimeout(function(){ window.location.reload(); }, 90000);</script>", unsafe_allow_html=True)

tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "Universe",
    "Daily Pre-scan",
    "Intraday Scanner",
    "Chart + Groww",
    "Risk Calculator"
])

with tab1:
    st.subheader("Loaded stock universe")
    st.write("This is the symbol universe the app can scan in batches.")
    st.dataframe(
        filtered_universe[["symbol", "name", "series"]].reset_index(drop=True),
        use_container_width=True,
        hide_index=True
    )
    csv = filtered_universe.to_csv(index=False).encode("utf-8")
    st.download_button(
        "Download selected universe CSV",
        data=csv,
        file_name="selected_stock_universe.csv",
        mime="text/csv"
    )

with tab2:
    st.subheader("Fast daily pre-scan")
    st.write("Use this to quickly shortlist stocks under your price limit before running intraday scan.")
    if run_daily_scan or "daily_scan" in st.session_state:
        if run_daily_scan:
            chunk = filtered_universe.iloc[int(start_at): int(start_at) + int(batch_size)].copy()
            with st.spinner("Running daily pre-scan..."):
                st.session_state["daily_scan"] = quick_daily_scan(chunk, max_price=max_price, limit=batch_size)
        daily = st.session_state.get("daily_scan", pd.DataFrame())
        if daily.empty:
            st.info("No daily matches found in this batch. Try another batch or reduce filters.")
        else:
            st.dataframe(daily, use_container_width=True, hide_index=True)
            st.download_button(
                "Download daily pre-scan CSV",
                data=daily.to_csv(index=False).encode("utf-8"),
                file_name="daily_pre_scan.csv",
                mime="text/csv"
            )
    else:
        st.info("Click **Fast daily pre-scan** in the sidebar.")

with tab3:
    st.subheader("Intraday batch scanner")
    st.write("This scans a selected batch. Scanning all stocks at once is slow without a paid market-data API.")
    if run_intraday_scan or "intraday_scan" in st.session_state:
        if run_intraday_scan:
            chunk = filtered_universe.iloc[int(start_at): int(start_at) + int(batch_size)].copy()
            rows = []
            chart_cache = {}
            progress = st.progress(0)
            status = st.empty()

            for i, (_, row) in enumerate(chunk.iterrows()):
                symbol = row["symbol"]
                name = row["name"]
                status.write(f"Scanning {i+1}/{len(chunk)}: {symbol}")
                result, df = analyze_intraday(
                    symbol,
                    name=name,
                    max_price=max_price,
                    min_volume=min_volume
                )
                if result:
                    rows.append(result)
                    chart_cache[symbol] = df
                progress.progress((i + 1) / max(len(chunk), 1))

            progress.empty()
            status.empty()
            out = pd.DataFrame(rows)
            if not out.empty:
                out = out.sort_values(["Score", "Volume"], ascending=[False, False]).reset_index(drop=True)
            st.session_state["intraday_scan"] = out
            st.session_state["chart_cache"] = chart_cache

        scan = st.session_state.get("intraday_scan", pd.DataFrame())
        if scan.empty:
            st.info("No intraday matches found. Try lower minimum volume, another batch, or a different time.")
        else:
            st.dataframe(
                scan.drop(columns=["Groww Search"]),
                use_container_width=True,
                hide_index=True
            )
            st.download_button(
                "Download intraday scan CSV",
                data=scan.to_csv(index=False).encode("utf-8"),
                file_name="intraday_scan.csv",
                mime="text/csv"
            )
    else:
        st.info("Click **Run intraday batch scan** in the sidebar.")

with tab4:
    st.subheader("Chart and manual Groww check")
    scan = st.session_state.get("intraday_scan", pd.DataFrame())
    chart_cache = st.session_state.get("chart_cache", {})
    if scan.empty:
        st.info("Run intraday scanner first.")
    else:
        selected = st.selectbox("Choose symbol", scan["Symbol"].tolist())
        chart_df = chart_cache.get(selected, pd.DataFrame())
        if chart_df.empty:
            chart_df = fetch_intraday(selected)
            if not chart_df.empty:
                chart_df["VWAP"] = calculate_vwap(chart_df)

        if chart_df.empty:
            st.error("Chart data unavailable for this symbol.")
        else:
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
                fig.add_trace(go.Scatter(x=chart_df.index, y=chart_df["VWAP"], mode="lines", name="VWAP"))
            fig.update_layout(height=550, xaxis_rangeslider_visible=False)
            st.plotly_chart(fig, use_container_width=True)

            row = scan[scan["Symbol"] == selected].iloc[0]
            c1, c2, c3, c4, c5 = st.columns(5)
            c1.metric("Signal", row["Signal"])
            c2.metric("Score", int(row["Score"]))
            c3.metric("Price", f"₹{row['Price']}")
            c4.metric("Stop idea", f"₹{row['Stop-loss idea']}")
            c5.metric("Target 1", f"₹{row['Target 1 idea']}")
            st.write("Reasons:", row["Reasons"])
            st.link_button("Open Groww search manually", row["Groww Search"])

with tab5:
    st.subheader("₹100 learning-mode risk calculator")
    entry = st.number_input("Entry price (₹)", min_value=0.01, value=50.0, step=0.05)
    stop = st.number_input("Stop-loss price (₹)", min_value=0.01, value=49.5, step=0.05)
    qty, risk_amount = quantity_from_risk(capital, max_risk_pct, entry, stop)
    st.write(f"Maximum risk amount: **₹{risk_amount:.2f}**")
    st.write(f"Suggested quantity: **{qty} share(s)**")
    if qty > 0:
        st.write(f"Approx order value: **₹{qty * entry:.2f}**")
        st.write(f"Approx loss if stop-loss hits: **₹{qty * max(entry - stop, 0):.2f}**")
    else:
        st.warning("Quantity is 0. Increase capital, reduce stock price, or reduce stop-loss distance.")

st.caption("Tip: Start with daily pre-scan, then intraday scan only the best batch. Full-market scanning is slow without a paid data feed.")
