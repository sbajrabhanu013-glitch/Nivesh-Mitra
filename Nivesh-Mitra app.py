import streamlit as st
import pandas as pd
import numpy as np
import yfinance as yf
import sqlite3
from io import BytesIO
from datetime import datetime
import plotly.graph_objects as go


# =========================================================
# PAGE CONFIG
# =========================================================
st.set_page_config(
    page_title="BAJRABHANU Trading Desk",
    page_icon="📈",
    layout="wide"
)


# =========================================================
# DATABASE
# =========================================================
DB_NAME = "trading_app.db"


def get_conn():
    return sqlite3.connect(DB_NAME, check_same_thread=False)


def init_db():
    conn = get_conn()
    cur = conn.cursor()

    cur.execute("""
        CREATE TABLE IF NOT EXISTS watchlist (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            symbol TEXT UNIQUE,
            added_on TEXT
        )
    """)

    cur.execute("""
        CREATE TABLE IF NOT EXISTS paper_trades (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            symbol TEXT,
            trade_type TEXT,
            entry_price REAL,
            quantity INTEGER,
            stop_loss REAL,
            target REAL,
            trade_date TEXT,
            status TEXT,
            created_on TEXT
        )
    """)

    conn.commit()
    conn.close()


init_db()


# =========================================================
# CSS UI
# =========================================================
st.markdown("""
<style>
    .stApp {
        background: linear-gradient(135deg, #07111f 0%, #0b1f35 45%, #101828 100%);
        color: #ffffff;
    }

    .main-title {
        font-size: 42px;
        font-weight: 900;
        color: #ffffff;
        margin-bottom: 0px;
    }

    .sub-title {
        color: #d6e4ff;
        font-size: 16px;
        margin-top: -5px;
        margin-bottom: 20px;
    }

    .premium-card {
        background: rgba(255, 255, 255, 0.08);
        padding: 22px;
        border-radius: 20px;
        border: 1px solid rgba(255, 255, 255, 0.12);
        box-shadow: 0px 12px 30px rgba(0,0,0,0.25);
        margin-bottom: 18px;
    }

    .metric-card {
        background: linear-gradient(135deg, rgba(255,153,51,0.18), rgba(19,136,8,0.15));
        padding: 18px;
        border-radius: 18px;
        border: 1px solid rgba(255,255,255,0.15);
        text-align: center;
    }

    .metric-label {
        color: #d6e4ff;
        font-size: 14px;
    }

    .metric-value {
        color: #ffffff;
        font-size: 28px;
        font-weight: 900;
    }

    .buy-box {
        padding: 16px;
        border-radius: 16px;
        background: rgba(0, 180, 100, 0.18);
        border: 1px solid rgba(0, 255, 150, 0.4);
        color: #d7ffe9;
        font-weight: 800;
        text-align: center;
    }

    .sell-box {
        padding: 16px;
        border-radius: 16px;
        background: rgba(220, 53, 69, 0.18);
        border: 1px solid rgba(255, 100, 100, 0.4);
        color: #ffe1e1;
        font-weight: 800;
        text-align: center;
    }

    .hold-box {
        padding: 16px;
        border-radius: 16px;
        background: rgba(255, 193, 7, 0.18);
        border: 1px solid rgba(255, 220, 100, 0.4);
        color: #fff4cc;
        font-weight: 800;
        text-align: center;
    }

    .risk-note {
        background: rgba(255,255,255,0.08);
        padding: 14px;
        border-radius: 14px;
        border-left: 5px solid #ff9933;
        color: #f8fafc;
    }

    div[data-testid="stMetric"] {
        background: rgba(255,255,255,0.08);
        padding: 16px;
        border-radius: 16px;
        border: 1px solid rgba(255,255,255,0.12);
    }
</style>
""", unsafe_allow_html=True)


# =========================================================
# HELPER FUNCTIONS
# =========================================================
def clean_symbol(symbol):
    symbol = str(symbol).strip().upper()
    if symbol and "." not in symbol:
        symbol = symbol + ".NS"
    return symbol


@st.cache_data(ttl=300)
def load_stock_data(symbol, period="6mo", interval="1d"):
    try:
        df = yf.download(
            symbol,
            period=period,
            interval=interval,
            progress=False,
            auto_adjust=False
        )

        if df is None or df.empty:
            return pd.DataFrame()

        if isinstance(df.columns, pd.MultiIndex):
            df.columns = df.columns.get_level_values(0)

        df = df.reset_index()

        required_cols = ["Open", "High", "Low", "Close", "Volume"]
        for col in required_cols:
            if col not in df.columns:
                return pd.DataFrame()

        df = df.dropna(subset=["Open", "High", "Low", "Close"])
        return df

    except Exception:
        return pd.DataFrame()


def calculate_rsi(df, period=14):
    delta = df["Close"].diff()

    gain = delta.where(delta > 0, 0)
    loss = -delta.where(delta < 0, 0)

    avg_gain = gain.rolling(period).mean()
    avg_loss = loss.rolling(period).mean()

    rs = avg_gain / avg_loss
    rsi = 100 - (100 / (1 + rs))

    return rsi


def add_indicators(df):
    df = df.copy()

    df["MA20"] = df["Close"].rolling(20).mean()
    df["MA50"] = df["Close"].rolling(50).mean()
    df["RSI"] = calculate_rsi(df)

    return df


def generate_signal(df):
    if df.empty or len(df) < 50:
        return "HOLD", "Not enough data for strong signal."

    latest = df.iloc[-1]

    close = latest["Close"]
    ma20 = latest["MA20"]
    ma50 = latest["MA50"]
    rsi = latest["RSI"]

    if pd.isna(ma20) or pd.isna(ma50) or pd.isna(rsi):
        return "HOLD", "Indicator data is incomplete."

    if close > ma20 and ma20 > ma50 and 50 <= rsi <= 70:
        return "BUY", "Price is above MA20 and MA50 with healthy RSI momentum."

    elif close < ma20 and ma20 < ma50 and rsi < 45:
        return "SELL", "Price is below key moving averages with weak RSI."

    elif rsi > 70:
        return "HOLD", "RSI is overbought. Avoid fresh entry without confirmation."

    elif rsi < 30:
        return "HOLD", "RSI is oversold. Watch for reversal confirmation."

    else:
        return "HOLD", "No clear high-confidence setup right now."


def make_candlestick_chart(df, symbol):
    fig = go.Figure()

    fig.add_trace(go.Candlestick(
        x=df["Date"] if "Date" in df.columns else df["Datetime"],
        open=df["Open"],
        high=df["High"],
        low=df["Low"],
        close=df["Close"],
        name="Candlestick"
    ))

    fig.add_trace(go.Scatter(
        x=df["Date"] if "Date" in df.columns else df["Datetime"],
        y=df["MA20"],
        mode="lines",
        name="MA20"
    ))

    fig.add_trace(go.Scatter(
        x=df["Date"] if "Date" in df.columns else df["Datetime"],
        y=df["MA50"],
        mode="lines",
        name="MA50"
    ))

    fig.update_layout(
        title=f"{symbol} Price Chart",
        xaxis_title="Date",
        yaxis_title="Price",
        height=560,
        template="plotly_dark",
        xaxis_rangeslider_visible=False,
        margin=dict(l=20, r=20, t=60, b=20)
    )

    return fig


def make_rsi_chart(df, symbol):
    fig = go.Figure()

    fig.add_trace(go.Scatter(
        x=df["Date"] if "Date" in df.columns else df["Datetime"],
        y=df["RSI"],
        mode="lines",
        name="RSI"
    ))

    fig.add_hline(y=70, line_dash="dash", annotation_text="Overbought 70")
    fig.add_hline(y=30, line_dash="dash", annotation_text="Oversold 30")

    fig.update_layout(
        title=f"{symbol} RSI",
        xaxis_title="Date",
        yaxis_title="RSI",
        height=300,
        template="plotly_dark",
        margin=dict(l=20, r=20, t=60, b=20)
    )

    return fig


def add_to_watchlist(symbol):
    conn = get_conn()
    cur = conn.cursor()

    try:
        cur.execute(
            "INSERT OR IGNORE INTO watchlist (symbol, added_on) VALUES (?, ?)",
            (symbol, datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
        )
        conn.commit()
    finally:
        conn.close()


def get_watchlist():
    conn = get_conn()
    df = pd.read_sql_query("SELECT * FROM watchlist ORDER BY id DESC", conn)
    conn.close()
    return df


def delete_watchlist_symbol(symbol):
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("DELETE FROM watchlist WHERE symbol = ?", (symbol,))
    conn.commit()
    conn.close()


def save_trade(symbol, trade_type, entry_price, quantity, stop_loss, target, trade_date):
    conn = get_conn()
    cur = conn.cursor()

    cur.execute("""
        INSERT INTO paper_trades 
        (symbol, trade_type, entry_price, quantity, stop_loss, target, trade_date, status, created_on)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, (
        symbol,
        trade_type,
        float(entry_price),
        int(quantity),
        float(stop_loss),
        float(target),
        str(trade_date),
        "OPEN",
        datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    ))

    conn.commit()
    conn.close()


def get_trades():
    conn = get_conn()
    df = pd.read_sql_query("SELECT * FROM paper_trades ORDER BY id DESC", conn)
    conn.close()
    return df


def update_trade_status(trade_id, status):
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("UPDATE paper_trades SET status = ? WHERE id = ?", (status, trade_id))
    conn.commit()
    conn.close()


def delete_trade(trade_id):
    conn = get_conn()
    cur = conn.cursor()
    cur.execute("DELETE FROM paper_trades WHERE id = ?", (trade_id,))
    conn.commit()
    conn.close()


def df_to_excel(download_sheets):
    output = BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for sheet_name, df in download_sheets.items():
            safe_name = sheet_name[:31]
            df.to_excel(writer, index=False, sheet_name=safe_name)

    output.seek(0)
    return output


def get_current_price(symbol):
    data = load_stock_data(symbol, period="5d", interval="1d")
    if data.empty:
        return None
    return float(data.iloc[-1]["Close"])


# =========================================================
# HEADER
# =========================================================
st.markdown("""
<div class="premium-card">
    <div class="main-title">📈 BAJRABHANU Trading Desk</div>
    <div class="sub-title">
        Stock Analysis • Watchlist • Educational Signal • Paper Trading • P/L Tracker
    </div>
</div>
""", unsafe_allow_html=True)

st.markdown("""
<div class="risk-note">
    ⚠️ <b>Important:</b> This application is for educational analysis and paper trading only. 
    It does not provide guaranteed profit or investment advice. Always use proper risk management.
</div>
<br>
""", unsafe_allow_html=True)


# =========================================================
# SIDEBAR
# =========================================================
st.sidebar.title("📌 Navigation")

page = st.sidebar.radio(
    "Select Module",
    [
        "Dashboard",
        "Stock Analysis",
        "Watchlist",
        "Paper Trading",
        "Reports"
    ]
)

st.sidebar.markdown("---")
st.sidebar.caption("Built for learning, analysis and disciplined trading.")


# =========================================================
# DASHBOARD
# =========================================================
if page == "Dashboard":
    st.subheader("📊 Dashboard Summary")

    watchlist_df = get_watchlist()
    trades_df = get_trades()

    total_watchlist = len(watchlist_df)
    total_trades = len(trades_df)
    open_trades = len(trades_df[trades_df["status"] == "OPEN"]) if not trades_df.empty else 0
    closed_trades = len(trades_df[trades_df["status"] == "CLOSED"]) if not trades_df.empty else 0

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Watchlist Stocks", total_watchlist)
    c2.metric("Total Paper Trades", total_trades)
    c3.metric("Open Trades", open_trades)
    c4.metric("Closed Trades", closed_trades)

    st.markdown("### 🇮🇳 Popular Indian Symbols")

    sample_symbols = pd.DataFrame({
        "Company": [
            "Reliance Industries",
            "TCS",
            "Infosys",
            "HDFC Bank",
            "ICICI Bank",
            "SBI",
            "Tata Motors",
            "Larsen & Toubro",
            "ITC",
            "Bharti Airtel"
        ],
        "Yahoo Symbol": [
            "RELIANCE.NS",
            "TCS.NS",
            "INFY.NS",
            "HDFCBANK.NS",
            "ICICIBANK.NS",
            "SBIN.NS",
            "TATAMOTORS.NS",
            "LT.NS",
            "ITC.NS",
            "BHARTIARTL.NS"
        ]
    })

    st.dataframe(sample_symbols, width="stretch", hide_index=True)

    st.markdown("### How to Use")
    st.info("""
    1. Go to Stock Analysis.
    2. Enter a stock symbol like RELIANCE.NS.
    3. Review chart, RSI, moving averages and educational signal.
    4. Add good stocks to Watchlist.
    5. Use Paper Trading before real trading.
    """)


# =========================================================
# STOCK ANALYSIS
# =========================================================
elif page == "Stock Analysis":
    st.subheader("🔎 Stock Analysis")

    col1, col2, col3 = st.columns([2, 1, 1])

    with col1:
        symbol_input = st.text_input(
            "Enter Stock Symbol",
            value="RELIANCE.NS",
            help="For Indian stocks use .NS, example: RELIANCE.NS, TCS.NS, INFY.NS"
        )

    with col2:
        period = st.selectbox(
            "Period",
            ["1mo", "3mo", "6mo", "1y", "2y", "5y"],
            index=2
        )

    with col3:
        interval = st.selectbox(
            "Interval",
            ["1d", "1wk", "1mo"],
            index=0
        )

    symbol = clean_symbol(symbol_input)

    if st.button("🚀 Start Analysis", type="primary"):
        st.session_state["selected_symbol"] = symbol
        st.session_state["selected_period"] = period
        st.session_state["selected_interval"] = interval

    if "selected_symbol" in st.session_state:
        symbol = st.session_state["selected_symbol"]
        period = st.session_state["selected_period"]
        interval = st.session_state["selected_interval"]

        df = load_stock_data(symbol, period, interval)

        if df.empty:
            st.error("No data found. Please check the stock symbol.")
        else:
            df = add_indicators(df)

            latest = df.iloc[-1]
            previous = df.iloc[-2] if len(df) > 1 else latest

            current_price = float(latest["Close"])
            previous_price = float(previous["Close"])
            change = current_price - previous_price
            change_pct = (change / previous_price) * 100 if previous_price else 0

            signal, reason = generate_signal(df)

            c1, c2, c3, c4 = st.columns(4)
            c1.metric("Current Price", f"₹{current_price:,.2f}", f"{change_pct:.2f}%")
            c2.metric("Volume", f"{float(latest['Volume']):,.0f}")
            c3.metric("RSI", f"{float(latest['RSI']):.2f}" if not pd.isna(latest["RSI"]) else "N/A")
            c4.metric("Signal", signal)

            if signal == "BUY":
                st.markdown(f"<div class='buy-box'>🟢 BUY WATCH: {reason}</div>", unsafe_allow_html=True)
            elif signal == "SELL":
                st.markdown(f"<div class='sell-box'>🔴 SELL / AVOID: {reason}</div>", unsafe_allow_html=True)
            else:
                st.markdown(f"<div class='hold-box'>🟡 HOLD / WAIT: {reason}</div>", unsafe_allow_html=True)

            st.plotly_chart(make_candlestick_chart(df, symbol), width="stretch")
            st.plotly_chart(make_rsi_chart(df, symbol), width="stretch")

            col_a, col_b = st.columns(2)

            with col_a:
                if st.button("➕ Add to Watchlist"):
                    add_to_watchlist(symbol)
                    st.success(f"{symbol} added to watchlist.")

            with col_b:
                st.download_button(
                    label="📥 Download Stock Data Excel",
                    data=df_to_excel({symbol: df}),
                    file_name=f"{symbol}_analysis.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

            with st.expander("View Raw Data"):
                st.dataframe(df.tail(100), width="stretch", hide_index=True)


# =========================================================
# WATCHLIST
# =========================================================
elif page == "Watchlist":
    st.subheader("⭐ Watchlist")

    col1, col2 = st.columns([3, 1])

    with col1:
        new_symbol = st.text_input("Add Stock Symbol", placeholder="Example: TCS.NS")

    with col2:
        st.write("")
        st.write("")
        if st.button("Add Stock"):
            if new_symbol.strip():
                symbol = clean_symbol(new_symbol)
                add_to_watchlist(symbol)
                st.success(f"{symbol} added.")

    watchlist_df = get_watchlist()

    if watchlist_df.empty:
        st.info("No stock added in watchlist yet.")
    else:
        enriched_rows = []

        for _, row in watchlist_df.iterrows():
            symbol = row["symbol"]
            price = get_current_price(symbol)

            enriched_rows.append({
                "Symbol": symbol,
                "Current Price": price if price is not None else "N/A",
                "Added On": row["added_on"]
            })

        display_df = pd.DataFrame(enriched_rows)
        st.dataframe(display_df, width="stretch", hide_index=True)

        delete_symbol = st.selectbox("Select Symbol to Remove", watchlist_df["symbol"].tolist())

        if st.button("🗑️ Remove Selected Symbol"):
            delete_watchlist_symbol(delete_symbol)
            st.success(f"{delete_symbol} removed. Please refresh if needed.")

        st.download_button(
            label="📥 Download Watchlist Excel",
            data=df_to_excel({"Watchlist": display_df}),
            file_name="watchlist.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )


# =========================================================
# PAPER TRADING
# =========================================================
elif page == "Paper Trading":
    st.subheader("🧾 Paper Trading")

    with st.form("paper_trade_form"):
        c1, c2, c3 = st.columns(3)

        with c1:
            symbol = clean_symbol(st.text_input("Stock Symbol", value="RELIANCE.NS"))
            trade_type = st.selectbox("Trade Type", ["BUY", "SELL"])

        with c2:
            entry_price = st.number_input("Entry Price", min_value=0.01, value=1000.00, step=1.00)
            quantity = st.number_input("Quantity", min_value=1, value=1, step=1)

        with c3:
            stop_loss = st.number_input("Stop Loss", min_value=0.01, value=950.00, step=1.00)
            target = st.number_input("Target", min_value=0.01, value=1100.00, step=1.00)

        trade_date = st.date_input("Trade Date", value=datetime.today())

        submitted = st.form_submit_button("💾 Save Paper Trade")

        if submitted:
            save_trade(symbol, trade_type, entry_price, quantity, stop_loss, target, trade_date)
            st.success("Paper trade saved successfully.")

    st.markdown("### 📌 Saved Paper Trades")

    trades_df = get_trades()

    if trades_df.empty:
        st.info("No paper trades saved yet.")
    else:
        report_rows = []

        for _, row in trades_df.iterrows():
            symbol = row["symbol"]
            current_price = get_current_price(symbol)

            entry_price = float(row["entry_price"])
            quantity = int(row["quantity"])

            if current_price is None:
                current_price = entry_price

            if row["trade_type"] == "BUY":
                pnl = (current_price - entry_price) * quantity
            else:
                pnl = (entry_price - current_price) * quantity

            investment = entry_price * quantity
            pnl_pct = (pnl / investment) * 100 if investment else 0

            risk_per_share = abs(entry_price - float(row["stop_loss"]))
            reward_per_share = abs(float(row["target"]) - entry_price)
            risk_reward = reward_per_share / risk_per_share if risk_per_share else 0

            report_rows.append({
                "ID": row["id"],
                "Symbol": symbol,
                "Type": row["trade_type"],
                "Entry": entry_price,
                "Current": round(current_price, 2),
                "Qty": quantity,
                "Investment": round(investment, 2),
                "P/L": round(pnl, 2),
                "P/L %": round(pnl_pct, 2),
                "Stop Loss": row["stop_loss"],
                "Target": row["target"],
                "Risk Reward": round(risk_reward, 2),
                "Status": row["status"],
                "Trade Date": row["trade_date"]
            })

        report_df = pd.DataFrame(report_rows)

        st.dataframe(report_df, width="stretch", hide_index=True)

        col1, col2, col3 = st.columns(3)

        total_pnl = report_df["P/L"].sum()
        total_investment = report_df["Investment"].sum()
        total_pnl_pct = (total_pnl / total_investment) * 100 if total_investment else 0

        col1.metric("Total Investment", f"₹{total_investment:,.2f}")
        col2.metric("Total P/L", f"₹{total_pnl:,.2f}")
        col3.metric("Total P/L %", f"{total_pnl_pct:.2f}%")

        st.markdown("### Manage Trade")

        selected_id = st.selectbox("Select Trade ID", report_df["ID"].tolist())

        c1, c2, c3 = st.columns(3)

        with c1:
            if st.button("Mark as OPEN"):
                update_trade_status(selected_id, "OPEN")
                st.success("Trade marked as OPEN.")

        with c2:
            if st.button("Mark as CLOSED"):
                update_trade_status(selected_id, "CLOSED")
                st.success("Trade marked as CLOSED.")

        with c3:
            if st.button("Delete Trade"):
                delete_trade(selected_id)
                st.success("Trade deleted.")

        st.download_button(
            label="📥 Download Paper Trading Report",
            data=df_to_excel({"Paper Trades": report_df}),
            file_name="paper_trading_report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )


# =========================================================
# REPORTS
# =========================================================
elif page == "Reports":
    st.subheader("📑 Reports & Export")

    watchlist_df = get_watchlist()
    trades_df = get_trades()

    st.markdown("### Watchlist")
    if watchlist_df.empty:
        st.info("No watchlist data available.")
    else:
        st.dataframe(watchlist_df, width="stretch", hide_index=True)

    st.markdown("### Paper Trades")
    if trades_df.empty:
        st.info("No paper trading data available.")
    else:
        st.dataframe(trades_df, width="stretch", hide_index=True)

    export_data = {
        "Watchlist": watchlist_df,
        "Paper Trades": trades_df
    }

    st.download_button(
        label="📥 Download Complete Trading Report",
        data=df_to_excel(export_data),
        file_name="BAJRABHANU_Trading_Report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    st.warning("""
    This report is based on paper trading and educational market analysis.
    It should not be treated as investment advice.
    """)


# =========================================================
# FOOTER
# =========================================================
st.markdown("""
<br>
<div style="text-align:center; color:#cbd5e1; font-size:13px;">
    © BAJRABHANU Trading Desk | Educational Trading & Analysis Application
</div>
""", unsafe_allow_html=True)
