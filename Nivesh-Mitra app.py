
import streamlit as st
import streamlit.components.v1 as components
import pandas as pd
import numpy as np
import yfinance as yf
import plotly.graph_objects as go
import plotly.express as px
import requests
import feedparser
from io import StringIO
from pathlib import Path
from datetime import datetime
import urllib.parse
import math

# ============================================================
# CONFIG
# ============================================================

st.set_page_config(
    page_title="Nivesh Mitra Live Market Dashboard",
    page_icon="🍋",
    layout="wide",
    initial_sidebar_state="expanded"
)

DATA_DIR = Path("data")
DATA_DIR.mkdir(exist_ok=True)
CACHE_FILE = DATA_DIR / "nse_universe_cache.csv"
NSE_SEC_LIST_URL = "https://nsearchives.nseindia.com/content/equities/sec_list.csv"

APP_DISCLAIMER = """
This dashboard is for learning, screening, and risk planning only. It does not connect to Groww,
does not ask for broker login, does not place orders, and does not guarantee profit. Free/public
market data may be delayed, rate-limited, or unavailable. Verify every price in your broker app
before trading.
"""

STARTER_SYMBOLS = [
    ("YESBANK.NS", "Yes Bank", "EQ", "Banking"),
    ("IDFCFIRSTB.NS", "IDFC First Bank", "EQ", "Banking"),
    ("SUZLON.NS", "Suzlon Energy", "EQ", "Renewable Energy"),
    ("SJVN.NS", "SJVN", "EQ", "Power"),
    ("UCOBANK.NS", "UCO Bank", "EQ", "Banking"),
    ("CENTRALBK.NS", "Central Bank of India", "EQ", "Banking"),
    ("IOB.NS", "Indian Overseas Bank", "EQ", "Banking"),
    ("NMDC.NS", "NMDC", "EQ", "Metals"),
    ("BANKMAHARAS.NS", "Bank of Maharashtra", "EQ", "Banking"),
    ("IDEA.NS", "Vodafone Idea", "EQ", "Telecom"),
    ("IRFC.NS", "Indian Railway Finance Corporation", "EQ", "Finance"),
    ("PNB.NS", "Punjab National Bank", "EQ", "Banking"),
    ("SAIL.NS", "Steel Authority of India", "EQ", "Metals"),
    ("HUDCO.NS", "HUDCO", "EQ", "Finance"),
    ("GAIL.NS", "GAIL", "EQ", "Energy"),
    ("NTPC.NS", "NTPC", "EQ", "Power"),
    ("TATASTEEL.NS", "Tata Steel", "EQ", "Metals"),
    ("ONGC.NS", "ONGC", "EQ", "Energy"),
    ("RECLTD.NS", "REC", "EQ", "Finance"),
    ("NATIONALUM.NS", "National Aluminium", "EQ", "Metals"),
]

POSITIVE_WORDS = [
    "profit", "profits", "growth", "upgrade", "upgraded", "order", "wins", "approval",
    "approved", "dividend", "bonus", "buyback", "strong", "record", "rises", "surges",
    "expansion", "launch", "partnership", "deal", "raises", "beats"
]
NEGATIVE_WORDS = [
    "loss", "losses", "downgrade", "downgraded", "probe", "investigation", "penalty",
    "fine", "default", "debt", "weak", "falls", "slumps", "resigns", "cuts", "misses",
    "decline", "declines", "fraud", "delay", "risk"
]

# ============================================================
# STYLE
# ============================================================

def inject_css():
    st.markdown(
        """
        <style>
        :root {
            --lime: #D7FF39;
            --lime2: #A8FF00;
            --bg1: #07130B;
            --bg2: #0D1F13;
            --card: rgba(16, 32, 22, 0.78);
            --muted: #9BAA9F;
            --text: #F2FFE4;
            --danger: #FF5C7A;
            --warn: #FFCB47;
            --ok: #71FF9A;
            --cyan: #68F5FF;
        }

        .stApp {
            background:
                radial-gradient(circle at top left, rgba(215,255,57,0.12), transparent 26%),
                radial-gradient(circle at top right, rgba(104,245,255,0.08), transparent 28%),
                linear-gradient(135deg, #050B07 0%, #07130B 48%, #0E1F12 100%);
            color: var(--text);
        }

        [data-testid="stSidebar"] {
            background: linear-gradient(180deg, rgba(8,19,10,0.98), rgba(12,31,18,0.98));
            border-right: 1px solid rgba(215,255,57,0.20);
        }

        .hero {
            padding: 24px 26px;
            border-radius: 26px;
            background: linear-gradient(135deg, rgba(215,255,57,0.14), rgba(104,245,255,0.06));
            border: 1px solid rgba(215,255,57,0.24);
            box-shadow: 0 18px 50px rgba(0,0,0,0.35);
            margin-bottom: 18px;
        }

        .hero-title {
            font-size: 42px;
            line-height: 1.05;
            font-weight: 900;
            letter-spacing: -1.4px;
            margin: 0;
            color: #F6FFE8;
        }

        .hero-title span {
            color: var(--lime);
            text-shadow: 0 0 28px rgba(215,255,57,0.45);
        }

        .hero-sub {
            margin-top: 10px;
            color: #BFD2C5;
            font-size: 16px;
        }

        .glass-card {
            padding: 18px 18px;
            border-radius: 22px;
            background: rgba(13,31,19,0.76);
            border: 1px solid rgba(215,255,57,0.14);
            box-shadow: 0 14px 34px rgba(0,0,0,0.25);
        }

        .metric-label {
            color: #9BAA9F;
            font-size: 13px;
            margin-bottom: 5px;
        }

        .metric-value {
            color: #F2FFE4;
            font-size: 28px;
            font-weight: 850;
        }

        .pill {
            display:inline-block;
            padding: 6px 12px;
            border-radius: 999px;
            font-size: 12px;
            font-weight: 800;
            letter-spacing: .4px;
            margin-right: 6px;
            border: 1px solid rgba(255,255,255,0.12);
        }

        .pill-buy { background: rgba(113,255,154,0.14); color: #71FF9A; border-color: rgba(113,255,154,0.36); }
        .pill-wait { background: rgba(255,203,71,0.14); color: #FFCB47; border-color: rgba(255,203,71,0.36); }
        .pill-avoid { background: rgba(255,92,122,0.14); color: #FF5C7A; border-color: rgba(255,92,122,0.36); }
        .pill-info { background: rgba(104,245,255,0.12); color: #68F5FF; border-color: rgba(104,245,255,0.28); }

        div[data-testid="stMetric"] {
            background: rgba(13,31,19,0.62);
            border: 1px solid rgba(215,255,57,0.13);
            padding: 14px 14px;
            border-radius: 18px;
        }

        .stDataFrame {
            border-radius: 20px;
            overflow: hidden;
        }

        .small-note {
            color: #9BAA9F;
            font-size: 13px;
        }

        .news-card {
            padding: 14px;
            margin-bottom: 10px;
            border-radius: 16px;
            background: rgba(16,32,22,0.70);
            border: 1px solid rgba(215,255,57,0.12);
        }

        .news-title {
            font-weight: 800;
            color: #F2FFE4;
            margin-bottom: 6px;
        }

        .news-meta {
            color: #9BAA9F;
            font-size: 12px;
        }

        .block-container { padding-top: 1.3rem; }
        </style>
        """,
        unsafe_allow_html=True,
    )

inject_css()

# ============================================================
# HELPERS
# ============================================================

def auto_refresh(seconds: int):
    """Simple browser refresh. Min recommended interval: 1 second."""
    if seconds >= 1:
        components.html(
            f"""
            <script>
            setTimeout(function(){{
                window.parent.location.reload();
            }}, {int(seconds * 1000)});
            </script>
            """,
            height=0,
        )

def starter_df():
    return pd.DataFrame(STARTER_SYMBOLS, columns=["symbol", "name", "series", "sector"])

def clean_symbol(symbol):
    symbol = str(symbol).strip().upper()
    if not symbol:
        return ""
    if "." not in symbol:
        symbol = symbol + ".NS"
    return symbol

def raw_symbol(symbol):
    return str(symbol).replace(".NS", "")

def groww_url(symbol):
    return f"https://groww.in/search?query={urllib.parse.quote(raw_symbol(symbol))}"

def tradingview_url(symbol):
    return f"https://www.tradingview.com/chart/?symbol=NSE%3A{urllib.parse.quote(raw_symbol(symbol))}"

def load_universe_fast():
    if CACHE_FILE.exists():
        try:
            df = pd.read_csv(CACHE_FILE)
            required = {"symbol", "name", "series"}
            if required.issubset(set(df.columns)) and not df.empty:
                if "sector" not in df.columns:
                    df["sector"] = "Unknown"
                return df, "Loaded saved NSE universe cache"
        except Exception:
            pass
    return starter_df(), "Loaded starter list. Click Sync NSE Universe to load more symbols."

def sync_nse_universe():
    headers = {
        "User-Agent": "Mozilla/5.0",
        "Accept": "text/csv,text/plain,*/*",
        "Referer": "https://www.nseindia.com/",
    }
    response = requests.get(NSE_SEC_LIST_URL, headers=headers, timeout=10)
    response.raise_for_status()
    raw = pd.read_csv(StringIO(response.text))
    raw.columns = [str(c).strip().upper().replace(" ", "_").replace("-", "_") for c in raw.columns]

    if "SYMBOL" not in raw.columns:
        raise ValueError("NSE file format changed. SYMBOL column not found.")

    name_col = "NAME_OF_COMPANY" if "NAME_OF_COMPANY" in raw.columns else None
    series_col = "SERIES" if "SERIES" in raw.columns else None

    df = pd.DataFrame()
    df["symbol"] = raw["SYMBOL"].astype(str).str.strip().str.upper().apply(clean_symbol)
    df["name"] = raw[name_col].astype(str) if name_col else df["symbol"].str.replace(".NS", "", regex=False)
    df["series"] = raw[series_col].astype(str).str.strip().str.upper() if series_col else "EQ"
    df["sector"] = "Unknown"
    df = df[df["symbol"].str.len() > 3].drop_duplicates("symbol").sort_values("symbol").reset_index(drop=True)
    df.to_csv(CACHE_FILE, index=False)
    return df

@st.cache_data(ttl=8)
def fetch_intraday(symbol, period="1d", interval="1m"):
    try:
        data = yf.download(
            symbol,
            period=period,
            interval=interval,
            progress=False,
            auto_adjust=False,
            threads=False,
            timeout=10,
        )
        if data is None or data.empty:
            return pd.DataFrame()
        if isinstance(data.columns, pd.MultiIndex):
            data.columns = data.columns.get_level_values(0)
        needed = {"Open", "High", "Low", "Close", "Volume"}
        if not needed.issubset(set(data.columns)):
            return pd.DataFrame()
        return data.dropna().copy()
    except Exception:
        return pd.DataFrame()

@st.cache_data(ttl=12)
def fetch_batch_intraday(symbols, period="1d", interval="1m"):
    symbols = list(dict.fromkeys([s for s in symbols if s]))
    if not symbols:
        return pd.DataFrame()

    try:
        data = yf.download(
            tickers=symbols,
            period=period,
            interval=interval,
            group_by="ticker",
            progress=False,
            auto_adjust=False,
            threads=True,
            timeout=15,
        )
        return data
    except Exception:
        return pd.DataFrame()

@st.cache_data(ttl=180)
def fetch_news(query, limit=8):
    try:
        url = "https://news.google.com/rss/search?q=" + urllib.parse.quote(query + " stock NSE") + "&hl=en-IN&gl=IN&ceid=IN:en"
        feed = feedparser.parse(url)
        items = []
        for entry in feed.entries[:limit]:
            title = getattr(entry, "title", "")
            link = getattr(entry, "link", "")
            published = getattr(entry, "published", "")
            source = ""
            if hasattr(entry, "source"):
                source = getattr(entry.source, "title", "")
            impact, sentiment_score = classify_news_impact(title)
            items.append({
                "title": title,
                "source": source,
                "published": published,
                "link": link,
                "impact": impact,
                "sentiment_score": sentiment_score,
            })
        return items
    except Exception:
        return []

def classify_news_impact(text):
    text_l = str(text).lower()
    pos = sum(1 for w in POSITIVE_WORDS if w in text_l)
    neg = sum(1 for w in NEGATIVE_WORDS if w in text_l)
    score = pos - neg
    if score >= 2:
        return "Potential positive impact", score
    if score == 1:
        return "Mild positive impact", score
    if score <= -2:
        return "Potential negative impact", score
    if score == -1:
        return "Mild negative impact", score
    return "Neutral / unclear impact", score

def get_symbol_frame(batch_data, symbol):
    if batch_data is None or batch_data.empty:
        return pd.DataFrame()
    try:
        if isinstance(batch_data.columns, pd.MultiIndex):
            if symbol not in batch_data.columns.get_level_values(0):
                return pd.DataFrame()
            df = batch_data[symbol].dropna().copy()
        else:
            df = batch_data.dropna().copy()
        needed = {"Open", "High", "Low", "Close", "Volume"}
        if not needed.issubset(set(df.columns)):
            return pd.DataFrame()
        return df
    except Exception:
        return pd.DataFrame()

def vwap(df):
    typical = (df["High"] + df["Low"] + df["Close"]) / 3
    volume = df["Volume"].replace(0, np.nan)
    return (typical * volume).cumsum() / volume.cumsum()

def rsi(series, period=14):
    delta = series.diff()
    gain = delta.clip(lower=0).rolling(period).mean()
    loss = (-delta.clip(upper=0)).rolling(period).mean()
    rs = gain / loss.replace(0, np.nan)
    out = 100 - (100 / (1 + rs))
    return out.fillna(50)

def macd(series):
    ema12 = series.ewm(span=12, adjust=False).mean()
    ema26 = series.ewm(span=26, adjust=False).mean()
    macd_line = ema12 - ema26
    signal = macd_line.ewm(span=9, adjust=False).mean()
    hist = macd_line - signal
    return macd_line, signal, hist

def atr(df, period=14):
    high_low = df["High"] - df["Low"]
    high_close = (df["High"] - df["Close"].shift()).abs()
    low_close = (df["Low"] - df["Close"].shift()).abs()
    tr = pd.concat([high_low, high_close, low_close], axis=1).max(axis=1)
    return tr.rolling(period).mean().fillna(tr.mean())

def add_indicators(df):
    if df.empty:
        return df
    out = df.copy()
    out["VWAP"] = vwap(out)
    out["EMA9"] = out["Close"].ewm(span=9, adjust=False).mean()
    out["EMA20"] = out["Close"].ewm(span=20, adjust=False).mean()
    out["EMA50"] = out["Close"].ewm(span=50, adjust=False).mean()
    out["RSI"] = rsi(out["Close"])
    out["MACD"], out["MACD_SIGNAL"], out["MACD_HIST"] = macd(out["Close"])
    out["ATR"] = atr(out)
    return out

def analyze_df(symbol, name, df, max_price=None, min_volume=0):
    if df.empty or len(df) < 10:
        return None

    df = add_indicators(df)
    last = df.iloc[-1]

    price = float(last["Close"])
    open_price = float(df["Open"].iloc[0])
    day_high = float(df["High"].max())
    day_low = float(df["Low"].min())
    volume = int(df["Volume"].sum())
    prev_close = float(df["Close"].iloc[-2]) if len(df) >= 2 else price
    pct_move = ((price - open_price) / open_price * 100) if open_price else 0

    if max_price is not None and price > max_price:
        return None
    if volume < min_volume:
        return None

    curr_vwap = float(last["VWAP"]) if not pd.isna(last["VWAP"]) else price
    ema9 = float(last["EMA9"])
    ema20 = float(last["EMA20"])
    ema50 = float(last["EMA50"])
    rsi_now = float(last["RSI"])
    macd_hist = float(last["MACD_HIST"])
    atr_now = float(last["ATR"]) if not pd.isna(last["ATR"]) and last["ATR"] > 0 else max(price * 0.01, 0.05)

    above_vwap = price > curr_vwap
    ema_bull = ema9 > ema20 > ema50
    ema_bear = ema9 < ema20 < ema50
    near_high = price >= day_high * 0.995
    near_low = price <= day_low * 1.005
    rsi_ok_buy = 45 <= rsi_now <= 72
    rsi_risky = rsi_now > 75 or rsi_now < 30
    momentum_ok = macd_hist > 0

    score = 0
    reasons = []

    if above_vwap:
        score += 22
        reasons.append("price above VWAP")
    else:
        reasons.append("price below VWAP")

    if ema_bull:
        score += 22
        reasons.append("EMA trend bullish")
    elif ema_bear:
        score -= 8
        reasons.append("EMA trend bearish")
    else:
        score += 6
        reasons.append("EMA trend mixed")

    if near_high:
        score += 16
        reasons.append("near day high")
    elif near_low:
        score -= 10
        reasons.append("near day low")
    else:
        score += 4
        reasons.append("middle of range")

    if rsi_ok_buy:
        score += 14
        reasons.append("RSI in tradable range")
    elif rsi_risky:
        score -= 8
        reasons.append("RSI stretched")
    else:
        score += 3
        reasons.append("RSI neutral")

    if momentum_ok:
        score += 14
        reasons.append("MACD momentum positive")
    else:
        reasons.append("MACD momentum weak")

    if pct_move > 0:
        score += 12
        reasons.append("positive intraday move")
    else:
        reasons.append("negative intraday move")

    score = max(0, min(100, int(score)))

    if score >= 78:
        signal = "BUY WATCH"
        risk_label = "Medium"
    elif score >= 58:
        signal = "WAIT / WATCH"
        risk_label = "Medium"
    elif score >= 40:
        signal = "NO TRADE"
        risk_label = "High"
    else:
        signal = "AVOID / WEAK"
        risk_label = "High"

    stop_loss = min(curr_vwap, price - atr_now * 0.8)
    if stop_loss >= price:
        stop_loss = price - max(price * 0.006, 0.05)
    target_1 = price + (price - stop_loss) * 1.5
    target_2 = price + (price - stop_loss) * 2.0

    sell_risk = "Low while price holds above VWAP/EMA20"
    if price < curr_vwap:
        sell_risk = "High: price below VWAP"
    elif rsi_now > 75:
        sell_risk = "Medium-high: RSI overbought"
    elif price < ema20:
        sell_risk = "Medium: price below EMA20"

    if signal == "BUY WATCH":
        outlook = "Bullish intraday bias if price holds above VWAP and EMA20. Weakness below stop-loss cancels setup."
    elif signal == "WAIT / WATCH":
        outlook = "Mixed setup. Wait for breakout above day high or pullback near VWAP with strong volume."
    else:
        outlook = "Weak or unclear setup. Better to avoid until price recovers above VWAP with momentum."

    return {
        "Symbol": symbol,
        "Name": name,
        "Price": round(price, 2),
        "Change from open %": round(pct_move, 2),
        "VWAP": round(curr_vwap, 2),
        "EMA9": round(ema9, 2),
        "EMA20": round(ema20, 2),
        "EMA50": round(ema50, 2),
        "RSI": round(rsi_now, 1),
        "MACD Hist": round(macd_hist, 4),
        "Day High": round(day_high, 2),
        "Day Low": round(day_low, 2),
        "Volume": volume,
        "Score": score,
        "Signal": signal,
        "Risk": risk_label,
        "Sell risk": sell_risk,
        "Stop-loss idea": round(stop_loss, 2),
        "Target 1 idea": round(target_1, 2),
        "Target 2 idea": round(target_2, 2),
        "Outlook": outlook,
        "Reasons": "; ".join(reasons),
        "Groww": groww_url(symbol),
        "TradingView": tradingview_url(symbol),
        "Last updated": datetime.now().strftime("%H:%M:%S"),
    }

def signal_pill(signal):
    s = str(signal).upper()
    if "BUY" in s:
        cls = "pill pill-buy"
    elif "WAIT" in s or "NO TRADE" in s:
        cls = "pill pill-wait"
    else:
        cls = "pill pill-avoid"
    return f'<span class="{cls}">{signal}</span>'

def risk_pill(risk):
    r = str(risk).upper()
    if "LOW" in r:
        cls = "pill pill-buy"
    elif "MEDIUM" in r:
        cls = "pill pill-wait"
    else:
        cls = "pill pill-avoid"
    return f'<span class="{cls}">{risk}</span>'

def qty_from_risk(capital, risk_pct, entry, stop):
    risk_amount = capital * risk_pct / 100
    per_share = max(entry - stop, 0.01)
    qty_by_risk = int(risk_amount // per_share)
    qty_by_cash = int(capital // max(entry, 0.01))
    qty = max(0, min(qty_by_risk, qty_by_cash))
    return qty, risk_amount, per_share

# ============================================================
# STATE
# ============================================================

if "universe" not in st.session_state:
    st.session_state["universe"], st.session_state["universe_status"] = load_universe_fast()

if "dashboard_rows" not in st.session_state:
    st.session_state["dashboard_rows"] = pd.DataFrame()

if "chart_cache" not in st.session_state:
    st.session_state["chart_cache"] = {}

# ============================================================
# HERO
# ============================================================

st.markdown(
    """
    <div class="hero">
        <div class="hero-title">🍋 Nivesh Mitra <span>Live Market</span> Dashboard</div>
        <div class="hero-sub">
            Fresh lime UI • watchlist price board • chart analysis • risk zones • news impact • no broker password
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)

st.warning(APP_DISCLAIMER)

# ============================================================
# SIDEBAR
# ============================================================

with st.sidebar:
    st.header("🍋 Market sync")

    st.caption(st.session_state["universe_status"])

    if st.button("Sync NSE Universe", type="primary", use_container_width=True):
        try:
            with st.spinner("Syncing NSE symbol universe..."):
                st.session_state["universe"] = sync_nse_universe()
                st.session_state["universe_status"] = f"NSE universe synced at {datetime.now().strftime('%H:%M:%S')}"
            st.success("NSE universe synced.")
        except Exception as e:
            st.error(f"NSE sync failed: {e}")
            st.info("You can continue with starter list or upload a CSV.")

    uploaded = st.file_uploader("Upload custom CSV watchlist", type=["csv"], help="CSV should contain symbol column. Optional: name, series, sector")
    if uploaded is not None:
        try:
            custom = pd.read_csv(uploaded)
            custom.columns = [c.lower().strip() for c in custom.columns]
            if "symbol" not in custom.columns:
                st.error("CSV must contain a symbol column.")
            else:
                custom["symbol"] = custom["symbol"].apply(clean_symbol)
                if "name" not in custom.columns:
                    custom["name"] = custom["symbol"]
                if "series" not in custom.columns:
                    custom["series"] = "EQ"
                if "sector" not in custom.columns:
                    custom["sector"] = "Custom"
                st.session_state["universe"] = custom[["symbol", "name", "series", "sector"]].drop_duplicates("symbol")
                st.session_state["universe_status"] = "Loaded your custom CSV"
                st.success("Custom watchlist loaded.")
        except Exception as e:
            st.error(f"CSV load failed: {e}")

    universe = st.session_state["universe"].copy()
    if "sector" not in universe.columns:
        universe["sector"] = "Unknown"

    st.write(f"Symbols loaded: **{len(universe):,}**")

    st.header("⚙️ Dashboard settings")
    series_values = sorted(universe["series"].astype(str).dropna().unique().tolist())
    default_series = ["EQ"] if "EQ" in series_values else series_values[:1]
    selected_series = st.multiselect("Series", series_values, default=default_series)

    search_text = st.text_input("Search stock", "")
    max_price = st.number_input("Max price filter ₹", min_value=1.0, value=100.0, step=1.0)
    min_volume = st.number_input("Min intraday volume", min_value=0, value=0, step=25000)

    batch_size = st.slider("Dashboard stocks to refresh", 5, 60, 20, 5)
    start_row = st.number_input("Start row", min_value=0, max_value=max(len(universe) - 1, 0), value=0, step=20)

    st.header("⚡ Auto update")
    auto = st.toggle("Auto-refresh dashboard", value=False)
    refresh_seconds = st.slider("Refresh interval seconds", 1, 30, 5, 1)
    st.caption("Realistic minimum is 1 second. Free/public data may still update slower.")

    st.header("💰 Risk settings")
    capital = st.number_input("Capital ₹", min_value=100.0, value=100.0, step=50.0)
    risk_pct = st.slider("Risk per trade %", 0.5, 5.0, 1.0, 0.5)

    run_dashboard = st.button("Refresh price dashboard", type="primary", use_container_width=True)

    if st.button("Clear data cache", use_container_width=True):
        st.cache_data.clear()
        st.success("Cache cleared.")

if auto:
    auto_refresh(refresh_seconds)

# Filter universe
filtered = universe[universe["series"].astype(str).isin(selected_series)].copy()
if search_text.strip():
    q = search_text.strip().upper()
    filtered = filtered[
        filtered["symbol"].astype(str).str.upper().str.contains(q, na=False) |
        filtered["name"].astype(str).str.upper().str.contains(q, na=False)
    ]

# ============================================================
# TABS
# ============================================================

tab_dash, tab_stock, tab_news, tab_risk, tab_setup = st.tabs([
    "⚡ Live Dashboard",
    "📈 Stock Lab",
    "📰 News Impact",
    "🛡️ Risk Console",
    "🚀 Setup Notes"
])

# ============================================================
# TAB 1: LIVE DASHBOARD
# ============================================================

with tab_dash:
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Selected universe", f"{len(filtered):,}")
    c2.metric("Batch refresh", f"{batch_size} stocks")
    c3.metric("Auto refresh", "ON" if auto else "OFF")
    c4.metric("Last refresh", datetime.now().strftime("%H:%M:%S"))

    st.markdown('<span class="pill pill-info">Realtime-style dashboard</span><span class="pill pill-wait">Public data speed depends on source</span>', unsafe_allow_html=True)

    if run_dashboard or auto:
        chunk = filtered.iloc[int(start_row): int(start_row) + int(batch_size)].copy()
        symbols = chunk["symbol"].tolist()
        batch = fetch_batch_intraday(symbols, period="1d", interval="1m")
        rows = []
        chart_cache = {}

        progress = st.progress(0)
        status = st.empty()

        for i, (_, row) in enumerate(chunk.iterrows()):
            symbol = row["symbol"]
            name = row.get("name", symbol)
            df = get_symbol_frame(batch, symbol)
            if df.empty:
                df = fetch_intraday(symbol, period="1d", interval="1m")
            if not df.empty:
                analysis = analyze_df(symbol, name, df, max_price=max_price, min_volume=min_volume)
                if analysis:
                    rows.append(analysis)
                    chart_cache[symbol] = add_indicators(df)
            status.write(f"Updating {i+1}/{len(chunk)}: {symbol}")
            progress.progress((i+1)/max(len(chunk), 1))

        progress.empty()
        status.empty()

        out = pd.DataFrame(rows)
        if not out.empty:
            out = out.sort_values(["Score", "Volume"], ascending=[False, False]).reset_index(drop=True)
        st.session_state["dashboard_rows"] = out
        st.session_state["chart_cache"].update(chart_cache)

    dashboard = st.session_state.get("dashboard_rows", pd.DataFrame())
    if dashboard.empty:
        st.info("Click **Refresh price dashboard** from the sidebar. Start with 20 stocks for speed.")
    else:
        display_cols = [
            "Symbol", "Name", "Price", "Change from open %", "VWAP", "RSI", "Score",
            "Signal", "Risk", "Stop-loss idea", "Target 1 idea", "Volume", "Last updated"
        ]
        st.dataframe(
            dashboard[display_cols],
            use_container_width=True,
            hide_index=True,
            column_config={
                "Price": st.column_config.NumberColumn("Price", format="₹%.2f"),
                "Change from open %": st.column_config.NumberColumn("Open %", format="%.2f%%"),
                "VWAP": st.column_config.NumberColumn("VWAP", format="₹%.2f"),
                "Stop-loss idea": st.column_config.NumberColumn("Stop", format="₹%.2f"),
                "Target 1 idea": st.column_config.NumberColumn("Target 1", format="₹%.2f"),
                "Volume": st.column_config.NumberColumn("Volume", format="%d"),
            }
        )

        st.download_button(
            "Download dashboard CSV",
            data=dashboard.to_csv(index=False).encode("utf-8"),
            file_name="nivesh_mitra_dashboard.csv",
            mime="text/csv"
        )

        top = dashboard.head(8).copy()
        if not top.empty:
            fig = px.bar(
                top,
                x="Score",
                y="Symbol",
                orientation="h",
                hover_data=["Name", "Price", "Signal", "Risk"],
                title="Top dashboard scores"
            )
            fig.update_layout(height=420, yaxis={"categoryorder": "total ascending"})
            st.plotly_chart(fig, use_container_width=True)

# ============================================================
# TAB 2: STOCK LAB
# ============================================================

with tab_stock:
    st.subheader("📈 Open any stock and analyse chart + trend + buy/sell risk")

    manual_symbol = st.text_input("Type NSE symbol", value="SUZLON", help="Example: SUZLON, YESBANK, RELIANCE, TCS")
    selected_symbol = clean_symbol(manual_symbol)

    colA, colB, colC = st.columns([1,1,1])
    period = colA.selectbox("Period", ["1d", "5d", "1mo", "3mo", "6mo"], index=0)
    interval = colB.selectbox("Interval", ["1m", "5m", "15m", "30m", "1d"], index=0)
    load_stock = colC.button("Analyse stock", type="primary", use_container_width=True)

    if load_stock:
        df = fetch_intraday(selected_symbol, period=period, interval=interval)
        if df.empty:
            st.error("Could not load price data. Try another interval like 5m or 1d.")
        else:
            df = add_indicators(df)
            name_match = universe[universe["symbol"] == selected_symbol]
            name = name_match["name"].iloc[0] if not name_match.empty else raw_symbol(selected_symbol)
            analysis = analyze_df(selected_symbol, name, df, max_price=None, min_volume=0)

            if analysis:
                st.session_state["selected_analysis"] = analysis
                st.session_state["selected_df"] = df

    analysis = st.session_state.get("selected_analysis")
    df = st.session_state.get("selected_df", pd.DataFrame())

    if analysis and not df.empty:
        st.markdown(
            f"""
            <div class="glass-card">
              <h3>{analysis['Name']} — {analysis['Symbol']}</h3>
              {signal_pill(analysis['Signal'])} {risk_pill(analysis['Risk'])}
              <span class="pill pill-info">Updated {analysis['Last updated']}</span>
              <p class="small-note">{analysis['Outlook']}</p>
            </div>
            """,
            unsafe_allow_html=True,
        )

        m1, m2, m3, m4, m5 = st.columns(5)
        m1.metric("Price", f"₹{analysis['Price']}")
        m2.metric("VWAP", f"₹{analysis['VWAP']}")
        m3.metric("RSI", analysis["RSI"])
        m4.metric("Score", analysis["Score"])
        m5.metric("Open move", f"{analysis['Change from open %']}%")

        fig = go.Figure()
        fig.add_trace(go.Candlestick(
            x=df.index, open=df["Open"], high=df["High"], low=df["Low"], close=df["Close"], name="Price"
        ))
        fig.add_trace(go.Scatter(x=df.index, y=df["VWAP"], name="VWAP", mode="lines"))
        fig.add_trace(go.Scatter(x=df.index, y=df["EMA20"], name="EMA20", mode="lines"))
        fig.add_trace(go.Scatter(x=df.index, y=df["EMA50"], name="EMA50", mode="lines"))
        fig.add_hline(y=analysis["Stop-loss idea"], line_dash="dash", annotation_text="Stop idea")
        fig.add_hline(y=analysis["Target 1 idea"], line_dash="dot", annotation_text="Target 1 idea")
        fig.update_layout(height=620, xaxis_rangeslider_visible=False, title=f"{analysis['Symbol']} chart analysis")
        st.plotly_chart(fig, use_container_width=True)

        r1, r2, r3 = st.columns(3)
        r1.metric("Stop-loss idea", f"₹{analysis['Stop-loss idea']}")
        r2.metric("Target 1 idea", f"₹{analysis['Target 1 idea']}")
        r3.metric("Target 2 idea", f"₹{analysis['Target 2 idea']}")

        st.markdown("#### Trend reading")
        st.write("Reasons:", analysis["Reasons"])
        st.write("Sell risk:", analysis["Sell risk"])

        qty, risk_amount, per_share = qty_from_risk(capital, risk_pct, analysis["Price"], analysis["Stop-loss idea"])
        st.markdown("#### Position size")
        st.write(f"With capital **₹{capital:.2f}** and risk **{risk_pct}%**, suggested quantity is **{qty} share(s)**.")
        st.caption("This is a risk calculator only, not an order instruction.")

        b1, b2 = st.columns(2)
        b1.link_button("Open Groww manually", analysis["Groww"])
        b2.link_button("Open TradingView chart", analysis["TradingView"])
    else:
        st.info("Enter a symbol and click **Analyse stock**.")

# ============================================================
# TAB 3: NEWS IMPACT
# ============================================================

with tab_news:
    st.subheader("📰 Live news + possible price impact")
    news_symbol = st.text_input("News symbol/company", value="SUZLON", key="news_symbol")
    news_query = raw_symbol(clean_symbol(news_symbol))

    if st.button("Fetch news impact", type="primary"):
        st.session_state["news_items"] = fetch_news(news_query, limit=10)

    news_items = st.session_state.get("news_items", [])
    if not news_items:
        st.info("Click **Fetch news impact**. This uses Google News RSS and a keyword-based impact model.")
    else:
        score_total = sum(item["sentiment_score"] for item in news_items)
        if score_total > 2:
            overall = "News tone looks positive, but confirm with price and volume."
            pill = "pill-buy"
        elif score_total < -2:
            overall = "News tone looks negative/cautious. Risk may increase."
            pill = "pill-avoid"
        else:
            overall = "News tone is mixed or unclear."
            pill = "pill-wait"

        st.markdown(f'<span class="pill {pill}">{overall}</span>', unsafe_allow_html=True)
        st.caption("News impact is heuristic. It is not a guaranteed price prediction.")

        for item in news_items:
            if "positive" in item["impact"].lower():
                cls = "pill-buy"
            elif "negative" in item["impact"].lower():
                cls = "pill-avoid"
            else:
                cls = "pill-wait"

            st.markdown(
                f"""
                <div class="news-card">
                    <div class="news-title">{item['title']}</div>
                    <span class="pill {cls}">{item['impact']}</span>
                    <div class="news-meta">{item['source']} • {item['published']}</div>
                    <a href="{item['link']}" target="_blank">Open news</a>
                </div>
                """,
                unsafe_allow_html=True,
            )

# ============================================================
# TAB 4: RISK CONSOLE
# ============================================================

with tab_risk:
    st.subheader("🛡️ Buy/sell risk calculator")

    c1, c2, c3 = st.columns(3)
    entry = c1.number_input("Entry price ₹", min_value=0.01, value=50.00, step=0.05)
    stop = c2.number_input("Stop-loss price ₹", min_value=0.01, value=49.50, step=0.05)
    target = c3.number_input("Target price ₹", min_value=0.01, value=51.00, step=0.05)

    qty, risk_amount, per_share = qty_from_risk(capital, risk_pct, entry, stop)
    reward_per_share = max(target - entry, 0)
    rr = reward_per_share / per_share if per_share > 0 else 0

    k1, k2, k3, k4 = st.columns(4)
    k1.metric("Quantity", qty)
    k2.metric("Max risk", f"₹{risk_amount:.2f}")
    k3.metric("Risk/share", f"₹{per_share:.2f}")
    k4.metric("Risk:Reward", f"1:{rr:.2f}")

    if rr < 1.5:
        st.warning("Risk:Reward is weak. Consider avoiding or improving entry/stop/target.")
    else:
        st.success("Risk:Reward is acceptable for a learning setup.")

    st.info("For ₹100 capital, expect very small profits/losses after charges. Treat it as practice.")

# ============================================================
# TAB 5: SETUP NOTES
# ============================================================

with tab_setup:
    st.subheader("🚀 Realistic setup notes")

    st.markdown(
        """
        ### What has been added

        - Fresh lime/cooler UI
        - Auto-refresh dashboard for a selected stock batch
        - Price board with signal, score, VWAP, RSI, stop-loss idea, and target idea
        - Open any stock and view chart with VWAP, EMA20, EMA50
        - Trend/risk analysis from chart indicators
        - News feed with possible impact classification
        - Scenario-based outlook, not guaranteed prediction
        - Groww and TradingView manual links

        ### Important realistic limitation

        A free no-API app cannot update **every stock in the market under 1 second**. That needs a licensed real-time market feed or broker/exchange WebSocket.  
        This dashboard supports fast refresh for selected batches/watchlists. For best performance, use 20–60 stocks per refresh.

        ### Recommended Streamlit Cloud settings

        - Python version: 3.11 or 3.12
        - Main file path: `app.py`
        - Keep the files extracted in GitHub, not as a ZIP only
        """
    )

    st.code(
        """
pip install -r requirements.txt
python -m streamlit run app.py
        """.strip()
    )

    st.download_button(
        "Download current universe CSV",
        data=st.session_state["universe"].to_csv(index=False).encode("utf-8"),
        file_name="current_universe.csv",
        mime="text/csv"
    )

st.caption("Nivesh Mitra v3 • Built for learning, screening, and manual decision support.")
