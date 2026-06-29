
# Nivesh Mitra Live Market Dashboard v3

A fresh lime UI stock-market dashboard for learning, screening, chart analysis, risk planning, and news impact.

## Features

- No Groww login
- No API key
- No order placement
- Fresh lime/cooler design
- Price dashboard for selected stock batches
- Auto-refresh dashboard
- Open any stock and analyze chart
- VWAP, EMA9, EMA20, EMA50, RSI, MACD, ATR
- Buy-watch / wait / avoid labels
- Stop-loss and target ideas
- Risk calculator
- News impact using Google News RSS
- Manual Groww and TradingView links

## Important realistic limitation

A no-API free public-data app cannot update every stock in the entire market in under one second.
True sub-second full-market data needs a licensed market-data feed or broker/exchange WebSocket.

This app gives a realistic alternative:
- Fast refresh for selected watchlists
- Batch scanning
- Manual verification in Groww before trading

## Streamlit Cloud setup

1. Upload extracted files to GitHub, not only the ZIP.
2. Main file path: `app.py`
3. Python version: 3.11 or 3.12
4. Requirements file: `requirements.txt`

## Run locally

```bash
pip install -r requirements.txt
python -m streamlit run app.py
```

## Files

- `app.py` — main Streamlit dashboard
- `requirements.txt` — dependencies
- `runtime.txt` — suggested Python version for cloud
- `.streamlit/config.toml` — theme/config
- `run_app.bat` — Windows launcher

## Warning

This tool is not a registered financial adviser. Signals are rule-based market-screening labels and not guaranteed buy/sell advice.
