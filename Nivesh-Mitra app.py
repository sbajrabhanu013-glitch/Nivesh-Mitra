import json
import base64
import re
import sqlite3
import pickle
from dataclasses import dataclass
from datetime import datetime, date
from io import BytesIO
from copy import copy, deepcopy
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import pandas as pd
import streamlit as st
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font


# =========================================================
# IMS RECON PRO — FULL WORKING VERSION
# Premium UI + Login + Purchase Register + IMS JSON + Reco + Action + Final GST JSON
# Copyright @BAJRABHANU
# =========================================================

APP_TITLE = "IMS Recon Pro"
APP_TAGLINE = "Intelligent GST IMS Reconciliation & Action Management Platform"
COPYRIGHT_OWNER = "@BAJRABHANU"
APP_DB = "ims_recon_pro.db"
ENGINE_VERSION = "2026.05.09-V10.5-IMS-CN-ITC-FIX"

IMS_SHEETS = ["B2B", "B2BA", "B2B-DN", "B2B-DNA", "B2B-CN", "B2B-CNA"]
ACTION_VALUES = ["No Action", "Accepted", "Rejected", "Pending", "Review"]
USER_MASTER = {
    "Admin": {"password": "Admin", "role": "Admin", "name": "Admin"},
    "User_1": {"password": "User1", "role": "User-1", "name": "User-1"},
    "User_2": {"password": "User2", "role": "User-2", "name": " "},
}

MONEY_COLS = ["invoice_value", "taxable_value", "igst", "cgst", "sgst", "cess"]
TAX_COLS = ["igst", "cgst", "sgst", "cess"]

COLUMN_ALIASES = {
    "supplier_gstin": [
        "supplier gstin", "gstin of supplier", "gstin", "ctin", "stin", "counterparty gstin",
        "vendor gstin", "party gstin", "gstin/uın of supplier", "gstin/uin of supplier"
    ],
    "supplier_name": [
        "supplier name", "trade/legal name", "trade legal name", "tradenm", "legal name",
        "vendor name", "party name", "name", "supplier legal name"
    ],
    "document_type": [
        "document type", "doc type", "invoice type", "type", "supply type",
        "nature of document", "note type"
    ],
    "document_no": [
        "document number", "document no", "doc no", "invoice number", "invoice no", "inum", "nt_num",
        "invoice", "note number", "note no", "bill number", "voucher number"
    ],
    "document_date": [
        "document date", "doc date", "invoice date", "idt", "nt_dt", "date", "note date", "bill date"
    ],
    "invoice_value": [
        "invoice value", "document value", "val", "total invoice value", "gross value",
        "total value", "invoice value(inr)", "invoice value(rs)", "total document value"
    ],
    "taxable_value": [
        "taxable value", "txval", "taxable amount", "taxable value(inr)", "taxable value(rs)",
        "assessable value", "net value", "taxable val"
    ],
    "igst": ["igst", "iamt", "integrated tax", "integrated tax amount", "igst amount"],
    "cgst": ["cgst", "camt", "central tax", "central tax amount", "cgst amount"],
    "sgst": ["sgst", "samt", "state tax", "state/ut tax", "utgst", "sgst amount"],
    "cess": ["cess", "cess amount"],
    "itc_available": ["itc available", "itc availability", "eligible itc", "itc eligibility", "eligible"],
    "ims_status": ["status", "ims status", "recipient status", "recipient action", "action"],
    "remarks": ["remarks", "remark", "reason", "comments", "comment"],
    "pos": ["place of supply", "pos", "state", "supply state"],
    "return_period": ["return period", "rtnprd", "tax period", "period", "month"],
}


# =========================================================
# STREAMLIT PAGE
# =========================================================

st.set_page_config(
    page_title=f"{APP_TITLE} | {COPYRIGHT_OWNER}",
    page_icon="Welcome IMS Reco.",
    layout="wide",
    initial_sidebar_state="expanded",
    menu_items=None,
)


# =========================================================
# DATABASE
# =========================================================

def get_conn():
    return sqlite3.connect(APP_DB, check_same_thread=False)


def init_db():
    with get_conn() as conn:
        conn.execute("""
            CREATE TABLE IF NOT EXISTS user_data (
                username TEXT NOT NULL,
                key TEXT NOT NULL,
                value BLOB NOT NULL,
                updated_at TEXT NOT NULL,
                PRIMARY KEY (username, key)
            )
        """)
        conn.execute("""
            CREATE TABLE IF NOT EXISTS audit_log (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                username TEXT,
                event_time TEXT,
                event_type TEXT,
                detail TEXT
            )
        """)


def db_save(username: str, key: str, value):
    blob = pickle.dumps(value)
    with get_conn() as conn:
        conn.execute(
            "INSERT OR REPLACE INTO user_data (username, key, value, updated_at) VALUES (?, ?, ?, ?)",
            (username, key, blob, datetime.now().isoformat(timespec="seconds")),
        )


def db_load(username: str, key: str, default=None):
    try:
        with get_conn() as conn:
            row = conn.execute("SELECT value FROM user_data WHERE username=? AND key=?", (username, key)).fetchone()
        if not row:
            return default
        return pickle.loads(row[0])
    except Exception:
        return default


def db_delete_user(username: str):
    with get_conn() as conn:
        conn.execute("DELETE FROM user_data WHERE username=?", (username,))
        conn.execute("DELETE FROM audit_log WHERE username=?", (username,))


def log_event(event_type: str, detail: str):
    username = st.session_state.get("username", "")
    try:
        with get_conn() as conn:
            conn.execute(
                "INSERT INTO audit_log (username, event_time, event_type, detail) VALUES (?, ?, ?, ?)",
                (username, datetime.now().isoformat(timespec="seconds"), event_type, detail),
            )
    except Exception:
        pass


def load_audit(username: str = "") -> pd.DataFrame:
    try:
        with get_conn() as conn:
            if username:
                return pd.read_sql(
                    "SELECT event_time, event_type, detail FROM audit_log WHERE username=? ORDER BY id DESC",
                    conn,
                    params=(username,),
                )
            return pd.read_sql(
                "SELECT username, event_time, event_type, detail FROM audit_log ORDER BY id DESC",
                conn,
            )
    except Exception:
        return pd.DataFrame()


# =========================================================
# SESSION
# =========================================================

def init_state():
    defaults = {
        "logged_in": False,
        "username": "",
        "role": "",
        "display_name": "",
        "page": "Dashboard",
        "client_name": "",
        "client_gstin": "",
        "return_period": datetime.today().strftime("%b-%Y"),
        "purchase_df": pd.DataFrame(),
        "ims_df": pd.DataFrame(),
        "ims_source": "",
        "ims_json_records": [],
        "ims_template_bytes": b"",
        "ims_auto_xlsm_bytes": b"",
        "ims_json_data": {},
        "ims_json_bytes": b"",
        "final_action_xlsm_bytes": b"",
        "final_json_bytes": b"",
        "final_json_summary": pd.DataFrame(),
        "recon_df": pd.DataFrame(),
        "action_df": pd.DataFrame(),
        "amount_tolerance": 5.0,
        "date_tolerance": 2,
        "include_amendments": True,
        "use_fuzzy": False,
    }
    for k, v in defaults.items():
        st.session_state.setdefault(k, v)


def load_user_state():
    username = st.session_state.get("username")
    if not username:
        return
    for key in [
        "client_name", "client_gstin", "return_period",
        "purchase_df", "ims_df", "ims_source", "ims_json_records", "ims_template_bytes", "ims_auto_xlsm_bytes", "ims_json_data", "ims_json_bytes", "final_action_xlsm_bytes", "final_json_bytes", "final_json_summary", "recon_df", "action_df"
    ]:
        st.session_state[key] = db_load(username, key, st.session_state.get(key))


def save_user_state(keys: Optional[List[str]] = None):
    username = st.session_state.get("username")
    if not username:
        return
    keys = keys or [
        "client_name", "client_gstin", "return_period",
        "purchase_df", "ims_df", "ims_source", "ims_json_records", "ims_template_bytes", "ims_auto_xlsm_bytes", "ims_json_data", "ims_json_bytes", "final_action_xlsm_bytes", "final_json_bytes", "final_json_summary", "recon_df", "action_df"
    ]
    for key in keys:
        db_save(username, key, st.session_state.get(key))


# =========================================================
# STYLING
# =========================================================

def inject_css():
    st.markdown("""
    <style>
        :root {
            --navy:#061a3e; --navy2:#0b2d66; --blue:#2563eb; --cyan:#38bdf8;
            --saffron:#ff9933; --green:#138808; --gold:#f3b34d; --red:#dc463f;
            --bg1:#d8e7f7; --bg2:#b9cde8; --card:#ffffff; --border:#c8d8ec;
            --text:#102244; --muted:#5c708f;
        }

        .stApp {
            background:
                radial-gradient(circle at 12% 8%, rgba(255,153,51,0.20), transparent 26%),
                radial-gradient(circle at 88% 16%, rgba(56,189,248,0.18), transparent 28%),
                radial-gradient(circle at 88% 82%, rgba(19,136,8,0.13), transparent 30%),
                linear-gradient(135deg, #dbeafe 0%, #c7d9ef 48%, #e4eef9 100%);
            color: var(--text);
        }

        header[data-testid="stHeader"] {
            background: rgba(224, 236, 249, 0.88) !important;
            backdrop-filter: blur(12px);
            border-bottom: 1px solid rgba(11,45,102,0.10);
        }

        div[data-testid="stToolbar"] { visibility:hidden; height:0; }
        #MainMenu {visibility: hidden;}
        footer {visibility: hidden;}

        .block-container {
            padding-top: 1.05rem;
            padding-bottom: 2rem;
            max-width: 1560px;
        }

        .gst-shell {
            border-radius: 28px;
            overflow: hidden;
            border: 1px solid rgba(198,216,236,0.95);
            background: #ffffff;
            box-shadow: 0 18px 45px rgba(7, 26, 61, 0.16);
            margin-bottom: 18px;
        }

        .gst-top-strip {
            min-height: 34px;
            background: linear-gradient(90deg, #061a3e 0%, #0b2d66 60%, #061a3e 100%);
            display:flex; align-items:center; justify-content:flex-end;
            gap:18px; padding: 0 22px;
            color:#dceaff; font-size: 12px; letter-spacing:.02em;
        }

        .gst-masthead {
            position:relative;
            background:
                linear-gradient(120deg, rgba(255,153,51,0.14), transparent 28%),
                linear-gradient(90deg, #08214f 0%, #0b3677 52%, #123d82 100%);
            min-height: 150px;
            padding: 28px 32px;
            display:flex; align-items:center; justify-content:space-between;
            gap:24px; overflow:hidden;
        }

        .gst-masthead::after {
            content:""; position:absolute; right:-130px; top:-170px;
            width:420px; height:420px; border-radius:50%;
            background: radial-gradient(circle, rgba(255,255,255,0.16), transparent 62%);
        }

        .gst-brand { position:relative; z-index:2; display:flex; align-items:center; gap:20px; min-width:0; }

        .gst-emblem {
            width:78px; height:78px; border-radius:22px;
            display:flex; align-items:center; justify-content:center;
            color:#ffffff; font-size:42px;
            background: linear-gradient(135deg, rgba(255,153,51,0.95), rgba(19,136,8,0.88));
            border:1px solid rgba(255,255,255,0.22);
            box-shadow:0 14px 30px rgba(0,0,0,0.20);
        }

        .gst-title {
            font-size: 36px; font-weight: 900; letter-spacing: -.02em;
            line-height: 1.05; color:#ffffff;
        }

        .gst-subtitle {
            margin-top: 8px; font-size: 18px; line-height: 1.3;
            color:#edf5ff; font-weight: 500;
        }

        .header-note { margin-top: 7px; font-size: 13px; color:#cfe0fb; }

        .gst-action-wrap {
            position:relative; z-index:2; display:flex; align-items:center;
            justify-content:flex-end; min-width: 220px;
        }

        .gst-floating-flag {
            width: 150px; height: 92px; border-radius: 14px;
            position: relative; overflow: hidden;
            background: linear-gradient(to bottom, #ff9933 0 33.33%, #ffffff 33.33% 66.66%, #138808 66.66% 100%);
            border:1px solid rgba(255,255,255,0.45);
            box-shadow: 0 14px 28px rgba(0,0,0,0.24);
            animation: gstFlagFloat 3.8s ease-in-out infinite;
        }

        .gst-floating-flag::after {
            content:"☸"; position:absolute; left:50%; top:50%;
            transform:translate(-50%, -50%);
            color:#0a3d91; font-size:28px; font-weight:800; z-index:2;
        }

        @keyframes gstFlagFloat {
            0%,100% { transform: translateY(0px); }
            50% { transform: translateY(-5px); }
        }

        .gst-meta-row {
            background: linear-gradient(90deg, #eef5ff 0%, #e7f0fb 100%);
            border-top: 1px solid #c9d9ee;
            padding: 16px 18px;
            display:grid;
            grid-template-columns: repeat(4, minmax(0, 1fr));
            gap:14px;
        }

        .gst-meta-card {
            background:#ffffff; border:1px solid #ceddf0; border-radius:18px;
            min-height:72px; padding:14px 16px;
            display:flex; align-items:center; gap:12px;
            box-shadow:0 8px 18px rgba(7,26,61,0.06);
        }

        .gst-meta-icon {
            width:44px; height:44px; border-radius:14px;
            background:linear-gradient(135deg,#eef6ff,#dbeafe);
            display:flex; align-items:center; justify-content:center;
            font-size:21px; flex-shrink:0;
        }

        .gst-meta-label {
            font-size:12px; color:var(--muted); font-weight:700;
            text-transform:uppercase; letter-spacing:.04em;
        }

        .gst-meta-value {
            font-size:15px; color:var(--text); font-weight:800;
            line-height:1.25; margin-top:3px;
        }

        .main-shell, .panel, .metric-card, .small-card {
            background: rgba(255,255,255,0.96);
            border: 1px solid rgba(200,216,236,0.98);
            box-shadow: 0 14px 35px rgba(7,26,61,0.12);
        }

        .main-shell { border-radius: 26px; overflow: hidden; margin-bottom: 18px; }
        .content-pad { padding: 30px; position:relative; }
        .panel, .metric-card, .small-card { border-radius: 24px; padding: 20px 22px; height: 100%; }
        .panel { background: linear-gradient(180deg,#ffffff,#f8fbff); }

        .metric-card { position:relative; overflow:hidden; }
        .metric-card::after {
            content:""; position:absolute; right:-36px; bottom:-48px;
            width:115px; height:115px; border-radius:50%;
            background: rgba(37,99,235,0.07);
        }

        .metric-top {display:flex;align-items:center;gap:15px;position:relative;z-index:2;}
        .metric-icon {
            width:58px;height:58px;border-radius:18px;
            display:flex;align-items:center;justify-content:center;
            font-size:25px;
            box-shadow: inset 0 0 0 1px rgba(255,255,255,0.55);
        }
        .metric-label {
            font-size:13px;color:#60748f;font-weight:800;
            text-transform:uppercase;letter-spacing:.04em;
        }
        .metric-value {font-size:33px;font-weight:900;color:#112244;line-height:1.15;}
        .metric-delta {font-size:13px;color:#12a150;margin-top:5px;font-weight:700;}
        .metric-delta.red {color:#e1563a;}

        .panel-title {font-size:20px;font-weight:900;color:#112244;}
        .section-title {font-size:28px;font-weight:950;color:#102244;margin:10px 0 4px 0;}
        .section-sub {font-size:15px;color:#60748f;margin-bottom:18px;}

        .headline {font-size:21px;color:#e98012;font-weight:900;}
        .main-title {font-size:34px;font-weight:950;color:#112244;line-height:1.18;margin-top:8px;}
        .subcopy {font-size:17px;color:#52637d;margin-top:12px;line-height:1.5;}

        .cta-dark,.cta-light {
            display:inline-block; padding:13px 24px; border-radius:16px;
            font-weight:900; text-decoration:none; font-size:15px;
            margin-right:10px; margin-top:20px;
        }
        .cta-dark {background:linear-gradient(135deg,#0b2d66,#2563eb);color:white;box-shadow:0 12px 22px rgba(11,42,93,.22);}
        .cta-light {background:white;color:#0b2a5d;border:1px solid #d0def0;}

        .feature-card {
            background:linear-gradient(180deg,#fffaf1,#fff7eb);
            border:1px solid #f0dfc0; border-radius:18px;
            padding:15px 17px; margin-bottom:13px;
            box-shadow:0 8px 18px rgba(7,26,61,0.06);
        }
        .feature-card.blue {background:linear-gradient(180deg,#f5f9ff,#eef5ff);border-color:#d6e4ff;}
        .feature-card.green {background:linear-gradient(180deg,#f5fbf3,#eff9ec);border-color:#d8ead0;}
        .feature-title {font-weight:900;color:#23385d;font-size:16px;}
        .feature-desc {font-size:13px;color:#5f6f89;line-height:1.4;margin-top:4px;}
        .shield-center {
            width:140px;height:140px;border-radius:50%;margin:0 auto 18px auto;
            background:radial-gradient(circle at 30% 30%,#fffef4,#f8f0d2 55%,#ead39f 100%);
            display:flex;align-items:center;justify-content:center;font-size:62px;
            box-shadow:inset 0 0 0 10px rgba(255,255,255,.65),0 12px 28px rgba(194,165,97,.18);
        }

        .watermark {
            position:absolute;left:50%;top:47%;transform:translate(-50%,-50%);
            font-size:230px;color:rgba(14,41,90,.035);pointer-events:none;
        }

        .stTextInput input, .stNumberInput input, .stDateInput input, .stTextArea textarea {
            border-radius: 14px !important;
            border: 1px solid #cdddf0 !important;
            background: #ffffff !important;
        }

        .stSelectbox div[data-baseweb="select"] > div {
            border-radius: 14px !important;
            border-color: #cdddf0 !important;
        }

        .stTabs [data-baseweb="tab-list"] {
            gap: 8px; background:#edf4ff; padding:8px;
            border-radius:18px; border:1px solid #cfdded;
        }
        .stTabs [data-baseweb="tab"] {
            border-radius:14px; padding:10px 16px; font-weight:800;
        }
        .stTabs [aria-selected="true"] {
            background:#ffffff !important; color:#0b2d66 !important;
            box-shadow:0 8px 16px rgba(7,26,61,0.08);
        }

        div[data-testid="stDataFrame"] {
            border-radius: 18px; overflow:hidden;
            border:1px solid #cfdff0;
            box-shadow:0 8px 18px rgba(7,26,61,0.06);
        }

        div[data-testid="stHorizontalBlock"] .stButton > button {
            border-radius: 16px !important;
            min-height: 48px !important;
            border: 1px solid #c9d9ee !important;
            background: linear-gradient(180deg, #ffffff, #f2f7ff) !important;
            color: #0d2d63 !important;
            font-weight: 800 !important;
            box-shadow: 0 6px 14px rgba(7,26,61,0.06) !important;
        }

        div[data-testid="stHorizontalBlock"] .stButton > button:hover {
            border-color: #2563eb !important;
            background: linear-gradient(180deg, #eef6ff, #dbeafe) !important;
            color: #071a3d !important;
            transform: translateY(-1px);
        }

        .footer-bar {
            margin-top: 20px; border-radius: 24px;
            background: linear-gradient(90deg,#061a3e 0%,#082b61 50%,#061a3e 100%);
            color:white; padding: 18px 22px;
            box-shadow: 0 14px 32px rgba(7,26,61,0.16);
        }
        .foot-item {display:flex;align-items:center;gap:10px;justify-content:center;}
        .foot-main {font-weight:900;}
        .foot-sub {font-size:13px;color:#d4e0ff;}

        .login-bg {
            min-height:calc(100vh - 35px);
            display:flex; align-items:center; justify-content:center;
            background:
                radial-gradient(circle at 18% 12%, rgba(255,153,51,.20), transparent 30%),
                radial-gradient(circle at 85% 80%, rgba(19,136,8,.16), transparent 32%),
                linear-gradient(135deg,#071a3d,#0d2d63);
            border-radius:30px; position:relative; overflow:hidden;
        }
        .login-bg::after {
            content:""; position:absolute;
            width:720px;height:720px;border-radius:50%;
            right:-220px;top:-260px;
            background:radial-gradient(circle, rgba(255,255,255,0.13), transparent 62%);
        }
        .login-card {
            width: 470px; background:rgba(255,255,255,.96);
            border:1px solid rgba(255,255,255,.62);
            border-radius:32px; padding:38px;
            box-shadow:0 35px 90px rgba(0,0,0,.30);
            position:relative; z-index:2;
        }
        .login-title {font-size:36px;font-weight:950;color:#071b4a;text-align:center;}
        .login-sub {font-size:15px;color:#566982;text-align:center;margin-bottom:24px;line-height:1.5;}

        .copyright-float {
            position:fixed;right:18px;bottom:14px;
            color:rgba(7,26,61,.20);
            font-weight:950;letter-spacing:.08em;z-index:99;
        }

        @media (max-width: 1100px) {
            .gst-masthead {flex-direction:column;align-items:flex-start;}
            .gst-meta-row {grid-template-columns: repeat(2, minmax(0, 1fr));}
            .gst-action-wrap {justify-content:flex-start;}
        }
        @media (max-width: 760px) {
            .gst-title {font-size:28px;}
            .gst-subtitle {font-size:15px;}
            .gst-meta-row {grid-template-columns: 1fr;}
            .block-container {padding-left:0.75rem;padding-right:0.75rem;}
        }
    
        /* ================= V9 SALEABLE UI EDITION ================= */
        .v9-workflow {
            background: rgba(255,255,255,0.96);
            border: 1px solid #c9d9ee;
            border-radius: 24px;
            padding: 16px 18px;
            margin: 0 0 18px 0;
            box-shadow: 0 14px 34px rgba(7,26,61,0.10);
        }
        .v9-workflow-title {
            font-size: 17px;
            font-weight: 950;
            color: #0b2d66;
            margin-bottom: 12px;
            display:flex;
            align-items:center;
            gap:8px;
        }
        .v9-step-grid {
            display:grid;
            grid-template-columns: repeat(6, minmax(0, 1fr));
            gap: 10px;
        }
        .v9-step {
            position:relative;
            min-height: 88px;
            border-radius: 18px;
            border: 1px solid #d1e0f2;
            background: linear-gradient(180deg,#ffffff,#f4f8ff);
            padding: 12px;
            overflow:hidden;
        }
        .v9-step.done {
            border-color: rgba(19,136,8,0.28);
            background: linear-gradient(180deg,#ffffff,#effaf1);
        }
        .v9-step.active {
            border-color: rgba(255,153,51,0.55);
            background: linear-gradient(180deg,#ffffff,#fff5e8);
            box-shadow: inset 0 0 0 1px rgba(255,153,51,0.16);
        }
        .v9-step.pending {
            border-color: rgba(120,140,170,0.25);
        }
        .v9-step-num {
            width:30px;height:30px;border-radius:10px;
            display:flex;align-items:center;justify-content:center;
            font-weight:900;font-size:13px;
            background:#eaf2ff;color:#0b2d66;margin-bottom:8px;
        }
        .v9-step.done .v9-step-num { background:#e9f9ed;color:#138808; }
        .v9-step.active .v9-step-num { background:#fff0da;color:#c76f00; }
        .v9-step-label {font-size:13px;font-weight:900;color:#102244;line-height:1.25;}
        .v9-step-status {font-size:11px;font-weight:800;color:#5c708f;margin-top:5px;}
        .v9-step.done .v9-step-status {color:#138808;}
        .v9-step.active .v9-step-status {color:#c76f00;}

        .v9-kpi-strip {
            display:grid;
            grid-template-columns: repeat(4, minmax(0,1fr));
            gap:14px;
            margin: 14px 0 18px 0;
        }
        .v9-kpi {
            background:linear-gradient(135deg,#ffffff,#f7fbff);
            border:1px solid #cdddf0;
            border-radius:22px;
            padding:18px;
            box-shadow:0 12px 26px rgba(7,26,61,0.09);
            min-height:112px;
            position:relative;
            overflow:hidden;
        }
        .v9-kpi::after {
            content:"";
            position:absolute;
            width:105px;height:105px;right:-38px;bottom:-48px;
            border-radius:50%;background:rgba(37,99,235,0.08);
        }
        .v9-kpi-label {font-size:12px;text-transform:uppercase;letter-spacing:.04em;font-weight:900;color:#60748f;}
        .v9-kpi-value {font-size:32px;font-weight:950;color:#102244;margin-top:8px;}
        .v9-kpi-note {font-size:12px;font-weight:800;color:#138808;margin-top:6px;}

        .v9-module-grid {
            display:grid;
            grid-template-columns: repeat(3, minmax(0,1fr));
            gap:16px;
            margin: 14px 0 20px 0;
        }
        .v9-module-card {
            background:#ffffff;
            border:1px solid #cdddf0;
            border-radius:24px;
            padding:20px;
            min-height:185px;
            box-shadow:0 14px 34px rgba(7,26,61,0.10);
            position:relative;
            overflow:hidden;
        }
        .v9-module-card::before {
            content:"";
            position:absolute;left:0;top:0;width:100%;height:6px;
            background:linear-gradient(90deg,#ff9933,#2563eb,#138808);
        }
        .v9-module-icon {
            width:52px;height:52px;border-radius:16px;
            display:flex;align-items:center;justify-content:center;
            background:linear-gradient(135deg,#eef6ff,#dbeafe);
            font-size:25px;margin-bottom:12px;
        }
        .v9-module-title {font-size:18px;font-weight:950;color:#102244;margin-bottom:7px;}
        .v9-module-desc {font-size:13px;color:#60748f;line-height:1.45;}
        .v9-module-badge {
            display:inline-block;margin-top:12px;padding:6px 11px;border-radius:999px;
            font-size:11px;font-weight:900;background:#eef6ff;color:#0b2d66;border:1px solid #d1e0f2;
        }

        .v9-readiness {
            background:linear-gradient(135deg,#071a3d,#0b3677);
            color:#ffffff;border-radius:26px;padding:24px;
            border:1px solid rgba(255,255,255,0.18);
            box-shadow:0 18px 40px rgba(7,26,61,0.18);
            margin: 14px 0 18px 0;
        }
        .v9-readiness-title {font-size:22px;font-weight:950;margin-bottom:12px;}
        .v9-check-grid {
            display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:12px;
        }
        .v9-check {
            background:rgba(255,255,255,0.10);
            border:1px solid rgba(255,255,255,0.14);
            border-radius:18px;padding:14px;
            min-height:82px;
        }
        .v9-check-icon {font-size:22px;margin-bottom:6px;}
        .v9-check-label {font-size:13px;font-weight:900;color:#eef5ff;line-height:1.3;}

        .v9-action-badge {
            display:inline-block;padding:5px 10px;border-radius:999px;
            font-size:12px;font-weight:900;border:1px solid transparent;
        }
        .v9-action-accepted {background:#e9f9ed;color:#138808;border-color:#bfe9c9;}
        .v9-action-pending {background:#fff3df;color:#b96b00;border-color:#ffd9a8;}
        .v9-action-rejected {background:#fff0ed;color:#d33a2f;border-color:#ffc8c0;}
        .v9-action-review {background:#f2ecff;color:#6d3bd1;border-color:#d8c9ff;}
        .v9-action-no {background:#f1f5f9;color:#475569;border-color:#d7e0ea;}

        .v9-help-box {
            background:#fffdf6;border:1px solid #f3d9a7;border-radius:20px;
            padding:16px 18px;margin:12px 0 18px 0;
            color:#62420d;box-shadow:0 10px 22px rgba(7,26,61,0.06);
        }
        .v9-help-title {font-weight:950;font-size:16px;margin-bottom:6px;color:#7a4b00;}
        .v9-help-text {font-size:13px;line-height:1.5;}

        .v9-report-grid {
            display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:16px;
            margin:16px 0;
        }
        .v9-report-card {
            background:#ffffff;border:1px solid #cdddf0;border-radius:22px;
            padding:18px;box-shadow:0 12px 26px rgba(7,26,61,0.09);
        }
        .v9-report-title {font-size:16px;font-weight:950;color:#102244;}
        .v9-report-desc {font-size:13px;color:#60748f;margin-top:7px;line-height:1.45;}

        @media (max-width: 1150px) {
            .v9-step-grid {grid-template-columns: repeat(3, minmax(0,1fr));}
            .v9-module-grid, .v9-report-grid {grid-template-columns: repeat(2,minmax(0,1fr));}
            .v9-kpi-strip, .v9-check-grid {grid-template-columns: repeat(2,minmax(0,1fr));}
        }
        @media (max-width: 760px) {
            .v9-step-grid, .v9-module-grid, .v9-report-grid, .v9-kpi-strip, .v9-check-grid {
                grid-template-columns: 1fr;
            }
        }

    
        /* ================= V10 ADVANCED SALEABLE UI ================= */
        .v10-command-center {
            background: linear-gradient(135deg,#071a3d,#0b3677);
            color:#ffffff;
            border-radius:28px;
            padding:24px;
            box-shadow:0 20px 48px rgba(7,26,61,0.22);
            border:1px solid rgba(255,255,255,0.16);
            margin: 14px 0 20px 0;
            position:relative;
            overflow:hidden;
        }
        .v10-command-center::after {
            content:"";
            position:absolute;
            right:-120px;
            top:-160px;
            width:360px;
            height:360px;
            border-radius:50%;
            background:radial-gradient(circle,rgba(255,255,255,0.15),transparent 62%);
        }
        .v10-command-title {
            font-size:26px;
            font-weight:950;
            margin-bottom:8px;
            position:relative;
            z-index:2;
        }
        .v10-command-sub {
            font-size:14px;
            color:#d9e8ff;
            margin-bottom:18px;
            position:relative;
            z-index:2;
        }
        .v10-action-grid {
            display:grid;
            grid-template-columns:repeat(4,minmax(0,1fr));
            gap:14px;
            position:relative;
            z-index:2;
        }
        .v10-action-card {
            background:rgba(255,255,255,0.10);
            border:1px solid rgba(255,255,255,0.16);
            border-radius:22px;
            padding:18px;
            min-height:126px;
            backdrop-filter: blur(4px);
        }
        .v10-action-icon {font-size:30px;margin-bottom:10px;}
        .v10-action-title {font-size:16px;font-weight:950;color:#fff;}
        .v10-action-desc {font-size:12px;color:#d7e7ff;line-height:1.45;margin-top:6px;}

        .v10-quality-grid {
            display:grid;
            grid-template-columns:repeat(2,minmax(0,1fr));
            gap:18px;
            margin:16px 0 18px 0;
        }
        .v10-quality-card {
            background:#ffffff;
            border:1px solid #cdddf0;
            border-radius:26px;
            padding:22px;
            box-shadow:0 14px 34px rgba(7,26,61,0.10);
        }
        .v10-quality-head {
            display:flex;
            justify-content:space-between;
            align-items:flex-start;
            gap:16px;
            margin-bottom:16px;
        }
        .v10-quality-title {
            font-size:20px;
            font-weight:950;
            color:#102244;
        }
        .v10-quality-score {
            min-width:86px;
            height:86px;
            border-radius:50%;
            display:flex;
            align-items:center;
            justify-content:center;
            font-size:24px;
            font-weight:950;
            color:#ffffff;
            background:linear-gradient(135deg,#138808,#28b463);
            box-shadow:0 12px 26px rgba(19,136,8,0.22);
        }
        .v10-quality-score.warn {background:linear-gradient(135deg,#ff9933,#f3b34d);}
        .v10-quality-score.bad {background:linear-gradient(135deg,#dc463f,#ef675b);}
        .v10-mini-grid {
            display:grid;
            grid-template-columns:repeat(4,minmax(0,1fr));
            gap:10px;
        }
        .v10-mini-stat {
            background:#f7fbff;
            border:1px solid #d6e4f5;
            border-radius:16px;
            padding:12px;
        }
        .v10-mini-label {font-size:11px;color:#60748f;font-weight:900;text-transform:uppercase;}
        .v10-mini-value {font-size:18px;color:#102244;font-weight:950;margin-top:5px;}

        .v10-control-room {
            display:grid;
            grid-template-columns:1.35fr .9fr;
            gap:18px;
            margin:16px 0 20px 0;
        }
        .v10-control-main,
        .v10-control-side {
            background:#ffffff;
            border:1px solid #cdddf0;
            border-radius:26px;
            padding:22px;
            box-shadow:0 14px 34px rgba(7,26,61,0.10);
        }
        .v10-control-title {
            font-size:22px;
            font-weight:950;
            color:#102244;
            margin-bottom:12px;
        }
        .v10-badge-row {
            display:flex;
            gap:10px;
            flex-wrap:wrap;
            margin:10px 0 14px 0;
        }
        .v10-filter-badge {
            display:inline-flex;
            align-items:center;
            gap:6px;
            padding:8px 12px;
            border-radius:999px;
            background:#eef6ff;
            border:1px solid #d1e0f2;
            color:#0b2d66;
            font-size:12px;
            font-weight:900;
        }
        .v10-filter-badge.green {background:#e9f9ed;color:#138808;border-color:#bfe9c9;}
        .v10-filter-badge.orange {background:#fff3df;color:#b96b00;border-color:#ffd9a8;}
        .v10-filter-badge.red {background:#fff0ed;color:#d33a2f;border-color:#ffc8c0;}
        .v10-filter-badge.purple {background:#f2ecff;color:#6d3bd1;border-color:#d8c9ff;}

        .v10-empty-state {
            background:linear-gradient(135deg,#ffffff,#f7fbff);
            border:1px dashed #a8bdd8;
            border-radius:26px;
            padding:28px;
            text-align:center;
            margin:16px 0;
            box-shadow:0 12px 28px rgba(7,26,61,0.07);
        }
        .v10-empty-icon {font-size:44px;margin-bottom:10px;}
        .v10-empty-title {font-size:22px;font-weight:950;color:#102244;}
        .v10-empty-text {font-size:14px;color:#60748f;margin-top:8px;}

        .v10-management-summary {
            background:linear-gradient(135deg,#ffffff,#f6faff);
            border:1px solid #cdddf0;
            border-radius:26px;
            padding:22px;
            box-shadow:0 14px 34px rgba(7,26,61,0.10);
            margin:16px 0 18px 0;
        }
        .v10-management-title {
            font-size:22px;
            font-weight:950;
            color:#102244;
            margin-bottom:10px;
        }
        .v10-management-text {
            font-size:15px;
            line-height:1.6;
            color:#334866;
        }

        .v10-json-review {
            background:#ffffff;
            border:1px solid #cdddf0;
            border-radius:28px;
            padding:24px;
            box-shadow:0 16px 38px rgba(7,26,61,0.12);
            margin:16px 0 18px 0;
        }
        .v10-json-title {
            font-size:24px;
            font-weight:950;
            color:#102244;
            margin-bottom:14px;
        }
        .v10-json-checks {
            display:grid;
            grid-template-columns:repeat(5,minmax(0,1fr));
            gap:12px;
        }
        .v10-json-check {
            background:#f7fbff;
            border:1px solid #d6e4f5;
            border-radius:18px;
            padding:14px;
            min-height:96px;
            text-align:center;
        }
        .v10-json-check-icon {font-size:24px;margin-bottom:7px;}
        .v10-json-check-label {font-size:12px;font-weight:900;color:#102244;line-height:1.35;}

        .v10-tooltip-grid {
            display:grid;
            grid-template-columns:repeat(4,minmax(0,1fr));
            gap:12px;
            margin:14px 0;
        }
        .v10-tooltip {
            background:#fffdf6;
            border:1px solid #f3d9a7;
            border-radius:18px;
            padding:14px;
        }
        .v10-tooltip-title {
            font-size:13px;
            font-weight:950;
            color:#7a4b00;
        }
        .v10-tooltip-text {
            font-size:12px;
            color:#62420d;
            line-height:1.45;
            margin-top:6px;
        }

        .v10-premium-divider {
            height:1px;
            background:linear-gradient(90deg,transparent,#9eb9dc,transparent);
            margin:20px 0;
        }

        @media (max-width:1150px) {
            .v10-action-grid {grid-template-columns:repeat(2,minmax(0,1fr));}
            .v10-quality-grid,.v10-control-room {grid-template-columns:1fr;}
            .v10-json-checks,.v10-tooltip-grid {grid-template-columns:repeat(2,minmax(0,1fr));}
        }
        @media (max-width:760px) {
            .v10-action-grid,.v10-mini-grid,.v10-json-checks,.v10-tooltip-grid {grid-template-columns:1fr;}
        }

    
        /* INALSA login logo + native quality card polish */
        .inalsa-login-logo {
            width: 100%;
            max-width: 335px;
            height: auto;
            display: block;
            margin: 0 auto 18px auto;
            object-fit: contain;
            background: #ffffff;
            border-radius: 18px;
            padding: 10px 14px;
            box-shadow: 0 10px 28px rgba(7,26,61,0.10);
            border: 1px solid #e4edf8;
        }
        .inalsa-login-welcome {
            text-align:center;
            font-size:30px;
            font-weight:950;
            color:#071b4a;
            margin-top:6px;
        }
        .inalsa-login-sub {
            text-align:center;
            font-size:14px;
            line-height:1.5;
            color:#566982;
            margin:8px 0 22px 0;
        }
        .inalsa-login-copy {
            text-align:center;
            font-size:12px;
            color:#6b7d96;
            margin-top:16px;
            border-top:1px solid #e3edf8;
            padding-top:13px;
            font-weight:700;
        }
        div[data-testid="stMetric"] {
            background: linear-gradient(180deg, #ffffff, #f7fbff);
            border: 1px solid #d6e4f5;
            border-radius: 16px;
            padding: 12px 14px;
            box-shadow: 0 6px 14px rgba(7,26,61,0.06);
        }
        div[data-testid="stMetricLabel"] {
            font-weight: 800;
            color: #60748f;
        }

    </style>
    """, unsafe_allow_html=True)



# =========================================================
# V10.2 QUALITY CARD HELPERS
# =========================================================

def v10_render_html(html: str):
    """Render premium HTML cards safely in Streamlit."""
    st.markdown(str(html), unsafe_allow_html=True)


def v10_empty_upload_notice(title: str, text: str):
    st.markdown(f"""
    <div class='v10-empty-state'>
        <div class='v10-empty-icon'>📤</div>
        <div class='v10-empty-title'>{title}</div>
        <div class='v10-empty-text'>{text}</div>
    </div>
    """, unsafe_allow_html=True)


def v10_df_loaded(df) -> bool:
    return isinstance(df, pd.DataFrame) and not df.empty


def v10_quality_cards_section():
    """Show Purchase/IMS premium quality cards only after data is actually uploaded."""
    p = st.session_state.get("purchase_df", pd.DataFrame())
    ims = st.session_state.get("ims_df", pd.DataFrame())

    if not v10_df_loaded(p) and not v10_df_loaded(ims):
        v10_empty_upload_notice(
            "Upload files to view quality summary",
            "Purchase Register and IMS JSON quality cards will appear here after you upload and process the files."
        )
        return

    st.markdown("### Upload Quality Summary")
    c1, c2 = st.columns(2)

    with c1:
        if v10_df_loaded(p):
            v10_render_html(v10_quality_card(p, "Purchase Register Quality"))
        else:
            v10_empty_upload_notice(
                "Purchase Register not processed",
                "Upload and process Purchase Register to view Records, Taxable Value, IGST, CGST, SGST, CESS and Total Tax."
            )

    with c2:
        if v10_df_loaded(ims):
            v10_render_html(v10_quality_card(ims, "IMS JSON Quality"))
        else:
            v10_empty_upload_notice(
                "IMS JSON not processed",
                "Upload and process GST IMS JSON to view Records, Taxable Value, IGST, CGST, SGST, CESS and Total Tax."
            )

# =========================================================
# DATA PROCESSING
# =========================================================

def clean_header(value) -> str:
    text = str(value or "").strip().lower()
    text = text.replace("\n", " ").replace("\r", " ")
    text = text.replace("₹", "rs")
    text = re.sub(r"\s+", " ", text)
    text = re.sub(r"[^a-z0-9%/() -]", "", text)
    return text.strip()


def find_col(df: pd.DataFrame, logical: str) -> Optional[str]:
    aliases = COLUMN_ALIASES.get(logical, [logical])
    norm = {clean_header(c): c for c in df.columns}
    for alias in aliases:
        ca = clean_header(alias)
        if ca in norm:
            return norm[ca]
    for alias in aliases:
        ca = clean_header(alias)
        for nc, orig in norm.items():
            if ca and ca in nc:
                return orig
    return None


def normalize_doc_no(x) -> str:
    text = str(x or "").strip().upper()
    text = re.sub(r"\.0$", "", text)
    return re.sub(r"[^A-Z0-9]", "", text)


def normalize_gstin(x) -> str:
    text = str(x or "").strip().upper()
    return re.sub(r"[^A-Z0-9]", "", text)


def validate_gstin(x) -> bool:
    return bool(re.match(r"^[0-9]{2}[A-Z]{5}[0-9]{4}[A-Z][1-9A-Z]Z[0-9A-Z]$", normalize_gstin(x)))


def to_number(s: pd.Series) -> pd.Series:
    if s is None:
        return pd.Series(dtype="float64")
    clean = (
        s.astype(str)
        .str.replace(",", "", regex=False)
        .str.replace("INR", "", case=False, regex=False)
        .str.replace("RS.", "", case=False, regex=False)
        .str.replace("RS", "", case=False, regex=False)
        .str.replace("₹", "", regex=False)
        .str.strip()
        .str.replace(r"^\((.*)\)$", r"-\1", regex=True)
    )
    clean = clean.replace({"": "0", "nan": "0", "None": "0", "NaT": "0"})
    return pd.to_numeric(clean, errors="coerce").fillna(0.0)


def to_date(s: pd.Series) -> pd.Series:
    if s is None:
        return pd.Series(dtype="datetime64[ns]")
    numeric = pd.to_numeric(s, errors="coerce")
    excel_dates = pd.to_datetime(numeric, unit="D", origin="1899-12-30", errors="coerce")
    parsed = pd.to_datetime(s, errors="coerce", dayfirst=True)
    return parsed.fillna(excel_dates)


def sign_for_doc_type(x) -> int:
    text = str(x or "").lower()
    if any(k in text for k in ["credit", "cn", "cdn", "refund"]):
        return -1
    return 1


def detect_header_row(raw: pd.DataFrame) -> int:
    alias_words = [clean_header(a) for v in COLUMN_ALIASES.values() for a in v if len(clean_header(a)) >= 3]
    best_idx = raw.index[0]
    best_score = -1
    for idx, row in raw.head(30).iterrows():
        row_text = " ".join(clean_header(c) for c in row.tolist())
        score = sum(1 for a in alias_words if a and a in row_text)
        if score > best_score:
            best_score = score
            best_idx = idx
    return int(best_idx)


def normalize_sheet(sheet_df: pd.DataFrame) -> pd.DataFrame:
    raw = sheet_df.dropna(how="all").dropna(how="all", axis=1)
    if raw.empty:
        return pd.DataFrame()
    header_idx = detect_header_row(raw)
    header_pos = list(raw.index).index(header_idx)
    headers = []
    seen = {}
    for i, h in enumerate(raw.loc[header_idx].fillna("").astype(str).tolist(), start=1):
        name = str(h or "").strip() or f"Column {i}"
        if name in seen:
            seen[name] += 1
            name = f"{name}_{seen[name]}"
        else:
            seen[name] = 0
        headers.append(name)
    body = raw.iloc[header_pos + 1:].copy()
    body.columns = headers
    body = body.dropna(how="all")
    body = body.loc[:, [c for c in body.columns if str(c).strip()]]
    return body


def read_excel_all_sheets(file, wanted_sheets: Optional[List[str]] = None) -> Dict[str, pd.DataFrame]:
    sheets = pd.read_excel(file, sheet_name=None, dtype=object, header=None, engine="openpyxl")
    out = {}
    for name, df in sheets.items():
        if wanted_sheets and name.strip().upper() not in [s.upper() for s in wanted_sheets]:
            continue
        clean = normalize_sheet(df)
        if not clean.empty:
            out[name.strip()] = clean
    return out


def standardize(df: pd.DataFrame, source_label: str, sheet_name: str = "", default_doc_type: str = "") -> pd.DataFrame:
    if df is None or df.empty:
        return pd.DataFrame()

    schema = {k: find_col(df, k) for k in COLUMN_ALIASES}
    n = len(df)

    def get(logical, default=""):
        col = schema.get(logical)
        if col and col in df.columns:
            return df[col]
        return pd.Series([default] * n, index=df.index)

    # =====================================================
    # IMPORTANT FIX FOR GST IMS JSON CREDIT NOTE / DEBIT NOTE
    # B2B Invoice fields: inum, idt, inv_typ
    # B2B-CN / B2B-DN fields: nt_num, nt_dt, ntty
    # In mixed JSON dataframe, inum/idt may exist but remain blank for CN rows.
    # Therefore, for CN/DN sheets, force priority to nt_num / nt_dt / ntty.
    # =====================================================
    sheet_upper = str(sheet_name or "").upper()
    default_doc_upper = str(default_doc_type or "").upper()

    is_note_section = (
        "CN" in sheet_upper
        or "DN" in sheet_upper
        or "CREDIT" in default_doc_upper
        or "DEBIT" in default_doc_upper
    )

    def get_note_safe_value(invoice_field: str, note_field: str, logical: str, default=""):
        """
        For Credit Note / Debit Note, prefer note fields like nt_num, nt_dt, ntty.
        For Invoice, use normal alias logic.
        """
        if is_note_section:
            if note_field in df.columns:
                return df[note_field]
            if note_field.upper() in df.columns:
                return df[note_field.upper()]
            if note_field.lower() in df.columns:
                return df[note_field.lower()]

        # Normal existing logic
        return get(logical, default)

    out = pd.DataFrame(index=df.index)

    out["supplier_gstin"] = get("supplier_gstin", "").map(normalize_gstin)
    out["supplier_name"] = get("supplier_name", "").astype(str).str.strip()

    # Document Type
    if is_note_section:
        if "ntty" in df.columns:
            out["document_type"] = df["ntty"].astype(str).str.strip()
        else:
            out["document_type"] = pd.Series([default_doc_type] * n, index=df.index)
    else:
        out["document_type"] = get("document_type", default_doc_type).astype(str).replace({"": default_doc_type})

    # Convert GST note type C/D into readable document type
    out["document_type"] = out["document_type"].replace({
        "C": "Credit Note",
        "c": "Credit Note",
        "D": "Debit Note",
        "d": "Debit Note",
        "": default_doc_type,
        "nan": default_doc_type,
        "None": default_doc_type,
    })

    # Document Number
    if is_note_section and "nt_num" in df.columns:
        out["document_no"] = df["nt_num"].astype(str).str.strip()
    elif is_note_section and "ntnum" in df.columns:
        out["document_no"] = df["ntnum"].astype(str).str.strip()
    else:
        out["document_no"] = get("document_no", "").astype(str).str.strip()

    # Clean blank-like values
    out["document_no"] = out["document_no"].replace({
        "None": "",
        "nan": "",
        "NaN": "",
        "<NA>": "",
    })

    out["document_norm"] = out["document_no"].map(normalize_doc_no)

    # Document Date
    if is_note_section and "nt_dt" in df.columns:
        out["document_date"] = to_date(df["nt_dt"])
    elif is_note_section and "ntdt" in df.columns:
        out["document_date"] = to_date(df["ntdt"])
    else:
        out["document_date"] = to_date(get("document_date", pd.NaT))

    out["invoice_value"] = to_number(get("invoice_value", 0))
    out["taxable_value"] = to_number(get("taxable_value", 0))
    out["igst"] = to_number(get("igst", 0))
    out["cgst"] = to_number(get("cgst", 0))
    out["sgst"] = to_number(get("sgst", 0))
    out["cess"] = to_number(get("cess", 0))
    out["total_tax"] = out[TAX_COLS].sum(axis=1)

    out["itc_available"] = get("itc_available", "Yes").astype(str)
    out["ims_status"] = get("ims_status", "No Action").astype(str)
    out["remarks"] = get("remarks", "").astype(str)
    out["pos"] = get("pos", "").astype(str)
    out["return_period"] = get("return_period", "").astype(str)
    out["source"] = source_label
    out["ims_sheet"] = sheet_name
    out["gstin_valid"] = out["supplier_gstin"].map(validate_gstin)
    out["data_quality"] = out.apply(row_quality_score, axis=1)

    sign = out["document_type"].map(sign_for_doc_type)

    for c in MONEY_COLS + ["total_tax"]:
        out[c] = out[c] * sign.where(out[c] >= 0, 1)

    out = out[
        [
            "source", "ims_sheet", "supplier_gstin", "supplier_name", "document_type", "document_no",
            "document_norm", "document_date", "invoice_value", "taxable_value", "igst", "cgst", "sgst", "cess",
            "total_tax", "itc_available", "ims_status", "remarks", "pos", "return_period", "gstin_valid",
            "data_quality"
        ]
    ]

    out = out[
        (out["supplier_gstin"].astype(str).str.len() > 0)
        | (out["document_norm"].astype(str).str.len() > 0)
    ]

    return out.reset_index(drop=True)


def row_quality_score(row) -> int:
    score = 100
    if not validate_gstin(row.get("supplier_gstin", "")) and str(row.get("supplier_gstin", "")).strip():
        score -= 25
    if not str(row.get("document_norm", "")).strip():
        score -= 30
    if pd.isna(pd.to_datetime(row.get("document_date"), errors="coerce")):
        score -= 15
    if abs(float(row.get("taxable_value", 0) or 0)) < 0.01 and abs(float(row.get("total_tax", 0) or 0)) > 0.01:
        score -= 20
    return max(0, min(100, score))


def read_purchase_file(file) -> pd.DataFrame:
    if file is None:
        return pd.DataFrame()
    suffix = Path(file.name).suffix.lower()
    if suffix == ".csv":
        raw = pd.read_csv(file, dtype=object)
        return standardize(raw, "Purchase Register")
    sheets = read_excel_all_sheets(file)
    frames = []
    for sheet, raw in sheets.items():
        std = standardize(raw, "Purchase Register", sheet)
        if not std.empty:
            frames.append(std)
    return pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()


def infer_doc_type_from_sheet(sheet: str) -> str:
    s = sheet.upper()
    if "CN" in s:
        return "Credit Note"
    if "DN" in s:
        return "Debit Note"
    if "A" in s and "B2B" in s:
        return "Amended Invoice"
    return "Invoice"


def read_ims_utility(file) -> pd.DataFrame:
    if file is None:
        return pd.DataFrame()
    sheets = read_excel_all_sheets(file, IMS_SHEETS)
    frames = []
    for sheet, raw in sheets.items():
        default_doc = infer_doc_type_from_sheet(sheet)
        std = standardize(raw, "IMS Utility", sheet, default_doc)
        if not std.empty:
            frames.append(std)
    return pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()


def flatten_json(obj, parent_key="", sep="_"):
    items = []
    if isinstance(obj, dict):
        for k, v in obj.items():
            new_key = f"{parent_key}{sep}{k}" if parent_key else str(k)
            items.extend(flatten_json(v, new_key, sep=sep).items())
    elif isinstance(obj, list):
        if all(isinstance(i, dict) for i in obj):
            return {parent_key: obj}
        return {parent_key: obj}
    else:
        return {parent_key: obj}
    return dict(items)


def extract_records_from_json(obj) -> List[dict]:
    records = []
    def walk(x, path=""):
        if isinstance(x, list):
            if x and all(isinstance(i, dict) for i in x):
                for item in x:
                    flat = flatten_json(item)
                    flat["_json_path"] = path
                    records.append(flat)
            else:
                for i, item in enumerate(x):
                    walk(item, f"{path}.{i}")
        elif isinstance(x, dict):
            for k, v in x.items():
                walk(v, f"{path}.{k}" if path else k)
    walk(obj)
    return records


def ims_json_records(data) -> List[dict]:
    section_map = {
        "b2b": "B2B", "b2ba": "B2BA", "b2bdn": "B2B-DN", "b2bdna": "B2B-DNA",
        "b2bcn": "B2B-CN", "b2bcna": "B2B-CNA", "cdnr": "B2B-CN", "cdnra": "B2B-CNA",
        "dn": "B2B-DN", "dna": "B2B-DNA", "cn": "B2B-CN", "cna": "B2B-CNA",
        "eco": "ECO", "ecoa": "ECOA"
    }
    found = []

    def walk(obj, path=""):
        if isinstance(obj, dict):
            for k, v in obj.items():
                key = str(k).lower().replace("_", "").replace("-", "")
                if key in section_map and isinstance(v, list):
                    for item in v:
                        if isinstance(item, dict):
                            rec = flatten_json(item)
                            rec["__section"] = section_map[key]
                            rec["__json_key"] = str(k)
                            found.append(rec)
                else:
                    walk(v, f"{path}.{k}" if path else str(k))
        elif isinstance(obj, list):
            for i, item in enumerate(obj):
                walk(item, f"{path}.{i}")

    walk(data)
    return found


def read_ims_json(file) -> pd.DataFrame:
    if file is None:
        return pd.DataFrame()
    data = json.load(file)
    records = ims_json_records(data)
    if not records:
        records = extract_records_from_json(data)
    if not records:
        return pd.DataFrame()
    raw = pd.DataFrame(records)
    section = raw.get("__section", pd.Series(["JSON"] * len(raw)))
    frames = []
    for sec, part in raw.groupby(section, dropna=False):
        default_doc = infer_doc_type_from_sheet(str(sec))
        std = standardize(part, "IMS JSON", str(sec), default_doc)
        if not std.empty:
            frames.append(std)
    return pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()


def parse_ims_json_bytes(file_bytes: bytes) -> Tuple[pd.DataFrame, List[dict], dict]:
    data = json.loads(file_bytes.decode("utf-8-sig"))
    records = ims_json_records(data)
    raw = pd.DataFrame(records) if records else pd.DataFrame()
    if raw.empty:
        return pd.DataFrame(), [], data
    frames = []
    for sec, part in raw.groupby(raw["__section"], dropna=False):
        std = standardize(part, "IMS JSON", str(sec), infer_doc_type_from_sheet(str(sec)))
        if not std.empty:
            frames.append(std)
    df = pd.concat(frames, ignore_index=True) if frames else pd.DataFrame()
    return df, records, data


def to_excel_date_string(value):
    if value in [None, "", pd.NaT]:
        return ""
    try:
        dt = pd.to_datetime(value, dayfirst=True, errors="coerce")
        if pd.isna(dt):
            return str(value)
        return dt.strftime("%d-%m-%Y")
    except Exception:
        return str(value)


def get_json_value(rec: dict, *keys, default=""):
    """
    Robust GST JSON value reader.

    GST portal/downloaded IMS JSON may use small variations in field names:
    - nt_num / ntnum / NT_NUM
    - nt_dt  / ntdt  / NT_DT
    - inv_typ / invtyp

    This function keeps the existing logic but also matches keys after removing
    spaces, underscores, hyphens and case differences. This is very important
    for B2B Credit Note / Debit Note matching.
    """
    if not isinstance(rec, dict):
        return default

    blank_values = [None, "", "nan", "NaN", "None", "<NA>"]

    def norm_key(value):
        return re.sub(r"[^a-z0-9]", "", str(value or "").lower())

    # 1) Exact key match first
    for key in keys:
        if key in rec and rec.get(key) not in blank_values:
            return rec.get(key)

    # 2) Case-insensitive match
    lowered = {str(k).lower(): v for k, v in rec.items()}
    for key in keys:
        k = str(key).lower()
        if k in lowered and lowered[k] not in blank_values:
            return lowered[k]

    # 3) Normalized match: ignores underscore, hyphen, space and case
    normalized = {norm_key(k): v for k, v in rec.items()}
    for key in keys:
        nk = norm_key(key)
        if nk in normalized and normalized[nk] not in blank_values:
            return normalized[nk]

    return default

def action_for_record(rec: dict, action_map: dict) -> str:
    gstin = normalize_gstin(get_json_value(rec, "stin", "ctin", "supplier_gstin"))
    doc_no = normalize_doc_no(get_json_value(rec, "inum", "nt_num", "ntnum", "document_no", "oinum"))
    return action_map.get((gstin, doc_no), "Pending")

def build_action_map(recon: pd.DataFrame) -> dict:
    if recon is None or recon.empty:
        return {}
    out = {}
    for _, row in recon.iterrows():
        gstin = normalize_gstin(row.get("supplier_gstin", ""))
        doc = normalize_doc_no(row.get("document_norm", ""))
        action = "Accepted" if str(row.get("mismatch_type", "")) == "Matched" else "Pending"
        if gstin and doc:
            out[(gstin, doc)] = action
    return out


def clear_utility_rows(ws, start_row: int, max_col: int):
    last = max(ws.max_row, start_row + 500)
    for row in ws.iter_rows(min_row=start_row, max_row=last, min_col=1, max_col=max_col):
        for cell in row:
            cell.value = None


def write_row_values(ws, row: int, values: dict):
    for col, value in values.items():
        ws.cell(row=row, column=col).value = value


def populate_ims_utility_xlsm(template_bytes: bytes, records: List[dict], recon: pd.DataFrame) -> bytes:
    if not template_bytes:
        raise ValueError("Please upload Inbuilt IMS Utility .xlsm template first.")
    if not records:
        raise ValueError("Please upload/process IMS JSON first.")

    wb = load_workbook(BytesIO(template_bytes), keep_vba=True)
    action_map = build_action_map(recon)
    groups: Dict[str, List[dict]] = {}
    for rec in records:
        groups.setdefault(str(rec.get("__section", "B2B")), []).append(rec)

    config = {
        "B2B": {"start": 7, "max_col": 23, "amend": False, "note": False},
        "B2B-DN": {"start": 7, "max_col": 23, "amend": False, "note": False},
        "B2B-CN": {"start": 7, "max_col": 35, "amend": False, "note": True},
        "B2BA": {"start": 8, "max_col": 37, "amend": True, "note": False},
        "B2B-DNA": {"start": 8, "max_col": 37, "amend": True, "note": False},
        "B2B-CNA": {"start": 8, "max_col": 37, "amend": True, "note": True},
    }

    for sheet_name, rows in groups.items():
        if sheet_name not in wb.sheetnames or sheet_name not in config:
            continue
        ws = wb[sheet_name]
        cfg = config[sheet_name]
        start = cfg["start"]
        clear_utility_rows(ws, start, cfg["max_col"])

        for idx, rec in enumerate(rows, start=start):
            status = action_for_record(rec, action_map)
            gstin = normalize_gstin(get_json_value(rec, "stin", "ctin", "supplier_gstin"))
            trade = get_json_value(rec, "tradenm", "supplier_name")
            doc_no = get_json_value(rec, "inum", "nt_num", "ntnum", "document_no")
            doc_type = get_json_value(rec, "inv_typ", "invtyp", "ntty", "document_type", default="R")
            doc_date = to_excel_date_string(get_json_value(rec, "idt", "nt_dt", "ntdt", "document_date"))
            doc_val = get_json_value(rec, "val", "invoice_value", default=0)
            pos = get_json_value(rec, "pos", default="")
            txval = get_json_value(rec, "txval", "taxable_value", default=0)
            iamt = get_json_value(rec, "iamt", "igst", default=0)
            camt = get_json_value(rec, "camt", "cgst", default=0)
            samt = get_json_value(rec, "samt", "sgst", default=0)
            cess = get_json_value(rec, "cess", default=0)
            remarks = "Auto Accepted by IMS Recon Pro" if status == "Accepted" else "Auto Pending - not matched in Purchase Register"
            src = get_json_value(rec, "srcform", default="")
            rtnprd = get_json_value(rec, "rtnprd", default="")
            filing = get_json_value(rec, "srcfilstatus", default="")
            pending_block = get_json_value(rec, "ispendactblocked", default="N")
            remarks_block = get_json_value(rec, "isRemarksBlocked", "isremarksblocked", default="N")

            if not cfg["amend"]:
                values = {
                    1: gstin, 2: trade, 3: doc_no, 4: doc_type, 5: doc_date, 6: float(doc_val or 0),
                    7: status, 8: pos, 9: float(txval or 0), 10: float(iamt or 0), 11: float(camt or 0),
                    12: float(samt or 0), 13: float(cess or 0),
                    14: remarks if sheet_name != "B2B-CN" else "No", 15: src, 16: rtnprd, 17: filing,
                    20: get_json_value(rec, "action", default="N"),
                }
                if sheet_name == "B2B-CN":
                    values.update({19: remarks, 20: src, 21: rtnprd, 22: filing, 25: get_json_value(rec, "action", default="N"), 32: pending_block, 33: remarks_block})
                else:
                    values.update({22: pending_block, 23: remarks_block})
            else:
                orig_no = get_json_value(rec, "oinum", "org_inum", "oinv_num", default="")
                orig_dt = to_excel_date_string(get_json_value(rec, "oidt", "org_idt", "oinv_dt", default=""))
                values = {
                    1: orig_no, 2: orig_dt, 3: gstin, 4: trade, 5: doc_no, 6: doc_type, 7: doc_date, 8: float(doc_val or 0),
                    9: status, 10: pos, 11: float(txval or 0), 12: float(iamt or 0), 13: float(camt or 0),
                    14: float(samt or 0), 15: float(cess or 0), 16: "No", 21: remarks, 22: src,
                    23: rtnprd, 24: filing, 27: get_json_value(rec, "action", default="N"),
                    34: pending_block, 35: remarks_block
                }
            write_row_values(ws, idx, values)
            fill = PatternFill("solid", fgColor="E2F0D9") if status == "Accepted" else PatternFill("solid", fgColor="FFF2CC")
            ws.cell(idx, 7 if not cfg["amend"] else 9).fill = fill
            ws.cell(idx, 7 if not cfg["amend"] else 9).font = Font(bold=True)

    if "Home" in wb.sheetnames:
        ws = wb["Home"]
        try:
            rtin = ""
            for rec in records:
                rtin = get_json_value(rec, "rtin", default="") or rtin
            ws["B5"] = "GSTIN"
            ws["C5"] = rtin or st.session_state.get("client_gstin", "")
        except Exception:
            pass

    buffer = BytesIO()
    wb.save(buffer)
    return buffer.getvalue()


def normalize_action_label(value) -> str:
    """Convert IMS Utility status/dropdown value into one of Accepted/Pending/Rejected/No Action."""
    text = str(value or "").strip().lower()
    if text in ["a", "accept", "accepted"]:
        return "Accepted"
    if text in ["p", "pending"]:
        return "Pending"
    if text in ["r", "reject", "rejected"]:
        return "Rejected"
    if text in ["n", "no", "no action", "noaction", "na", ""]:
        return "No Action"
    if "accept" in text:
        return "Accepted"
    if "pend" in text or "review" in text:
        return "Pending"
    if "reject" in text:
        return "Rejected"
    return "Pending"


def action_label_to_gst_code(action: str) -> str:
    """GST IMS JSON action code.

    Important for GST upload JSON:
    - Accepted -> A
    - Pending  -> P
    - Rejected -> R
    - No Action -> N internally only.

    In the final GST upload file, records with No Action/N are skipped, because
    the official utility output sample contains actioned records only.
    """
    label = normalize_action_label(action)
    return {"Accepted": "A", "Pending": "P", "Rejected": "R", "No Action": "N"}.get(label, "P")


def utility_sheet_config() -> Dict[str, dict]:
    return {
        "B2B": {"start": 7, "gstin_col": 1, "doc_col": 3, "status_col": 7},
        "B2B-DN": {"start": 7, "gstin_col": 1, "doc_col": 3, "status_col": 7},
        "B2B-CN": {"start": 7, "gstin_col": 1, "doc_col": 3, "status_col": 7},
        "B2BA": {"start": 8, "gstin_col": 3, "doc_col": 5, "status_col": 9},
        "B2B-DNA": {"start": 8, "gstin_col": 3, "doc_col": 5, "status_col": 9},
        "B2B-CNA": {"start": 8, "gstin_col": 3, "doc_col": 5, "status_col": 9},
    }


def read_action_status_from_utility_xlsm(xlsm_bytes: bytes) -> Tuple[dict, pd.DataFrame]:
    """Read final action/status selected by user from Inbuilt IMS Utility .xlsm."""
    if not xlsm_bytes:
        raise ValueError("Please upload the final/edited IMS JSON file first.")
    wb = load_workbook(BytesIO(xlsm_bytes), keep_vba=True, data_only=False)
    action_map = {}
    rows = []

    for sheet_name, cfg in utility_sheet_config().items():
        if sheet_name not in wb.sheetnames:
            continue
        ws = wb[sheet_name]
        blank_streak = 0
        for r in range(cfg["start"], ws.max_row + 1):
            gstin = normalize_gstin(ws.cell(r, cfg["gstin_col"]).value)
            doc_no_raw = ws.cell(r, cfg["doc_col"]).value
            doc_norm = normalize_doc_no(doc_no_raw)
            status_raw = ws.cell(r, cfg["status_col"]).value
            status = normalize_action_label(status_raw)

            if not gstin and not doc_norm:
                blank_streak += 1
                if blank_streak >= 25:
                    break
                continue
            blank_streak = 0
            if not doc_norm:
                continue

            action_map[(sheet_name, gstin, doc_norm)] = status
            action_map[(gstin, doc_norm)] = status
            rows.append({
                "Sheet": sheet_name,
                "Supplier GSTIN": gstin,
                "Document No": doc_no_raw,
                "Normalized Document No": doc_norm,
                "Utility Status": status,
                "GST JSON Action Code": action_label_to_gst_code(status),
            })

    summary = pd.DataFrame(rows)
    return action_map, summary


def get_json_section_map() -> dict:
    """
    Map downloaded IMS section names to the exact upload JSON section keys.

    Important: GST portal upload schema is case-sensitive. The official utility output
    uses lowercase keys inside `invdata` like `b2b`, not display names like `B2B`.
    """
    return {
        "b2b": "b2b",
        "b2ba": "b2ba",
        "b2bdn": "b2bdn",
        "b2bdna": "b2bdna",
        "b2bcn": "b2bcn",
        "b2bcna": "b2bcna",
        "cdnr": "b2bcn",
        "cdnra": "b2bcna",
        "dn": "b2bdn",
        "dna": "b2bdna",
        "cn": "b2bcn",
        "cna": "b2bcna",
        "eco": "eco",
        "ecoa": "ecoa",
    }


def update_ims_json_actions_from_utility(original_json: dict, action_map: dict) -> Tuple[dict, pd.DataFrame]:
    """Preserve GST portal JSON structure and update only action field based on utility status."""
    if not isinstance(original_json, dict) or not original_json:
        raise ValueError("Original IMS JSON is not available. Please process IMS JSON first.")
    if not action_map:
        raise ValueError("No status/action found in the uploaded IMS JSON.")

    data = deepcopy(original_json)
    section_map = get_json_section_map()
    updated_rows = []

    def walk(obj):
        if isinstance(obj, dict):
            for k, v in obj.items():
                normalized_key = str(k).lower().replace("_", "").replace("-", "")
                if normalized_key in section_map and isinstance(v, list):
                    section = section_map[normalized_key]
                    for item in v:
                        if not isinstance(item, dict):
                            continue
                        gstin = normalize_gstin(get_json_value(item, "stin", "ctin", "supplier_gstin"))
                        doc_no = get_json_value(item, "inum", "nt_num", "ntnum", "document_no", "oinum")
                        doc_norm = normalize_doc_no(doc_no)
                        status = action_map.get((section, gstin, doc_norm)) or action_map.get((gstin, doc_norm))
                        if status:
                            item["action"] = action_label_to_gst_code(status)
                            if "remarks" in item and not str(item.get("remarks") or "").strip():
                                item["remarks"] = f"Auto {status} by IMS Recon Pro"
                            updated_rows.append({
                                "Sheet": section,
                                "Supplier GSTIN": gstin,
                                "Document No": doc_no,
                                "Utility Status": status,
                                "GST JSON Action Code": item.get("action", ""),
                            })
                else:
                    walk(v)
        elif isinstance(obj, list):
            for item in obj:
                walk(item)

    walk(data)
    return data, pd.DataFrame(updated_rows)


def generate_gst_upload_json_bytes(original_json: dict, final_xlsm_bytes: bytes) -> Tuple[bytes, pd.DataFrame, pd.DataFrame]:
    action_map, utility_status_df = read_action_status_from_utility_xlsm(final_xlsm_bytes)
    updated_json, updated_summary = update_ims_json_actions_from_utility(original_json, action_map)
    json_bytes = json.dumps(updated_json, ensure_ascii=False, indent=2).encode("utf-8")
    return json_bytes, utility_status_df, updated_summary



def _compact_json_number(value):
    """Return Python number/string without converting zero/decimal fields into invalid schema values."""
    if value is None:
        return value
    # bool is subclass of int; do not treat it as number here.
    if isinstance(value, bool):
        return value
    if isinstance(value, (int, float)):
        try:
            f = float(value)
            if f.is_integer():
                return int(f)
            return round(f, 2)
        except Exception:
            return value
    return value


def _clean_ims_upload_record(source_item: dict, action_code: str, section: str = "") -> dict:
    """
    Create one GST IMS upload record in schema-safe mode.

    Important GST portal rule:
    Preserve the exact field names/structure received from the GST portal JSON
    and only update the action field. Do not invent or rename fields such as
    ntty/ntnum/ntdt, because that can cause GST schema validation failure.
    """

    if not isinstance(source_item, dict):
        source_item = {}

    blocked_keys = {
        "tradenm", "tradeNm", "trade_name", "supplier_name", "hash",
        "remarks", "remark", "comments", "comment",
        "recommended_action", "final_user_action", "user_remarks",
        "mismatch_type", "risk_level", "risk_score", "reason",
        "match_status", "confidence_score", "source", "ims_sheet",
        "document_norm", "data_quality", "gstin_valid",
        "__section", "__json_key", "_json_path",
    }

    # Keep normal GST order where those fields exist, but never create a field
    # that was not present in the downloaded portal JSON.
    preferred_order = [
        "stin",
        "inum", "inv_typ", "idt",
        "val", "action", "pos", "txval",
        "iamt", "camt", "samt", "cess",
        "srcform", "rtnprd", "srcfilstatus",
        "ispendactblocked", "isRemarksBlocked",
        "nt_num", "nt_dt", "itcRedReqBlocked",
    ]

    out = {}

    for key in preferred_order:
        if key == "action":
            out["action"] = action_code
            continue
        if key in source_item and key not in blocked_keys:
            out[key] = _compact_json_number(source_item.get(key))

    # Append any other GST-source keys exactly as received. This preserves
    # amendment-specific and section-specific keys without breaking schema.
    for key, value in source_item.items():
        if key in blocked_keys or key in out or key == "action":
            continue
        if str(key).startswith("_"):
            continue
        out[key] = _compact_json_number(value)

    # Only action is forcibly changed. No ntty / renamed fields are added.
    out["action"] = action_code

    # =========================================================
    # GST PORTAL IMS B2B-CN IMPORTANT FIX
    # =========================================================
    # The official IMS Offline Tool has one extra taxpayer input field for
    # Accepted Credit Notes: "Whether ITC to be reduced". In JSON this is
    # linked with the portal field family "itcRedReq" / "itcRedReqBlocked".
    #
    # Without this field, GST portal may process the file but show error like
    # "No action taken in B2B Credit Notes" because the CN action is incomplete.
    #
    # Apply it ONLY for accepted Credit Note sections. For Pending/Rejected/No
    # Action, do not add it because GST utility instructions say this field is
    # applicable only for Accepted records in relevant sections.
    section_norm = str(section or "").lower().replace("_", "").replace("-", "")
    is_credit_note_section = section_norm in ["b2bcn", "b2bcna", "cdnr", "cdnra", "cn", "cna"] or "cn" in section_norm

    if is_credit_note_section and str(action_code).upper() == "A":
        if not get_json_value(out, "itcRedReq", "itcredreq", default=""):
            out["itcRedReq"] = "Y"

    return out

def _record_missing_mandatory(upload_item: dict, section: str = "") -> list:
    """
    Validate only that the existing portal-style record has enough identity
    fields. For B2B-CN/B2B-DN, GST JSON may use nt_num/nt_dt with inv_typ and
    may not contain ntty. Therefore ntty is NOT mandatory here.
    """

    missing = []
    section_norm = str(section or "").lower().replace("_", "").replace("-", "")
    is_note_section = (
        section_norm in ["b2bcn", "b2bcna", "b2bdn", "b2bdna", "cdnr", "cdnra", "cn", "cna", "dn", "dna"]
        or "cn" in section_norm
        or "dn" in section_norm
    )

    common_mandatory = ["stin", "val", "action", "pos", "txval", "srcform", "rtnprd", "srcfilstatus"]
    for key in common_mandatory:
        if get_json_value(upload_item, key) in [None, ""]:
            missing.append(key)

    if is_note_section:
        if get_json_value(upload_item, "nt_num", "ntnum", "inum", "document_no") in [None, ""]:
            missing.append("nt_num/inum")
        if get_json_value(upload_item, "nt_dt", "ntdt", "idt", "document_date") in [None, ""]:
            missing.append("nt_dt/idt")
        # GST downloaded b2bcn commonly has inv_typ='R' and no ntty.
        if get_json_value(upload_item, "inv_typ", "invtyp", "ntty", "document_type") in [None, ""]:
            missing.append("inv_typ/ntty")
    else:
        if get_json_value(upload_item, "inum", "document_no", "nt_num", "ntnum", "oinum") in [None, ""]:
            missing.append("inum")
        if get_json_value(upload_item, "idt", "document_date", "nt_dt", "ntdt", "oidt") in [None, ""]:
            missing.append("idt")
        if get_json_value(upload_item, "inv_typ", "invtyp", "document_type", "ntty") in [None, ""]:
            missing.append("inv_typ")

    return missing


def ims_json_section_counts(original_json: dict) -> pd.DataFrame:
    """Return section-wise counts from the uploaded GST IMS JSON / generated upload JSON."""
    section_map = get_json_section_map()
    rows = []
    if not isinstance(original_json, dict) or not original_json:
        return pd.DataFrame(columns=["Section", "Records"])
    root = original_json.get("imsDetails", original_json.get("invdata", original_json))

    def normalized_section_name(key: str) -> str:
        raw = str(key or "").strip()
        norm = raw.lower().replace("_", "").replace("-", "")
        return section_map.get(norm, raw.lower())

    def walk(obj):
        if isinstance(obj, dict):
            for key, value in obj.items():
                norm = str(key).lower().replace("_", "").replace("-", "")
                if norm in section_map and isinstance(value, list):
                    rows.append({"Section": normalized_section_name(key), "Records": len(value)})
                elif isinstance(value, (dict, list)):
                    walk(value)
        elif isinstance(obj, list):
            for item in obj:
                walk(item)

    walk(root)
    if not rows:
        return pd.DataFrame(columns=["Section", "Records"])
    out = pd.DataFrame(rows).groupby("Section", as_index=False)["Records"].sum()
    return out.sort_values("Section").reset_index(drop=True)


def generated_json_action_counts(json_bytes: bytes) -> pd.DataFrame:
    """Return section/action counts from the final GST upload JSON bytes."""
    try:
        data = json.loads(json_bytes.decode("utf-8") if isinstance(json_bytes, (bytes, bytearray)) else str(json_bytes))
    except Exception:
        return pd.DataFrame(columns=["Section", "Action", "Records"])
    rows = []
    invdata = data.get("invdata", {}) if isinstance(data, dict) else {}
    if isinstance(invdata, dict):
        for section, records in invdata.items():
            if isinstance(records, list):
                for rec in records:
                    if isinstance(rec, dict):
                        rows.append({"Section": section, "Action": rec.get("action", ""), "Records": 1})
    if not rows:
        return pd.DataFrame(columns=["Section", "Action", "Records"])
    return pd.DataFrame(rows).groupby(["Section", "Action"], as_index=False)["Records"].sum()

def generate_gst_upload_json_from_final_actions(original_json: dict, action_df: pd.DataFrame) -> Tuple[bytes, pd.DataFrame]:
    """
    Generate GST Portal IMS upload JSON in official utility style.

    The user's official output sample confirms this wrapper:
    {
      "rtin": "...",
      "reqtyp": "SAVE",
      "invdata": {"b2b": [ ... ]}
    }

    This V6 generator is amendment-safe and GST-utility style:
    - lowercase invdata section keys
    - no imsDetails wrapper
    - no tradenm/hash/remarks/internal fields, but preserves amendment mandatory fields
    - every original IMS JSON record is actively actioned by default
    - Matched/manual Accepted => A
    - Unmatched/missing/No Action/Review => P (Pending), so records do not remain No Action on portal
    - invalid identity records are skipped and shown in summary
    """
    if not isinstance(original_json, dict) or not original_json:
        raise ValueError("Original IMS JSON is not available. Please process IMS JSON first.")
    if action_df is None or action_df.empty:
        raise ValueError("Final action table is empty. Please run reconciliation and save final actions first.")

    action_map = {}
    for _, row in action_df.iterrows():
        gstin = normalize_gstin(row.get("supplier_gstin", ""))
        doc_norm = normalize_doc_no(row.get("document_norm", ""))
        if not gstin or not doc_norm:
            continue
        action = normalize_action_label(row.get("final_user_action", row.get("recommended_action", "Pending")))
        section = str(row.get("ims_sheet_ims", row.get("ims_sheet", "")) or "").strip().lower()
        action_map[(gstin, doc_norm)] = action
        if section:
            action_map[(section, gstin, doc_norm)] = action

    section_map = get_json_section_map()
    invdata = {}
    updated_rows = []
    skipped_rows = []

    source_root = original_json.get("imsDetails", original_json.get("invdata", original_json))

    def normalized_section_name(key: str) -> str:
        raw = str(key or "").strip()
        norm = raw.lower().replace("_", "").replace("-", "")
        return section_map.get(norm, raw.lower())

    def process_rows(section_key: str, rows: list):
        section = normalized_section_name(section_key)
        if not isinstance(rows, list):
            return

        for item in rows:
            if not isinstance(item, dict):
                continue

            gstin = normalize_gstin(get_json_value(item, "stin", "ctin", "supplier_gstin"))
            doc_no = get_json_value(item, "inum", "nt_num", "ntnum", "document_no", "oinum")
            doc_norm = normalize_doc_no(doc_no)
            if not gstin or not doc_norm:
                skipped_rows.append({"Section": section, "Supplier GSTIN": gstin, "Document No": doc_no, "Reason": "Missing GSTIN or document number"})
                continue

            status = (
                action_map.get((section, gstin, doc_norm))
                or action_map.get((str(section_key).lower(), gstin, doc_norm))
                or action_map.get((gstin, doc_norm))
            )

            # Very important GST IMS rule for this project:
            # after reconciliation, every IMS invoice should be actioned.
            # Matched records become Accepted. All records not found/mapped in the
            # Action Center are treated as Pending, not No Action, because the
            # portal otherwise continues to show them under No Action.
            if not status:
                status = "Pending"

            action_code = action_label_to_gst_code(status)

            # For GST upload, do not leave JSON records as N/No Action.
            # If the app/user action is No Action or Review, keep it Pending so
            # the portal moves it out of No Action and the user can take final
            # action later from IMS dashboard if required.
            if action_code == "N":
                status = "Pending"
                action_code = "P"

            upload_item = _clean_ims_upload_record(item, action_code, section)
            missing = _record_missing_mandatory(upload_item, section)
            if missing:
                skipped_rows.append({"Section": section, "Supplier GSTIN": gstin, "Document No": doc_no, "Reason": "Missing mandatory fields: " + ", ".join(missing)})
                continue

            invdata.setdefault(section, []).append(upload_item)
            updated_rows.append({
                "Section": section,
                "Supplier GSTIN": gstin,
                "Document No": doc_no,
                "Final Action": status,
                "GST JSON Action Code": action_code,
                "Validation": "Included in GST upload JSON",
            })

    def walk_sections(obj):
        if isinstance(obj, dict):
            for key, value in obj.items():
                norm = str(key).lower().replace("_", "").replace("-", "")
                if norm in section_map and isinstance(value, list):
                    process_rows(key, value)
                elif isinstance(value, (dict, list)):
                    walk_sections(value)
        elif isinstance(obj, list):
            for item in obj:
                walk_sections(item)

    walk_sections(source_root)

    # Remove any empty section defensively.
    invdata = {k: v for k, v in invdata.items() if isinstance(v, list) and len(v) > 0}

    if not invdata:
        raise ValueError("No valid IMS records found for GST JSON generation. Please check the uploaded GST IMS JSON.")

    upload_json = {
        "rtin": str(original_json.get("rtin", st.session_state.get("client_gstin", ""))).strip(),
        "reqtyp": "SAVE",
        "invdata": invdata,
    }

    json_bytes = json.dumps(upload_json, ensure_ascii=False, indent=3).encode("utf-8")
    included = pd.DataFrame(updated_rows)
    skipped = pd.DataFrame(skipped_rows)
    if not skipped.empty:
        skipped.insert(0, "GST JSON Action Code", "SKIPPED")
    summary = pd.concat([included, skipped], ignore_index=True) if not skipped.empty else included
    return json_bytes, summary



# =========================================================
# V7 STRONGER RECONCILIATION + VALIDATION HELPERS
# =========================================================

def safe_float_value(value) -> float:
    try:
        if pd.isna(value):
            return 0.0
        return float(value)
    except Exception:
        return 0.0


def approx_equal(a, b, tolerance: float) -> bool:
    return abs(safe_float_value(a) - safe_float_value(b)) <= float(tolerance or 0)


def date_gap_days(a, b) -> Optional[int]:
    da = pd.to_datetime(a, errors="coerce")
    db = pd.to_datetime(b, errors="coerce")
    if pd.isna(da) or pd.isna(db):
        return None
    return int(abs((da - db).days))


def is_credit_note_text(value) -> bool:
    text = str(value or "").lower()
    return any(x in text for x in ["credit", "cn", "cdn", "b2b-cn", "b2bcna"])


def make_recon_key(gstin, doc_norm) -> str:
    return f"{normalize_gstin(gstin)}|{normalize_doc_no(doc_norm)}"


def enhance_recon_row(row, amount_tol: float, date_tol: int) -> pd.Series:
    """Classify exact key matches using amount, tax-head and date checks."""
    merge_status = str(row.get("_merge", ""))
    if merge_status == "left_only":
        row["mismatch_type"] = "Only in Purchase Register"
        row["match_level"] = "L0 Not in IMS"
        return row
    if merge_status == "right_only":
        row["mismatch_type"] = "Only in IMS"
        row["match_level"] = "L0 Not in Purchase Register"
        return row

    amount_ok = (
        abs(safe_float_value(row.get("taxable_value_diff"))) <= amount_tol
        and abs(safe_float_value(row.get("total_tax_diff"))) <= amount_tol
        and abs(safe_float_value(row.get("invoice_value_diff"))) <= max(amount_tol, 1)
    )
    tax_head_ok = all(abs(safe_float_value(row.get(f"{c}_diff"))) <= amount_tol for c in ["igst", "cgst", "sgst", "cess"])
    gap = row.get("date_diff_days")
    try:
        date_ok = int(gap) <= int(date_tol)
    except Exception:
        date_ok = True

    row["match_level"] = "L1 Exact GSTIN + Invoice No"
    if amount_ok and tax_head_ok and date_ok:
        row["mismatch_type"] = "Matched"
    elif amount_ok and not tax_head_ok:
        row["mismatch_type"] = "Tax Head Mismatch"
    elif (not amount_ok) and date_ok:
        row["mismatch_type"] = "Value / Tax Mismatch"
    elif amount_ok and tax_head_ok and not date_ok:
        row["mismatch_type"] = "Date Mismatch"
    else:
        row["mismatch_type"] = "Value and Date Mismatch"
    return row


def add_probable_match_flags(result: pd.DataFrame, p_agg: pd.DataFrame, i_agg: pd.DataFrame, amount_tol: float, date_tol: int) -> pd.DataFrame:
    """Mark IMS-only rows where supplier/value/date indicates a likely invoice-number difference."""
    if result.empty or p_agg.empty or i_agg.empty:
        return result
    p_lookup = p_agg.copy()
    i_only_mask = result["_merge"].astype(str).eq("right_only")
    if not i_only_mask.any():
        return result

    for idx, row in result[i_only_mask].iterrows():
        gstin = row.get("supplier_gstin", "")
        candidates = p_lookup[p_lookup["supplier_gstin"].astype(str).eq(str(gstin))].copy()
        if candidates.empty:
            continue
        best_score = 999999.0
        best = None
        for _, p in candidates.iterrows():
            tax_gap = abs(safe_float_value(p.get("total_tax_purchase")) - safe_float_value(row.get("total_tax_ims")))
            taxable_gap = abs(safe_float_value(p.get("taxable_value_purchase")) - safe_float_value(row.get("taxable_value_ims")))
            inv_gap = abs(safe_float_value(p.get("invoice_value_purchase")) - safe_float_value(row.get("invoice_value_ims")))
            dg = date_gap_days(p.get("document_date_purchase"), row.get("document_date_ims"))
            dg_score = 9999 if dg is None else dg
            if taxable_gap <= amount_tol and tax_gap <= amount_tol and inv_gap <= max(amount_tol, 1) and dg_score <= date_tol:
                score = taxable_gap + tax_gap + inv_gap + dg_score
                if score < best_score:
                    best_score = score
                    best = p
        if best is not None:
            result.loc[idx, "mismatch_type"] = "Probable Match - Invoice No Difference"
            result.loc[idx, "match_level"] = "L3 GSTIN + Date + Value"
            result.loc[idx, "confidence_score"] = 78
            result.loc[idx, "reason"] = "GSTIN, date and values are close but invoice/document number differs. Review invoice number format before accepting."
            for col in ["supplier_name_purchase", "document_type_purchase", "document_no_purchase", "document_date_purchase", "invoice_value_purchase", "taxable_value_purchase", "igst_purchase", "cgst_purchase", "sgst_purchase", "cess_purchase", "total_tax_purchase"]:
                if col in best.index:
                    result.loc[idx, col] = best[col]
            for c in MONEY_COLS + ["total_tax"]:
                result.loc[idx, f"{c}_diff"] = safe_float_value(result.loc[idx].get(f"{c}_purchase")) - safe_float_value(result.loc[idx].get(f"{c}_ims"))
            result.loc[idx, "date_diff_days"] = date_gap_days(result.loc[idx].get("document_date_purchase"), result.loc[idx].get("document_date_ims")) or 0
    return result


def upload_quality_summary(df: pd.DataFrame, label: str) -> pd.DataFrame:
    """Upload validation with value columns for taxation review."""
    value_cols = ["taxable_value", "igst", "cgst", "sgst", "cess", "total_tax"]

    def blank_series(index=None):
        return pd.Series([False] * (0 if index is None else len(index)), index=index)

    def amount_row(check: str, subset: pd.DataFrame) -> dict:
        subset = subset if subset is not None else pd.DataFrame()
        row = {
            "Check": check,
            "Records": int(len(subset)),
            "Taxable Value": round(safe_float_value(subset.get("taxable_value", pd.Series(dtype=float)).sum()) if "taxable_value" in subset else 0, 2),
            "Invoice Value": round(safe_float_value(subset.get("invoice_value", pd.Series(dtype=float)).sum()) if "invoice_value" in subset else 0, 2),
            "IGST": round(safe_float_value(subset.get("igst", pd.Series(dtype=float)).sum()) if "igst" in subset else 0, 2),
            "CGST": round(safe_float_value(subset.get("cgst", pd.Series(dtype=float)).sum()) if "cgst" in subset else 0, 2),
            "SGST": round(safe_float_value(subset.get("sgst", pd.Series(dtype=float)).sum()) if "sgst" in subset else 0, 2),
            "CESS": round(safe_float_value(subset.get("cess", pd.Series(dtype=float)).sum()) if "cess" in subset else 0, 2),
            "Total Tax": round(safe_float_value(subset.get("total_tax", pd.Series(dtype=float)).sum()) if "total_tax" in subset else 0, 2),
        }
        return row

    if df is None or df.empty:
        return pd.DataFrame([amount_row(f"{label} records", pd.DataFrame())])

    work = df.copy()
    rows = []
    rows.append(amount_row(f"{label} records", work))

    if "gstin_valid" in work.columns:
        valid_mask = work["gstin_valid"].fillna(False).astype(bool)
        rows.append(amount_row("Valid GSTIN", work[valid_mask]))
        rows.append(amount_row("Invalid GSTIN", work[~valid_mask]))
    else:
        rows.append(amount_row("Valid GSTIN", pd.DataFrame()))
        rows.append(amount_row("Invalid GSTIN", pd.DataFrame()))

    if "document_norm" in work.columns:
        blank_inv = work["document_norm"].astype(str).eq("")
        rows.append(amount_row("Blank invoice/document no", work[blank_inv]))
    else:
        rows.append(amount_row("Blank invoice/document no", pd.DataFrame()))

    if {"supplier_gstin", "document_norm"}.issubset(work.columns):
        dup_mask = work.duplicated(["supplier_gstin", "document_norm"], keep=False)
        rows.append(amount_row("Duplicate GSTIN + invoice/document no", work[dup_mask]))
    else:
        rows.append(amount_row("Duplicate GSTIN + invoice/document no", pd.DataFrame()))

    if "document_date" in work.columns:
        blank_date = pd.to_datetime(work["document_date"], errors="coerce").isna()
        rows.append(amount_row("Blank document date", work[blank_date]))
    else:
        rows.append(amount_row("Blank document date", pd.DataFrame()))

    # Additional tax-wise check helpful for GST reconciliation review
    if all(c in work.columns for c in ["igst", "cgst", "sgst"]):
        tax_blank = work[["igst", "cgst", "sgst"]].fillna(0).abs().sum(axis=1).le(0.009)
        rows.append(amount_row("Zero IGST/CGST/SGST", work[tax_blank]))

    return pd.DataFrame(rows)


def duplicate_report(df: pd.DataFrame, label: str) -> pd.DataFrame:
    if df is None or df.empty or not {"supplier_gstin", "document_norm"}.issubset(df.columns):
        return pd.DataFrame()
    dup = df[df.duplicated(["supplier_gstin", "document_norm"], keep=False)].copy()
    if dup.empty:
        return pd.DataFrame()
    cols = [
        c for c in [
            "supplier_gstin", "supplier_name", "document_no", "document_date",
            "invoice_value", "taxable_value", "igst", "cgst", "sgst", "cess", "total_tax",
            "source", "ims_sheet"
        ] if c in dup.columns
    ]
    out = dup[cols].copy()
    rename_map = {
        "invoice_value": "Invoice Value",
        "taxable_value": "Taxable Value",
        "igst": "IGST",
        "cgst": "CGST",
        "sgst": "SGST",
        "cess": "CESS",
        "total_tax": "Total Tax",
        "supplier_gstin": "Supplier GSTIN",
        "supplier_name": "Supplier Name",
        "document_no": "Invoice/Document No",
        "document_date": "Document Date",
        "source": "Source",
        "ims_sheet": "IMS Section",
    }
    out = out.rename(columns=rename_map)
    out.insert(0, "Dataset", label)
    return out


def final_json_review_table(action_df: pd.DataFrame) -> pd.DataFrame:
    if action_df is None or action_df.empty:
        return pd.DataFrame()
    work = action_df.copy()
    work["GST JSON Code"] = work.get("final_user_action", "Pending").apply(action_label_to_gst_code)
    return work.groupby(["final_user_action", "GST JSON Code"], dropna=False).size().reset_index(name="Records")


def split_report_sheets(p: pd.DataFrame, ims: pd.DataFrame, recon: pd.DataFrame, action: pd.DataFrame) -> Dict[str, pd.DataFrame]:
    sheets = {
        "Summary": recon_summary(recon),
        "Final Action Report": action,
        "Reconciliation": recon,
        "Purchase Standardized": p,
        "IMS JSON Standardized": ims,
        "Purchase Quality": upload_quality_summary(p, "Purchase Register"),
        "IMS Quality": upload_quality_summary(ims, "IMS JSON"),
        "Purchase Duplicates": duplicate_report(p, "Purchase Register"),
        "IMS Duplicates": duplicate_report(ims, "IMS JSON"),
        "Audit Log": load_audit(st.session_state.username),
    }
    if recon is not None and not recon.empty:
        sheets.update({
            "Matched": recon[recon["mismatch_type"].eq("Matched")],
            "Pending Cases": action[action.get("final_user_action", pd.Series()).eq("Pending")] if action is not None and not action.empty else pd.DataFrame(),
            "Rejected Cases": action[action.get("final_user_action", pd.Series()).eq("Rejected")] if action is not None and not action.empty else pd.DataFrame(),
            "Only in IMS": recon[recon["mismatch_type"].eq("Only in IMS")],
            "Only in Purchase": recon[recon["mismatch_type"].eq("Only in Purchase Register")],
            "Value Mismatch": recon[recon["mismatch_type"].astype(str).str.contains("Value|Tax Head", case=False, na=False)],
            "High Risk": recon[recon["risk_level"].isin(["High", "Critical"])],
            "Probable Matches": recon[recon["mismatch_type"].astype(str).str.contains("Probable", case=False, na=False)],
        })
    return sheets

def aggregate(df: pd.DataFrame, label: str) -> pd.DataFrame:
    if df is None or df.empty:
        return pd.DataFrame()
    work = df.copy()
    work = work[(work["supplier_gstin"].astype(str) != "") & (work["document_norm"].astype(str) != "")]
    if work.empty:
        return pd.DataFrame()

    agg_map = {
        "supplier_name": "first",
        "document_type": "first",
        "document_no": "first",
        "document_date": "min",
        "invoice_value": "sum",
        "taxable_value": "sum",
        "igst": "sum",
        "cgst": "sum",
        "sgst": "sum",
        "cess": "sum",
        "total_tax": "sum",
        "itc_available": "first",
        "ims_status": "first",
        "remarks": "first",
        "ims_sheet": "first",
        "data_quality": "mean",
    }
    out = work.groupby(["supplier_gstin", "document_norm"], dropna=False).agg(agg_map).reset_index()
    rename = {col: f"{col}_{label}" for col in out.columns if col not in ["supplier_gstin", "document_norm"]}
    return out.rename(columns=rename)


def calculate_recon(purchase: pd.DataFrame, ims: pd.DataFrame, amount_tol: float, date_tol: int, include_amendments: bool) -> pd.DataFrame:
    """V7 stronger reconciliation engine.

    Matching levels:
    L1: GSTIN + cleaned invoice/document number
    L3: GSTIN + date + values where invoice number differs (probable match flag)
    All exact matches are then tested for value/date/tax-head mismatch.
    """
    if purchase is None or ims is None or purchase.empty or ims.empty:
        return pd.DataFrame()

    ims_work = ims.copy()
    if not include_amendments and "ims_sheet" in ims_work.columns:
        ims_work = ims_work[~ims_work["ims_sheet"].astype(str).str.upper().isin(["B2BA", "B2B-DNA", "B2B-CNA", "B2BDN", "B2BCN"])]

    p = aggregate(purchase, "purchase")
    i = aggregate(ims_work, "ims")
    if p.empty or i.empty:
        return pd.DataFrame()

    m = p.merge(i, on=["supplier_gstin", "document_norm"], how="outer", indicator=True)

    for c in MONEY_COLS + ["total_tax"]:
        m[f"{c}_diff"] = m.get(f"{c}_purchase", 0).fillna(0) - m.get(f"{c}_ims", 0).fillna(0)

    pdate = pd.to_datetime(m.get("document_date_purchase"), errors="coerce")
    idate = pd.to_datetime(m.get("document_date_ims"), errors="coerce")
    m["date_diff_days"] = (pdate - idate).dt.days.abs().fillna(0).astype("Int64")

    m["recon_key"] = m.apply(lambda r: make_recon_key(r.get("supplier_gstin"), r.get("document_norm")), axis=1)
    m["mismatch_type"] = "Review"
    m["match_level"] = "L0 Not matched"
    m = m.apply(lambda r: enhance_recon_row(r, amount_tol, date_tol), axis=1)

    # Mark likely matches where invoice number differs but GSTIN/date/value are close.
    m = add_probable_match_flags(m, p, i, amount_tol, date_tol)

    # Data-quality and duplicate flags.
    purchase_dups = set()
    ims_dups = set()
    if {"supplier_gstin", "document_norm"}.issubset(purchase.columns):
        purchase_dups = set(purchase[purchase.duplicated(["supplier_gstin", "document_norm"], keep=False)].apply(lambda r: make_recon_key(r.get("supplier_gstin"), r.get("document_norm")), axis=1))
    if {"supplier_gstin", "document_norm"}.issubset(ims_work.columns):
        ims_dups = set(ims_work[ims_work.duplicated(["supplier_gstin", "document_norm"], keep=False)].apply(lambda r: make_recon_key(r.get("supplier_gstin"), r.get("document_norm")), axis=1))
    m["duplicate_flag"] = m["recon_key"].apply(lambda k: "Purchase Duplicate" if k in purchase_dups else ("IMS Duplicate" if k in ims_dups else ""))
    m.loc[m["duplicate_flag"].ne("") & m["mismatch_type"].eq("Matched"), "mismatch_type"] = "Duplicate Review"

    # Presentation columns
    m["supplier_name"] = m.get("supplier_name_purchase").fillna(m.get("supplier_name_ims"))
    m["document_type"] = m.get("document_type_purchase").fillna(m.get("document_type_ims"))
    m["document_no"] = m.get("document_no_purchase").fillna(m.get("document_no_ims"))
    m["document_date"] = pd.to_datetime(m.get("document_date_purchase"), errors="coerce").fillna(pd.to_datetime(m.get("document_date_ims"), errors="coerce"))

    m["risk_score"] = m.apply(risk_score, axis=1)
    m["risk_level"] = m["risk_score"].map(risk_level)
    m["recommended_action"] = m.apply(recommend_action, axis=1)
    m["reason"] = m.apply(recommend_reason, axis=1)
    m.loc[m["mismatch_type"].eq("Probable Match - Invoice No Difference"), "recommended_action"] = "Pending"
    m.loc[m["mismatch_type"].eq("Probable Match - Invoice No Difference"), "reason"] = "Probable match by GSTIN/date/value. Keep Pending until invoice number is confirmed."
    m.loc[m["mismatch_type"].eq("Duplicate Review"), "recommended_action"] = "Pending"
    m.loc[m["mismatch_type"].eq("Duplicate Review"), "reason"] = "Duplicate document key detected. Review before final IMS action."
    m["vendor_followup_required"] = m["mismatch_type"].isin(["Only in Purchase Register", "Value / Tax Mismatch", "Tax Head Mismatch", "Value and Date Mismatch", "Only in IMS", "Probable Match - Invoice No Difference", "Duplicate Review"])
    m["final_user_action"] = m["mismatch_type"].apply(lambda x: "Accepted" if x == "Matched" else "Pending")
    m["user_remarks"] = ""
    m["confidence_score"] = m.apply(confidence_score, axis=1)
    m.loc[m["mismatch_type"].eq("Probable Match - Invoice No Difference"), "confidence_score"] = 78
    m.loc[m["mismatch_type"].eq("Duplicate Review"), "confidence_score"] = 45
    m["json_action_code"] = m["final_user_action"].apply(action_label_to_gst_code)

    priority_order = {"Critical": 0, "High": 1, "Medium": 2, "Low": 3}
    m["_risk_order"] = m["risk_level"].map(priority_order).fillna(9)
    m = m.sort_values(["_risk_order", "mismatch_type", "supplier_gstin", "document_norm"]).drop(columns=["_risk_order"])
    return m.reset_index(drop=True)


def risk_score(row) -> int:
    score = 0
    typ = str(row.get("mismatch_type", ""))
    total_tax_diff = abs(float(row.get("total_tax_diff", 0) or 0))
    taxable_diff = abs(float(row.get("taxable_value_diff", 0) or 0))
    if typ == "Only in IMS":
        score += 30
    if typ == "Only in Purchase Register":
        score += 25
    if "Value" in typ:
        score += 25
    if "Tax Head" in typ:
        score += 20
    if "Date" in typ:
        score += 10
    if total_tax_diff >= 100000 or taxable_diff >= 500000:
        score += 25
    if "CN" in str(row.get("ims_sheet_ims", "")).upper() or "credit" in str(row.get("document_type_ims", "")).lower():
        score += 15
    return min(100, score)


def risk_level(score) -> str:
    score = int(score or 0)
    if score >= 75:
        return "Critical"
    if score >= 50:
        return "High"
    if score >= 21:
        return "Medium"
    return "Low"


def recommend_action(row) -> str:
    typ = str(row.get("mismatch_type", ""))
    itc = str(row.get("itc_available_purchase", "Yes")).lower()
    if "no" in itc or "ineligible" in itc or "not" in itc:
        return "Rejected"
    if typ == "Matched":
        return "Accepted"
    if typ == "Only in IMS":
        return "Pending"
    if typ == "Only in Purchase Register":
        return "No Action"
    if typ in ["Value / Tax Mismatch", "Tax Head Mismatch", "Value and Date Mismatch", "Date Mismatch"]:
        return "Pending"
    return "Review"


def recommend_reason(row) -> str:
    typ = str(row.get("mismatch_type", ""))
    if typ == "Matched":
        return "Purchase Register and IMS values are matched within selected tolerance."
    if typ == "Only in IMS":
        return "Document appears in IMS but is not found in Purchase Register. Keep Pending until booking/vendor confirmation."
    if typ == "Only in Purchase Register":
        return "Document exists in books but not in IMS. No IMS action possible; follow up with vendor if ITC expected."
    if typ == "Tax Head Mismatch":
        return "Total tax may be close but IGST/CGST/SGST/Cess split differs. Check POS and tax head classification."
    if typ == "Value / Tax Mismatch":
        return "Amount difference detected. Compare invoice copy, credit/debit note treatment and amendment."
    if typ == "Date Mismatch":
        return "Invoice matched by GSTIN and document number, but document date differs beyond tolerance."
    if typ == "Value and Date Mismatch":
        return "Both amount and date differences detected. Manual review required before IMS action."
    return "Review required."


def confidence_score(row) -> int:
    typ = str(row.get("mismatch_type", ""))
    if typ == "Matched":
        return 100
    if typ == "Date Mismatch":
        return 82
    if typ == "Tax Head Mismatch":
        return 75
    if typ == "Value / Tax Mismatch":
        return 65
    if typ.startswith("Only"):
        return 35
    return 50


def recon_summary(df: pd.DataFrame) -> pd.DataFrame:
    if df.empty:
        return pd.DataFrame()
    return df.groupby(["mismatch_type", "risk_level", "recommended_action"], dropna=False).agg(
        Records=("mismatch_type", "size"),
        Taxable_Diff=("taxable_value_diff", "sum"),
        Tax_Diff=("total_tax_diff", "sum"),
        Avg_Confidence=("confidence_score", "mean"),
    ).reset_index().round(2)


def to_excel_bytes(sheets: Dict[str, pd.DataFrame]) -> bytes:
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        meta = pd.DataFrame([
            {"Field": "Application", "Value": APP_TITLE},
            {"Field": "Owner", "Value": COPYRIGHT_OWNER},
            {"Field": "Engine Version", "Value": ENGINE_VERSION},
            {"Field": "Generated On", "Value": datetime.now().strftime("%d-%b-%Y %H:%M:%S")},
            {"Field": "Client", "Value": st.session_state.get("client_name", "")},
            {"Field": "GSTIN", "Value": st.session_state.get("client_gstin", "")},
            {"Field": "Return Period", "Value": st.session_state.get("return_period", "")},
            {"Field": "Generated By", "Value": st.session_state.get("username", "")},
        ])
        meta.to_excel(writer, index=False, sheet_name="BAJRABHANU")
        used = {"BAJRABHANU"}
        for sheet_name, df in sheets.items():
            safe = re.sub(r"[\[\]:*?/\\]", "_", str(sheet_name))[:31] or "Sheet"
            base = safe
            i = 1
            while safe in used:
                safe = f"{base[:27]}_{i}"
                i += 1
            used.add(safe)
            data = df if isinstance(df, pd.DataFrame) and not df.empty else pd.DataFrame({"Message": ["No data available"]})
            data.to_excel(writer, index=False, sheet_name=safe)
            ws = writer.sheets[safe]
            ws.freeze_panes = "A2"
            ws.auto_filter.ref = ws.dimensions
            for col_cells in ws.columns:
                max_len = 0
                letter = col_cells[0].column_letter
                for cell in col_cells:
                    max_len = max(max_len, len(str(cell.value)) if cell.value is not None else 0)
                ws.column_dimensions[letter].width = min(max_len + 2, 38)
    return buffer.getvalue()


def safe_display_df(df: pd.DataFrame, limit: int = 1000) -> pd.DataFrame:
    if df is None or df.empty:
        return pd.DataFrame()
    out = df.head(limit).copy()
    for col in out.columns:
        if pd.api.types.is_datetime64_any_dtype(out[col]):
            out[col] = out[col].dt.strftime("%d-%b-%Y").fillna("")
    return out


# =========================================================
# UI HELPERS
# =========================================================

def top_header():
    client_name = st.session_state.get("client_name", "") or "Not set"
    client_gstin = st.session_state.get("client_gstin", "") or "GSTIN pending"
    st.markdown(f"""
    <div class='gst-shell'>
        <div class='gst-top-strip'>
            <span>IMS Recon Pro</span>
            <span>•</span>
            <span>Smart GST IMS Workflow</span>
            <span>•</span>
            <span>Designed for India</span>
        </div>
        <div class='gst-masthead'>
            <div class='gst-brand'>
                <div class='gst-emblem'>🧾</div>
                <div>
                    <div class='gst-title'>Goods and Services Tax</div>
                    <div class='gst-subtitle'>IMS Recon Pro — Reconciliation, Action & GST JSON Platform</div>
                    <div class='header-note'>Premium compliance workspace for Purchase Register vs IMS JSON review</div>
                </div>
            </div>
            <div class='gst-action-wrap'>
                <div class='gst-floating-flag' title='India'></div>
            </div>
        </div>
        <div class='gst-meta-row'>
            <div class='gst-meta-card'>
                <div class='gst-meta-icon'>🗓️</div>
                <div><div class='gst-meta-label'>Today</div><div class='gst-meta-value'>{datetime.today().strftime("%d %b %Y")} • {datetime.today().strftime("%A")}</div></div>
            </div>
            <div class='gst-meta-card'>
                <div class='gst-meta-icon'>👤</div>
                <div><div class='gst-meta-label'>Logged in user</div><div class='gst-meta-value'>{st.session_state.get("display_name", "User")} • {st.session_state.get("role", "")}</div></div>
            </div>
            <div class='gst-meta-card'>
                <div class='gst-meta-icon'>🏢</div>
                <div><div class='gst-meta-label'>Client / GSTIN</div><div class='gst-meta-value'>{client_name} • {client_gstin}</div></div>
            </div>
            <div class='gst-meta-card'>
                <div class='gst-meta-icon'>🛡️</div>
                <div><div class='gst-meta-label'>System</div><div class='gst-meta-value'>GST JSON Logic Protected</div></div>
            </div>
        </div>
    </div>
    """, unsafe_allow_html=True)


def horizontal_nav():
    pages = [
        ("🏠", "Dashboard"),
        ("🔐", "Client Setup"),
        ("📤", "Upload Center"),
        ("🧾", "IMS Data Viewer"),
        ("🔄", "Reconciliation Workspace"),
        ("✅", "Action Center"),
        ("⚠️", "Risk Center"),
        ("📨", "Vendor Follow-up"),
        ("📊", "Reports & Export"),
        ("🧠", "AI Insight Desk"),
        ("👑", "Admin Panel"),
    ]
    st.markdown("<div class='gst-nav-title'>GST IMS Services</div>", unsafe_allow_html=True)
    st.markdown("<div class='gst-nav-panel'>", unsafe_allow_html=True)
    row1 = pages[:6]
    row2 = pages[6:]
    cols = st.columns(len(row1))
    for col, (icon, page) in zip(cols, row1):
        with col:
            label = f"{icon} {page}"
            if st.button(label, key=f"nav_top_{page}", use_container_width=True):
                st.session_state.page = page
                st.rerun()
    cols = st.columns(len(row2))
    for col, (icon, page) in zip(cols, row2):
        with col:
            label = f"{icon} {page}"
            if st.button(label, key=f"nav_top_{page}", use_container_width=True):
                st.session_state.page = page
                st.rerun()
    st.markdown("</div>", unsafe_allow_html=True)


def hero_dashboard():
    st.markdown("<div class='main-shell'><div class='content-pad'><div class='watermark'>◉</div>", unsafe_allow_html=True)

    col1, col2, col3 = st.columns([3.6, 1.7, 1.5])
    with col1:
        st.markdown(f"""
        <div class='headline'>☀️ Namaste, {st.session_state.get("display_name", "User")}! 🙏</div>
        <div class='main-title'>Reconcile Today. Stay Compliant.<br>Drive Confidence.</div>
        <div class='subcopy'>AI-powered IMS reconciliation with accuracy,<br>automation & actionable insights.</div>
        """, unsafe_allow_html=True)

        b1, b2 = st.columns([1, 1])
        with b1:
            if st.button("🚀 Go to Workspace", use_container_width=True, key="hero_go_workspace"):
                st.session_state.page = "Reconciliation Workspace"
                st.rerun()
        with b2:
            if st.button("📤 Upload IMS Data", use_container_width=True, key="hero_upload_ims"):
                st.session_state.page = "Upload Center"
                st.rerun()

    with col2:
        st.markdown("<div class='shield-center'>🛡️</div>", unsafe_allow_html=True)
    with col3:
        st.markdown("""
        <div class='feature-card'><div class='feature-title'>🧠 Smart Reconciliation</div><div class='feature-desc'>AI-driven matching with high accuracy</div></div>
        <div class='feature-card blue'><div class='feature-title'>🛡️ Risk Detection</div><div class='feature-desc'>Identify mismatches & compliance risks</div></div>
        <div class='feature-card green'><div class='feature-title'>📈 Actionable Insights</div><div class='feature-desc'>Real-time dashboards for better decisions</div></div>
        """, unsafe_allow_html=True)
    st.markdown("</div></div>", unsafe_allow_html=True)


def metric_card(icon, label, value, delta="", bg="#edf4ff", fg="#4d8df7", red=False):
    st.markdown(f"""
    <div class='metric-card'>
        <div class='metric-top'>
            <div class='metric-icon' style='background:{bg};color:{fg};'>{icon}</div>
            <div>
                <div class='metric-label'>{label}</div>
                <div class='metric-value'>{value}</div>
                <div class='metric-delta {"red" if red else ""}'>{delta}</div>
            </div>
        </div>
    </div>
    """, unsafe_allow_html=True)


def page_title(title: str, subtitle: str):
    st.markdown(
        f"""
        <div style="
            background:linear-gradient(135deg,#ffffff,#f3f8ff);
            border:1px solid #cbdced;
            border-radius:22px;
            padding:20px 24px;
            margin:8px 0 18px 0;
            box-shadow:0 12px 28px rgba(7,26,61,0.09);
        ">
            <div class='section-title' style='margin:0;'>{title}</div>
            <div class='section-sub' style='margin:6px 0 0 0;'>{subtitle}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def show_df(df: pd.DataFrame, limit=1000):
    if df is None or df.empty:
        st.info("No data available.")
    else:
        if len(df) > limit:
            st.caption(f"Showing first {limit:,} rows out of {len(df):,}. Use export for full data.")
        st.dataframe(safe_display_df(df, limit), use_container_width=True, hide_index=True)


# =========================================================
# LOGIN
# =========================================================

def login_page():
    """
    Login page — fixed version.
    This function intentionally uses native Streamlit elements for the welcome text,
    so HTML code cannot appear on the login screen.
    GST JSON generation logic is untouched.
    """
    logo_b64 = "iVBORw0KGgoAAAANSUhEUgAAAlgAAAEsCAYAAAAfPc2WAAAAGXRFWHRTb2Z0d2FyZQBBZG9iZSBJbWFnZVJlYWR5ccllPAAAA3JpVFh0WE1MOmNvbS5hZG9iZS54bXAAAAAAADw/eHBhY2tldCBiZWdpbj0i77u/IiBpZD0iVzVNME1wQ2VoaUh6cmVTek5UY3prYzlkIj8+IDx4OnhtcG1ldGEgeG1sbnM6eD0iYWRvYmU6bnM6bWV0YS8iIHg6eG1wdGs9IkFkb2JlIFhNUCBDb3JlIDUuNi1jMTMyIDc5LjE1OTI4NCwgMjAxNi8wNC8xOS0xMzoxMzo0MCAgICAgICAgIj4gPHJkZjpSREYgeG1sbnM6cmRmPSJodHRwOi8vd3d3LnczLm9yZy8xOTk5LzAyLzIyLXJkZi1zeW50YXgtbnMjIj4gPHJkZjpEZXNjcmlwdGlvbiByZGY6YWJvdXQ9IiIgeG1sbnM6eG1wTU09Imh0dHA6Ly9ucy5hZG9iZS5jb20veGFwLzEuMC9tbS8iIHhtbG5zOnN0UmVmPSJodHRwOi8vbnMuYWRvYmUuY29tL3hhcC8xLjAvc1R5cGUvUmVzb3VyY2VSZWYjIiB4bWxuczp4bXA9Imh0dHA6Ly9ucy5hZG9iZS5jb20veGFwLzEuMC8iIHhtcE1NOk9yaWdpbmFsRG9jdW1lbnRJRD0ieG1wLmRpZDo5ZmJkY2I0YS1mY2ZkLTg0NDYtODY5Yi04ODI3OGE5MDgwYjkiIHhtcE1NOkRvY3VtZW50SUQ9InhtcC5kaWQ6NTM1QUJFMzIzMTc0MTFFNzlGRDFGQzlFQTVDQTQ0RjgiIHhtcE1NOkluc3RhbmNlSUQ9InhtcC5paWQ6NTM1QUJFMzEzMTc0MTFFNzlGRDFGQzlFQTVDQTQ0RjgiIHhtcDpDcmVhdG9yVG9vbD0iQWRvYmUgUGhvdG9zaG9wIENDIDIwMTUgKFdpbmRvd3MpIj4gPHhtcE1NOkRlcml2ZWRGcm9tIHN0UmVmOmluc3RhbmNlSUQ9InhtcC5paWQ6ZWQwZWIyNDAtOTgwNS0wMjQzLTliYzQtZWQ1ZjUxNmQ2YWY0IiBzdFJlZjpkb2N1bWVudElEPSJ4bXAuZGlkOkQ3NUZCNDYzQUZBRDExRTU5NUQ2RDNCNDYwQzhBMzI5Ii8+IDwv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"

    left, center, right = st.columns([1.15, 1.1, 1.15])
    with center:
        with st.container(border=True):
            if logo_b64:
                try:
                    st.image(BytesIO(base64.b64decode(logo_b64)), use_container_width=True)
                except Exception:
                    pass

            st.markdown(
                "<h2 style='text-align:center; color:#071b4a; margin-bottom:4px;'>Welcome</h2>",
                unsafe_allow_html=True,
            )
            st.markdown(
                "<div style='text-align:center; color:#566982; font-size:14px; margin-bottom:18px;'>"
                "GST IMS Reconciliation Workspace<br>Secure compliance access panel"
                "</div>",
                unsafe_allow_html=True,
            )

            username = st.text_input("User ID", placeholder="Enter your User ID")
            password = st.text_input("Password", type="password", placeholder="Enter password")

            if st.button("🔐 Login Securely", use_container_width=True):
                user = USER_MASTER.get(username)
                if user and password == user["password"]:
                    st.session_state.logged_in = True
                    st.session_state.username = username
                    st.session_state.role = user["role"]
                    st.session_state.display_name = user["name"]
                    log_event("Login", "Successful login")
                    load_user_state()
                    st.rerun()
                else:
                    st.error("Invalid User ID or Password. Please check case-sensitive credentials.")

            st.markdown(
                "<div style='text-align:center; font-size:12px; color:#6b7d96; "
                "margin-top:16px; border-top:1px solid #e3edf8; padding-top:13px; font-weight:700;'>"
                "© @BAJRABHANU • INALSA Home Appliances • IMS Recon Workspace"
                "</div>",
                unsafe_allow_html=True,
            )


# =========================================================
# PAGES
# =========================================================

def dashboard_page():
    v10_command_center()
    v10_help_tooltips()
    hero_dashboard()
    v9_saleable_kpis()
    v9_home_modules()
    v9_json_readiness_panel()

    p, ims, recon = st.session_state.purchase_df, st.session_state.ims_df, st.session_state.recon_df
    total_itc = float(ims["total_tax"].sum()) if not ims.empty else 0
    matched = int((recon["mismatch_type"] == "Matched").sum()) if not recon.empty else 0
    pending = int((recon["recommended_action"] == "Pending").sum()) if not recon.empty else 0
    accepted = int((recon["recommended_action"] == "Accepted").sum()) if not recon.empty else 0
    highrisk = int(recon["risk_level"].isin(["High", "Critical"]).sum()) if not recon.empty else 0

    cols = st.columns(5)
    with cols[0]: metric_card("📚", "Purchase Rows", f"{len(p):,}", "Books data", "#ffefe2", "#ec8b24")
    with cols[1]: metric_card("📥", "IMS Rows", f"{len(ims):,}", st.session_state.get("ims_source",""), "#ecfaef", "#27a857")
    with cols[2]: metric_card("✅", "Matched", f"{matched:,}", "Accepted ready", "#edf4ff", "#4d8df7")
    with cols[3]: metric_card("📌", "Pending", f"{pending:,}", "Needs action", "#f4eefe", "#8b6cf7")
    with cols[4]: metric_card("⚠️", "High Risk", f"{highrisk:,}", "Review required", "#fff0ed", "#e1563a", True)

    c1, c2, c3 = st.columns([2, 2, 1.3])
    with c1:
        st.markdown("<div class='panel'><div class='panel-title'>Reconciliation Summary</div>", unsafe_allow_html=True)
        if recon.empty:
            st.info("Upload data and start reconciliation.")
        else:
            show_df(recon_summary(recon), 20)
        st.markdown("</div>", unsafe_allow_html=True)

    with c2:
        st.markdown("<div class='panel'><div class='panel-title'>Top Mismatch Reasons</div>", unsafe_allow_html=True)
        if recon.empty:
            st.info("No reconciliation data.")
        else:
            summary = recon["mismatch_type"].value_counts().reset_index()
            summary.columns = ["Mismatch Type", "Count"]
            show_df(summary, 20)
        st.markdown("</div>", unsafe_allow_html=True)

    with c3:
        st.markdown("<div class='small-card'><div class='panel-title'>Compliance Health</div>", unsafe_allow_html=True)
        health = 0 if recon.empty else round((matched / max(len(recon), 1)) * 100, 2)
        st.markdown(f"<div style='font-size:48px;font-weight:900;color:#112244;text-align:center;margin:20px 0;'>{health}%</div>", unsafe_allow_html=True)
        st.progress(int(min(100, health)))
        st.caption("Based on matched records against total reconciliation records.")
        st.markdown("</div>", unsafe_allow_html=True)

        st.markdown("<div class='small-card'><div class='panel-title'>AI Insight</div><br>", unsafe_allow_html=True)
        if recon.empty:
            st.write("Upload data to generate AI-like insights.")
        else:
            st.write(generate_ai_insight())
        st.markdown("</div>", unsafe_allow_html=True)


def client_setup_page():
    page_title("Client Setup", "Set client GSTIN, return period and review controls.")
    with st.form("client_form"):
        c1, c2, c3 = st.columns(3)
        with c1:
            client_name = st.text_input("Client Name", st.session_state.client_name)
        with c2:
            client_gstin = st.text_input("Client GSTIN", st.session_state.client_gstin).upper()
        with c3:
            return_period = st.text_input("Return Period", st.session_state.return_period)

        submitted = st.form_submit_button("💾 Save Client Setup", use_container_width=True)
        if submitted:
            st.session_state.client_name = client_name
            st.session_state.client_gstin = normalize_gstin(client_gstin)
            st.session_state.return_period = return_period
            save_user_state(["client_name", "client_gstin", "return_period"])
            log_event("Client Setup", "Client details saved")
            st.success("Client setup saved.")

    if st.session_state.client_gstin:
        if validate_gstin(st.session_state.client_gstin):
            st.success("GSTIN format is valid.")
        else:
            st.warning("GSTIN format appears invalid.")


def upload_center_page():
    v10_quality_dashboard()
    v10_help_tooltips()
    v9_help_box('Upload Guidance', 'Upload Purchase Register and GST IMS JSON. Review quality checks before reconciliation to avoid wrong action selection.')
    page_title("Upload Center", "Upload Purchase Register and GST IMS JSON only. The GST utility is now built inside this app.")

    st.markdown("### Step 1 — Upload source files")
    c1, c2 = st.columns(2)

    with c1:
        st.markdown("<div class='panel'><div class='panel-title'>📚 Purchase Register</div>", unsafe_allow_html=True)
        file = st.file_uploader("Upload Purchase Register", type=["xlsx", "xls", "csv"], key="purchase_upload")
        if file and st.button("Process Purchase Register", use_container_width=True):
            try:
                df = read_purchase_file(file)
                st.session_state.purchase_df = df
                st.session_state.recon_df = pd.DataFrame()
                st.session_state.action_df = pd.DataFrame()
                st.session_state.final_json_bytes = b""
                st.session_state.final_json_summary = pd.DataFrame()
                save_user_state(["purchase_df", "recon_df", "action_df", "final_json_bytes", "final_json_summary"])
                log_event("Upload", f"Purchase Register uploaded: {len(df):,} rows")
                st.success(f"Purchase Register processed: {len(df):,} rows.")
            except Exception as e:
                st.error(f"Unable to process Purchase Register: {e}")
        st.markdown("</div>", unsafe_allow_html=True)

    with c2:
        st.markdown("<div class='panel'><div class='panel-title'>🧬 GST IMS JSON</div>", unsafe_allow_html=True)
        json_file = st.file_uploader("Upload IMS JSON downloaded from GST Portal", type=["json"], key="ims_json_upload")
        if json_file and st.button("Process IMS JSON", use_container_width=True):
            try:
                raw_bytes = json_file.getvalue()
                df, records, data = parse_ims_json_bytes(raw_bytes)
                st.session_state.ims_df = df
                st.session_state.ims_source = "IMS JSON"
                st.session_state.ims_json_records = records
                st.session_state.ims_json_data = data
                st.session_state.ims_json_bytes = raw_bytes
                st.session_state.final_json_bytes = b""
                st.session_state.final_json_summary = pd.DataFrame()
                st.session_state.recon_df = pd.DataFrame()
                st.session_state.action_df = pd.DataFrame()
                if isinstance(data, dict) and data.get("rtin"):
                    st.session_state.client_gstin = normalize_gstin(data.get("rtin"))
                save_user_state(["ims_df", "ims_source", "ims_json_records", "ims_json_data", "ims_json_bytes", "final_json_bytes", "final_json_summary", "client_gstin", "recon_df", "action_df"])
                log_event("Upload", f"IMS JSON uploaded: {len(df):,} rows")
                st.success(f"IMS JSON processed: {len(df):,} rows.")
            except Exception as e:
                st.error(f"Unable to process IMS JSON: {e}")
        st.markdown("</div>", unsafe_allow_html=True)

    st.markdown("---")
    st.markdown("### Step 2 — Reconcile inside the app")
    st.info("Final process: Purchase Register + IMS JSON → in-app reconciliation → in-app action/remarks → final GST upload JSON. No .xlsm utility is required.")

    c4, c5, c6 = st.columns(3)
    with c4:
        st.session_state.amount_tolerance = st.number_input("Amount tolerance ₹", min_value=0.0, value=float(st.session_state.amount_tolerance), step=1.0, key="upload_amount_tol")
    with c5:
        st.session_state.date_tolerance = st.number_input("Date tolerance days", min_value=0, value=int(st.session_state.date_tolerance), step=1, key="upload_date_tol")
    with c6:
        st.session_state.include_amendments = st.checkbox("Include amendment records", value=bool(st.session_state.include_amendments), key="upload_include_amend")

    ready_reco = not st.session_state.purchase_df.empty and not st.session_state.ims_df.empty
    if st.button("🚀 Run Reconciliation from JSON", type="primary", use_container_width=True, disabled=not ready_reco):
        with st.spinner("Reconciling Purchase Register with IMS JSON..."):
            recon = calculate_recon(
                st.session_state.purchase_df,
                st.session_state.ims_df,
                st.session_state.amount_tolerance,
                st.session_state.date_tolerance,
                st.session_state.include_amendments,
            )
            # As per final business rule:
            # Matched = Accepted, Unmatched/Mismatch = Pending. User may manually change later.
            recon["final_user_action"] = recon["mismatch_type"].apply(lambda x: "Accepted" if x == "Matched" else "Pending")
            recon["json_action_code"] = recon["final_user_action"].apply(action_label_to_gst_code)
            st.session_state.recon_df = recon
            st.session_state.action_df = recon.copy()
            st.session_state.final_json_bytes = b""
            st.session_state.final_json_summary = pd.DataFrame()
            save_user_state(["recon_df", "action_df", "final_json_bytes", "final_json_summary"])
            log_event("Reconciliation", f"JSON reconciliation completed: {len(recon):,} rows")
        st.success(f"Reconciliation completed: {len(st.session_state.recon_df):,} rows. Now go to Action Center to review/edit actions and remarks.")

    st.markdown("---")
    page_title("Data Health Check", "Upload status and quality summary.")
    h1, h2, h3, h4 = st.columns(4)
    p = st.session_state.purchase_df
    ims = st.session_state.ims_df
    with h1: metric_card("📚", "Purchase Rows", f"{len(p):,}", "", "#ffefe2", "#ec8b24")
    with h2: metric_card("📥", "IMS JSON Rows", f"{len(ims):,}", st.session_state.ims_source, "#ecfaef", "#27a857")
    with h3:
        invalid = int((~p["gstin_valid"]).sum()) if not p.empty and "gstin_valid" in p else 0
        metric_card("⚠️", "Purchase Invalid GSTIN", f"{invalid:,}", "", "#fff0ed", "#e1563a", True)
    with h4:
        invalid = int((~ims["gstin_valid"]).sum()) if not ims.empty and "gstin_valid" in ims else 0
        metric_card("🛡️", "IMS Invalid GSTIN", f"{invalid:,}", "", "#edf4ff", "#4d8df7")

    tabs = st.tabs(["Purchase Preview", "IMS JSON Preview", "Reconciliation Preview"])
    with tabs[0]:
        show_df(st.session_state.purchase_df.head(100))
    with tabs[1]:
        show_df(st.session_state.ims_df.head(100))
    with tabs[2]:
        show_df(st.session_state.recon_df.head(100))

    st.markdown("### V10.1 Upload Validation — Taxable Value and Tax Head Wise")
    vtab1, vtab2, vtab3 = st.tabs(["Purchase Quality", "IMS Quality", "Duplicate Report"])
    with vtab1:
        show_df(upload_quality_summary(st.session_state.purchase_df, "Purchase Register"), 50)
    with vtab2:
        show_df(upload_quality_summary(st.session_state.ims_df, "IMS JSON"), 50)
        if st.session_state.ims_json_data:
            st.markdown("**IMS JSON section count**")
            show_df(ims_json_section_counts(st.session_state.ims_json_data), 50)
    with vtab3:
        dup = pd.concat([duplicate_report(st.session_state.purchase_df, "Purchase Register"), duplicate_report(st.session_state.ims_df, "IMS JSON")], ignore_index=True)
        show_df(dup, 500)

def ims_data_viewer_page():
    page_title("IMS Data Viewer", "Review uploaded and standardized data before reconciliation.")
    tabs = st.tabs(["Purchase Register", "IMS Combined", "IMS Sheet Summary"])
    with tabs[0]:
        show_df(st.session_state.purchase_df)
    with tabs[1]:
        show_df(st.session_state.ims_df)
    with tabs[2]:
        ims = st.session_state.ims_df
        if ims.empty:
            st.info("No IMS data uploaded.")
        else:
            show_df(ims.groupby(["ims_sheet", "document_type"], dropna=False).agg(
                Records=("document_no", "size"),
                Taxable=("taxable_value", "sum"),
                Tax=("total_tax", "sum"),
            ).reset_index().round(2))


def reconciliation_page():
    v10_reco_control_room()
    v9_help_box('Reconciliation Control Room', 'Run reconciliation only after both Purchase Register and IMS JSON are uploaded and validated. Review mismatch categories carefully.')
    page_title("Reconciliation Workspace", "Run IMS reconciliation only when you click Start Reconciliation.")

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        st.session_state.amount_tolerance = st.number_input("Amount Tolerance ₹", min_value=0.0, value=float(st.session_state.amount_tolerance), step=1.0)
    with c2:
        st.session_state.date_tolerance = st.number_input("Date Tolerance Days", min_value=0, value=int(st.session_state.date_tolerance), step=1)
    with c3:
        st.session_state.include_amendments = st.checkbox("Include Amendment Sheets", value=bool(st.session_state.include_amendments))
    with c4:
        st.session_state.use_fuzzy = st.checkbox("Fuzzy Matching Later", value=False, disabled=True)

    ready = not st.session_state.purchase_df.empty and not st.session_state.ims_df.empty
    if not ready:
        st.warning("Upload Purchase Register and IMS data first.")

    if st.button("🚀 Start IMS Reconciliation", type="primary", use_container_width=True, disabled=not ready):
        with st.spinner("Running IMS reconciliation..."):
            recon = calculate_recon(
                st.session_state.purchase_df,
                st.session_state.ims_df,
                st.session_state.amount_tolerance,
                st.session_state.date_tolerance,
                st.session_state.include_amendments,
            )
            st.session_state.recon_df = recon
            st.session_state.action_df = recon.copy()
            save_user_state(["recon_df", "action_df"])
            log_event("Reconciliation", f"Reconciliation completed: {len(recon):,} rows")
        st.success(f"Reconciliation completed: {len(st.session_state.recon_df):,} rows.")

    recon = st.session_state.recon_df
    if not recon.empty:
        st.markdown("---")
        tabs = st.tabs(["All Results", "Matched", "Value Mismatch", "Only in IMS", "Only in Purchase", "High Risk"])
        with tabs[0]: show_df(recon)
        with tabs[1]: show_df(recon[recon["mismatch_type"] == "Matched"])
        with tabs[2]: show_df(recon[recon["mismatch_type"].isin(["Value / Tax Mismatch", "Tax Head Mismatch", "Value and Date Mismatch"])])
        with tabs[3]: show_df(recon[recon["mismatch_type"] == "Only in IMS"])
        with tabs[4]: show_df(recon[recon["mismatch_type"] == "Only in Purchase Register"])
        with tabs[5]: show_df(recon[recon["risk_level"].isin(["High", "Critical"])])


def action_center_page():
    v10_action_header()
    v9_help_box('Action Center Guidance', 'Use filters, bulk actions and remarks to finalize invoice-wise IMS action before generating GST upload JSON.')
    page_title("IMS Action Center", "Filter, bulk-update and finalize invoice-wise action/remarks before GST JSON generation.")
    df = st.session_state.action_df
    if df.empty:
        st.info("Run reconciliation first.")
        return

    # Ensure required columns exist.
    df = df.copy()
    if "final_user_action" not in df.columns:
        df["final_user_action"] = df.get("recommended_action", "Pending")
    if "user_remarks" not in df.columns:
        df["user_remarks"] = ""
    if "json_action_code" not in df.columns:
        df["json_action_code"] = df["final_user_action"].apply(action_label_to_gst_code)

    c1, c2, c3, c4, c5 = st.columns(5)
    with c1: metric_card("✅", "Accepted", f"{(df['final_user_action']=='Accepted').sum():,}", "JSON A", "#ecfaef", "#27a857")
    with c2: metric_card("📌", "Pending", f"{(df['final_user_action']=='Pending').sum():,}", "JSON P", "#fff7ed", "#f4a62a")
    with c3: metric_card("🚫", "Rejected", f"{(df['final_user_action']=='Rejected').sum():,}", "JSON R", "#fff0ed", "#e1563a", True)
    with c4: metric_card("🕘", "No Action", f"{(df['final_user_action']=='No Action').sum():,}", "converted safely", "#edf4ff", "#4d8df7")
    with c5: metric_card("⚠️", "High Risk", f"{df['risk_level'].isin(['High','Critical']).sum():,}" if 'risk_level' in df else "0", "review first", "#f4eefe", "#8b6cf7")

    st.markdown("### Filters")
    f1, f2, f3, f4 = st.columns(4)
    with f1:
        action_filter = st.selectbox("Final Action", ["All"] + ACTION_VALUES, key="action_filter_v7")
    with f2:
        mismatch_options = ["All"] + sorted(df.get("mismatch_type", pd.Series(dtype=str)).dropna().astype(str).unique().tolist())
        mismatch_filter = st.selectbox("Mismatch Type", mismatch_options, key="mismatch_filter_v7")
    with f3:
        risk_options = ["All"] + sorted(df.get("risk_level", pd.Series(dtype=str)).dropna().astype(str).unique().tolist())
        risk_filter = st.selectbox("Risk Level", risk_options, key="risk_filter_v7")
    with f4:
        search_text = st.text_input("Search GSTIN / Invoice / Vendor", key="action_search_v7")

    view = df.copy()
    if action_filter != "All":
        view = view[view["final_user_action"].astype(str).eq(action_filter)]
    if mismatch_filter != "All" and "mismatch_type" in view:
        view = view[view["mismatch_type"].astype(str).eq(mismatch_filter)]
    if risk_filter != "All" and "risk_level" in view:
        view = view[view["risk_level"].astype(str).eq(risk_filter)]
    if search_text:
        stext = search_text.lower().strip()
        combined = view[[c for c in ["supplier_gstin", "supplier_name", "document_no", "document_norm"] if c in view.columns]].astype(str).agg(" ".join, axis=1).str.lower()
        view = view[combined.str.contains(re.escape(stext), na=False)]

    st.caption(f"Showing {len(view):,} rows out of {len(df):,}. Tick Select for bulk action, or directly edit Final User Action / Remarks.")

    view_cols = [
        "supplier_gstin", "supplier_name", "document_type", "document_no", "document_date",
        "taxable_value_ims", "total_tax_ims", "taxable_value_diff", "total_tax_diff",
        "mismatch_type", "match_level", "risk_level", "confidence_score", "recommended_action",
        "final_user_action", "json_action_code", "reason", "user_remarks"
    ]
    exist_cols = [c for c in view_cols if c in view.columns]
    edit_df = view[exist_cols].copy()
    edit_df.insert(0, "_row_id", edit_df.index)
    edit_df.insert(0, "Select", False)

    edited = st.data_editor(
        edit_df,
        use_container_width=True,
        hide_index=True,
        num_rows="fixed",
        column_config={
            "Select": st.column_config.CheckboxColumn("Select"),
            "_row_id": st.column_config.NumberColumn("Row ID", disabled=True),
            "final_user_action": st.column_config.SelectboxColumn("Final User Action", options=ACTION_VALUES),
            "user_remarks": st.column_config.TextColumn("User Remarks"),
        },
        disabled=[c for c in edit_df.columns if c not in ["Select", "final_user_action", "user_remarks"]],
        key="action_editor_v7",
    )

    st.markdown("### Bulk Action for Selected Rows")
    b1, b2, b3 = st.columns([1, 1, 2])
    with b1:
        bulk_action = st.selectbox("Bulk final action", ACTION_VALUES, index=ACTION_VALUES.index("Pending"), key="bulk_action_v7")
    with b2:
        apply_bulk = st.button("Apply to selected", use_container_width=True)
    with b3:
        bulk_remarks = st.text_input("Optional common remarks", key="bulk_remarks_v7")

    updated = df.copy()
    # Save direct edits first.
    for _, erow in edited.iterrows():
        rid = int(erow["_row_id"])
        if rid in updated.index:
            updated.loc[rid, "final_user_action"] = erow.get("final_user_action", updated.loc[rid, "final_user_action"])
            updated.loc[rid, "user_remarks"] = erow.get("user_remarks", updated.loc[rid, "user_remarks"])

    if apply_bulk:
        selected_ids = edited.loc[edited["Select"] == True, "_row_id"].astype(int).tolist()
        if not selected_ids:
            st.warning("Please tick Select for at least one row.")
        else:
            updated.loc[selected_ids, "final_user_action"] = bulk_action
            if bulk_remarks.strip():
                updated.loc[selected_ids, "user_remarks"] = bulk_remarks.strip()
            updated["json_action_code"] = updated["final_user_action"].apply(action_label_to_gst_code)
            st.session_state.action_df = updated
            st.session_state.final_json_bytes = b""
            st.session_state.final_json_summary = pd.DataFrame()
            save_user_state(["action_df", "final_json_bytes", "final_json_summary"])
            log_event("Action Center", f"Bulk action applied to {len(selected_ids)} rows: {bulk_action}")
            st.success(f"Bulk action applied to {len(selected_ids):,} rows.")
            st.rerun()

    if st.button("💾 Save Final Actions / Remarks", type="primary", use_container_width=True):
        updated["json_action_code"] = updated["final_user_action"].apply(action_label_to_gst_code)
        st.session_state.action_df = updated
        st.session_state.final_json_bytes = b""
        st.session_state.final_json_summary = pd.DataFrame()
        save_user_state(["action_df", "final_json_bytes", "final_json_summary"])
        log_event("Action Center", "Final user actions and remarks updated")
        st.success("Final actions and remarks saved. Now go to Reports & Export for final review and GST upload JSON generation.")


def risk_center_page():
    page_title("Risk & Exception Center", "Focused review of high-risk IMS records.")
    df = st.session_state.recon_df
    if df.empty:
        st.info("Run reconciliation first.")
        return
    high = df[df["risk_level"].isin(["High", "Critical"])]
    show_df(high)
    if not high.empty:
        st.download_button(
            "Download High Risk Report",
            data=to_excel_bytes({"High Risk": high}),
            file_name="IMS_High_Risk_Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )


def vendor_followup_page():
    v10_management_summary()
    page_title("Vendor Follow-up", "Vendor-wise pending, mismatch and follow-up list.")
    df = st.session_state.recon_df
    if df.empty:
        st.info("Run reconciliation first.")
        return
    follow = df[df["vendor_followup_required"] == True].copy()
    if follow.empty:
        st.success("No vendor follow-up cases identified.")
        return

    summary = follow.groupby(["supplier_gstin", "supplier_name"], dropna=False).agg(
        Exceptions=("mismatch_type", "size"),
        Taxable_Diff=("taxable_value_diff", "sum"),
        Tax_Diff=("total_tax_diff", "sum"),
        High_Risk=("risk_level", lambda x: int(x.isin(["High", "Critical"]).sum())),
    ).reset_index().round(2)
    show_df(summary)

    selected = st.selectbox("Generate email draft for vendor", summary["supplier_gstin"].astype(str).tolist())
    part = follow[follow["supplier_gstin"].astype(str) == str(selected)]
    if not part.empty:
        vendor_name = str(part["supplier_name"].dropna().iloc[0]) if "supplier_name" in part else ""
        email = f"""Dear {vendor_name or 'Vendor'},

During our GST IMS reconciliation for {st.session_state.get('return_period', '')}, the following documents require clarification:

Total exception cases: {len(part)}
Total tax difference / exposure: ₹{float(part['total_tax_diff'].abs().sum()):,.2f}

Request you to kindly review the invoices/credit notes/debit notes and share clarification or corrective action at the earliest.

Regards,
{st.session_state.get('display_name', '')}
"""
        st.text_area("Vendor Email Draft", email, height=220)


def reports_page():
    v10_final_json_review_ui()
    v10_management_summary()
    v9_json_readiness_panel()
    v9_report_cards()
    page_title("Reports & Final GST Upload JSON", "Final review, workpaper export and GST portal upload JSON generation.")
    p, ims, recon, action = st.session_state.purchase_df, st.session_state.ims_df, st.session_state.recon_df, st.session_state.action_df

    st.markdown("### Final Review Before GST JSON")
    if action.empty:
        st.info("No final action report available yet. Run reconciliation and save actions first.")
    else:
        review = final_json_review_table(action)
        r1, r2 = st.columns([1.2, 2])
        with r1:
            show_df(review, 20)
        with r2:
            st.info("GST JSON generation logic is the stable V6 amendment-safe logic. It remains unchanged in V7. Only your final action values are used to update GST action codes.")
            risky = action[action.get("risk_level", pd.Series(dtype=str)).isin(["High", "Critical"])] if "risk_level" in action else pd.DataFrame()
            if not risky.empty:
                st.warning(f"{len(risky):,} high/critical risk rows exist. Review them before generating JSON.")

    c1, c2 = st.columns([1, 1])
    with c1:
        st.markdown("<div class='panel'><div class='panel-title'>📊 Complete Excel Workpaper</div>", unsafe_allow_html=True)
        sheets = split_report_sheets(p, ims, recon, action)
        st.download_button(
            "📥 Download Complete IMS Workpaper",
            data=to_excel_bytes(sheets),
            file_name=f"IMS_JSON_Recon_Workpaper_V7_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
        st.markdown("</div>", unsafe_allow_html=True)

    with c2:
        st.markdown("<div class='panel'><div class='panel-title'>🧬 Final GST Portal Upload JSON</div>", unsafe_allow_html=True)
        st.caption("Stable V6 GST schema: rtin + reqtyp SAVE + invdata. Amendment-safe fields are preserved from the original GST IMS JSON.")
        if not st.session_state.ims_json_data:
            st.warning("Please upload/process IMS JSON first.")
        elif action.empty:
            st.warning("Please run reconciliation and save final actions first.")
        else:
            source_counts = ims_json_section_counts(st.session_state.ims_json_data)
            if not source_counts.empty:
                st.markdown("**Uploaded IMS JSON section count**")
                st.dataframe(source_counts, use_container_width=True, hide_index=True)

            confirm = st.checkbox("I have reviewed final actions and want to generate GST upload JSON", key="confirm_json_v7")
            if st.button("⚙️ Generate Final GST Upload JSON", type="primary", use_container_width=True, disabled=not confirm):
                try:
                    # DO NOT CHANGE: stable V6 amendment-safe generator.
                    json_bytes, summary = generate_gst_upload_json_from_final_actions(st.session_state.ims_json_data, st.session_state.action_df)
                    st.session_state.final_json_bytes = json_bytes
                    st.session_state.final_json_summary = summary
                    save_user_state(["final_json_bytes", "final_json_summary"])
                    log_event("GST JSON", f"Final GST upload JSON generated: {len(summary):,} records")
                    st.success(f"Final GST upload JSON generated. Records included: {len(summary):,}")
                except Exception as e:
                    st.error(f"Unable to generate final JSON: {e}")

            if st.session_state.final_json_bytes:
                action_counts = generated_json_action_counts(st.session_state.final_json_bytes)
                if not action_counts.empty:
                    st.markdown("**Generated GST upload JSON action count**")
                    st.dataframe(action_counts, use_container_width=True, hide_index=True)
                st.download_button(
                    "⬇️ Download Final GST Portal Upload JSON",
                    data=st.session_state.final_json_bytes,
                    file_name=f"IMS_Final_Action_Upload_{datetime.now().strftime('%Y%m%d_%H%M')}.json",
                    mime="application/json",
                    use_container_width=True,
                )
        st.markdown("</div>", unsafe_allow_html=True)

    st.markdown("---")
    st.markdown("### Upload & Duplicate Validation")
    t1, t2, t3 = st.tabs(["Purchase Quality", "IMS Quality", "Duplicates"])
    with t1:
        show_df(upload_quality_summary(p, "Purchase Register"), 50)
    with t2:
        show_df(upload_quality_summary(ims, "IMS JSON"), 50)
    with t3:
        dup = pd.concat([duplicate_report(p, "Purchase Register"), duplicate_report(ims, "IMS JSON")], ignore_index=True)
        show_df(dup, 500)

    if isinstance(st.session_state.final_json_summary, pd.DataFrame) and not st.session_state.final_json_summary.empty:
        st.markdown("### Final JSON Update Summary")
        show_df(st.session_state.final_json_summary.groupby(["Section", "Final Action", "GST JSON Action Code"], dropna=False).size().reset_index(name="Records"), 100)


def ai_insight_page():
    v10_management_summary()
    v10_help_tooltips()
    page_title("AI Insight Desk", "Rule-based smart GST IMS insights without external API.")
    st.markdown("<div class='panel'>", unsafe_allow_html=True)
    st.write(generate_ai_insight(long=True))
    st.markdown("</div>", unsafe_allow_html=True)

    q = st.text_input("Ask IMS Insight Desk", placeholder="Example: Which invoices should I accept?")
    if q:
        ql = q.lower()
        df = st.session_state.recon_df
        if df.empty:
            st.info("Run reconciliation first.")
        elif "accept" in ql:
            show_df(df[df["recommended_action"] == "Accepted"])
        elif "pending" in ql:
            show_df(df[df["recommended_action"] == "Pending"])
        elif "reject" in ql:
            show_df(df[df["recommended_action"] == "Rejected"])
        elif "risk" in ql:
            show_df(df[df["risk_level"].isin(["High", "Critical"])])
        elif "vendor" in ql:
            vendor_followup_page()
        else:
            st.write(generate_ai_insight(long=True))


def generate_ai_insight(long: bool = False) -> str:
    df = st.session_state.recon_df
    if df.empty:
        return "No reconciliation has been run yet. Upload Purchase Register and IMS data, then click Start IMS Reconciliation."
    total = len(df)
    matched = int((df["mismatch_type"] == "Matched").sum())
    pending = int((df["recommended_action"] == "Pending").sum())
    high = int(df["risk_level"].isin(["High", "Critical"]).sum())
    only_ims = int((df["mismatch_type"] == "Only in IMS").sum())
    only_purchase = int((df["mismatch_type"] == "Only in Purchase Register").sum())
    health = round((matched / max(total, 1)) * 100, 2)
    base = (
        f"IMS Health Score is {health}%. Out of {total:,} reconciled documents, "
        f"{matched:,} are clean matches, {pending:,} should be kept pending/reviewed, "
        f"and {high:,} are high/critical risk cases."
    )
    if not long:
        return base
    return (
        base
        + f"\n\nKey observations:\n"
        + f"- {only_ims:,} documents are appearing in IMS but not in Purchase Register.\n"
        + f"- {only_purchase:,} documents are booked in Purchase Register but not found in IMS.\n"
        + "- Accept matched invoices first, keep value/tax mismatch cases pending, and send vendor follow-up for books-only cases.\n"
        + "- Review credit notes and amendment sheets separately before final upload through final GST upload JSON."
    )


def admin_page():
    page_title("Admin Panel", "Data reset, audit log and user controls.")

    st.markdown("### Audit Log")
    show_df(load_audit("" if st.session_state.role == "Main Admin" else st.session_state.username), 500)

    st.markdown("### Reset Data")
    st.warning("Reset will delete saved data for the current logged-in user.")
    confirm = st.text_input("Type DELETE to confirm reset")
    if st.button("🗑️ Reset My Complete Data", use_container_width=True):
        if confirm == "DELETE":
            username = st.session_state.username
            db_delete_user(username)
            log_event("Reset", "User data reset")
            for key in ["purchase_df", "ims_df", "recon_df", "action_df"]:
                st.session_state[key] = pd.DataFrame()
            st.session_state.ims_source = ""
            st.success("Your saved data has been deleted.")
        else:
            st.error("Please type DELETE exactly.")




# =========================================================
# V9 SALEABLE UI HELPERS — UI ONLY, NO GST JSON LOGIC CHANGE
# =========================================================

def v9_status_bool(value) -> bool:
    try:
        if isinstance(value, pd.DataFrame):
            return not value.empty
        return bool(value)
    except Exception:
        return False


def v9_workflow_tracker():
    p_ready = v9_status_bool(st.session_state.get("purchase_df", pd.DataFrame()))
    ims_ready = v9_status_bool(st.session_state.get("ims_df", pd.DataFrame()))
    recon_ready = v9_status_bool(st.session_state.get("recon_df", pd.DataFrame()))
    action_ready = v9_status_bool(st.session_state.get("action_df", pd.DataFrame()))
    client_ready = bool(st.session_state.get("client_gstin", ""))

    steps = [
        ("01", "Client Setup", client_ready, st.session_state.get("page") == "Client Setup"),
        ("02", "Upload Purchase", p_ready, st.session_state.get("page") == "Upload Center"),
        ("03", "Upload IMS JSON", ims_ready, st.session_state.get("page") == "Upload Center"),
        ("04", "Reconciliation", recon_ready, st.session_state.get("page") == "Reconciliation Workspace"),
        ("05", "Action Review", action_ready, st.session_state.get("page") == "Action Center"),
        ("06", "GST JSON", False, st.session_state.get("page") == "Reports & Export"),
    ]

    html = ["<div class='v9-workflow'><div class='v9-workflow-title'>🚀 Guided IMS Workflow</div><div class='v9-step-grid'>"]
    for num, label, done, active in steps:
        cls = "done" if done else ("active" if active else "pending")
        status = "Completed" if done else ("Action Required" if active else "Pending")
        icon = "✓" if done else num
        html.append(f"""
        <div class='v9-step {cls}'>
            <div class='v9-step-num'>{icon}</div>
            <div class='v9-step-label'>{label}</div>
            <div class='v9-step-status'>{status}</div>
        </div>
        """)
    html.append("</div></div>")
    st.markdown("".join(html), unsafe_allow_html=True)


def v9_saleable_kpis():
    p = st.session_state.get("purchase_df", pd.DataFrame())
    ims = st.session_state.get("ims_df", pd.DataFrame())
    recon = st.session_state.get("recon_df", pd.DataFrame())
    action = st.session_state.get("action_df", pd.DataFrame())

    matched = int((recon.get("mismatch_type", pd.Series(dtype=str)) == "Matched").sum()) if isinstance(recon, pd.DataFrame) and not recon.empty else 0
    pending = int((action.get("final_user_action", pd.Series(dtype=str)) == "Pending").sum()) if isinstance(action, pd.DataFrame) and not action.empty else 0
    accepted = int((action.get("final_user_action", pd.Series(dtype=str)) == "Accepted").sum()) if isinstance(action, pd.DataFrame) and not action.empty else 0
    highrisk = int(action.get("risk_level", pd.Series(dtype=str)).isin(["High", "Critical"]).sum()) if isinstance(action, pd.DataFrame) and not action.empty and "risk_level" in action else 0

    st.markdown(f"""
    <div class='v9-kpi-strip'>
        <div class='v9-kpi'><div class='v9-kpi-label'>Purchase Register</div><div class='v9-kpi-value'>{len(p):,}</div><div class='v9-kpi-note'>Books records loaded</div></div>
        <div class='v9-kpi'><div class='v9-kpi-label'>IMS JSON</div><div class='v9-kpi-value'>{len(ims):,}</div><div class='v9-kpi-note'>Portal records loaded</div></div>
        <div class='v9-kpi'><div class='v9-kpi-label'>Accepted / Matched</div><div class='v9-kpi-value'>{accepted:,}</div><div class='v9-kpi-note'>{matched:,} system matches</div></div>
        <div class='v9-kpi'><div class='v9-kpi-label'>Pending / Risk</div><div class='v9-kpi-value'>{pending:,}</div><div class='v9-kpi-note'>{highrisk:,} high-risk cases</div></div>
    </div>
    """, unsafe_allow_html=True)


def v9_home_modules():
    st.markdown("""
    <div class='v9-module-grid'>
        <div class='v9-module-card'>
            <div class='v9-module-icon'>📤</div>
            <div class='v9-module-title'>Smart Upload Center</div>
            <div class='v9-module-desc'>Upload Purchase Register and GST IMS JSON with quality checks, duplicate review and section-wise visibility.</div>
            <div class='v9-module-badge'>Upload → Validate</div>
        </div>
        <div class='v9-module-card'>
            <div class='v9-module-icon'>🔄</div>
            <div class='v9-module-title'>Reconciliation Control Room</div>
            <div class='v9-module-desc'>Review matched, pending, mismatch, duplicate and risk cases in a structured and user-friendly flow.</div>
            <div class='v9-module-badge'>Recon → Review</div>
        </div>
        <div class='v9-module-card'>
            <div class='v9-module-icon'>✅</div>
            <div class='v9-module-title'>Action Center</div>
            <div class='v9-module-desc'>Use filters, manual actions, remarks and bulk review to finalize invoice-wise IMS actions.</div>
            <div class='v9-module-badge'>Action → Finalize</div>
        </div>
        <div class='v9-module-card'>
            <div class='v9-module-icon'>⚠️</div>
            <div class='v9-module-title'>Risk Desk</div>
            <div class='v9-module-desc'>Identify value mismatches, tax-head mismatches, only-in-IMS cases and vendor follow-up requirements.</div>
            <div class='v9-module-badge'>Risk → Resolve</div>
        </div>
        <div class='v9-module-card'>
            <div class='v9-module-icon'>📊</div>
            <div class='v9-module-title'>Professional Reports</div>
            <div class='v9-module-desc'>Download Excel workpapers with summary, final action, mismatch, pending, risk and audit reports.</div>
            <div class='v9-module-badge'>Report → Export</div>
        </div>
        <div class='v9-module-card'>
            <div class='v9-module-icon'>🧾</div>
            <div class='v9-module-title'>GST JSON Output</div>
            <div class='v9-module-desc'>Generate GST portal-ready JSON after final review while keeping the confirmed JSON logic protected.</div>
            <div class='v9-module-badge'>Review → JSON</div>
        </div>
    </div>
    """, unsafe_allow_html=True)


def v9_json_readiness_panel():
    action = st.session_state.get("action_df", pd.DataFrame())
    ready = isinstance(action, pd.DataFrame) and not action.empty
    if ready:
        accepted = int((action.get("final_user_action", pd.Series(dtype=str)) == "Accepted").sum())
        pending = int((action.get("final_user_action", pd.Series(dtype=str)) == "Pending").sum())
        rejected = int((action.get("final_user_action", pd.Series(dtype=str)) == "Rejected").sum())
        status_line = f"Accepted: {accepted:,} • Pending: {pending:,} • Rejected: {rejected:,}"
    else:
        status_line = "Run reconciliation and review actions first"

    st.markdown(f"""
    <div class='v9-readiness'>
        <div class='v9-readiness-title'>🛡️ Final GST JSON Readiness</div>
        <div class='v9-check-grid'>
            <div class='v9-check'><div class='v9-check-icon'>✅</div><div class='v9-check-label'>GST upload structure protected</div></div>
            <div class='v9-check'><div class='v9-check-icon'>🧾</div><div class='v9-check-label'>rtin / reqtyp / invdata preserved</div></div>
            <div class='v9-check'><div class='v9-check-icon'>🔐</div><div class='v9-check-label'>Amendment-safe section handling</div></div>
            <div class='v9-check'><div class='v9-check-icon'>📌</div><div class='v9-check-label'>{status_line}</div></div>
        </div>
    </div>
    """, unsafe_allow_html=True)


def v9_help_box(title, text):
    st.markdown(f"""
    <div class='v9-help-box'>
        <div class='v9-help-title'>{title}</div>
        <div class='v9-help-text'>{text}</div>
    </div>
    """, unsafe_allow_html=True)


def v9_report_cards():
    st.markdown("""
    <div class='v9-report-grid'>
        <div class='v9-report-card'><div class='v9-report-title'>📘 Final Action Report</div><div class='v9-report-desc'>Invoice-wise final action, remarks and recommended action summary.</div></div>
        <div class='v9-report-card'><div class='v9-report-title'>⚠️ Risk & Exception Report</div><div class='v9-report-desc'>High-risk cases, mismatches, duplicates and vendor follow-up items.</div></div>
        <div class='v9-report-card'><div class='v9-report-title'>🧾 JSON Upload Summary</div><div class='v9-report-desc'>Action summary and records prepared for GST portal JSON generation.</div></div>
    </div>
    """, unsafe_allow_html=True)





# =========================================================
# V10 ADVANCED SALEABLE UI HELPERS — UI ONLY
# Final GST JSON generation logic is intentionally untouched.
# =========================================================

def v10_df_len(key: str) -> int:
    try:
        df = st.session_state.get(key, pd.DataFrame())
        return len(df) if isinstance(df, pd.DataFrame) else 0
    except Exception:
        return 0


def v10_safe_sum(df: pd.DataFrame, col: str) -> float:
    try:
        if isinstance(df, pd.DataFrame) and col in df.columns:
            return float(pd.to_numeric(df[col], errors="coerce").fillna(0).sum())
    except Exception:
        pass
    return 0.0


def v10_quality_score(df: pd.DataFrame) -> int:
    if not isinstance(df, pd.DataFrame) or df.empty:
        return 0
    total = max(len(df), 1)
    score = 100
    for col in ["supplier_gstin", "document_no", "document_date"]:
        if col in df.columns:
            blanks = df[col].astype(str).str.strip().isin(["", "nan", "None", "NaT"]).sum()
            score -= int((blanks / total) * 18)
    if "supplier_gstin" in df.columns:
        invalid = (~df["supplier_gstin"].astype(str).str.upper().str.match(r"^[0-9]{2}[A-Z]{5}[0-9]{4}[A-Z][1-9A-Z]Z[0-9A-Z]$", na=False)).sum()
        score -= int((invalid / total) * 20)
    if {"supplier_gstin", "document_no"}.issubset(df.columns):
        dups = df.duplicated(["supplier_gstin", "document_no"], keep=False).sum()
        score -= int((dups / total) * 15)
    return max(0, min(100, score))


def v10_score_class(score: int) -> str:
    if score >= 90:
        return ""
    if score >= 70:
        return "warn"
    return "bad"


def v10_command_center():
    st.markdown("""
    <div class='v10-command-center'>
        <div class='v10-command-title'>⚡ IMS Recon Pro Command Center</div>
        <div class='v10-command-sub'>A premium workflow built for GST IMS reconciliation, exception review, action control, reporting and final GST upload preparation.</div>
        <div class='v10-action-grid'>
            <div class='v10-action-card'><div class='v10-action-icon'>🚀</div><div class='v10-action-title'>Start New Reconciliation</div><div class='v10-action-desc'>Begin client setup, upload data and run IMS matching in a guided workflow.</div></div>
            <div class='v10-action-card'><div class='v10-action-icon'>📤</div><div class='v10-action-title'>Upload & Validate</div><div class='v10-action-desc'>Check GSTIN, invoice details, duplicates and tax values before matching.</div></div>
            <div class='v10-action-card'><div class='v10-action-icon'>✅</div><div class='v10-action-title'>Review Actions</div><div class='v10-action-desc'>Finalize Accepted, Pending and Rejected actions with remarks and filters.</div></div>
            <div class='v10-action-card'><div class='v10-action-icon'>🧾</div><div class='v10-action-title'>Generate Output</div><div class='v10-action-desc'>Prepare reports and GST portal-ready JSON after final review.</div></div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        if st.button("📌 Client Setup", use_container_width=True, key="v10_client_setup"):
            st.session_state.page = "Client Setup"
            st.rerun()
    with c2:
        if st.button("📤 Upload Data", use_container_width=True, key="v10_upload_data"):
            st.session_state.page = "Upload Center"
            st.rerun()
    with c3:
        if st.button("🔄 Run Reco", use_container_width=True, key="v10_run_reco"):
            st.session_state.page = "Reconciliation Workspace"
            st.rerun()
    with c4:
        if st.button("🧾 Final JSON", use_container_width=True, key="v10_final_json"):
            st.session_state.page = "Reports & Export"
            st.rerun()


def v10_quality_dashboard():
    """
    Upload Data Quality Dashboard — native Streamlit implementation.
    No raw HTML is used here, so HTML code cannot appear on screen.
    UI-only function. GST JSON generation logic is untouched.
    """
    p = st.session_state.get("purchase_df", pd.DataFrame())
    ims = st.session_state.get("ims_df", pd.DataFrame())

    def amount_summary(df):
        if not isinstance(df, pd.DataFrame) or df.empty:
            return {
                "records": 0,
                "taxable": 0.0,
                "invoice_value": 0.0,
                "igst": 0.0,
                "cgst": 0.0,
                "sgst": 0.0,
                "cess": 0.0,
                "total_tax": 0.0,
            }

        igst = v10_safe_sum(df, "igst")
        cgst = v10_safe_sum(df, "cgst")
        sgst = v10_safe_sum(df, "sgst")
        cess = v10_safe_sum(df, "cess")
        return {
            "records": len(df),
            "taxable": v10_safe_sum(df, "taxable_value"),
            "invoice_value": v10_safe_sum(df, "invoice_value"),
            "igst": igst,
            "cgst": cgst,
            "sgst": sgst,
            "cess": cess,
            "total_tax": igst + cgst + sgst + cess,
        }

    def quality_status(score):
        if score >= 90:
            return "Strong data quality"
        if score >= 70:
            return "Review recommended"
        return "Upload/check data"

    def render_quality(title, df, icon):
        score = v10_quality_score(df)
        data = amount_summary(df)

        st.markdown(f"#### {icon} {title}")
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Data Quality", f"{score}%", quality_status(score))
        c2.metric("Records", f"{data['records']:,}")
        c3.metric("Taxable Value", f"₹{data['taxable']:,.0f}")
        c4.metric("Invoice Value", f"₹{data['invoice_value']:,.0f}")

        c5, c6, c7, c8 = st.columns(4)
        c5.metric("IGST", f"₹{data['igst']:,.0f}")
        c6.metric("CGST", f"₹{data['cgst']:,.0f}")
        c7.metric("SGST", f"₹{data['sgst']:,.0f}")
        c8.metric("CESS", f"₹{data['cess']:,.0f}")

        c9, _, _, _ = st.columns(4)
        c9.metric("Total Tax", f"₹{data['total_tax']:,.0f}")

    st.markdown("### 📊 Upload Data Quality")
    st.caption("Data health, amount and tax summary for Purchase Register and IMS JSON.")

    left, right = st.columns(2)
    with left:
        with st.container(border=True):
            render_quality("Purchase Register Quality", p, "📘")

    with right:
        with st.container(border=True):
            render_quality("IMS JSON Quality", ims, "🧾")


def v10_empty_state(title: str, text: str, icon: str = "📭"):
    st.markdown(f"""
    <div class='v10-empty-state'>
        <div class='v10-empty-icon'>{icon}</div>
        <div class='v10-empty-title'>{title}</div>
        <div class='v10-empty-text'>{text}</div>
    </div>
    """, unsafe_allow_html=True)


def v10_reco_control_room():
    recon = st.session_state.get("recon_df", pd.DataFrame())
    action = st.session_state.get("action_df", pd.DataFrame())

    if not isinstance(recon, pd.DataFrame) or recon.empty:
        v10_empty_state("Reconciliation not started yet", "Upload Purchase Register and IMS JSON, then run reconciliation to open the control room.", "🔄")
        return

    def count_col(df, col, value):
        try:
            return int((df[col] == value).sum()) if col in df.columns else 0
        except Exception:
            return 0

    matched = count_col(recon, "mismatch_type", "Matched")
    only_ims = count_col(recon, "mismatch_type", "Only in IMS")
    only_purchase = count_col(recon, "mismatch_type", "Only in Purchase")
    value_mismatch = int(recon.get("mismatch_type", pd.Series(dtype=str)).astype(str).str.contains("Value", case=False, na=False).sum()) if "mismatch_type" in recon.columns else 0
    highrisk = int(action.get("risk_level", pd.Series(dtype=str)).isin(["High", "Critical"]).sum()) if isinstance(action, pd.DataFrame) and not action.empty and "risk_level" in action.columns else 0

    st.markdown(f"""
    <div class='v10-control-room'>
        <div class='v10-control-main'>
            <div class='v10-control-title'>🎛️ Reconciliation Control Room</div>
            <div style='font-size:14px;color:#60748f;line-height:1.5;'>Focus on exception areas first. Use the Action Center to finalize Accepted, Pending and Rejected actions.</div>
            <div class='v10-badge-row'>
                <span class='v10-filter-badge green'>✅ Matched: {matched:,}</span>
                <span class='v10-filter-badge orange'>🟠 Only in IMS: {only_ims:,}</span>
                <span class='v10-filter-badge purple'>📘 Only in Purchase: {only_purchase:,}</span>
                <span class='v10-filter-badge red'>⚠️ Value Mismatch: {value_mismatch:,}</span>
                <span class='v10-filter-badge red'>🔥 High Risk: {highrisk:,}</span>
            </div>
        </div>
        <div class='v10-control-side'>
            <div class='v10-control-title'>🧭 Suggested Review Order</div>
            <div class='v10-badge-row'>
                <span class='v10-filter-badge red'>1. High Risk</span>
                <span class='v10-filter-badge orange'>2. Value Mismatch</span>
                <span class='v10-filter-badge purple'>3. Only in IMS</span>
                <span class='v10-filter-badge green'>4. Matched</span>
            </div>
            <div style='font-size:13px;color:#60748f;line-height:1.55;'>This review flow helps users reach final action faster and reduces manual checking effort.</div>
        </div>
    </div>
    """, unsafe_allow_html=True)


def v10_action_header():
    action = st.session_state.get("action_df", pd.DataFrame())
    if not isinstance(action, pd.DataFrame) or action.empty:
        v10_empty_state("No action table available", "Run reconciliation first, then review invoice-wise actions here.", "✅")
        return

    def c(value):
        try:
            return int((action["final_user_action"] == value).sum()) if "final_user_action" in action.columns else 0
        except Exception:
            return 0

    accepted, pending, rejected = c("Accepted"), c("Pending"), c("Rejected")
    review = c("Review")
    no_action = c("No Action")

    st.markdown(f"""
    <div class='v10-control-main' style='margin:14px 0 18px 0;'>
        <div class='v10-control-title'>✅ Action Center Command Bar</div>
        <div style='font-size:14px;color:#60748f;line-height:1.5;'>Use filters, manual action and remarks to finalize invoices before GST JSON generation.</div>
        <div class='v10-badge-row'>
            <span class='v10-filter-badge green'>Accepted: {accepted:,}</span>
            <span class='v10-filter-badge orange'>Pending: {pending:,}</span>
            <span class='v10-filter-badge red'>Rejected: {rejected:,}</span>
            <span class='v10-filter-badge purple'>Review: {review:,}</span>
            <span class='v10-filter-badge'>No Action: {no_action:,}</span>
        </div>
    </div>
    """, unsafe_allow_html=True)


def v10_final_json_review_ui():
    action = st.session_state.get("action_df", pd.DataFrame())
    ims = st.session_state.get("ims_df", pd.DataFrame())

    accepted = pending = rejected = no_action = review = 0
    if isinstance(action, pd.DataFrame) and not action.empty and "final_user_action" in action.columns:
        accepted = int((action["final_user_action"] == "Accepted").sum())
        pending = int((action["final_user_action"] == "Pending").sum())
        rejected = int((action["final_user_action"] == "Rejected").sum())
        no_action = int((action["final_user_action"] == "No Action").sum())
        review = int((action["final_user_action"] == "Review").sum())

    sections = 0
    try:
        if isinstance(ims, pd.DataFrame) and "ims_section" in ims.columns:
            sections = ims["ims_section"].nunique()
    except Exception:
        sections = 0

    st.markdown(f"""
    <div class='v10-json-review'>
        <div class='v10-json-title'>🧾 Final GST Upload JSON Review</div>
        <div class='v10-json-checks'>
            <div class='v10-json-check'><div class='v10-json-check-icon'>✅</div><div class='v10-json-check-label'>Accepted<br>{accepted:,}</div></div>
            <div class='v10-json-check'><div class='v10-json-check-icon'>🟠</div><div class='v10-json-check-label'>Pending<br>{pending:,}</div></div>
            <div class='v10-json-check'><div class='v10-json-check-icon'>🔴</div><div class='v10-json-check-label'>Rejected<br>{rejected:,}</div></div>
            <div class='v10-json-check'><div class='v10-json-check-icon'>📦</div><div class='v10-json-check-label'>IMS Sections<br>{sections:,}</div></div>
            <div class='v10-json-check'><div class='v10-json-check-icon'>🛡️</div><div class='v10-json-check-label'>GST JSON Logic<br>Protected</div></div>
        </div>
    </div>
    """, unsafe_allow_html=True)


def v10_management_summary():
    p = st.session_state.get("purchase_df", pd.DataFrame())
    ims = st.session_state.get("ims_df", pd.DataFrame())
    action = st.session_state.get("action_df", pd.DataFrame())

    accepted = pending = rejected = 0
    if isinstance(action, pd.DataFrame) and not action.empty and "final_user_action" in action.columns:
        accepted = int((action["final_user_action"] == "Accepted").sum())
        pending = int((action["final_user_action"] == "Pending").sum())
        rejected = int((action["final_user_action"] == "Rejected").sum())

    summary = (
        f"For the selected period, Purchase Register has {len(p):,} records and IMS JSON has {len(ims):,} records. "
        f"Based on current action review, {accepted:,} records are marked Accepted, {pending:,} records are marked Pending, "
        f"and {rejected:,} records are marked Rejected. The final GST upload JSON should be generated only after completing invoice-wise review."
    )

    st.markdown(f"""
    <div class='v10-management-summary'>
        <div class='v10-management-title'>📝 Management Summary</div>
        <div class='v10-management-text'>{summary}</div>
    </div>
    """, unsafe_allow_html=True)


def v10_help_tooltips():
    st.markdown("""
    <div class='v10-tooltip-grid'>
        <div class='v10-tooltip'><div class='v10-tooltip-title'>What is IMS JSON?</div><div class='v10-tooltip-text'>The file downloaded from GST portal containing inward supply records and action fields.</div></div>
        <div class='v10-tooltip'><div class='v10-tooltip-title'>When to mark Pending?</div><div class='v10-tooltip-text'>Use Pending for unmatched, disputed or review-required invoices.</div></div>
        <div class='v10-tooltip'><div class='v10-tooltip-title'>When to Accept?</div><div class='v10-tooltip-text'>Use Accepted when invoice details match books and credit/action is acceptable.</div></div>
        <div class='v10-tooltip'><div class='v10-tooltip-title'>Final JSON?</div><div class='v10-tooltip-text'>Generated after action review and uploaded back to GST portal.</div></div>
    </div>
    """, unsafe_allow_html=True)



# =========================================================
# MAIN
# =========================================================

def main():
    init_db()
    init_state()
    inject_css()

    if not st.session_state.logged_in:
        login_page()
        return

    top_header()
    horizontal_nav()

    page = st.session_state.page
    if page == "Dashboard":
        dashboard_page()
    elif page == "Client Setup":
        client_setup_page()
    elif page == "Upload Center":
        upload_center_page()
    elif page == "IMS Data Viewer":
        ims_data_viewer_page()
    elif page == "Reconciliation Workspace":
        reconciliation_page()
    elif page == "Action Center":
        action_center_page()
    elif page == "Risk Center":
        risk_center_page()
    elif page == "Vendor Follow-up":
        vendor_followup_page()
    elif page == "Reports & Export":
        reports_page()
    elif page == "AI Insight Desk":
        ai_insight_page()
    elif page == "Admin Panel":
        admin_page()

    st.markdown(f"""
    <div class='footer-bar'>
        <div style='display:flex;justify-content:space-around;gap:20px;flex-wrap:wrap;'>
            <div class='foot-item'><div style='font-size:26px;'>🛡️</div><div><div class='foot-main'>Secure</div><div class='foot-sub'>Enterprise-grade control</div></div></div>
            <div class='foot-item'><div style='font-size:26px;'>✅</div><div><div class='foot-main'>Compliant</div><div class='foot-sub'>GSTN workflow aligned</div></div></div>
            <div class='foot-item'><div style='font-size:26px;'>🔄</div><div><div class='foot-main'>Reliable</div><div class='foot-sub'>JSON + export ready</div></div></div>
            <div class='foot-item'><div style='font-size:26px;'>✨</div><div><div class='foot-main'>Smart</div><div class='foot-sub'>AI-like insights</div></div></div>
        </div>
    </div>
    <div style='text-align:center;color:#5d718e;font-size:14px;margin-top:14px;padding-bottom:10px;'>
        © 2026 IMS Recon Pro • {COPYRIGHT_OWNER} • Designed for India • Built for Compliance
    </div>
    """, unsafe_allow_html=True)



# =========================================================
# UI ONLY OVERRIDE PACK — PREMIUM REPLIT-INSPIRED STREAMLIT SKIN
# IMPORTANT:
# This block overrides only visual/rendering helper functions.
# GST JSON generation, reconciliation formulas, IMS parsing, CN/DN logic,
# action mapping and export functions above are intentionally untouched.
# =========================================================

def _ui_df_len(df) -> int:
    try:
        return int(len(df)) if isinstance(df, pd.DataFrame) else 0
    except Exception:
        return 0


def _ui_sum(df, col: str) -> float:
    try:
        if isinstance(df, pd.DataFrame) and not df.empty and col in df.columns:
            return float(pd.to_numeric(df[col], errors="coerce").fillna(0).sum())
    except Exception:
        pass
    return 0.0


def _ui_money(value) -> str:
    try:
        value = float(value or 0)
        abs_value = abs(value)
        sign = "-" if value < 0 else ""
        if abs_value >= 10000000:
            return f"{sign}₹{abs_value/10000000:,.2f} Cr"
        if abs_value >= 100000:
            return f"{sign}₹{abs_value/100000:,.2f} L"
        return f"{sign}₹{abs_value:,.0f}"
    except Exception:
        return "₹0"


def _ui_action_count(df, action: str) -> int:
    try:
        if isinstance(df, pd.DataFrame) and not df.empty and "final_user_action" in df.columns:
            return int((df["final_user_action"].astype(str) == action).sum())
        if isinstance(df, pd.DataFrame) and not df.empty and "recommended_action" in df.columns:
            return int((df["recommended_action"].astype(str) == action).sum())
    except Exception:
        pass
    return 0


def _ui_recon_count(df, mismatch: str) -> int:
    try:
        if isinstance(df, pd.DataFrame) and not df.empty and "mismatch_type" in df.columns:
            return int((df["mismatch_type"].astype(str) == mismatch).sum())
    except Exception:
        pass
    return 0


def _ui_safe_unique(df, col: str) -> int:
    try:
        if isinstance(df, pd.DataFrame) and not df.empty and col in df.columns:
            return int(df[col].astype(str).replace("", pd.NA).dropna().nunique())
    except Exception:
        pass
    return 0


def inject_css():
    st.markdown("""
    <style>
        :root {
            --ui-bg:#eaf1fb;
            --ui-ink:#0b1f3a;
            --ui-muted:#607086;
            --ui-navy:#071a3d;
            --ui-navy2:#0b2d66;
            --ui-blue:#2563eb;
            --ui-cyan:#38bdf8;
            --ui-saffron:#ff9933;
            --ui-green:#138808;
            --ui-red:#dc463f;
            --ui-purple:#6d3bd1;
            --ui-border:#c8d8ec;
            --ui-card:#ffffff;
        }

        html, body, [class*="css"] {
            font-family: Inter, "Segoe UI", Roboto, Arial, sans-serif;
        }

        .stApp {
            background:
                radial-gradient(circle at 10% 5%, rgba(255,153,51,.22), transparent 26%),
                radial-gradient(circle at 88% 8%, rgba(37,99,235,.16), transparent 28%),
                radial-gradient(circle at 78% 92%, rgba(19,136,8,.13), transparent 28%),
                linear-gradient(135deg,#edf5ff 0%,#dbeafe 46%,#f7fbff 100%);
            color: var(--ui-ink);
        }

        header[data-testid="stHeader"] {
            background: rgba(237,245,255,.84) !important;
            backdrop-filter: blur(16px);
            border-bottom: 1px solid rgba(11,45,102,.10);
        }

        div[data-testid="stToolbar"], #MainMenu, footer { visibility:hidden; height:0; }
        .block-container { max-width: 1600px; padding-top: 1rem; padding-bottom: 2.2rem; }

        section[data-testid="stSidebar"] {
            background: linear-gradient(180deg,#071a3d 0%,#0b2d66 55%,#071a3d 100%) !important;
            border-right: 1px solid rgba(255,255,255,.10);
        }
        section[data-testid="stSidebar"] * { color: #eef6ff !important; }
        section[data-testid="stSidebar"] div[data-testid="stMarkdownContainer"] p {
            color: #d7e7ff !important;
        }

        .ui-shell {
            border-radius: 30px;
            overflow: hidden;
            background: rgba(255,255,255,.92);
            border: 1px solid rgba(200,216,236,.98);
            box-shadow: 0 24px 60px rgba(7,26,61,.14);
            margin-bottom: 18px;
        }
        .ui-topbar {
            min-height: 38px;
            background: linear-gradient(90deg,#071a3d,#0b2d66,#071a3d);
            color:#dceaff;
            display:flex;
            align-items:center;
            justify-content:space-between;
            gap:12px;
            padding: 0 22px;
            font-size: 12px;
            letter-spacing:.02em;
        }
        .ui-hero {
            position: relative;
            padding: 28px 30px;
            background:
                radial-gradient(circle at 80% 20%, rgba(255,255,255,.16), transparent 32%),
                linear-gradient(135deg,#08214f 0%,#0b3677 56%,#123d82 100%);
            color: white;
            display:flex;
            justify-content:space-between;
            gap:22px;
            align-items:center;
            overflow:hidden;
        }
        .ui-hero::after {
            content:"@BAJRABHANU";
            position:absolute;
            right:22px;
            bottom:-8px;
            font-size:58px;
            font-weight:950;
            letter-spacing:.04em;
            color:rgba(255,255,255,.055);
        }
        .ui-brand-wrap { display:flex; align-items:center; gap:18px; position:relative; z-index:2; min-width:0; }
        .ui-logo-badge {
            width:76px; height:76px; border-radius:24px;
            background: linear-gradient(135deg,#ff9933 0%,#ffffff 50%,#138808 100%);
            display:flex; align-items:center; justify-content:center;
            color:#0b2d66; font-size:34px; font-weight:950;
            box-shadow: 0 18px 38px rgba(0,0,0,.22);
            border: 1px solid rgba(255,255,255,.38);
        }
        .ui-title { font-size:38px; font-weight:950; line-height:1.03; letter-spacing:-.03em; }
        .ui-subtitle { margin-top:8px; font-size:16px; color:#e7f0ff; line-height:1.45; max-width: 820px; }
        .ui-badges { display:flex; gap:8px; flex-wrap:wrap; margin-top:12px; }
        .ui-badge {
            display:inline-flex; align-items:center; gap:6px;
            padding:7px 11px; border-radius:999px;
            background: rgba(255,255,255,.12);
            border:1px solid rgba(255,255,255,.18);
            color:#eef6ff; font-size:12px; font-weight:800;
        }
        .ui-flag {
            width:148px; height:92px; border-radius:18px; position:relative; flex:0 0 auto;
            background: linear-gradient(to bottom,#ff9933 0 33.33%,#fff 33.33% 66.66%,#138808 66.66% 100%);
            box-shadow: 0 18px 40px rgba(0,0,0,.24);
            border:1px solid rgba(255,255,255,.38);
            animation: uiFloat 4s ease-in-out infinite;
            z-index:2;
        }
        .ui-flag::after {
            content:"☸"; position:absolute; inset:0;
            display:flex; align-items:center; justify-content:center;
            color:#0a3d91; font-size:28px; font-weight:900;
        }
        @keyframes uiFloat { 0%,100%{transform:translateY(0)} 50%{transform:translateY(-6px)} }

        .ui-meta-grid {
            display:grid;
            grid-template-columns: repeat(4, minmax(0,1fr));
            gap:14px;
            padding:16px 18px;
            background:linear-gradient(90deg,#f7fbff,#edf5ff);
        }
        .ui-meta-card {
            display:flex; gap:12px; align-items:center;
            background:#fff;
            border:1px solid #d3e2f4;
            border-radius:20px;
            padding:14px 15px;
            box-shadow: 0 10px 24px rgba(7,26,61,.07);
            min-height:74px;
        }
        .ui-meta-icon {
            width:44px; height:44px; border-radius:15px;
            display:flex; align-items:center; justify-content:center;
            background:linear-gradient(135deg,#eef6ff,#dbeafe);
            font-size:21px; flex:0 0 auto;
        }
        .ui-meta-label {
            font-size:11px; font-weight:900; color:var(--ui-muted);
            text-transform:uppercase; letter-spacing:.055em;
        }
        .ui-meta-value {
            font-size:14px; font-weight:950; color:var(--ui-ink);
            margin-top:3px; white-space:nowrap; overflow:hidden; text-overflow:ellipsis;
        }

        .ui-nav-wrap {
            background: rgba(255,255,255,.88);
            border:1px solid #c8d8ec;
            border-radius:24px;
            padding:12px;
            box-shadow:0 14px 34px rgba(7,26,61,.10);
            margin: 0 0 18px 0;
        }
        .ui-nav-label {
            font-size:12px; font-weight:950; color:#0b2d66;
            text-transform:uppercase; letter-spacing:.07em;
            margin: 0 0 8px 4px;
        }

        .main-shell, .panel, .metric-card, .small-card {
            background: rgba(255,255,255,.96);
            border: 1px solid rgba(200,216,236,.98);
            box-shadow: 0 16px 40px rgba(7,26,61,.10);
        }
        .main-shell { border-radius: 28px; overflow:hidden; margin-bottom:18px; }
        .content-pad { padding: 30px; position:relative; }
        .panel, .small-card {
            border-radius: 26px;
            padding: 20px 22px;
            height:100%;
            background: linear-gradient(180deg,#ffffff,#f8fbff);
        }
        .panel-title {font-size:20px;font-weight:950;color:#102244;margin-bottom:8px;}
        .section-title {font-size:30px;font-weight:950;color:#0b1f3a;margin:8px 0 4px 0;}
        .section-sub {font-size:15px;color:#607086;margin-bottom:18px;}

        .ui-page-title {
            border-radius:28px;
            padding:22px 24px;
            border:1px solid #c8d8ec;
            background:
                radial-gradient(circle at right top, rgba(56,189,248,.14), transparent 26%),
                linear-gradient(135deg,#ffffff,#f7fbff);
            box-shadow:0 16px 40px rgba(7,26,61,.10);
            margin-bottom:16px;
            position:relative;
            overflow:hidden;
        }
        .ui-page-title::after {
            content:"";
            position:absolute; right:-50px; top:-70px;
            width:190px; height:190px; border-radius:50%;
            background:rgba(37,99,235,.08);
        }
        .ui-page-kicker {
            color:#ff8a00; font-size:12px; font-weight:950;
            text-transform:uppercase; letter-spacing:.08em;
        }
        .ui-page-heading {
            font-size:30px; font-weight:950; color:#0b1f3a;
            line-height:1.12; margin-top:5px;
        }
        .ui-page-sub {
            font-size:15px; color:#607086; margin-top:7px; max-width:900px;
        }

        .metric-card {
            border-radius: 24px;
            padding:18px;
            min-height: 122px;
            position:relative;
            overflow:hidden;
            transition: all .18s ease;
        }
        .metric-card:hover { transform: translateY(-2px); box-shadow:0 20px 46px rgba(7,26,61,.14); }
        .metric-card::after {
            content:""; position:absolute; right:-36px; bottom:-44px;
            width:116px; height:116px; border-radius:50%;
            background: rgba(37,99,235,.075);
        }
        .metric-top {display:flex;align-items:center;gap:14px;position:relative;z-index:2;}
        .metric-icon {
            width:56px;height:56px;border-radius:18px;display:flex;align-items:center;justify-content:center;
            font-size:25px; box-shadow: inset 0 0 0 1px rgba(255,255,255,.55);
        }
        .metric-label {font-size:12px;color:#607086;font-weight:950;text-transform:uppercase;letter-spacing:.055em;}
        .metric-value {font-size:32px;font-weight:950;color:#0b1f3a;line-height:1.15;margin-top:5px;}
        .metric-delta {font-size:12px;color:#138808;margin-top:5px;font-weight:900;}
        .metric-delta.red {color:#dc463f;}

        .ui-command {
            border-radius: 30px;
            padding: 24px;
            color:white;
            background:
                radial-gradient(circle at 90% 20%, rgba(255,255,255,.18), transparent 30%),
                linear-gradient(135deg,#071a3d,#0b3677);
            border:1px solid rgba(255,255,255,.16);
            box-shadow: 0 24px 58px rgba(7,26,61,.20);
            margin: 14px 0 20px 0;
            overflow:hidden;
            position:relative;
        }
        .ui-command::after {
            content:"GST"; position:absolute; right:24px; top:8px;
            font-size:90px; font-weight:950; color:rgba(255,255,255,.055);
        }
        .ui-command-title {font-size:28px;font-weight:950;position:relative;z-index:2;}
        .ui-command-sub {font-size:14px;color:#d9e8ff;margin:8px 0 18px 0;position:relative;z-index:2;max-width:950px;}
        .ui-action-grid {
            display:grid;grid-template-columns:repeat(4,minmax(0,1fr));gap:14px;position:relative;z-index:2;
        }
        .ui-action-card {
            background:rgba(255,255,255,.105);
            border:1px solid rgba(255,255,255,.16);
            border-radius:22px;
            padding:16px;
            min-height:126px;
            backdrop-filter: blur(5px);
        }
        .ui-action-icon {font-size:28px;margin-bottom:9px;}
        .ui-action-title {font-size:16px;font-weight:950;color:#fff;}
        .ui-action-desc {font-size:12px;color:#d7e7ff;line-height:1.45;margin-top:6px;}

        .ui-workflow {
            display:grid;
            grid-template-columns: repeat(8, minmax(0,1fr));
            gap:10px;
            margin: 12px 0 18px 0;
        }
        .ui-step {
            background:#ffffff;
            border:1px solid #d6e4f5;
            border-radius:18px;
            padding:12px;
            min-height:92px;
            box-shadow: 0 10px 22px rgba(7,26,61,.06);
            position:relative;
            overflow:hidden;
        }
        .ui-step.done { background:linear-gradient(180deg,#ffffff,#effaf1); border-color:#bfe9c9; }
        .ui-step.active { background:linear-gradient(180deg,#ffffff,#fff5e8); border-color:#ffd9a8; }
        .ui-step-num {
            width:30px;height:30px;border-radius:11px;background:#eaf2ff;color:#0b2d66;
            display:flex;align-items:center;justify-content:center;font-size:12px;font-weight:950;margin-bottom:8px;
        }
        .ui-step.done .ui-step-num {background:#e9f9ed;color:#138808;}
        .ui-step.active .ui-step-num {background:#fff0da;color:#c76f00;}
        .ui-step-label {font-size:12px;font-weight:950;color:#102244;line-height:1.25;}
        .ui-step-status {font-size:11px;font-weight:800;color:#607086;margin-top:5px;}

        .v10-command-center, .v9-workflow, .v9-readiness, .v10-json-review {
            border-radius: 28px !important;
        }

        div[data-testid="stMetric"] {
            background: linear-gradient(180deg,#ffffff,#f7fbff);
            border:1px solid #d6e4f5;
            border-radius:18px;
            padding:13px 14px;
            box-shadow:0 8px 18px rgba(7,26,61,.07);
        }
        div[data-testid="stMetricLabel"] {font-weight:900;color:#607086;}
        div[data-testid="stMetricValue"] {font-weight:950;color:#0b1f3a;}

        div[data-testid="stDataFrame"] {
            border-radius:20px;
            overflow:hidden;
            border:1px solid #cfdff0;
            box-shadow:0 10px 24px rgba(7,26,61,.07);
        }

        .stTabs [data-baseweb="tab-list"] {
            gap:8px;
            background:#edf4ff;
            padding:8px;
            border-radius:20px;
            border:1px solid #cfdded;
        }
        .stTabs [data-baseweb="tab"] {
            border-radius:15px;
            padding:10px 16px;
            font-weight:900;
        }
        .stTabs [aria-selected="true"] {
            background:#ffffff !important;
            color:#0b2d66 !important;
            box-shadow:0 8px 16px rgba(7,26,61,.08);
        }

        .stTextInput input, .stNumberInput input, .stTextArea textarea, .stDateInput input {
            border-radius:16px !important;
            border:1px solid #cdddf0 !important;
            background:#ffffff !important;
            box-shadow: inset 0 1px 0 rgba(255,255,255,.8);
        }
        .stSelectbox div[data-baseweb="select"] > div {
            border-radius:16px !important;
            border-color:#cdddf0 !important;
            background:#ffffff !important;
        }

        .stFileUploader {
            background: linear-gradient(180deg,#ffffff,#f7fbff);
            border: 1px dashed #9eb9dc;
            border-radius: 22px;
            padding: 12px;
        }

        .stButton > button, .stDownloadButton > button, .stFormSubmitButton > button {
            border-radius:16px !important;
            min-height:46px !important;
            border:1px solid #c8d8ec !important;
            background:linear-gradient(180deg,#ffffff,#f2f7ff) !important;
            color:#0b2d66 !important;
            font-weight:950 !important;
            box-shadow:0 7px 16px rgba(7,26,61,.07) !important;
            transition: all .16s ease !important;
        }
        .stButton > button:hover, .stDownloadButton > button:hover, .stFormSubmitButton > button:hover {
            transform: translateY(-1px);
            border-color:#2563eb !important;
            background:linear-gradient(180deg,#eef6ff,#dbeafe) !important;
            color:#071a3d !important;
        }
        .stButton > button[kind="primary"], .stFormSubmitButton > button[kind="primary"] {
            background:linear-gradient(135deg,#0b2d66,#2563eb) !important;
            color:white !important;
            border-color:rgba(255,255,255,.18) !important;
        }

        .ui-help {
            background:#fffdf6;
            border:1px solid #f3d9a7;
            border-radius:22px;
            padding:16px 18px;
            color:#62420d;
            box-shadow:0 10px 22px rgba(7,26,61,.06);
            margin:12px 0 18px 0;
        }
        .ui-help-title {font-weight:950;font-size:16px;color:#7a4b00;margin-bottom:6px;}
        .ui-help-text {font-size:13px;line-height:1.5;}

        .footer-bar {
            margin-top:22px;
            border-radius:26px;
            background:linear-gradient(90deg,#071a3d 0%,#082b61 50%,#071a3d 100%);
            color:white;
            padding:18px 22px;
            box-shadow:0 16px 38px rgba(7,26,61,.16);
        }
        .foot-item {display:flex;align-items:center;gap:10px;justify-content:center;}
        .foot-main {font-weight:950;}
        .foot-sub {font-size:13px;color:#d4e0ff;}

        .ui-login-bg {
            min-height: calc(100vh - 34px);
            border-radius:34px;
            display:flex;
            align-items:center;
            justify-content:center;
            padding:24px;
            position:relative;
            overflow:hidden;
            background:
                radial-gradient(circle at 18% 15%, rgba(255,153,51,.25), transparent 28%),
                radial-gradient(circle at 82% 84%, rgba(19,136,8,.18), transparent 30%),
                linear-gradient(135deg,#071a3d,#0b2d66 56%,#071a3d);
        }
        .ui-login-bg::before {
            content:"IMS";
            position:absolute;
            left:-30px;
            bottom:-46px;
            font-size:190px;
            font-weight:950;
            color:rgba(255,255,255,.045);
        }
        .ui-login-card {
            width:min(1040px, 100%);
            display:grid;
            grid-template-columns: 1.05fr .88fr;
            gap:18px;
            position:relative;
            z-index:2;
        }
        .ui-login-left, .ui-login-panel {
            border-radius:32px;
            border:1px solid rgba(255,255,255,.18);
            box-shadow:0 34px 90px rgba(0,0,0,.28);
        }
        .ui-login-left {
            color:white;
            padding:34px;
            min-height:520px;
            background:
                radial-gradient(circle at 80% 25%, rgba(255,255,255,.13), transparent 30%),
                rgba(255,255,255,.08);
            backdrop-filter: blur(8px);
        }
        .ui-login-panel {
            background:rgba(255,255,255,.96);
            padding:34px;
        }
        .ui-login-logo {
            width:76px;height:76px;border-radius:24px;
            background:linear-gradient(135deg,#ff9933 0%,#fff 50%,#138808 100%);
            color:#0b2d66;display:flex;align-items:center;justify-content:center;
            font-size:34px;font-weight:950;box-shadow:0 18px 38px rgba(0,0,0,.22);
            margin-bottom:22px;
        }
        .ui-login-title {font-size:42px;font-weight:950;line-height:1.04;letter-spacing:-.035em;}
        .ui-login-sub {font-size:16px;color:#dceaff;line-height:1.6;margin-top:14px;max-width:520px;}
        .ui-login-feature-grid {display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:12px;margin-top:26px;}
        .ui-login-feature {
            background:rgba(255,255,255,.10);
            border:1px solid rgba(255,255,255,.14);
            border-radius:18px;padding:14px;
        }
        .ui-login-feature-title {font-weight:950;color:white;font-size:14px;}
        .ui-login-feature-text {font-size:12px;color:#d7e7ff;margin-top:5px;line-height:1.4;}
        .ui-login-panel-title {font-size:30px;font-weight:950;color:#071b4a;text-align:center;}
        .ui-login-panel-sub {font-size:14px;color:#566982;text-align:center;line-height:1.5;margin:8px 0 22px 0;}
        .ui-login-copy {
            text-align:center;font-size:12px;color:#6b7d96;margin-top:16px;
            border-top:1px solid #e3edf8;padding-top:13px;font-weight:800;
        }

        @media (max-width: 1150px) {
            .ui-action-grid { grid-template-columns:repeat(2,minmax(0,1fr)); }
            .ui-workflow { grid-template-columns:repeat(4,minmax(0,1fr)); }
            .ui-meta-grid { grid-template-columns:repeat(2,minmax(0,1fr)); }
            .ui-login-card { grid-template-columns:1fr; }
        }
        @media (max-width: 760px) {
            .block-container { padding-left:.75rem; padding-right:.75rem; }
            .ui-hero { flex-direction:column; align-items:flex-start; }
            .ui-title { font-size:30px; }
            .ui-flag { width:126px;height:78px; }
            .ui-action-grid, .ui-workflow, .ui-meta-grid, .ui-login-feature-grid { grid-template-columns:1fr; }
            .ui-login-left { min-height:auto; }
        }
    </style>
    """, unsafe_allow_html=True)


def top_header():
    client = st.session_state.get("client_name") or "Client not selected"
    gstin = st.session_state.get("client_gstin") or "GSTIN pending"
    period = st.session_state.get("return_period") or datetime.today().strftime("%b-%Y")
    user = st.session_state.get("display_name") or st.session_state.get("username") or "User"
    role = st.session_state.get("role") or "Role"

    p = st.session_state.get("purchase_df", pd.DataFrame())
    ims = st.session_state.get("ims_df", pd.DataFrame())
    recon = st.session_state.get("recon_df", pd.DataFrame())
    action = st.session_state.get("action_df", pd.DataFrame())

    st.markdown(f"""
    <div class="ui-shell">
        <div class="ui-topbar">
            <div>🛡️ Secure GST IMS Workspace • Engine {ENGINE_VERSION}</div>
            <div>{datetime.now().strftime("%d-%b-%Y %I:%M %p")} • {COPYRIGHT_OWNER}</div>
        </div>
        <div class="ui-hero">
            <div class="ui-brand-wrap">
                <div class="ui-logo-badge">☸</div>
                <div>
                    <div class="ui-title">{APP_TITLE}</div>
                    <div class="ui-subtitle">{APP_TAGLINE}<br>Upload → Validate → Reconcile → Explain → Act → Review → Sign-off → Generate GST JSON</div>
                    <div class="ui-badges">
                        <span class="ui-badge">🔒 JSON Logic Protected</span>
                        <span class="ui-badge">🧾 IMS CN/DN Ready</span>
                        <span class="ui-badge">📊 Action Control</span>
                        <span class="ui-badge">🇮🇳 India GST Workflow</span>
                    </div>
                </div>
            </div>
            <div class="ui-flag"></div>
        </div>
        <div class="ui-meta-grid">
            <div class="ui-meta-card"><div class="ui-meta-icon">🏢</div><div><div class="ui-meta-label">Client</div><div class="ui-meta-value">{client}</div></div></div>
            <div class="ui-meta-card"><div class="ui-meta-icon">🔢</div><div><div class="ui-meta-label">GSTIN</div><div class="ui-meta-value">{gstin}</div></div></div>
            <div class="ui-meta-card"><div class="ui-meta-icon">📅</div><div><div class="ui-meta-label">Return Period</div><div class="ui-meta-value">{period}</div></div></div>
            <div class="ui-meta-card"><div class="ui-meta-icon">👤</div><div><div class="ui-meta-label">Logged In</div><div class="ui-meta-value">{user} • {role}</div></div></div>
        </div>
    </div>
    """, unsafe_allow_html=True)


def horizontal_nav():
    nav_items = [
        ("Dashboard", "🏠"),
        ("Client Setup", "🏢"),
        ("Upload Center", "📤"),
        ("IMS Data Viewer", "📑"),
        ("Reconciliation Workspace", "🔄"),
        ("Action Center", "✅"),
        ("Risk Center", "⚠️"),
        ("Vendor Follow-up", "✉️"),
        ("Reports & Export", "🧾"),
        ("AI Insight Desk", "🧠"),
        ("Admin Panel", "⚙️"),
    ]

    st.markdown("<div class='ui-nav-wrap'><div class='ui-nav-label'>Workspace Navigation</div>", unsafe_allow_html=True)
    rows = [nav_items[:6], nav_items[6:]]
    for row_idx, row in enumerate(rows):
        cols = st.columns(len(row))
        for i, (label, icon) in enumerate(row):
            with cols[i]:
                active = st.session_state.get("page") == label
                button_label = f"{icon} {label}" if not active else f"● {icon} {label}"
                if st.button(button_label, use_container_width=True, key=f"ui_nav_{row_idx}_{i}_{label}"):
                    st.session_state.page = label
                    st.rerun()
    st.markdown("</div>", unsafe_allow_html=True)


def page_title(title: str, subtitle: str = ""):
    st.markdown(f"""
    <div class="ui-page-title">
        <div class="ui-page-kicker">IMS Recon Pro Workspace</div>
        <div class="ui-page-heading">{title}</div>
        <div class="ui-page-sub">{subtitle}</div>
    </div>
    """, unsafe_allow_html=True)


def metric_card(icon, label, value, delta="", bg="#ffffff", color="#2563eb", red=False):
    red_class = " red" if red else ""
    st.markdown(f"""
    <div class="metric-card" style="background:linear-gradient(135deg,{bg},#ffffff);">
        <div class="metric-top">
            <div class="metric-icon" style="background:{bg}; color:{color};">{icon}</div>
            <div>
                <div class="metric-label">{label}</div>
                <div class="metric-value">{value}</div>
                <div class="metric-delta{red_class}">{delta}</div>
            </div>
        </div>
    </div>
    """, unsafe_allow_html=True)


def _ui_step_class(done: bool, active: bool) -> str:
    if done:
        return "done"
    if active:
        return "active"
    return ""


def v9_workflow_tracker():
    """Native Streamlit workflow tracker.
    UI-only change: avoids exposing raw HTML tags while keeping all workflow logic untouched.
    """
    p = st.session_state.get("purchase_df", pd.DataFrame())
    ims = st.session_state.get("ims_df", pd.DataFrame())
    recon = st.session_state.get("recon_df", pd.DataFrame())
    action = st.session_state.get("action_df", pd.DataFrame())
    final_json = bool(st.session_state.get("final_json_bytes", b""))

    final_summary = st.session_state.get("final_json_summary", pd.DataFrame())
    review_ready = isinstance(final_summary, pd.DataFrame) and not final_summary.empty

    steps = [
        ("Client", bool(st.session_state.get("client_gstin")), "GSTIN/period"),
        ("Upload PR", not p.empty, f"{_ui_df_len(p):,} rows"),
        ("Upload IMS", not ims.empty, f"{_ui_df_len(ims):,} rows"),
        ("Validate", (not p.empty or not ims.empty), "quality check"),
        ("Reconcile", not recon.empty, f"{_ui_df_len(recon):,} rows"),
        ("Act", not action.empty, "actions ready"),
        ("Review", review_ready, "sign-off"),
        ("JSON", final_json, "download"),
    ]

    first_pending = next((i for i, (_, done, _) in enumerate(steps) if not done), len(steps) - 1)

    st.markdown("### 🚀 GST IMS Workflow")
    cols = st.columns(8)
    for idx, (label, done, note) in enumerate(steps, start=1):
        is_active = (idx - 1) == first_pending and not done
        if done:
            icon = "✅"
            status = "Done"
            delta = note
        elif is_active:
            icon = "🟠"
            status = "Active"
            delta = note
        else:
            icon = "⚪"
            status = "Pending"
            delta = note

        with cols[idx - 1]:
            st.metric(
                label=f"{icon} {idx}. {label}",
                value=status,
                delta=delta,
            )

def hero_dashboard():
    p = st.session_state.get("purchase_df", pd.DataFrame())
    ims = st.session_state.get("ims_df", pd.DataFrame())
    recon = st.session_state.get("recon_df", pd.DataFrame())
    matched = _ui_recon_count(recon, "Matched")
    total = _ui_df_len(recon)
    health = round((matched / max(total, 1)) * 100, 1) if total else 0
    total_itc = _ui_sum(ims, "total_tax")

    st.markdown(f"""
    <div class="main-shell">
        <div class="content-pad">
            <div class="headline">GST IMS COMMAND DASHBOARD</div>
            <div class="main-title">Your complete IMS reconciliation control room is ready.</div>
            <div class="subcopy">
                Monitor uploads, reconciliation health, mismatch buckets, action readiness and final GST JSON safety from one premium workspace.
            </div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        metric_card("📚", "Purchase Rows", f"{_ui_df_len(p):,}", "books register", "#fff3e8", "#ff8a00")
    with c2:
        metric_card("🧬", "IMS Rows", f"{_ui_df_len(ims):,}", "GST portal data", "#e9f9ed", "#138808")
    with c3:
        metric_card("🧾", "IMS ITC Value", _ui_money(total_itc), "tax summary", "#edf4ff", "#2563eb")
    with c4:
        metric_card("🛡️", "Compliance Health", f"{health}%", "matched ratio", "#f2ecff", "#6d3bd1")

    v9_workflow_tracker()


def v10_command_center():
    p = st.session_state.get("purchase_df", pd.DataFrame())
    ims = st.session_state.get("ims_df", pd.DataFrame())
    recon = st.session_state.get("recon_df", pd.DataFrame())
    action = st.session_state.get("action_df", pd.DataFrame())
    accepted = _ui_action_count(action, "Accepted")
    pending = _ui_action_count(action, "Pending")
    highrisk = 0
    try:
        if isinstance(recon, pd.DataFrame) and not recon.empty and "risk_level" in recon.columns:
            highrisk = int(recon["risk_level"].isin(["High", "Critical"]).sum())
    except Exception:
        highrisk = 0

    st.markdown(f"""
    <div class="ui-command">
        <div class="ui-command-title">⚡ GST IMS Command Center</div>
        <div class="ui-command-sub">
            A Replit-inspired, premium Streamlit control room built on your existing Python engine.
            Your reconciliation formula and GST JSON generation logic are not changed.
        </div>
        <div class="ui-action-grid">
            <div class="ui-action-card"><div class="ui-action-icon">📤</div><div class="ui-action-title">Upload & Validate</div><div class="ui-action-desc">Purchase rows: {_ui_df_len(p):,}<br>IMS rows: {_ui_df_len(ims):,}</div></div>
            <div class="ui-action-card"><div class="ui-action-icon">🔄</div><div class="ui-action-title">Reconciliation</div><div class="ui-action-desc">Reco rows: {_ui_df_len(recon):,}<br>Matched: {_ui_recon_count(recon, "Matched"):,}</div></div>
            <div class="ui-action-card"><div class="ui-action-icon">✅</div><div class="ui-action-title">IMS Actions</div><div class="ui-action-desc">Accepted: {accepted:,}<br>Pending: {pending:,}</div></div>
            <div class="ui-action-card"><div class="ui-action-icon">⚠️</div><div class="ui-action-title">Risk Review</div><div class="ui-action-desc">High/Critical rows: {highrisk:,}<br>Review before JSON</div></div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        if st.button("🏢 Client Setup", use_container_width=True, key="ui_cmd_client"):
            st.session_state.page = "Client Setup"; st.rerun()
    with c2:
        if st.button("📤 Upload Data", use_container_width=True, key="ui_cmd_upload"):
            st.session_state.page = "Upload Center"; st.rerun()
    with c3:
        if st.button("🔄 Run Reco", use_container_width=True, key="ui_cmd_reco"):
            st.session_state.page = "Reconciliation Workspace"; st.rerun()
    with c4:
        if st.button("🧾 Final JSON", use_container_width=True, key="ui_cmd_json"):
            st.session_state.page = "Reports & Export"; st.rerun()


def v10_help_tooltips():
    st.markdown("""
    <div class="ui-help">
        <div class="ui-help-title">🛡️ Safe UI Upgrade Notice</div>
        <div class="ui-help-text">
            This screen uses the upgraded UI layer only. Your IMS parsing, CN/DN handling, reconciliation conditions, action mapping and GST JSON generation remain controlled by the existing backend functions.
        </div>
    </div>
    """, unsafe_allow_html=True)


def v9_help_box(title, text):
    st.markdown(f"""
    <div class="ui-help">
        <div class="ui-help-title">{title}</div>
        <div class="ui-help-text">{text}</div>
    </div>
    """, unsafe_allow_html=True)


def v9_saleable_kpis():
    p = st.session_state.get("purchase_df", pd.DataFrame())
    ims = st.session_state.get("ims_df", pd.DataFrame())
    recon = st.session_state.get("recon_df", pd.DataFrame())
    action = st.session_state.get("action_df", pd.DataFrame())
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        metric_card("💰", "Purchase Tax", _ui_money(_ui_sum(p, "total_tax")), "from books", "#fff3e8", "#ff8a00")
    with c2:
        metric_card("📥", "IMS Tax", _ui_money(_ui_sum(ims, "total_tax")), "GST portal", "#e9f9ed", "#138808")
    with c3:
        metric_card("🧩", "Mismatch Rows", f"{max(_ui_df_len(recon) - _ui_recon_count(recon, 'Matched'), 0):,}", "needs review", "#fff0ed", "#dc463f", True)
    with c4:
        metric_card("✅", "Accepted Actions", f"{_ui_action_count(action, 'Accepted'):,}", "JSON A", "#edf4ff", "#2563eb")


def v9_home_modules():
    st.markdown("""
    <div class="v9-module-grid">
        <div class="v9-module-card"><div class="v9-module-icon">🏢</div><div class="v9-module-title">Client Control</div><div class="v9-module-desc">Set GSTIN, period and workspace controls before upload.</div><div class="v9-module-badge">Setup</div></div>
        <div class="v9-module-card"><div class="v9-module-icon">📤</div><div class="v9-module-title">Upload Quality Room</div><div class="v9-module-desc">Review records, taxable value, tax heads and data readiness.</div><div class="v9-module-badge">Validate</div></div>
        <div class="v9-module-card"><div class="v9-module-icon">🔄</div><div class="v9-module-title">Reco Engine</div><div class="v9-module-desc">Run your existing matching engine with tolerance controls.</div><div class="v9-module-badge">Reconcile</div></div>
        <div class="v9-module-card"><div class="v9-module-icon">✅</div><div class="v9-module-title">IMS Action Manager</div><div class="v9-module-desc">Filter, edit and finalize invoice-wise actions and remarks.</div><div class="v9-module-badge">Act</div></div>
        <div class="v9-module-card"><div class="v9-module-icon">⚠️</div><div class="v9-module-title">Risk & ITC Desk</div><div class="v9-module-desc">Focus on critical mismatch and ITC recovery risk items.</div><div class="v9-module-badge">Review</div></div>
        <div class="v9-module-card"><div class="v9-module-icon">🧾</div><div class="v9-module-title">Final GST JSON</div><div class="v9-module-desc">Generate portal upload JSON only after final sign-off.</div><div class="v9-module-badge">Export</div></div>
    </div>
    """, unsafe_allow_html=True)


def v9_json_readiness_panel():
    ims_json_ok = bool(st.session_state.get("ims_json_data"))
    action_ok = isinstance(st.session_state.get("action_df", pd.DataFrame()), pd.DataFrame) and not st.session_state.get("action_df", pd.DataFrame()).empty
    final_ok = bool(st.session_state.get("final_json_bytes", b""))
    st.markdown(f"""
    <div class="v9-readiness">
        <div class="v9-readiness-title">🧾 Final JSON Safety Readiness</div>
        <div class="v9-check-grid">
            <div class="v9-check"><div class="v9-check-icon">{"✅" if ims_json_ok else "⏳"}</div><div class="v9-check-label">Original IMS JSON<br>{"Loaded" if ims_json_ok else "Pending"}</div></div>
            <div class="v9-check"><div class="v9-check-icon">{"✅" if action_ok else "⏳"}</div><div class="v9-check-label">Final Actions<br>{"Available" if action_ok else "Pending"}</div></div>
            <div class="v9-check"><div class="v9-check-icon">🛡️</div><div class="v9-check-label">GST Schema<br>Protected</div></div>
            <div class="v9-check"><div class="v9-check-icon">{"✅" if final_ok else "⏳"}</div><div class="v9-check-label">Download JSON<br>{"Ready" if final_ok else "After Sign-off"}</div></div>
        </div>
    </div>
    """, unsafe_allow_html=True)


def v9_report_cards():
    st.markdown("""
    <div class="v9-report-grid">
        <div class="v9-report-card"><div class="v9-report-title">📊 Complete Workpaper</div><div class="v9-report-desc">Purchase, IMS, reconciliation and action sheets in Excel format.</div></div>
        <div class="v9-report-card"><div class="v9-report-title">⚠️ Exception Review</div><div class="v9-report-desc">High risk, mismatch and pending cases for management sign-off.</div></div>
        <div class="v9-report-card"><div class="v9-report-title">🧬 GST JSON Output</div><div class="v9-report-desc">Portal-ready output generated from the protected backend function.</div></div>
    </div>
    """, unsafe_allow_html=True)


def v10_quality_dashboard():
    p = st.session_state.get("purchase_df", pd.DataFrame())
    ims = st.session_state.get("ims_df", pd.DataFrame())

    def render_quality(title, df, icon):
        records = _ui_df_len(df)
        taxable = _ui_sum(df, "taxable_value")
        invoice_value = _ui_sum(df, "invoice_value")
        igst = _ui_sum(df, "igst")
        cgst = _ui_sum(df, "cgst")
        sgst = _ui_sum(df, "sgst")
        cess = _ui_sum(df, "cess")
        total_tax = igst + cgst + sgst + cess
        score = v10_quality_score(df) if "v10_quality_score" in globals() else (100 if records else 0)

        st.markdown(f"#### {icon} {title}")
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Quality", f"{score}%", "ready" if records else "upload pending")
        c2.metric("Records", f"{records:,}")
        c3.metric("Taxable", _ui_money(taxable))
        c4.metric("Invoice Value", _ui_money(invoice_value))
        c5, c6, c7, c8 = st.columns(4)
        c5.metric("IGST", _ui_money(igst))
        c6.metric("CGST", _ui_money(cgst))
        c7.metric("SGST", _ui_money(sgst))
        c8.metric("CESS", _ui_money(cess))
        st.metric("Total Tax", _ui_money(total_tax))

    page_title("Upload Quality Room", "Data health, tax summary and readiness before reconciliation.")
    left, right = st.columns(2)
    with left:
        with st.container(border=True):
            render_quality("Purchase Register Quality", p, "📚")
    with right:
        with st.container(border=True):
            render_quality("IMS JSON Quality", ims, "🧬")


def v10_reco_control_room():
    p = st.session_state.get("purchase_df", pd.DataFrame())
    ims = st.session_state.get("ims_df", pd.DataFrame())
    recon = st.session_state.get("recon_df", pd.DataFrame())
    st.markdown(f"""
    <div class="ui-command">
        <div class="ui-command-title">🔄 Reconciliation Control Room</div>
        <div class="ui-command-sub">Use your existing matching rules with controlled tolerance. Start reconciliation only when both datasets are loaded.</div>
        <div class="ui-action-grid">
            <div class="ui-action-card"><div class="ui-action-icon">📚</div><div class="ui-action-title">Purchase Register</div><div class="ui-action-desc">{_ui_df_len(p):,} records loaded</div></div>
            <div class="ui-action-card"><div class="ui-action-icon">🧬</div><div class="ui-action-title">IMS JSON</div><div class="ui-action-desc">{_ui_df_len(ims):,} records loaded</div></div>
            <div class="ui-action-card"><div class="ui-action-icon">✅</div><div class="ui-action-title">Matched</div><div class="ui-action-desc">{_ui_recon_count(recon, "Matched"):,} records matched</div></div>
            <div class="ui-action-card"><div class="ui-action-icon">⚠️</div><div class="ui-action-title">Exceptions</div><div class="ui-action-desc">{max(_ui_df_len(recon) - _ui_recon_count(recon, "Matched"), 0):,} records need review</div></div>
        </div>
    </div>
    """, unsafe_allow_html=True)


def v10_action_header():
    action = st.session_state.get("action_df", pd.DataFrame())
    st.markdown(f"""
    <div class="ui-command">
        <div class="ui-command-title">✅ IMS Bulk Action Manager</div>
        <div class="ui-command-sub">Review, filter and finalize invoice-wise actions. The UI helps you control actions; the backend action-code mapping remains your existing logic.</div>
        <div class="ui-action-grid">
            <div class="ui-action-card"><div class="ui-action-icon">✅</div><div class="ui-action-title">Accepted</div><div class="ui-action-desc">{_ui_action_count(action, "Accepted"):,} records</div></div>
            <div class="ui-action-card"><div class="ui-action-icon">📌</div><div class="ui-action-title">Pending</div><div class="ui-action-desc">{_ui_action_count(action, "Pending"):,} records</div></div>
            <div class="ui-action-card"><div class="ui-action-icon">🚫</div><div class="ui-action-title">Rejected</div><div class="ui-action-desc">{_ui_action_count(action, "Rejected"):,} records</div></div>
            <div class="ui-action-card"><div class="ui-action-icon">🕘</div><div class="ui-action-title">No Action / Review</div><div class="ui-action-desc">{(_ui_action_count(action, "No Action") + _ui_action_count(action, "Review")):,} records</div></div>
        </div>
    </div>
    """, unsafe_allow_html=True)


def v10_final_json_review_ui():
    action = st.session_state.get("action_df", pd.DataFrame())
    accepted = _ui_action_count(action, "Accepted")
    pending = _ui_action_count(action, "Pending")
    rejected = _ui_action_count(action, "Rejected")
    no_action = _ui_action_count(action, "No Action")
    review = _ui_action_count(action, "Review")
    ims_sections = _ui_safe_unique(st.session_state.get("ims_df", pd.DataFrame()), "ims_sheet")
    st.markdown(f"""
    <div class="v10-json-review">
        <div class="v10-json-title">🧾 Final GST Upload JSON Review</div>
        <div class="v10-json-checks">
            <div class="v10-json-check"><div class="v10-json-check-icon">✅</div><div class="v10-json-check-label">Accepted<br>{accepted:,}</div></div>
            <div class="v10-json-check"><div class="v10-json-check-icon">🟠</div><div class="v10-json-check-label">Pending<br>{pending:,}</div></div>
            <div class="v10-json-check"><div class="v10-json-check-icon">🔴</div><div class="v10-json-check-label">Rejected<br>{rejected:,}</div></div>
            <div class="v10-json-check"><div class="v10-json-check-icon">📦</div><div class="v10-json-check-label">IMS Sections<br>{ims_sections:,}</div></div>
            <div class="v10-json-check"><div class="v10-json-check-icon">🛡️</div><div class="v10-json-check-label">JSON Engine<br>Untouched</div></div>
        </div>
    </div>
    """, unsafe_allow_html=True)


def v10_management_summary():
    recon = st.session_state.get("recon_df", pd.DataFrame())
    action = st.session_state.get("action_df", pd.DataFrame())
    total = _ui_df_len(recon)
    matched = _ui_recon_count(recon, "Matched")
    pending = _ui_action_count(action, "Pending")
    accepted = _ui_action_count(action, "Accepted")
    high = 0
    try:
        if isinstance(recon, pd.DataFrame) and not recon.empty and "risk_level" in recon.columns:
            high = int(recon["risk_level"].isin(["High", "Critical"]).sum())
    except Exception:
        high = 0
    text = (
        f"Total reconciliation rows are {total:,}. Matched rows are {matched:,}. "
        f"Accepted action rows are {accepted:,}, pending rows are {pending:,}, and high/critical risk rows are {high:,}. "
        "Please complete review and sign-off before final GST JSON generation."
        if total else
        "Upload Purchase Register and IMS JSON, then run reconciliation to generate a management summary."
    )
    st.markdown(f"""
    <div class="v10-management-summary">
        <div class="v10-management-title">📋 Management Summary</div>
        <div class="v10-management-text">{text}</div>
    </div>
    """, unsafe_allow_html=True)


def login_page():
    st.markdown("""
    <div class="ui-login-bg">
        <div class="ui-login-card">
            <div class="ui-login-left">
                <div class="ui-login-logo">☸</div>
                <div class="ui-login-title">IMS Recon Pro<br>Compliance Workspace</div>
                <div class="ui-login-sub">
                    Premium GST IMS reconciliation interface for upload validation, mismatch review,
                    action control and final GST upload JSON preparation.
                </div>
                <div class="ui-login-feature-grid">
                    <div class="ui-login-feature"><div class="ui-login-feature-title">🧬 IMS JSON Ready</div><div class="ui-login-feature-text">Built for B2B, amendments, CN and DN sections.</div></div>
                    <div class="ui-login-feature"><div class="ui-login-feature-title">🛡️ Safe Engine</div><div class="ui-login-feature-text">UI layer is separated from GST calculation logic.</div></div>
                    <div class="ui-login-feature"><div class="ui-login-feature-title">✅ Action Control</div><div class="ui-login-feature-text">Accepted, Pending, Rejected, Review and No Action handling.</div></div>
                    <div class="ui-login-feature"><div class="ui-login-feature-title">🇮🇳 India Theme</div><div class="ui-login-feature-text">Designed for Indian GST compliance teams.</div></div>
                </div>
            </div>
            <div class="ui-login-panel">
                <div class="ui-login-panel-title">Welcome</div>
                <div class="ui-login-panel-sub">Secure compliance access panel<br>Login to continue your IMS reconciliation workflow.</div>
    """, unsafe_allow_html=True)

    username = st.text_input("User ID", placeholder="Enter your User ID")
    password = st.text_input("Password", type="password", placeholder="Enter password")

    if st.button("🔐 Login Securely", use_container_width=True):
        user = USER_MASTER.get(username)
        if user and password == user["password"]:
            st.session_state.logged_in = True
            st.session_state.username = username
            st.session_state.role = user["role"]
            st.session_state.display_name = user["name"]
            log_event("Login", "Successful login")
            load_user_state()
            st.rerun()
        else:
            st.error("Invalid User ID or Password. Please check case-sensitive credentials.")

    st.markdown(f"""
                <div class="ui-login-copy">© {COPYRIGHT_OWNER} • IMS Recon Pro • GST IMS Reconciliation Workspace</div>
            </div>
        </div>
    </div>
    """, unsafe_allow_html=True)



if __name__ == "__main__":
    main()
