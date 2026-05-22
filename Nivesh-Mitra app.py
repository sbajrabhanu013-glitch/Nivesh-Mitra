import json
import base64
import re
import sqlite3
import pickle
import logging
from dataclasses import dataclass
from datetime import datetime
from io import BytesIO
from copy import deepcopy
from pathlib import Path
from typing import Dict, List, Optional, Tuple, Any

import pandas as pd
import numpy as np
import streamlit as st
from pydantic import BaseModel, Field
from rapidfuzz import process, fuzz
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font

# =========================================================
# ⚙️ ENTERPRISE CONFIGURATION & LOGGING
# =========================================================
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger("IMS_Recon_Pro")

class AppConfig:
    TITLE = "IMS Recon Pro Enterprise"
    TAGLINE = "Intelligent GST IMS Reconciliation & Action Management Platform"
    COPYRIGHT = "@BAJRABHANU"
    DB_NAME = "ims_recon_pro_enterprise.db"
    VERSION = "2026.05.09-V11-AI-FUZZY-OOP"
    
    IMS_SHEETS = ["B2B", "B2BA", "B2B-DN", "B2B-DNA", "B2B-CN", "B2B-CNA"]
    ACTION_VALUES = ["No Action", "Accepted", "Rejected", "Pending", "Review"]
    MONEY_COLS = ["invoice_value", "taxable_value", "igst", "cgst", "sgst", "cess"]
    TAX_COLS = ["igst", "cgst", "sgst", "cess"]
    
    USER_MASTER = {
        "Admin": {"password": "Admin", "role": "Admin", "name": "System Admin"},
        "User_1": {"password": "User1", "role": "Analyst", "name": "User-1"},
    }

class ReconSettings(BaseModel):
    """Strict type validation for reconciliation engine settings."""
    amount_tolerance: float = Field(default=5.0, ge=0.0)
    date_tolerance: int = Field(default=2, ge=0)
    include_amendments: bool = Field(default=True)
    fuzzy_threshold: float = Field(default=85.0, ge=0.0, le=100.0)

# =========================================================
# 🗄️ DATABASE MANAGER
# =========================================================
class DatabaseManager:
    def __init__(self, db_path: str = AppConfig.DB_NAME):
        self.db_path = db_path
        self._init_db()

    def get_conn(self):
        return sqlite3.connect(self.db_path, check_same_thread=False)

    def _init_db(self):
        with self.get_conn() as conn:
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

    def save_state(self, username: str, key: str, value: Any):
        blob = pickle.dumps(value)
        with self.get_conn() as conn:
            conn.execute(
                "INSERT OR REPLACE INTO user_data (username, key, value, updated_at) VALUES (?, ?, ?, ?)",
                (username, key, blob, datetime.now().isoformat(timespec="seconds")),
            )

    def load_state(self, username: str, key: str, default: Any = None):
        try:
            with self.get_conn() as conn:
                row = conn.execute("SELECT value FROM user_data WHERE username=? AND key=?", (username, key)).fetchone()
            return pickle.loads(row[0]) if row else default
        except Exception as e:
            logger.error(f"DB Load Error: {e}")
            return default

    def log_event(self, username: str, event_type: str, detail: str):
        try:
            with self.get_conn() as conn:
                conn.execute(
                    "INSERT INTO audit_log (username, event_time, event_type, detail) VALUES (?, ?, ?, ?)",
                    (username, datetime.now().isoformat(timespec="seconds"), event_type, detail),
                )
        except Exception as e:
            logger.error(f"Audit Log Error: {e}")

# =========================================================
# 🧠 AI RECONCILIATION ENGINE
# =========================================================
class ReconciliationEngine:
    """Enterprise-Grade Reconciliation Engine using Vectorization and Fuzzy Matching."""
    
    def __init__(self, purchase_df: pd.DataFrame, ims_df: pd.DataFrame, settings: ReconSettings):
        self.purchase = purchase_df.copy()
        self.ims = ims_df.copy()
        self.settings = settings
        self.money_cols = AppConfig.MONEY_COLS + ["total_tax"]

    def _preprocess(self) -> Tuple[pd.DataFrame, pd.DataFrame]:
        logger.info("Vectorizing and aggregating data...")
        if not self.settings.include_amendments and "ims_sheet" in self.ims.columns:
            amendments = ["B2BA", "B2B-DNA", "B2B-CNA", "B2BDN", "B2BCN"]
            self.ims = self.ims[~self.ims["ims_sheet"].astype(str).str.upper().isin(amendments)]

        agg_rules = {col: "sum" for col in self.money_cols}
        agg_rules.update({"supplier_name": "first", "document_type": "first", "document_no": "first", "document_date": "min"})

        p_agg = self.purchase.groupby(["supplier_gstin", "document_norm"], dropna=False).agg(agg_rules).reset_index()
        i_agg = self.ims.groupby(["supplier_gstin", "document_norm"], dropna=False).agg(agg_rules).reset_index()

        return p_agg.add_suffix("_purchase"), i_agg.add_suffix("_ims")

    def _apply_fuzzy_matching(self, merged: pd.DataFrame) -> pd.DataFrame:
        logger.info("Running NLP Fuzzy Matching algorithms...")
        unmatched_ims = merged[merged["_merge"] == "right_only"].copy()
        unmatched_pur = merged[merged["_merge"] == "left_only"].copy()

        if unmatched_ims.empty or unmatched_pur.empty:
            return merged

        pur_invoices = unmatched_pur["document_norm_purchase"].dropna().tolist()

        def get_best_match(ims_inv):
            if pd.isna(ims_inv): return None, 0
            match = process.extractOne(ims_inv, pur_invoices, scorer=fuzz.WRatio)
            return (match[0], match[1]) if match else (None, 0)

        unmatched_ims[["fuzzy_match", "fuzzy_score"]] = unmatched_ims["document_norm_ims"].apply(
            lambda x: pd.Series(get_best_match(x))
        )
        
        typo_mask = unmatched_ims["fuzzy_score"] >= self.settings.fuzzy_threshold
        typo_indices = unmatched_ims[typo_mask].index
        
        merged.loc[typo_indices, "mismatch_type"] = "AI Probable Match - Typo"
        merged.loc[typo_indices, "risk_level"] = "Medium"
        merged.loc[typo_indices, "recommended_action"] = "Pending"
        
        return merged

    def run(self) -> pd.DataFrame:
        logger.info("Executing High-Speed Reconciliation...")
        p_agg, i_agg = self._preprocess()
        
        p_agg.rename(columns={"supplier_gstin_purchase": "supplier_gstin", "document_norm_purchase": "document_norm"}, inplace=True)
        i_agg.rename(columns={"supplier_gstin_ims": "supplier_gstin", "document_norm_ims": "document_norm"}, inplace=True)

        merged = pd.merge(p_agg, i_agg, on=["supplier_gstin", "document_norm"], how="outer", indicator=True)

        # Vectorized Math
        for col in self.money_cols:
            merged[f"{col}_diff"] = merged[f"{col}_purchase"].fillna(0) - merged[f"{col}_ims"].fillna(0)

        merged["date_diff_days"] = (pd.to_datetime(merged["document_date_purchase"], errors="coerce") - 
                                    pd.to_datetime(merged["document_date_ims"], errors="coerce")).dt.days.abs().fillna(9999)

        # Vectorized Categorization
        conditions = [
            merged["_merge"] == "left_only",
            merged["_merge"] == "right_only",
            (merged["taxable_value_diff"].abs() <= self.settings.amount_tolerance) & (merged["date_diff_days"] <= self.settings.date_tolerance),
            (merged["taxable_value_diff"].abs() <= self.settings.amount_tolerance) & (merged["date_diff_days"] > self.settings.date_tolerance)
        ]
        choices = ["Only in Purchase Register", "Only in IMS", "Matched", "Date Mismatch"]
        merged["mismatch_type"] = np.select(conditions, choices, default="Value / Tax Mismatch")

        # Risk Classification
        merged["risk_level"] = np.where(
            (merged["mismatch_type"] == "Only in IMS") | (merged["total_tax_diff"].abs() > 50000), "Critical",
            np.where(merged["mismatch_type"] == "Matched", "Low", "Medium")
        )

        merged["recommended_action"] = np.where(merged["mismatch_type"] == "Matched", "Accepted", "Pending")

        merged = self._apply_fuzzy_matching(merged)
        return merged.reset_index(drop=True)

# =========================================================
# 🎨 ENTERPRISE UI MANAGER
# =========================================================
class UIManager:
    @staticmethod
    def inject_premium_css():
        st.markdown("""
        <style>
            :root { --navy:#061a3e; --blue:#2563eb; --saffron:#ff9933; --green:#138808; --bg1:#d8e7f7; --text:#102244; }
            .stApp { background: linear-gradient(135deg, #dbeafe 0%, #c7d9ef 48%, #e4eef9 100%); color: var(--text); }
            header[data-testid="stHeader"] { background: rgba(224, 236, 249, 0.88) !important; backdrop-filter: blur(12px); }
            .block-container { padding-top: 1rem; max-width: 1600px; }
            
            /* Premium V10 Cards */
            .v10-command-center { background: linear-gradient(135deg,#071a3d,#0b3677); color:#ffffff; border-radius:28px; padding:24px; box-shadow:0 20px 48px rgba(7,26,61,0.22); margin: 14px 0 20px 0; }
            .v10-action-grid { display:grid; grid-template-columns:repeat(4,minmax(0,1fr)); gap:14px; }
            .v10-action-card { background:rgba(255,255,255,0.10); border:1px solid rgba(255,255,255,0.16); border-radius:22px; padding:18px; backdrop-filter: blur(4px); }
            .metric-card { background: #ffffff; border-radius: 24px; padding: 20px 22px; box-shadow: 0 14px 35px rgba(7,26,61,0.12); display:flex; align-items:center; gap:15px; }
            .metric-value { font-size:33px; font-weight:900; color:#112244; }
        </style>
        """, unsafe_allow_html=True)

    @staticmethod
    def render_metric(icon: str, label: str, value: str, subtext: str = ""):
        st.markdown(f"""
        <div class='metric-card'>
            <div style='font-size:30px;'>{icon}</div>
            <div>
                <div style='font-size:13px; font-weight:800; color:#60748f; text-transform:uppercase;'>{label}</div>
                <div class='metric-value'>{value}</div>
                <div style='font-size:12px; color:#12a150;'>{subtext}</div>
            </div>
        </div>
        """, unsafe_allow_html=True)

# =========================================================
# 🏗️ APP STATE & ROUTING
# =========================================================
db = DatabaseManager()

def init_session():
    defaults = {
        "logged_in": False, "username": "", "role": "", "page": "Dashboard",
        "purchase_df": pd.DataFrame(), "ims_df": pd.DataFrame(), 
        "recon_df": pd.DataFrame(), "action_df": pd.DataFrame()
    }
    for k, v in defaults.items():
        st.session_state.setdefault(k, v)

def login_page():
    st.markdown("<h2 style='text-align:center;'>Welcome to IMS Recon Pro Enterprise</h2>", unsafe_allow_html=True)
    with st.container(border=True):
        user = st.text_input("User ID")
        pwd = st.text_input("Password", type="password")
        if st.button("Secure Login", use_container_width=True):
            if user in AppConfig.USER_MASTER and pwd == AppConfig.USER_MASTER[user]["password"]:
                st.session_state.update({"logged_in": True, "username": user, "role": AppConfig.USER_MASTER[user]["role"]})
                db.log_event(user, "Login", "Success")
                st.rerun()
            else:
                st.error("Invalid Credentials")

def dashboard_page():
    st.markdown("""
    <div class='v10-command-center'>
        <h2>⚡ IMS Recon Pro Command Center</h2>
        <p>Enterprise AI-Powered Workflow for GST Compliance.</p>
    </div>
    """, unsafe_allow_html=True)
    
    c1, c2, c3, c4 = st.columns(4)
    with c1: UIManager.render_metric("📚", "Purchase Records", f"{len(st.session_state.purchase_df):,}", "Books data")
    with c2: UIManager.render_metric("📥", "IMS Records", f"{len(st.session_state.ims_df):,}", "Portal data")
    
    if not st.session_state.recon_df.empty:
        matched = len(st.session_state.recon_df[st.session_state.recon_df['mismatch_type'] == 'Matched'])
        with c3: UIManager.render_metric("✅", "Matched", f"{matched:,}", "AI Verified")
    
    st.markdown("### Next Actions")
    col1, col2 = st.columns(2)
    with col1:
        if st.button("📤 Upload Data Workspace", use_container_width=True):
            st.session_state.page = "Upload"
            st.rerun()
    with col2:
        if st.button("🔄 Run AI Reconciliation", use_container_width=True):
            st.session_state.page = "Recon"
            st.rerun()

def main():
    st.set_page_config(page_title=AppConfig.TITLE, layout="wide", initial_sidebar_state="collapsed")
    UIManager.inject_premium_css()
    init_session()

    if not st.session_state.logged_in:
        login_page()
        return

    st.sidebar.title("Navigation")
    pages = ["Dashboard", "Upload", "Recon", "Export"]
    for p in pages:
        if st.sidebar.button(p, use_container_width=True):
            st.session_state.page = p
            st.rerun()

    if st.session_state.page == "Dashboard":
        dashboard_page()
    elif st.session_state.page == "Upload":
        st.write("## 📤 Upload Enterprise Data")
        # File uploaders logic goes here...
        st.info("Module ready for Integration.")
    elif st.session_state.page == "Recon":
        st.write("## 🔄 AI Reconciliation Engine")
        if st.button("Execute Vectorized Reconciliation", type="primary"):
            settings = ReconSettings()
            engine = ReconciliationEngine(st.session_state.purchase_df, st.session_state.ims_df, settings)
            st.session_state.recon_df = engine.run()
            st.success("High-Speed Reconciliation Complete.")
            st.dataframe(st.session_state.recon_df.head(100))

if __name__ == "__main__":
    main()
