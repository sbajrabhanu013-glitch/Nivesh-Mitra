from __future__ import annotations

import os
import sqlite3
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Callable, Dict, List, Optional

import pandas as pd
import streamlit as st


# -----------------------------------------------------------------------------
# Configuration and constants
#

# Title and tagline displayed in the app header
APP_TITLE: str = "IMS Recon Pro"
APP_TAGLINE: str = "Intelligent GST IMS Reconciliation & Action Platform"

# Database location; can be overridden via environment variable
APP_DB: str = os.environ.get("IMS_DB", "ims_recon_pro.db")

# Session state defaults.  When adding new keys, update this dict
SESSION_DEFAULTS: Dict[str, object] = {
    "logged_in": False,
    "username": "",
    "role": "",
    "client_name": "",
    "client_gstin": "",
    "return_period": datetime.today().strftime("%b-%Y"),
    "purchase_df": pd.DataFrame(),
    "ims_df": pd.DataFrame(),
    "recon_df": pd.DataFrame(),
    "action_df": pd.DataFrame(),
}

# Basic user registry for authentication; in production this should come
# from a secure source such as a hashed credential store.  The role can
# be used to gate functionality (e.g. admin vs. standard user).
USER_MASTER: Dict[str, Dict[str, str]] = {
    "Admin": {"password": "Admin", "role": "Admin", "name": "Administrator"},
    "User_1": {"password": "User1", "role": "User", "name": "User 1"},
    "User_2": {"password": "User2", "role": "User", "name": "User 2"},
}

# Possible action values for reconciliation items
ACTION_VALUES: List[str] = [
    "No Action",
    "Accepted",
    "Rejected",
    "Pending",
    "Review",
]

# Columns that represent monetary values (for summaries)
MONEY_COLS: List[str] = [
    "invoice_value",
    "taxable_value",
    "igst",
    "cgst",
    "sgst",
    "cess",
]


# -----------------------------------------------------------------------------
# Database helper functions
#
def get_conn(db_path: str = APP_DB) -> sqlite3.Connection:
    """Return a connection to the SQLite database specified by ``db_path``.

    The connection uses ``check_same_thread=False`` so it can be accessed
    across Streamlit reruns.  Callers are responsible for closing
    connections where appropriate.
    """
    return sqlite3.connect(db_path, check_same_thread=False)


def init_db() -> None:
    """Initialise the local database with required tables if they do not exist."""
    with get_conn() as conn:
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS audit_log (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                username TEXT,
                event_time TEXT,
                event_type TEXT,
                detail TEXT
            )
            """
        )


def log_event(username: str, event_type: str, detail: str) -> None:
    """Record a user event in the audit log.

    Parameters
    ----------
    username: str
        Name of the user who triggered the event.
    event_type: str
        Type of event (e.g. "login", "load_file").
    detail: str
        Additional information about the event.
    """
    try:
        with get_conn() as conn:
            conn.execute(
                "INSERT INTO audit_log (username, event_time, event_type, detail)"
                " VALUES (?, ?, ?, ?)",
                (username, datetime.now().isoformat(timespec="seconds"), event_type, detail),
            )
    except Exception:
        # Swallow errors silently to avoid interfering with user flow
        pass


# -----------------------------------------------------------------------------
# Session state helpers
#
def initialise_session_state() -> None:
    """Ensure all keys defined in ``SESSION_DEFAULTS`` are initialised in the session."""
    for key, default in SESSION_DEFAULTS.items():
        if key not in st.session_state:
            # Avoid sharing references like DataFrame across sessions
            st.session_state[key] = default if not isinstance(default, pd.DataFrame) else default.copy()


def require_login() -> bool:
    """Return True if the user is authenticated; otherwise show the login form.

    This helper displays a simple login form when the user is not
    authenticated.  Valid credentials are looked up in ``USER_MASTER``.
    On successful login the session state is updated.
    """
    if st.session_state.get("logged_in"):
        return True

    st.title(APP_TITLE)
    st.write(APP_TAGLINE)
    st.subheader("Log in")
    username = st.text_input("Username")
    password = st.text_input("Password", type="password")
    if st.button("Login"):
        user = USER_MASTER.get(username)
        if user and password == user.get("password"):
            st.session_state.logged_in = True
            st.session_state.username = username
            st.session_state.role = user.get("role", "User")
            st.success(f"Welcome, {user.get('name', username)}!")
            log_event(username, "login", "User logged in")
            return True
        else:
            st.error("Invalid credentials")
    return False


# -----------------------------------------------------------------------------
# Data loading and reconciliation logic
#
@st.cache_data(show_spinner=False)
def load_purchase_file(data: bytes, filename: str) -> pd.DataFrame:
    """Load the purchase register into a DataFrame.

    Supports CSV and Excel (xls/xlsx) formats.  The DataFrame is
    standardised to include lower‑case column names for easier matching.
    """
    if filename.endswith(('.xls', '.xlsx')):
        df = pd.read_excel(data)
    else:
        df = pd.read_csv(data)
    df.columns = [c.strip().lower() for c in df.columns]
    return df


@st.cache_data(show_spinner=False)
def load_ims_file(data: bytes, filename: str) -> pd.DataFrame:
    """Load the IMS data into a DataFrame.

    Supports CSV and Excel formats.  Column names are normalised.
    """
    if filename.endswith(('.xls', '.xlsx')):
        df = pd.read_excel(data)
    else:
        df = pd.read_csv(data)
    df.columns = [c.strip().lower() for c in df.columns]
    return df


@st.cache_data(show_spinner=False)
def perform_reconciliation(purchase_df: pd.DataFrame, ims_df: pd.DataFrame) -> pd.DataFrame:
    """Return a DataFrame highlighting reconciliation results between purchase and IMS.

    The algorithm performs a basic inner join on document number (case‑insensitive)
    and computes differences in invoice value.  Unmatched purchase records and
    unmatched IMS records are flagged separately.  This function is cached to
    avoid recomputation when the inputs have not changed.
    """
    if purchase_df.empty or ims_df.empty:
        return pd.DataFrame()

    # Standardise keys used for matching
    left = purchase_df.copy()
    right = ims_df.copy()
    for df in (left, right):
        if "document_no" not in df.columns:
            # Fall back to 'inum' or similar typical invoice column names
            for alt in ("invoice number", "invoice no", "document number", "inum"):
                if alt.lower() in df.columns:
                    df.rename(columns={alt.lower(): "document_no"}, inplace=True)
                    break
        df["document_no"] = df["document_no"].astype(str).str.strip().str.lower()
        if "invoice_value" in df.columns:
            df["invoice_value"] = pd.to_numeric(df["invoice_value"], errors='coerce')

    # Inner join on document number
    merged = pd.merge(
        left,
        right,
        on="document_no",
        suffixes=("_pur", "_ims"),
        how="outer",
        indicator=True,
    )
    # Determine reconciliation status
    def compute_status(row):
        if row['_merge'] == 'both':
            # Compare invoice values if available
            if pd.isna(row.get('invoice_value_pur')) or pd.isna(row.get('invoice_value_ims')):
                return 'Matched'
            return 'Matched' if abs(row['invoice_value_pur'] - row['invoice_value_ims']) < 1e-2 else 'Value Mismatch'
        elif row['_merge'] == 'left_only':
            return 'Purchase Only'
        else:
            return 'IMS Only'

    merged['recon_status'] = merged.apply(compute_status, axis=1)
    return merged


def generate_json_placeholder(recon_df: pd.DataFrame) -> bytes:
    """Placeholder for JSON generation logic.

    The existing JSON generation logic from the original script should be
    integrated here.  This function simply serialises the reconciliation
    result to JSON bytes for demonstration purposes.  Replace this with
    your actual JSON generation when integrating.
    """
    return recon_df.to_json(orient="records").encode()


# -----------------------------------------------------------------------------
# UI rendering functions
#
def inject_custom_css() -> None:
    """Inject a minimal custom CSS to enhance the look and feel of the app."""
    css = """
    .stApp {
        background-color: #f7f9fc;
        color: #113366;
    }
    .stApp header {
        background-color: #ffffff;
        border-bottom: 1px solid #e0e6f1;
    }
    .metric-card {
        border: 1px solid #d6e4ff;
        border-radius: 16px;
        padding: 16px;
        background: #ffffff;
        box-shadow: 0 2px 6px rgba(0,0,0,0.05);
    }
    """
    st.markdown(f"<style>{css}</style>", unsafe_allow_html=True)


def show_dashboard() -> None:
    """Display high‑level metrics summarising the loaded data."""
    st.header("Dashboard")
    # Use columns to lay out metric cards
    cols = st.columns(3)
    purchase_count = len(st.session_state.purchase_df)
    ims_count = len(st.session_state.ims_df)
    recon_count = len(st.session_state.recon_df)
    with cols[0]:
        st.subheader("Purchase Records")
        st.metric(label="Count", value=purchase_count)
    with cols[1]:
        st.subheader("IMS Records")
        st.metric(label="Count", value=ims_count)
    with cols[2]:
        st.subheader("Reconciled Entries")
        st.metric(label="Count", value=recon_count)

    # Visualise reconciliation status distribution
    if not st.session_state.recon_df.empty and 'recon_status' in st.session_state.recon_df.columns:
        status_counts = st.session_state.recon_df['recon_status'].value_counts()
        st.bar_chart(status_counts)


def show_reconciliation_page() -> None:
    """Allow the user to upload files and perform reconciliation."""
    st.header("Data Upload & Reconciliation")
    st.write("Upload your purchase register and IMS files. Supported formats: CSV, XLS, XLSX.")
    pur_file = st.file_uploader("Purchase Register", type=["csv", "xls", "xlsx"], key="purchase_uploader")
    ims_file = st.file_uploader("IMS File", type=["csv", "xls", "xlsx"], key="ims_uploader")
    # Only act when both files are uploaded
    if pur_file and ims_file:
        with st.spinner("Loading files..."):
            st.session_state.purchase_df = load_purchase_file(pur_file, pur_file.name)
            st.session_state.ims_df = load_ims_file(ims_file, ims_file.name)
        st.success("Files loaded successfully.")
        if st.button("Perform Reconciliation"):
            st.session_state.recon_df = perform_reconciliation(
                st.session_state.purchase_df, st.session_state.ims_df
            )
            log_event(st.session_state.username, "reconcile", "Performed reconciliation")
            st.success("Reconciliation complete.")
        if not st.session_state.recon_df.empty:
            st.subheader("Reconciliation Results")
            # Provide simple filtering
            search_term = st.text_input("Filter by document number or status", key="recon_filter")
            recon = st.session_state.recon_df
            if search_term:
                term = search_term.strip().lower()
                recon = recon[recon['document_no'].str.contains(term) | recon['recon_status'].str.contains(term)]
            st.dataframe(recon)
            # Provide option to download JSON
            json_bytes = generate_json_placeholder(recon)
            st.download_button(
                label="Download JSON",
                data=json_bytes,
                file_name="recon_results.json",
                mime="application/json",
            )
    elif pur_file or ims_file:
        st.info("Please upload both files before proceeding.")


def show_actions_page() -> None:
    """Display reconciliation entries and allow the user to assign actions."""
    st.header("Review Actions")
    recon = st.session_state.recon_df.copy()
    if recon.empty:
        st.info("No reconciliation data available. Please perform reconciliation first.")
        return
    # Only show rows that require user action
    actionable = recon[recon['recon_status'] != 'Matched'].reset_index(drop=True)
    if actionable.empty:
        st.success("All records matched. No actions needed.")
        return
    st.write(f"{len(actionable)} record(s) require action.")
    # Create a select box for each record
    for idx, row in actionable.iterrows():
        with st.expander(f"Document {row['document_no']} - {row['recon_status']}"):
            st.write(row.drop(labels=['recon_status', '_merge']))
            action_key = f"action_{idx}"
            default_action = st.session_state.action_df.loc[idx, 'action'] if not st.session_state.action_df.empty else ACTION_VALUES[0]
            selected = st.selectbox("Select action", ACTION_VALUES, index=ACTION_VALUES.index(default_action), key=action_key)
            # Save the selected action into session state DataFrame
            if st.session_state.action_df.empty or idx >= len(st.session_state.action_df):
                # Extend DataFrame
                st.session_state.action_df = pd.concat([
                    st.session_state.action_df,
                    pd.DataFrame({'action': [selected]})
                ], ignore_index=True)
            else:
                st.session_state.action_df.at[idx, 'action'] = selected
    st.success("Actions recorded. You may implement further processing as needed.")


# Mapping of page names to functions
PAGES: Dict[str, Callable[[], None]] = {
    "Dashboard": show_dashboard,
    "Reconciliation": show_reconciliation_page,
    "Actions": show_actions_page,
}


def main() -> None:
    """Main entry point for the Streamlit application."""
    st.set_page_config(
        page_title=APP_TITLE,
        page_icon="🧮",
        layout="wide",
        initial_sidebar_state="expanded",
    )
    inject_custom_css()
    initialise_session_state()
    # Initialise database (no harm if called multiple times)
    init_db()
    if not require_login():
        return
    # Sidebar navigation
    page = st.sidebar.radio("Navigation", list(PAGES.keys()))
    PAGES[page]()


if __name__ == "__main__":
    main()
