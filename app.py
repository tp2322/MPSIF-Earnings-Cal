"""
MPSIF Earnings Calendar
A Streamlit app for tracking earnings dates across fund holdings.
Supports Fidelity CSV portfolio export with per-analyst assignment.
"""

import streamlit as st
import pandas as pd
import yfinance as yf
from datetime import datetime, date
import calendar
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import json
import os
import io

# ─────────────────────────────────────────────
# PAGE CONFIG
# ─────────────────────────────────────────────
st.set_page_config(
    page_title="MPSIF Earnings Calendar",
    page_icon="📅",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ─────────────────────────────────────────────
# CUSTOM CSS
# ─────────────────────────────────────────────
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;600&family=IBM+Plex+Sans:wght@300;400;500;600&display=swap');

    html, body, [class*="css"] { font-family: 'IBM Plex Sans', sans-serif; }

    .stApp { background-color: #0d1117; color: #e6edf3; }

    section[data-testid="stSidebar"] {
        background-color: #161b22;
        border-right: 1px solid #30363d;
    }
    section[data-testid="stSidebar"] .stMarkdown h2 {
        color: #58a6ff;
        font-family: 'IBM Plex Mono', monospace;
        font-size: 0.85rem;
        letter-spacing: 0.1em;
        text-transform: uppercase;
        margin-bottom: 0.5rem;
    }

    .main-header {
        font-family: 'IBM Plex Mono', monospace;
        font-size: 1.8rem;
        font-weight: 600;
        color: #58a6ff;
        letter-spacing: -0.02em;
        margin-bottom: 0;
    }
    .main-subheader {
        font-family: 'IBM Plex Sans', sans-serif;
        font-size: 0.95rem;
        color: #8b949e;
        margin-top: 0.2rem;
        margin-bottom: 1.5rem;
    }

    .metric-card {
        background: #161b22;
        border: 1px solid #30363d;
        border-radius: 8px;
        padding: 1rem 1.25rem;
        text-align: center;
    }
    .metric-value {
        font-family: 'IBM Plex Mono', monospace;
        font-size: 1.8rem;
        font-weight: 600;
        color: #58a6ff;
    }
    .metric-label {
        font-size: 0.75rem;
        color: #8b949e;
        text-transform: uppercase;
        letter-spacing: 0.08em;
    }

    .badge {
        display: inline-block;
        padding: 2px 8px;
        border-radius: 4px;
        font-size: 0.75rem;
        font-family: 'IBM Plex Mono', monospace;
        font-weight: 600;
    }
    .badge-soon    { background:#1f3a1f; color:#3fb950; border:1px solid #3fb950; }
    .badge-upcoming{ background:#1f2d3f; color:#58a6ff; border:1px solid #58a6ff; }
    .badge-past    { background:#1e1e1e; color:#8b949e; border:1px solid #30363d; }
    .badge-bmc     { background:#2d1f0e; color:#d29922; border:1px solid #d29922; }

    .cal-cell {
        background: #161b22;
        border: 1px solid #30363d;
        border-radius: 6px;
        padding: 0.4rem;
        min-height: 80px;
        vertical-align: top;
    }
    .cal-cell-today   { border-color: #58a6ff; background: #0d2044; }
    .cal-day-num      { font-family:'IBM Plex Mono',monospace; font-size:0.75rem; color:#8b949e; margin-bottom:4px; }
    .cal-day-num-today{ color:#58a6ff; font-weight:700; }
    .cal-event {
        background: #1f3a1f;
        color: #3fb950;
        border-radius: 3px;
        padding: 1px 4px;
        font-size: 0.7rem;
        font-family: 'IBM Plex Mono', monospace;
        margin-bottom: 2px;
        white-space: nowrap;
        overflow: hidden;
        text-overflow: ellipsis;
        max-width: 100%;
    }
    .cal-event-past { background:#1e1e1e; color:#8b949e; }

    .section-header {
        font-family: 'IBM Plex Mono', monospace;
        font-size: 0.8rem;
        color: #8b949e;
        text-transform: uppercase;
        letter-spacing: 0.12em;
        border-bottom: 1px solid #30363d;
        padding-bottom: 0.5rem;
        margin-bottom: 1rem;
    }

    .upload-box {
        background: #161b22;
        border: 2px dashed #30363d;
        border-radius: 12px;
        padding: 2.5rem 2rem;
        text-align: center;
        margin: 1.5rem 0;
    }
    .upload-box h3 {
        font-family: 'IBM Plex Mono', monospace;
        color: #58a6ff;
        margin-bottom: 0.5rem;
    }
    .upload-box p { color: #8b949e; font-size: 0.9rem; }

    .cache-pill {
        display: inline-block;
        padding: 3px 10px;
        border-radius: 20px;
        font-family: 'IBM Plex Mono', monospace;
        font-size: 0.72rem;
        background: #1f2d3f;
        color: #58a6ff;
        border: 1px solid #58a6ff;
        margin-left: 8px;
    }
    .cache-pill-warn {
        background: #2d1f0e;
        color: #d29922;
        border-color: #d29922;
    }

    #MainMenu {visibility:hidden;}
    footer      {visibility:hidden;}
    header      {visibility:hidden;}

    .stDataFrame { border:1px solid #30363d; border-radius:8px; overflow:hidden; }

    .stButton > button {
        background: #21262d;
        color: #e6edf3;
        border: 1px solid #30363d;
        border-radius: 6px;
        font-family: 'IBM Plex Mono', monospace;
        font-size: 0.8rem;
    }
    .stButton > button:hover {
        background: #30363d;
        border-color: #58a6ff;
        color: #58a6ff;
    }

    .stTabs [data-baseweb="tab-list"] { background:transparent; border-bottom:1px solid #30363d; }
    .stTabs [data-baseweb="tab"]      { font-family:'IBM Plex Mono',monospace; font-size:0.8rem; color:#8b949e; }
    .stTabs [aria-selected="true"]    { color:#58a6ff !important; border-bottom:2px solid #58a6ff !important; }
</style>
""", unsafe_allow_html=True)


# ─────────────────────────────────────────────
# CONSTANTS
# ─────────────────────────────────────────────
ANALYSTS = [
    "Augustine", "Zach", "Kartik", "Nihir",
    "Boid", "Jake", "Tejas", "Sina", "Leah", "Rachel",
    "Unassigned",
]

HOLDINGS_FILE       = "mpsif_holdings.json"
ANALYST_EMAILS_FILE = "mpsif_analyst_emails.json"
UPLOAD_META_FILE    = "mpsif_upload_meta.json"

DEFAULT_ANALYST_EMAILS = {a: f"{a.lower()}@stern.nyu.edu" for a in ANALYSTS if a != "Unassigned"}

# Symbols to skip in Fidelity exports (money market, cash, etc.)
SKIP_SYMBOLS = {"SPAXX**", "SPAXX", "FDRXX", "FZFXX", "FCASH"}


# ─────────────────────────────────────────────
# PERSISTENCE HELPERS
# ─────────────────────────────────────────────
def load_holdings():
    if os.path.exists(HOLDINGS_FILE):
        with open(HOLDINGS_FILE) as f:
            return json.load(f)
    return []

def save_holdings(holdings):
    with open(HOLDINGS_FILE, "w") as f:
        json.dump(holdings, f, indent=2)

def load_analyst_emails():
    if os.path.exists(ANALYST_EMAILS_FILE):
        with open(ANALYST_EMAILS_FILE) as f:
            return json.load(f)
    return DEFAULT_ANALYST_EMAILS.copy()

def save_analyst_emails(emails):
    with open(ANALYST_EMAILS_FILE, "w") as f:
        json.dump(emails, f, indent=2)

def load_upload_meta():
    if os.path.exists(UPLOAD_META_FILE):
        with open(UPLOAD_META_FILE) as f:
            return json.load(f)
    return {}

def save_upload_meta(meta):
    with open(UPLOAD_META_FILE, "w") as f:
        json.dump(meta, f, indent=2)


# ─────────────────────────────────────────────
# FIDELITY CSV PARSER
# ─────────────────────────────────────────────
def parse_fidelity_csv(file_bytes: bytes) -> pd.DataFrame:
    """
    Parse a Fidelity portfolio CSV export.
    Stops at the blank line that precedes the footer disclaimers.
    Returns cleaned DataFrame with equity positions only.
    """
    text = file_bytes.decode("utf-8-sig")   # strips BOM

    data_lines = []
    for line in text.splitlines():
        stripped = line.strip()
        if not stripped:
            break           # blank line = start of footer
        if stripped.startswith('"'):
            break           # quoted disclaimer
        data_lines.append(line)

    if not data_lines:
        raise ValueError("Could not find data rows in the uploaded CSV.")

    # index_col=False is critical: prevents pandas from treating the repeated
    # Account Number column as a row index, which shifts all columns left by one.
    df = pd.read_csv(io.StringIO("\n".join(data_lines)), index_col=False)

    # Normalise column names — strip whitespace and any BOM remnants
    df.columns = [c.strip().lstrip("\ufeff").strip() for c in df.columns]

    # Cast Symbol and Description to str early so .str accessor always works
    df["Symbol"]      = df["Symbol"].fillna("").astype(str).str.strip()
    df["Description"] = df["Description"].fillna("").astype(str).str.strip()

    # Drop non-equity rows (money market, pending activity, blank symbols)
    df = df[df["Symbol"] != ""]
    df = df[~df["Symbol"].isin(SKIP_SYMBOLS)]
    df = df[~df["Description"].str.lower().str.contains("pending|money market", na=False)]

    # Sanitise numeric columns — strip $, +, commas, % before casting
    for col in ["Last Price", "Current Value", "Percent Of Account", "Quantity"]:
        if col in df.columns:
            df[col] = (
                df[col].fillna("").astype(str)
                .str.replace(r"[$+,%]", "", regex=True)
                .str.strip()
                .replace("", float("nan"))
            )
            df[col] = pd.to_numeric(df[col], errors="coerce")

    # Title-case company names now that Description is guaranteed a str column
    df["Description"] = df["Description"].str.title()

    keep = ["Symbol", "Description", "Quantity", "Last Price",
            "Current Value", "Percent Of Account"]
    # Only keep columns that actually exist (future-proof against export format changes)
    keep = [c for c in keep if c in df.columns]
    return df[keep].reset_index(drop=True)


# ─────────────────────────────────────────────
# YAHOO FINANCE HELPERS
# ─────────────────────────────────────────────
@st.cache_data(ttl=3600, show_spinner=False)
def fetch_earnings_date(ticker: str):
    """
    Try four yfinance methods in order of reliability.
    Returns (date | None, call_time_str, is_estimate_bool).
    """
    try:
        tk = yf.Ticker(ticker)
        today_naive = date.today()
        today_ts    = pd.Timestamp.today(tz="UTC")

        # ── Method 1: calendar dict (most reliable for next event) ──
        try:
            cal = tk.calendar
            if isinstance(cal, dict):
                raw = cal.get("Earnings Date")
                if raw:
                    # can be a list of timestamps or a single value
                    if not isinstance(raw, (list, tuple)):
                        raw = [raw]
                    for val in raw:
                        try:
                            d = pd.Timestamp(val).date()
                            if d >= today_naive - pd.Timedelta(days=5).to_pytimedelta():
                                return d, "", True
                        except Exception:
                            pass
        except Exception:
            pass

        # ── Method 2: earnings_dates DataFrame ──
        try:
            edf = tk.earnings_dates
            if edf is not None and not edf.empty:
                window = today_ts + pd.Timedelta(days=400)
                future = edf[
                    (edf.index >= today_ts - pd.Timedelta(days=5)) &
                    (edf.index <= window)
                ]
                if not future.empty:
                    next_ts   = future.index.min()
                    eps_est   = future.loc[next_ts, "EPS Estimate"] if "EPS Estimate" in future.columns else None
                    is_est    = True if eps_est is None else pd.isna(eps_est)
                    call_time = ""
                    if "Call Time" in future.columns:
                        v = future.loc[next_ts, "Call Time"]
                        call_time = "" if (v is None or pd.isna(v)) else str(v)
                    return next_ts.date(), call_time, is_est
        except Exception:
            pass

        # ── Method 3: info dict timestamp fields ──
        try:
            info = tk.info or {}
            for key in ("earningsDate", "earningsTimestamp", "earningsTimestampStart"):
                raw = info.get(key)
                if raw is None:
                    continue
                if isinstance(raw, (list, tuple)):
                    raw = raw[0] if raw else None
                if raw:
                    try:
                        d = datetime.fromtimestamp(int(raw)).date()
                        if d >= today_naive:
                            return d, "", True
                    except Exception:
                        pass
        except Exception:
            pass

        # ── Method 4: fast_info (newer yfinance versions) ──
        try:
            fi = tk.fast_info
            raw = getattr(fi, "earnings_date", None) or getattr(fi, "earningsDate", None)
            if raw:
                d = pd.Timestamp(raw).date()
                if d >= today_naive:
                    return d, "", True
        except Exception:
            pass

    except Exception:
        pass

    return None, "", True


@st.cache_data(ttl=3600, show_spinner=False)
def fetch_stock_info(ticker: str):
    try:
        info = yf.Ticker(ticker).info
        return {
            "sector":     info.get("sector", "—"),
            "market_cap": info.get("marketCap"),
            "name":       info.get("shortName", ticker),
        }
    except Exception:
        return {"sector": "—", "market_cap": None, "name": ticker}


def days_until(d):
    return (d - date.today()).days if d else None

def fmt_market_cap(mc):
    if mc is None: return "—"
    if mc >= 1e12: return f"${mc/1e12:.1f}T"
    if mc >= 1e9:  return f"${mc/1e9:.1f}B"
    return f"${mc/1e6:.0f}M"

def earnings_badge(d):
    if d is None:
        return '<span class="badge badge-past">TBD</span>'
    n = days_until(d)
    if n is None:  return '<span class="badge badge-past">TBD</span>'
    if n < 0:      return f'<span class="badge badge-past">{d.strftime("%b %d")}</span>'
    if n <= 7:     return f'<span class="badge badge-soon">in {n}d</span>'
    if n <= 30:    return f'<span class="badge badge-upcoming">in {n}d</span>'
    return f'<span class="badge badge-bmc">{d.strftime("%b %d")}</span>'


# ─────────────────────────────────────────────
# EMAIL HELPER
# ─────────────────────────────────────────────
def send_earnings_reminder(smtp_host, smtp_port, smtp_user, smtp_pass,
                           to_email, analyst_name, holdings_rows):
    msg = MIMEMultipart("alternative")
    msg["Subject"] = f"MPSIF Earnings Reminder — {analyst_name}"
    msg["From"]    = smtp_user
    msg["To"]      = to_email

    rows_html = ""
    for row in holdings_rows:
        d = row["earnings_date"]
        n = days_until(d) if d else None
        days_str = (
            f"in {n} day{'s' if n != 1 else ''}"
            if (n is not None and n >= 0)
            else (d.strftime("%b %d") if d else "TBD")
        )
        rows_html += f"""
        <tr>
          <td style="padding:8px 12px;font-family:monospace;font-weight:600;">{row['ticker']}</td>
          <td style="padding:8px 12px;">{row['company']}</td>
          <td style="padding:8px 12px;font-family:monospace;">{d.strftime('%B %d, %Y') if d else 'TBD'}</td>
          <td style="padding:8px 12px;color:#3fb950;font-family:monospace;">{days_str}</td>
        </tr>"""

    html = f"""<html><body style="background:#0d1117;color:#e6edf3;
        font-family:'IBM Plex Sans',sans-serif;padding:24px;">
      <h2 style="color:#58a6ff;font-family:monospace;">MPSIF — Earnings Reminder</h2>
      <p>Hi {analyst_name},</p>
      <p>Here are the upcoming earnings dates for your holdings:</p>
      <table style="border-collapse:collapse;width:100%;background:#161b22;
          border:1px solid #30363d;border-radius:8px;">
        <thead>
          <tr style="background:#21262d;color:#8b949e;font-size:0.8rem;
              text-transform:uppercase;letter-spacing:0.08em;">
            <th style="padding:8px 12px;text-align:left;">Ticker</th>
            <th style="padding:8px 12px;text-align:left;">Company</th>
            <th style="padding:8px 12px;text-align:left;">Earnings Date</th>
            <th style="padding:8px 12px;text-align:left;">When</th>
          </tr>
        </thead>
        <tbody>{rows_html}</tbody>
      </table>
      <p style="color:#8b949e;font-size:0.85rem;margin-top:24px;">
        — MPSIF Earnings Calendar</p>
    </body></html>"""

    msg.attach(MIMEText(html, "html"))
    with smtplib.SMTP(smtp_host, int(smtp_port)) as server:
        server.starttls()
        server.login(smtp_user, smtp_pass)
        server.sendmail(smtp_user, to_email, msg.as_string())


# ─────────────────────────────────────────────
# SESSION STATE INIT
# ─────────────────────────────────────────────
if "holdings"       not in st.session_state:
    st.session_state.holdings       = load_holdings()
if "analyst_emails" not in st.session_state:
    st.session_state.analyst_emails = load_analyst_emails()
if "upload_meta"    not in st.session_state:
    st.session_state.upload_meta    = load_upload_meta()
if "show_upload"    not in st.session_state:
    st.session_state.show_upload    = len(st.session_state.holdings) == 0


# ─────────────────────────────────────────────
# HEADER
# ─────────────────────────────────────────────
st.markdown('<h1 class="main-header">MPSIF Earnings Calendar</h1>', unsafe_allow_html=True)

meta = st.session_state.upload_meta
if meta.get("filename"):
    st.markdown(
        f'<p class="main-subheader">Long-only equity fund · NYU Stern · '
        f'Positions from <span class="cache-pill">{meta["filename"]}</span> '
        f'<span class="cache-pill cache-pill-warn">uploaded {meta.get("uploaded_at","")}</span></p>',
        unsafe_allow_html=True,
    )
else:
    st.markdown(
        '<p class="main-subheader">Long-only equity fund · NYU Stern · '
        'Upload a Fidelity positions CSV to get started</p>',
        unsafe_allow_html=True,
    )


# ─────────────────────────────────────────────
# UPLOAD / IMPORT FLOW
# ─────────────────────────────────────────────
has_holdings = len(st.session_state.holdings) > 0

if has_holdings:
    col_toggle, _ = st.columns([2.5, 6])
    with col_toggle:
        btn_label = "✕ Cancel" if st.session_state.show_upload else "📂 Upload new positions file"
        if st.button(btn_label, key="toggle_upload"):
            st.session_state.show_upload = not st.session_state.show_upload
            st.rerun()

if st.session_state.show_upload:
    st.markdown("""
    <div class="upload-box">
      <h3>📂 Upload Positions File</h3>
      <p>Export your portfolio from Fidelity as a CSV and drop it here.<br>
      Go to <b>Accounts → Positions → Download (↓)</b> and choose <b>CSV</b> format.</p>
    </div>
    """, unsafe_allow_html=True)

    uploaded = st.file_uploader(
        "Drop your Fidelity CSV here",
        type=["csv"],
        label_visibility="collapsed",
        key="csv_uploader",
    )

    if uploaded is not None:
        try:
            parsed_df = parse_fidelity_csv(uploaded.read())
            new_holdings = [
                {"ticker": row["Symbol"], "company": row["Description"]}
                for _, row in parsed_df.iterrows()
            ]
            new_meta = {
                "filename":    uploaded.name,
                "uploaded_at": datetime.now().strftime("%b %d, %Y %H:%M ET"),
                "count":       len(new_holdings),
            }
            st.session_state.holdings    = new_holdings
            st.session_state.upload_meta = new_meta
            st.session_state.show_upload = False
            save_holdings(new_holdings)
            save_upload_meta(new_meta)
            st.rerun()
        except Exception as e:
            st.error(f"Could not parse file: {e}")

    st.divider()


# ─────────────────────────────────────────────
# GUARD: nothing loaded yet
# ─────────────────────────────────────────────
if not st.session_state.holdings:
    st.info("No positions loaded yet. Upload a Fidelity CSV above to get started.")
    st.stop()


# ─────────────────────────────────────────────
# SIDEBAR
# ─────────────────────────────────────────────
with st.sidebar:
    st.markdown("## 🏦 MPSIF")

    show_past = st.checkbox("Show past earnings", value=False)

    tickers_all = [h["ticker"] for h in st.session_state.holdings]

    st.markdown('<div class="section-header" style="margin-top:1.5rem;">Manual Edits</div>', unsafe_allow_html=True)

    with st.expander("➕ Add Position Manually"):
        new_ticker  = st.text_input("Ticker", key="new_ticker").upper().strip()
        new_company = st.text_input("Company Name", key="new_company")
        if st.button("Add"):
            if new_ticker:
                if new_ticker in tickers_all:
                    st.error("Ticker already exists.")
                else:
                    st.session_state.holdings.append({
                        "ticker": new_ticker, "company": new_company or new_ticker,
                    })
                    save_holdings(st.session_state.holdings)
                    st.success(f"Added {new_ticker}")
                    st.rerun()

    with st.expander("🗑️ Remove Position"):
        remove_ticker = st.selectbox("Select ticker", tickers_all, key="remove_ticker")
        if st.button("Remove"):
            st.session_state.holdings = [h for h in st.session_state.holdings if h["ticker"] != remove_ticker]
            save_holdings(st.session_state.holdings)
            st.success(f"Removed {remove_ticker}")
            st.rerun()

    with st.expander("⚠️ Clear All Data"):
        st.caption("Wipes cached holdings and forces a fresh CSV upload.")
        if st.button("Clear & Re-upload", type="primary"):
            st.session_state.holdings    = []
            st.session_state.upload_meta = {}
            st.session_state.show_upload = True
            for f in [HOLDINGS_FILE, UPLOAD_META_FILE]:
                if os.path.exists(f):
                    os.remove(f)
            st.rerun()

    st.markdown('<div class="section-header" style="margin-top:1.5rem;">Email Reminders</div>', unsafe_allow_html=True)

    with st.expander("📧 SMTP Settings"):
        smtp_host = st.text_input("SMTP Host", value="smtp.gmail.com", key="smtp_host")
        smtp_port = st.text_input("SMTP Port", value="587", key="smtp_port")
        smtp_user = st.text_input("From Email", key="smtp_user")
        smtp_pass = st.text_input("App Password", type="password", key="smtp_pass")

    with st.expander("📬 Member Emails"):
        updated_emails = {}
        for a in [x for x in ANALYSTS if x != "Unassigned"]:
            updated_emails[a] = st.text_input(
                a, value=st.session_state.analyst_emails.get(a, ""), key=f"email_{a}"
            )
        if st.button("Save Emails"):
            st.session_state.analyst_emails = updated_emails
            save_analyst_emails(updated_emails)
            st.success("Saved.")

    with st.expander("📨 Send Reminders"):
        send_to     = st.multiselect("Send to", [a for a in ANALYSTS if a != "Unassigned"],
                                     default=[a for a in ANALYSTS if a != "Unassigned"], key="send_to")
        days_filter = st.slider("Holdings with earnings within X days", 1, 90, 30)
        if st.button("Send Reminder Emails"):
            if not smtp_user or not smtp_pass:
                st.error("Please fill in SMTP settings above.")
            else:
                sent = failed = 0
                for analyst in send_to:
                    email = st.session_state.analyst_emails.get(analyst, "")
                    if not email:
                        continue
                    email_rows = []
                    for h in st.session_state.holdings:
                        ed, _, _ = fetch_earnings_date(h["ticker"])
                        n = days_until(ed)
                        if n is not None and 0 <= n <= days_filter:
                            email_rows.append({**h, "earnings_date": ed})
                    if not email_rows:
                        continue
                    try:
                        send_earnings_reminder(
                            smtp_host, smtp_port, smtp_user, smtp_pass,
                            email, analyst, email_rows,
                        )
                        sent += 1
                    except Exception as e:
                        failed += 1
                        st.error(f"Failed for {analyst}: {e}")
                st.success(f"Sent {sent} emails. {failed} failed.")


# ─────────────────────────────────────────────
# BUILD ENRICHED TABLE
# ─────────────────────────────────────────────
with st.spinner("Fetching earnings dates from Yahoo Finance…"):
    rows = []
    for h in st.session_state.holdings:
        ed, call_time, is_est = fetch_earnings_date(h["ticker"])
        rows.append({
            "ticker":        h["ticker"],
            "company":       h.get("company") or h["ticker"],
            "earnings_date": ed,
            "call_time":     call_time,
            "is_estimate":   is_est,
            "days_until":    days_until(ed),
        })


# ─────────────────────────────────────────────
# SUMMARY METRICS
# ─────────────────────────────────────────────
today      = date.today()
this_week  = [r for r in rows if r["days_until"] is not None and 0 <= r["days_until"] <= 7]
this_month = [r for r in rows if r["days_until"] is not None and 0 <= r["days_until"] <= 30]
no_date    = [r for r in rows if r["earnings_date"] is None]

c1, c2, c3, c4 = st.columns(4)
for col, val, label in [
    (c1, len(rows),       "Total Holdings"),
    (c2, len(this_week),  "Earnings This Week"),
    (c3, len(this_month), "Earnings This Month"),
    (c4, len(no_date),    "Date TBD"),
]:
    col.markdown(f"""
    <div class="metric-card">
      <div class="metric-value">{val}</div>
      <div class="metric-label">{label}</div>
    </div>""", unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)


# ─────────────────────────────────────────────
# TABS
# ─────────────────────────────────────────────
tab_list, tab_cal = st.tabs(["📋  List View", "📅  Calendar View"])


# ── LIST VIEW ────────────────────────────────
with tab_list:
    sort_col, _ = st.columns([2, 6])
    with sort_col:
        sort_by = st.selectbox(
            "Sort by",
            ["Earnings Date (soonest)", "Ticker A→Z"],
            label_visibility="collapsed",
            key="sort_by",
        )

    sort_map = {
        "Earnings Date (soonest)": "days_until",
        "Ticker A→Z":              "ticker",
    }
    sort_key = sort_map[sort_by]

    display_rows = [r for r in rows if show_past or r["days_until"] is None or r["days_until"] >= 0]
    display_rows = sorted(
        display_rows,
        key=lambda r: (r[sort_key] is None, r[sort_key])
        if sort_key != "days_until"
        else (r["days_until"] is None, r["days_until"] if r["days_until"] is not None else 9999),
    )

    if not display_rows:
        st.info("No holdings match current filters.")
    else:
        hcols = st.columns([1.5, 5, 2.5, 2])
        for col, hdr in zip(hcols, ["TICKER", "COMPANY", "EARNINGS DATE", "WHEN"]):
            col.markdown(
                f'<div class="section-header" style="margin-bottom:0.25rem;">{hdr}</div>',
                unsafe_allow_html=True,
            )
        st.markdown('<hr style="border:none;border-top:1px solid #30363d;margin:0 0 0.5rem 0;">', unsafe_allow_html=True)

        for r in display_rows:
            rcols = st.columns([1.5, 5, 2.5, 2])
            rcols[0].markdown(f"**`{r['ticker']}`**")
            rcols[1].markdown(r["company"])
            ed      = r["earnings_date"]
            est_tag = ' <small style="color:#d29922">~est</small>' if r["is_estimate"] and ed else ""
            rcols[2].markdown(
                f"{ed.strftime('%b %d, %Y') if ed else '—'}{est_tag}",
                unsafe_allow_html=True,
            )
            rcols[3].markdown(earnings_badge(ed), unsafe_allow_html=True)
            st.markdown('<hr style="border:none;border-top:1px solid #21262d;margin:0.25rem 0;">', unsafe_allow_html=True)


# ── CALENDAR VIEW ────────────────────────────
with tab_cal:
    cal_col1, cal_col2, _ = st.columns([1.5, 1.5, 5])
    with cal_col1:
        cal_month = st.selectbox(
            "Month", range(1, 13),
            index=today.month - 1,
            format_func=lambda m: calendar.month_name[m],
            label_visibility="collapsed",
            key="cal_month",
        )
    with cal_col2:
        cal_year = st.selectbox(
            "Year", range(today.year, today.year + 2),
            label_visibility="collapsed",
            key="cal_year",
        )

    event_map: dict = {}
    for r in rows:
        ed = r["earnings_date"]
        if ed and ed.month == cal_month and ed.year == cal_year:
            event_map.setdefault(ed, []).append(r["ticker"])

    cal_matrix = calendar.monthcalendar(cal_year, cal_month)
    day_names  = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]

    hcols = st.columns(7)
    for col, dn in zip(hcols, day_names):
        col.markdown(
            f'<div style="text-align:center;font-family:\'IBM Plex Mono\',monospace;'
            f'font-size:0.75rem;color:#8b949e;padding-bottom:4px;">{dn}</div>',
            unsafe_allow_html=True,
        )

    for week in cal_matrix:
        wcols = st.columns(7)
        for col, day in zip(wcols, week):
            if day == 0:
                col.markdown(
                    '<div class="cal-cell" style="background:transparent;border-color:transparent;"></div>',
                    unsafe_allow_html=True,
                )
                continue
            this_date  = date(cal_year, cal_month, day)
            is_today   = this_date == today
            cell_class = "cal-cell-today" if is_today else "cal-cell"
            day_class  = "cal-day-num-today" if is_today else "cal-day-num"
            events_html = "".join(
                f'<div class="cal-event{" cal-event-past" if this_date < today else ""}" '
                f'title="{t}">{t}</div>'
                for t in event_map.get(this_date, [])
            )
            col.markdown(
                f'<div class="{cell_class}"><div class="{day_class}">{day}</div>{events_html}</div>',
                unsafe_allow_html=True,
            )

    st.markdown("""
    <div style="margin-top:1rem;display:flex;gap:1rem;align-items:center;">
      <span style="font-size:0.75rem;color:#8b949e;">Legend:</span>
      <span class="cal-event" style="position:static;">TICKER — future</span>
      <span class="cal-event cal-event-past" style="position:static;">TICKER — past</span>
      <div style="width:12px;height:12px;border:1px solid #58a6ff;background:#0d2044;
          border-radius:2px;display:inline-block;"></div>
      <span style="font-size:0.75rem;color:#8b949e;">Today</span>
    </div>""", unsafe_allow_html=True)

    if event_map:
        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown('<div class="section-header">Events This Month</div>', unsafe_allow_html=True)
        for d in sorted(event_map.keys()):
            for ticker in event_map[d]:
                row = next((x for x in rows if x["ticker"] == ticker), {})
                st.markdown(
                    f"**`{ticker}`** &nbsp;·&nbsp; {row.get('company','')}"
                    f" &nbsp;·&nbsp; {d.strftime('%A, %B %d')} &nbsp; {earnings_badge(d)}",
                    unsafe_allow_html=True,
                )


# ─────────────────────────────────────────────
# FOOTER
# ─────────────────────────────────────────────
st.markdown("<br><br>", unsafe_allow_html=True)
st.markdown(
    '<div style="font-family:\'IBM Plex Mono\',monospace;font-size:0.7rem;color:#30363d;text-align:center;">'
    f'MPSIF Earnings Calendar · Data via Yahoo Finance · Earnings dates may be estimates · '
    f'Refreshed: {datetime.now().strftime("%Y-%m-%d %H:%M")} ET'
    '</div>',
    unsafe_allow_html=True,
)
