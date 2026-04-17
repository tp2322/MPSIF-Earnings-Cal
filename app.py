"""
MPSIF Earnings Calendar
Light-theme redesign with EPS + revenue consensus estimates.
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
# CUSTOM CSS — light gray theme
# ─────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&family=JetBrains+Mono:wght@400;600&display=swap');

/* ── Base ── */
html, body, [class*="css"] {
    font-family: 'Inter', sans-serif;
    color: #1a1a2e;
}
.stApp {
    background-color: #f0f2f6;
}

/* ── Sidebar ── */
section[data-testid="stSidebar"] {
    background-color: #ffffff;
    border-right: 1px solid #e2e8f0;
    box-shadow: 2px 0 8px rgba(0,0,0,0.04);
}
section[data-testid="stSidebar"] .stMarkdown p,
section[data-testid="stSidebar"] label {
    color: #475569;
    font-size: 0.85rem;
}

/* ── Header ── */
.mpsif-header {
    background: linear-gradient(135deg, #1e3a5f 0%, #2d5986 100%);
    border-radius: 14px;
    padding: 1.5rem 2rem;
    margin-bottom: 1.5rem;
    display: flex;
    align-items: center;
    justify-content: space-between;
}
.mpsif-title {
    font-size: 2.2rem;
    font-weight: 700;
    color: #ffffff;
    letter-spacing: -0.03em;
    margin: 0;
}
.mpsif-subtitle {
    font-size: 0.83rem;
    color: rgba(255,255,255,0.6);
    margin: 6px 0 0 0;
}
.mpsif-logo {
    font-family: 'JetBrains Mono', monospace;
    font-size: 2rem;
    font-weight: 600;
    color: rgba(255,255,255,0.15);
    letter-spacing: -0.05em;
}

/* ── Metric cards ── */
.metric-card {
    background: #ffffff;
    border: 1px solid #e2e8f0;
    border-radius: 12px;
    padding: 1.2rem 1.4rem;
    text-align: center;
    box-shadow: 0 1px 4px rgba(0,0,0,0.05);
    transition: box-shadow 0.2s;
}
.metric-card:hover { box-shadow: 0 4px 12px rgba(0,0,0,0.08); }
.metric-value {
    font-family: 'JetBrains Mono', monospace;
    font-size: 2rem;
    font-weight: 600;
    color: #1e3a5f;
    line-height: 1;
}
.metric-label {
    font-size: 0.7rem;
    font-weight: 600;
    color: #94a3b8;
    text-transform: uppercase;
    letter-spacing: 0.1em;
    margin-top: 6px;
}
.metric-card.accent-green .metric-value { color: #16a34a; }
.metric-card.accent-amber .metric-value { color: #d97706; }
.metric-card.accent-red   .metric-value { color: #dc2626; }

/* ── Table card ── */
.table-card {
    background: #ffffff;
    border: 1px solid #e2e8f0;
    border-radius: 12px;
    padding: 0;
    box-shadow: 0 1px 4px rgba(0,0,0,0.05);
    overflow: hidden;
}
.table-header-row {
    background: #f8fafc;
    border-bottom: 2px solid #e2e8f0;
    padding: 0.65rem 1rem;
}
.col-label {
    font-size: 0.68rem;
    font-weight: 700;
    color: #94a3b8;
    text-transform: uppercase;
    letter-spacing: 0.1em;
}
.table-row {
    padding: 0.75rem 1rem;
    border-bottom: 1px solid #f1f5f9;
    transition: background 0.1s;
}
.table-row:last-child { border-bottom: none; }
.table-row:hover { background: #f8fafc; }

/* ── Ticker chip ── */
.ticker-chip {
    display: inline-block;
    background: #1e3a5f;
    color: #ffffff;
    font-family: 'JetBrains Mono', monospace;
    font-size: 0.75rem;
    font-weight: 600;
    padding: 3px 8px;
    border-radius: 5px;
    letter-spacing: 0.03em;
}

/* ── Company name ── */
.company-name {
    font-size: 0.9rem;
    font-weight: 500;
    color: #1e293b;
}

/* ── Estimate value ── */
.est-value {
    font-family: 'JetBrains Mono', monospace;
    font-size: 0.82rem;
    color: #334155;
}
.est-na {
    font-size: 0.8rem;
    color: #cbd5e1;
}

/* ── Badges ── */
.badge {
    display: inline-block;
    padding: 3px 9px;
    border-radius: 20px;
    font-size: 0.72rem;
    font-weight: 600;
    letter-spacing: 0.02em;
}
.badge-soon     { background:#dcfce7; color:#16a34a; border:1px solid #bbf7d0; }
.badge-upcoming { background:#dbeafe; color:#1d4ed8; border:1px solid #bfdbfe; }
.badge-future   { background:#fef9c3; color:#b45309; border:1px solid #fde68a; }
.badge-past     { background:#f1f5f9; color:#94a3b8; border:1px solid #e2e8f0; }
.badge-tbd      { background:#f1f5f9; color:#94a3b8; border:1px solid #e2e8f0; }

/* ── Date cell ── */
.date-text {
    font-family: 'JetBrains Mono', monospace;
    font-size: 0.82rem;
    color: #475569;
}
.est-marker {
    font-size: 0.68rem;
    color: #d97706;
    margin-left: 3px;
}

/* ── Upload box ── */
.upload-box {
    background: #ffffff;
    border: 2px dashed #cbd5e1;
    border-radius: 14px;
    padding: 2.5rem 2rem;
    text-align: center;
    margin: 1rem 0 1.5rem 0;
    transition: border-color 0.2s;
}
.upload-box:hover { border-color: #1e3a5f; }
.upload-box h3 {
    font-size: 1.1rem;
    font-weight: 600;
    color: #1e3a5f;
    margin: 0 0 0.4rem 0;
}
.upload-box p { color: #64748b; font-size: 0.88rem; margin: 0; }

/* ── Cache pill ── */
.cache-pill {
    display: inline-block;
    padding: 2px 9px;
    border-radius: 20px;
    font-family: 'JetBrains Mono', monospace;
    font-size: 0.68rem;
    font-weight: 600;
    background: #dbeafe;
    color: #1d4ed8;
    border: 1px solid #bfdbfe;
    margin-left: 6px;
}
.cache-pill-warn {
    background: #fef9c3;
    color: #b45309;
    border-color: #fde68a;
}

/* ── Section label ── */
.section-label {
    font-size: 0.68rem;
    font-weight: 700;
    color: #94a3b8;
    text-transform: uppercase;
    letter-spacing: 0.1em;
    border-bottom: 1px solid #e2e8f0;
    padding-bottom: 0.5rem;
    margin-bottom: 1rem;
}

/* ── Calendar ── */
.cal-cell {
    background: #ffffff;
    border: 1px solid #e2e8f0;
    border-radius: 8px;
    padding: 0.4rem 0.5rem;
    min-height: 80px;
}
.cal-cell-today {
    border-color: #1e3a5f;
    background: #eff6ff;
}
.cal-day-num       { font-family:'JetBrains Mono',monospace; font-size:0.72rem; color:#94a3b8; margin-bottom:4px; }
.cal-day-num-today { color:#1e3a5f; font-weight:700; }
.cal-event {
    background: #dbeafe;
    color: #1d4ed8;
    border-radius: 3px;
    padding: 1px 5px;
    font-size: 0.68rem;
    font-family: 'JetBrains Mono', monospace;
    font-weight: 600;
    margin-bottom: 2px;
    white-space: nowrap;
    overflow: hidden;
    text-overflow: ellipsis;
}
.cal-event-past { background:#f1f5f9; color:#94a3b8; }

/* ── Streamlit overrides ── */
#MainMenu { visibility: hidden; }
footer    { visibility: hidden; }
header    { visibility: hidden; }

.stTabs [data-baseweb="tab-list"] {
    background: transparent;
    border-bottom: 2px solid #e2e8f0;
    gap: 0.5rem;
}
.stTabs [data-baseweb="tab"] {
    font-family: 'Inter', sans-serif;
    font-size: 0.85rem;
    font-weight: 500;
    color: #64748b;
    padding: 0.5rem 1rem;
}
.stTabs [aria-selected="true"] {
    color: #1e3a5f !important;
    border-bottom: 2px solid #1e3a5f !important;
    font-weight: 600 !important;
}

div[data-testid="stSelectbox"] > div > div {
    background: #ffffff;
    border: 1px solid #e2e8f0;
    border-radius: 8px;
    font-size: 0.85rem;
}

.stButton > button {
    background: #ffffff;
    color: #334155;
    border: 1px solid #e2e8f0;
    border-radius: 8px;
    font-family: 'Inter', sans-serif;
    font-size: 0.82rem;
    font-weight: 500;
    padding: 0.4rem 0.9rem;
    transition: all 0.15s;
}
.stButton > button:hover {
    background: #f8fafc;
    border-color: #1e3a5f;
    color: #1e3a5f;
}

.stCheckbox label { font-size: 0.85rem; color: #475569; }

div[data-testid="stExpander"] {
    background: #ffffff;
    border: 1px solid #e2e8f0 !important;
    border-radius: 10px !important;
    margin-bottom: 0.5rem;
}
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
GITHUB_URL          = "https://github.com/tp2322/MPSIF-Earnings-Calendar"

DEFAULT_ANALYST_EMAILS = {a: f"{a.lower()}@stern.nyu.edu" for a in ANALYSTS if a != "Unassigned"}
SKIP_SYMBOLS = {"SPAXX**", "SPAXX", "FDRXX", "FZFXX", "FCASH"}


# ─────────────────────────────────────────────
# PERSISTENCE HELPERS
# ─────────────────────────────────────────────
def load_holdings():
    if os.path.exists(HOLDINGS_FILE):
        with open(HOLDINGS_FILE) as f:
            return json.load(f)
    return []

def save_holdings(h):
    with open(HOLDINGS_FILE, "w") as f:
        json.dump(h, f, indent=2)

def load_analyst_emails():
    if os.path.exists(ANALYST_EMAILS_FILE):
        with open(ANALYST_EMAILS_FILE) as f:
            return json.load(f)
    return DEFAULT_ANALYST_EMAILS.copy()

def save_analyst_emails(e):
    with open(ANALYST_EMAILS_FILE, "w") as f:
        json.dump(e, f, indent=2)

def load_upload_meta():
    if os.path.exists(UPLOAD_META_FILE):
        with open(UPLOAD_META_FILE) as f:
            return json.load(f)
    return {}

def save_upload_meta(m):
    with open(UPLOAD_META_FILE, "w") as f:
        json.dump(m, f, indent=2)


# ─────────────────────────────────────────────
# COMPANY NAME CLEANER
# ─────────────────────────────────────────────
import re as _re

# Suffixes/noise to strip from Fidelity description strings
_STRIP_SUFFIXES = _re.compile(
    r"\s+("
    r"Inc\.?|Corp(?:oration)?\.?|Ltd\.?|Llc\.?|Plc\.?|"
    r"Sa/Nv|S\.A\.|N\.?V\.?|"
    r"Com(?:\s+USD[\d.]+)?|Common\s+Stock|"
    r"Ord\s+Npv.*|Adr\b.*|Spon\s+Ads.*|"
    r"Cl\s+[A-C]\b|New\s+Cl\s+[A-C]\b|New\s+Com.*|"
    r"Com\s+USD\d.*|Com\s+NPV.*|USD[\d.]+\b.*|"
    r"Hldgs?\b|Bancorp\b|Bancshares?\b|"
    r"Nv\s+Isin.*|Sedol.*|\(Di\)|Eah\b.*|Rep\s+\d+.*"
    r")(\s+.*)?$",
    _re.IGNORECASE
)

def clean_company_name(raw: str) -> str:
    """
    Strips legal/share-class noise from Fidelity description strings.
    'Northrop Grumman Corp Com Usd1' → 'Northrop Grumman'
    'Medpace Hldgs Inc Com' → 'Medpace'
    'Anheuser-Busch Inbev Sa/Nv Adr...' → 'Anheuser-Busch Inbev'
    """
    if not raw:
        return raw
    name = raw.strip()
    # Iteratively strip trailing noise
    for _ in range(6):
        new = _STRIP_SUFFIXES.sub("", name).strip()
        if new == name:
            break
        name = new
    # Capitalize each word, preserving hyphens
    def cap_word(w):
        return "-".join(p.capitalize() for p in w.split("-"))
    name = " ".join(cap_word(w) for w in name.split())
    return name or raw.strip()


# ─────────────────────────────────────────────
# FIDELITY CSV PARSER
# ─────────────────────────────────────────────
def parse_fidelity_csv(file_bytes: bytes) -> pd.DataFrame:
    text = file_bytes.decode("utf-8-sig")
    data_lines = []
    for line in text.splitlines():
        s = line.strip()
        if not s or s.startswith('"'):
            break
        data_lines.append(line)
    if not data_lines:
        raise ValueError("No data rows found.")

    df = pd.read_csv(io.StringIO("\n".join(data_lines)), index_col=False)
    df.columns = [c.strip().lstrip("\ufeff").strip() for c in df.columns]

    df["Symbol"]      = df["Symbol"].fillna("").astype(str).str.strip()
    df["Description"] = df["Description"].fillna("").astype(str).str.strip()

    df = df[df["Symbol"] != ""]
    df = df[~df["Symbol"].isin(SKIP_SYMBOLS)]
    df = df[~df["Description"].str.lower().str.contains("pending|money market", na=False)]

    for col in ["Last Price", "Current Value", "Percent Of Account", "Quantity"]:
        if col in df.columns:
            df[col] = (
                df[col].fillna("").astype(str)
                .str.replace(r"[$+,%]", "", regex=True).str.strip()
                .replace("", float("nan"))
            )
            df[col] = pd.to_numeric(df[col], errors="coerce")

    df["Description"] = df["Description"].apply(clean_company_name)
    keep = [c for c in ["Symbol", "Description", "Quantity", "Last Price",
                         "Current Value", "Percent Of Account"] if c in df.columns]
    return df[keep].reset_index(drop=True)


# ─────────────────────────────────────────────
# YAHOO FINANCE — EARNINGS DATE
# ─────────────────────────────────────────────
@st.cache_data(ttl=3600, show_spinner=False)
def fetch_earnings_date(ticker: str):
    try:
        tk = yf.Ticker(ticker)
        today_naive = date.today()
        today_ts    = pd.Timestamp.today(tz="UTC")

        # Method 1: calendar dict
        try:
            cal = tk.calendar
            if isinstance(cal, dict):
                raw = cal.get("Earnings Date")
                if raw:
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

        # Method 2: earnings_dates DataFrame
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

        # Method 3: info dict
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

        # Method 4: fast_info
        try:
            fi  = tk.fast_info
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


# ─────────────────────────────────────────────
# YAHOO FINANCE — NEXT QUARTER ESTIMATES
# ─────────────────────────────────────────────
@st.cache_data(ttl=3600, show_spinner=False)
def fetch_estimates(ticker: str) -> dict:
    """
    Pulls next-quarter (0q) consensus EPS and revenue estimates.
    Falls back to current-quarter (0q) from earnings_estimate DataFrame.
    """
    out = {"eps_est": None, "rev_est": None}
    try:
        tk = yf.Ticker(ticker)

        # Primary: earnings_estimate DataFrame has rows like "-1q","0q","1q","2q"
        # "0q" = current quarter being reported (the one tied to the upcoming earnings date)
        try:
            ee = tk.earnings_estimate
            if ee is not None and not ee.empty:
                # Try current quarter first, then next quarter
                for period in ("0q", "+1q", "0y"):
                    if period in ee.index:
                        row = ee.loc[period]
                        # EPS avg estimate
                        for col in ("avg", "Avg", "average", "epsEstimate"):
                            if col in row.index and pd.notna(row[col]):
                                out["eps_est"] = float(row[col])
                                break
                        # Revenue avg estimate
                        break  # only need the period row
        except Exception:
            pass

        # Revenue: revenue_estimate DataFrame mirrors earnings_estimate
        try:
            re = tk.revenue_estimate
            if re is not None and not re.empty:
                for period in ("0q", "+1q", "0y"):
                    if period in re.index:
                        row = re.loc[period]
                        for col in ("avg", "Avg", "average"):
                            if col in row.index and pd.notna(row[col]):
                                out["rev_est"] = float(row[col])
                                break
                        break
        except Exception:
            pass

        # Fallback for EPS: calendar dict has the consensus EPS estimate for next event
        if out["eps_est"] is None:
            try:
                cal = tk.calendar
                if isinstance(cal, dict):
                    eps_raw = cal.get("Earnings Average") or cal.get("EPS Estimate")
                    if eps_raw is not None:
                        out["eps_est"] = float(eps_raw)
            except Exception:
                pass

        # Fallback for Revenue: calendar dict
        if out["rev_est"] is None:
            try:
                cal = tk.calendar
                if isinstance(cal, dict):
                    rev_raw = cal.get("Revenue Average") or cal.get("Revenue Estimate")
                    if rev_raw is not None:
                        out["rev_est"] = float(rev_raw)
            except Exception:
                pass

    except Exception:
        pass
    return out


# ─────────────────────────────────────────────
# FORMATTING HELPERS
# ─────────────────────────────────────────────
def days_until(d):
    return (d - date.today()).days if d else None

def fmt_eps(v):
    if v is None: return None
    return f"${v:.2f}"

def fmt_revenue(v):
    if v is None: return None
    if v >= 1e9:  return f"${v/1e9:.1f}B"
    if v >= 1e6:  return f"${v/1e6:.0f}M"
    return f"${v:,.0f}"

def fmt_growth(v):
    if v is None: return None
    sign = "+" if v >= 0 else ""
    return f"{sign}{v:.1f}%"

def earnings_badge(d):
    if d is None:
        return '<span class="badge badge-tbd">TBD</span>'
    n = days_until(d)
    if n is None:      return '<span class="badge badge-tbd">TBD</span>'
    if n < 0:          return f'<span class="badge badge-past">{d.strftime("%b %d")}</span>'
    if n == 0:         return '<span class="badge badge-soon">TODAY</span>'
    if n <= 7:         return f'<span class="badge badge-soon">in {n}d</span>'
    if n <= 30:        return f'<span class="badge badge-upcoming">in {n}d</span>'
    return f'<span class="badge badge-future">{d.strftime("%b %d")}</span>'


# ─────────────────────────────────────────────
# EMAIL — WEEKLY DIGEST
# ─────────────────────────────────────────────
def get_next_week_range():
    """Returns (monday, sunday) for the upcoming Monday–Sunday."""
    today = date.today()
    days_ahead = (7 - today.weekday()) % 7
    if days_ahead == 0:
        days_ahead = 7   # if today is Monday, next week
    monday = today + pd.Timedelta(days=days_ahead)
    sunday = monday + pd.Timedelta(days=6)
    return monday, sunday


def get_this_week_range():
    """Returns (monday, sunday) for the current week."""
    today  = date.today()
    monday = today - pd.Timedelta(days=today.weekday())
    sunday = monday + pd.Timedelta(days=6)
    return monday, sunday


def build_digest_html(smtp_user, to_emails, earnings_rows, week_start, week_end):
    week_label = f"{week_start.strftime('%b %d')} – {week_end.strftime('%b %d, %Y')}"
    n = len(earnings_rows)
    if not earnings_rows:
        body = "<p style='color:#64748b;'>No MPSIF holdings report earnings this week.</p>"
    else:
        rows_html = ""
        for row in sorted(earnings_rows, key=lambda r: r["earnings_date"]):
            d       = row["earnings_date"]
            day_str = d.strftime("%A, %b %d")
            nd      = days_until(d)
            timing  = ("TODAY" if nd == 0 else f"in {nd}d" if nd and nd > 0 else d.strftime("%b %d"))
            eps_str = fmt_eps(row.get("eps_est")) or "—"
            rev_str = fmt_revenue(row.get("rev_est")) or "—"
            rows_html += f"""
            <tr style="border-bottom:1px solid #f1f5f9;">
              <td style="padding:10px 14px;">
                <span style="background:#1e3a5f;color:#fff;font-family:monospace;
                  font-weight:600;font-size:13px;padding:2px 8px;border-radius:4px;">{row["ticker"]}</span>
              </td>
              <td style="padding:10px 14px;font-weight:500;color:#1e293b;">{row["company"]}</td>
              <td style="padding:10px 14px;font-family:monospace;font-size:13px;color:#475569;">{day_str}</td>
              <td style="padding:10px 14px;color:#16a34a;font-weight:600;font-size:13px;">{timing}</td>
              <td style="padding:10px 14px;font-family:monospace;font-size:13px;color:#334155;">{eps_str}</td>
              <td style="padding:10px 14px;font-family:monospace;font-size:13px;color:#334155;">{rev_str}</td>
            </tr>"""
        body = f"""
        <table style="border-collapse:collapse;width:100%;background:#fff;
          border:1px solid #e2e8f0;border-radius:10px;overflow:hidden;">
          <thead>
            <tr style="background:#f8fafc;color:#94a3b8;font-size:11px;text-transform:uppercase;letter-spacing:.08em;">
              <th style="padding:10px 14px;text-align:left;">Ticker</th>
              <th style="padding:10px 14px;text-align:left;">Company</th>
              <th style="padding:10px 14px;text-align:left;">Date</th>
              <th style="padding:10px 14px;text-align:left;">When</th>
              <th style="padding:10px 14px;text-align:left;">EPS Est (Q)</th>
              <th style="padding:10px 14px;text-align:left;">Rev Est (Q)</th>
            </tr>
          </thead>
          <tbody>{rows_html}</tbody>
        </table>"""

    return f"""<html><body style="background:#f0f2f6;font-family:Arial,sans-serif;padding:32px;color:#1a1a2e;">
      <div style="max-width:660px;margin:0 auto;">
        <div style="background:linear-gradient(135deg,#1e3a5f,#2d5986);border-radius:12px;
          padding:24px 28px;margin-bottom:24px;">
          <h1 style="color:#fff;margin:0;font-size:22px;font-weight:700;">
            📅 MPSIF — Earnings This Week
          </h1>
          <p style="color:rgba(255,255,255,.65);margin:6px 0 0;font-size:13px;">
            {week_label} &nbsp;·&nbsp; {n} position{"s" if n != 1 else ""} reporting
          </p>
        </div>
        {body}
        <p style="color:#94a3b8;font-size:11px;margin-top:20px;text-align:center;">
          MPSIF Earnings Calendar · Data via Yahoo Finance · Estimates are next-quarter consensus
        </p>
      </div>
    </body></html>"""


def send_digest(smtp_host, smtp_port, smtp_user, smtp_pass,
                to_emails, earnings_rows, week_start, week_end):
    week_label = f"{week_start.strftime('%b %d')} – {week_end.strftime('%b %d, %Y')}"
    msg = MIMEMultipart("alternative")
    msg["Subject"] = f"MPSIF Earnings This Week ({week_label})"
    msg["From"]    = smtp_user
    msg["To"]      = ", ".join(to_emails)
    msg.attach(MIMEText(build_digest_html(smtp_user, to_emails, earnings_rows, week_start, week_end), "html"))
    with smtplib.SMTP(smtp_host, int(smtp_port)) as s:
        s.starttls()
        s.login(smtp_user, smtp_pass)
        s.sendmail(smtp_user, to_emails, msg.as_string())


# ─────────────────────────────────────────────
# SESSION STATE
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
# HEADER BANNER
# ─────────────────────────────────────────────
meta = st.session_state.upload_meta
if meta.get("uploaded_at"):
    upload_line = f'Positions last uploaded: <b style="color:rgba(255,255,255,0.9);">{meta["uploaded_at"]}</b>'
else:
    upload_line = 'No positions file uploaded yet'

st.markdown(f"""
<div class="mpsif-header">
  <div>
    <p class="mpsif-title">📅 MPSIF Earnings Calendar</p>
    <p class="mpsif-subtitle">{upload_line}</p>
  </div>
  <div class="mpsif-logo">MPSIF</div>
</div>
""", unsafe_allow_html=True)


# ─────────────────────────────────────────────
# UPLOAD FLOW
# ─────────────────────────────────────────────
has_holdings = len(st.session_state.holdings) > 0

if has_holdings:
    col_toggle, _ = st.columns([2.2, 6])
    with col_toggle:
        btn_label = "✕ Cancel" if st.session_state.show_upload else "📂 Upload new positions file"
        if st.button(btn_label, key="toggle_upload"):
            st.session_state.show_upload = not st.session_state.show_upload
            st.rerun()

if st.session_state.show_upload:
    st.markdown("""
    <div class="upload-box">
      <h3>📂 Upload Positions File</h3>
      <p>Export your portfolio from Fidelity → <b>Accounts → Positions → Download (↓) → CSV</b></p>
    </div>
    """, unsafe_allow_html=True)

    uploaded = st.file_uploader(
        "Drop CSV here", type=["csv"],
        label_visibility="collapsed", key="csv_uploader",
    )

    if uploaded is not None:
        try:
            parsed_df    = parse_fidelity_csv(uploaded.read())
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

if not st.session_state.holdings:
    st.info("No positions loaded. Upload a Fidelity CSV above to get started.")
    st.stop()


# ─────────────────────────────────────────────
# SIDEBAR
# ─────────────────────────────────────────────
with st.sidebar:
    st.markdown("### 🏦 MPSIF")
    st.markdown("---")

    show_past = st.checkbox("Show past earnings", value=False)

    tickers_all = [h["ticker"] for h in st.session_state.holdings]

    st.markdown("**Manual Edits**")

    with st.expander("➕ Add Position"):
        new_ticker  = st.text_input("Ticker", key="new_ticker").upper().strip()
        new_company = st.text_input("Company Name", key="new_company")
        if st.button("Add"):
            if new_ticker:
                if new_ticker in tickers_all:
                    st.error("Already exists.")
                else:
                    st.session_state.holdings.append({"ticker": new_ticker, "company": new_company or new_ticker})
                    save_holdings(st.session_state.holdings)
                    st.rerun()

    with st.expander("🗑️ Remove Position"):
        rm = st.selectbox("Ticker", tickers_all, key="remove_ticker")
        if st.button("Remove"):
            st.session_state.holdings = [h for h in st.session_state.holdings if h["ticker"] != rm]
            save_holdings(st.session_state.holdings)
            st.rerun()

    with st.expander("⚠️ Clear All Data"):
        st.caption("Wipes cache and forces fresh upload.")
        if st.button("Clear & Re-upload", type="primary"):
            st.session_state.holdings    = []
            st.session_state.upload_meta = {}
            st.session_state.show_upload = True
            for f in [HOLDINGS_FILE, UPLOAD_META_FILE]:
                if os.path.exists(f): os.remove(f)
            st.rerun()

    st.markdown("---")
    st.markdown("**📧 Email Reminders**")

    with st.expander("SMTP Settings"):
        smtp_host = st.text_input("Host", value="smtp.gmail.com", key="smtp_host")
        smtp_port = st.text_input("Port", value="587", key="smtp_port")
        smtp_user = st.text_input("From Email", key="smtp_user")
        smtp_pass = st.text_input("App Password", type="password", key="smtp_pass")

    with st.expander("Member Emails"):
        updated_emails = {}
        for a in [x for x in ANALYSTS if x != "Unassigned"]:
            updated_emails[a] = st.text_input(a, value=st.session_state.analyst_emails.get(a, ""), key=f"email_{a}")
        if st.button("Save Emails"):
            st.session_state.analyst_emails = updated_emails
            save_analyst_emails(updated_emails)
            st.success("Saved.")

    with st.expander("📨 Weekly Digest"):
        st.caption(
            "Sends ONE email to the whole team listing all positions "
            "reporting in the upcoming week (Mon–Sun). "
            "Set this up to run every Monday morning via GitHub Actions."
        )
        digest_week = st.radio(
            "Which week?",
            ["Next week (recommended)", "This week"],
            key="digest_week",
        )
        if st.button("Send Digest Now"):
            if not smtp_user or not smtp_pass:
                st.error("Fill in SMTP settings first.")
            else:
                all_emails = [
                    st.session_state.analyst_emails.get(a, "")
                    for a in ANALYSTS if a != "Unassigned"
                ]
                all_emails = [e for e in all_emails if e]
                if not all_emails:
                    st.error("No analyst emails saved. Fill in Member Emails above.")
                else:
                    if digest_week == "Next week (recommended)":
                        week_start, week_end = get_next_week_range()
                    else:
                        week_start, week_end = get_this_week_range()

                    # Find all holdings with earnings in that window
                    digest_rows = []
                    for h in st.session_state.holdings:
                        ed, _, _ = fetch_earnings_date(h["ticker"])
                        if ed and week_start <= ed <= week_end:
                            ests = fetch_estimates(h["ticker"])
                            digest_rows.append({
                                **h,
                                "earnings_date": ed,
                                "eps_est": ests["eps_est"],
                                "rev_est": ests["rev_est"],
                            })

                    wl = f"{week_start.strftime('%b %d')}–{week_end.strftime('%b %d')}"
                    try:
                        send_digest(smtp_host, smtp_port, smtp_user, smtp_pass,
                                    all_emails, digest_rows, week_start, week_end)
                        if digest_rows:
                            st.success(
                                f"✓ Digest sent to {len(all_emails)} members — "
                                f"{len(digest_rows)} position(s) reporting {wl}."
                            )
                        else:
                            st.info(f"Digest sent — no holdings report earnings {wl}.")
                    except Exception as e:
                        st.error(f"Failed to send: {e}")


# ─────────────────────────────────────────────
# FETCH DATA
# ─────────────────────────────────────────────
with st.spinner("Pulling earnings dates and estimates from Yahoo Finance…"):
    rows = []
    for h in st.session_state.holdings:
        ed, call_time, is_est = fetch_earnings_date(h["ticker"])
        ests                  = fetch_estimates(h["ticker"])
        rows.append({
            "ticker":        h["ticker"],
            "company":       h.get("company") or h["ticker"],
            "earnings_date": ed,
            "is_estimate":   is_est,
            "days_until":    days_until(ed),
            "eps_est":       ests["eps_est"],
            "rev_est":       ests["rev_est"],
        })


# ─────────────────────────────────────────────
# SUMMARY METRICS
# ─────────────────────────────────────────────
today      = date.today()
this_week  = [r for r in rows if r["days_until"] is not None and 0 <= r["days_until"] <= 7]
this_month = [r for r in rows if r["days_until"] is not None and 0 <= r["days_until"] <= 30]
no_date    = [r for r in rows if r["earnings_date"] is None]

c1, c2, c3, c4 = st.columns(4)
for col, val, label, accent in [
    (c1, len(rows),       "Total Holdings",        ""),
    (c2, len(this_week),  "Earnings This Week",     "accent-green"),
    (c3, len(this_month), "Earnings This Month",    "accent-amber"),
    (c4, len(no_date),    "Date TBD",               ""),
]:
    col.markdown(f"""
    <div class="metric-card {accent}">
      <div class="metric-value">{val}</div>
      <div class="metric-label">{label}</div>
    </div>""", unsafe_allow_html=True)

st.markdown("<br>", unsafe_allow_html=True)


# ─────────────────────────────────────────────
# TABS
# ─────────────────────────────────────────────
tab_cal, tab_list = st.tabs(["  📅  Calendar View  ", "  📋  List View  "])


# ── LIST VIEW ────────────────────────────────
with tab_list:
    sort_col, _ = st.columns([2.5, 6])
    with sort_col:
        sort_by = st.selectbox(
            "Sort", ["Earnings Date (soonest)", "Ticker A→Z"],
            label_visibility="collapsed", key="sort_by",
        )

    display_rows = [r for r in rows if show_past or r["days_until"] is None or r["days_until"] >= 0]
    if sort_by == "Ticker A→Z":
        display_rows = sorted(display_rows, key=lambda r: r["ticker"])
    else:
        display_rows = sorted(display_rows,
            key=lambda r: (r["days_until"] is None, r["days_until"] if r["days_until"] is not None else 9999))

    if not display_rows:
        st.info("No holdings to display.")
    else:
        # ── Header row ──
        st.markdown('<div class="table-card">', unsafe_allow_html=True)
        hcols = st.columns([1.4, 3.8, 2, 1.6, 1.6, 1.8])
        header_labels = ["TICKER", "COMPANY", "EARNINGS DATE", "EPS EST", "REV EST", "WHEN"]
        st.markdown('<div class="table-header-row">', unsafe_allow_html=True)
        for col, lbl in zip(hcols, header_labels):
            col.markdown(f'<div class="col-label">{lbl}</div>', unsafe_allow_html=True)
        st.markdown('</div>', unsafe_allow_html=True)

        # ── Data rows ──
        for r in display_rows:
            ed       = r["earnings_date"]
            est_tag  = '<span class="est-marker">~est</span>' if r["is_estimate"] and ed else ""
            date_str = f"{ed.strftime('%b %d, %Y')}{est_tag}" if ed else "—"

            eps_str  = fmt_eps(r["eps_est"])
            rev_str  = fmt_revenue(r["rev_est"])
            eps_html = f'<span class="est-value">{eps_str}</span>' if eps_str else '<span class="est-na">—</span>'
            rev_html = f'<span class="est-value">{rev_str}</span>' if rev_str else '<span class="est-na">—</span>'

            rcols = st.columns([1.4, 3.8, 2, 1.6, 1.6, 1.8])
            rcols[0].markdown(f'<span class="ticker-chip">{r["ticker"]}</span>', unsafe_allow_html=True)
            rcols[1].markdown(f'<span class="company-name">{r["company"]}</span>', unsafe_allow_html=True)
            rcols[2].markdown(f'<span class="date-text">{date_str}</span>', unsafe_allow_html=True)
            rcols[3].markdown(eps_html, unsafe_allow_html=True)
            rcols[4].markdown(rev_html, unsafe_allow_html=True)
            rcols[5].markdown(earnings_badge(ed), unsafe_allow_html=True)
            st.markdown('<hr style="border:none;border-top:1px solid #f1f5f9;margin:0.1rem 0;">', unsafe_allow_html=True)

        st.markdown('</div>', unsafe_allow_html=True)

    st.markdown("""
    <div style="margin-top:0.75rem;font-size:0.72rem;color:#94a3b8;">
      <b>~est</b> = date is estimated · <b>EPS Est</b> = forward consensus EPS ·
      <b>Rev Est</b> = forward consensus revenue · Data via Yahoo Finance, refreshed hourly
    </div>""", unsafe_allow_html=True)


# ── CALENDAR VIEW ────────────────────────────
with tab_cal:
    cal_c1, cal_c2, _ = st.columns([1.5, 1.5, 5])
    with cal_c1:
        cal_month = st.selectbox("Month", range(1, 13), index=today.month - 1,
                                 format_func=lambda m: calendar.month_name[m],
                                 label_visibility="collapsed", key="cal_month")
    with cal_c2:
        cal_year = st.selectbox("Year", range(today.year, today.year + 2),
                                label_visibility="collapsed", key="cal_year")

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
            f'<div style="text-align:center;font-size:0.72rem;font-weight:600;'
            f'color:#94a3b8;padding-bottom:6px;letter-spacing:0.05em;">{dn}</div>',
            unsafe_allow_html=True)

    for week in cal_matrix:
        wcols = st.columns(7)
        for col, day in zip(wcols, week):
            if day == 0:
                col.markdown('<div style="min-height:80px;"></div>', unsafe_allow_html=True)
                continue
            this_date  = date(cal_year, cal_month, day)
            is_today   = this_date == today
            cell_class = "cal-cell-today" if is_today else "cal-cell"
            day_class  = "cal-day-num-today" if is_today else "cal-day-num"
            events_html = "".join(
                f'<div class="cal-event{" cal-event-past" if this_date < today else ""}">{t}</div>'
                for t in event_map.get(this_date, [])
            )
            col.markdown(
                f'<div class="{cell_class}"><div class="{day_class}">{day}</div>{events_html}</div>',
                unsafe_allow_html=True)

    st.markdown("""
    <div style="margin-top:1rem;display:flex;gap:1.2rem;align-items:center;flex-wrap:wrap;">
      <span style="font-size:0.72rem;color:#94a3b8;font-weight:600;">LEGEND</span>
      <span class="cal-event" style="position:static;">Upcoming</span>
      <span class="cal-event cal-event-past" style="position:static;">Past</span>
      <div style="display:flex;align-items:center;gap:4px;">
        <div style="width:12px;height:12px;border:2px solid #1e3a5f;background:#eff6ff;border-radius:2px;"></div>
        <span style="font-size:0.72rem;color:#94a3b8;">Today</span>
      </div>
    </div>""", unsafe_allow_html=True)

    two_weeks_out = today + pd.Timedelta(days=14)
    upcoming_events = {
        d: tickers for d, tickers in event_map.items()
        if today <= d <= two_weeks_out
    }

    if upcoming_events:
        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown('<div class="section-label">Events In Next Two Weeks</div>', unsafe_allow_html=True)
        for d in sorted(upcoming_events.keys()):
            for ticker in upcoming_events[d]:
                row = next((x for x in rows if x["ticker"] == ticker), {})
                st.markdown(
                    f'<span class="ticker-chip">{ticker}</span>'
                    f'&nbsp; <span class="company-name" style="font-size:0.85rem;">{row.get("company","")}</span>'
                    f'&nbsp;·&nbsp; <span class="date-text">{d.strftime("%A, %B %d")}</span>'
                    f'&nbsp; {earnings_badge(d)}',
                    unsafe_allow_html=True)
                st.markdown('<div style="height:6px;"></div>', unsafe_allow_html=True)
    elif event_map:
        st.markdown("<br>", unsafe_allow_html=True)
        st.markdown(
            '<div class="section-label">Events In Next Two Weeks</div>'
            '<p style="color:#94a3b8;font-size:0.85rem;">No holdings report earnings in the next 14 days.</p>',
            unsafe_allow_html=True)


# ─────────────────────────────────────────────
# GITHUB BUTTON (fixed top-right)
# ─────────────────────────────────────────────
st.markdown(f"""
<a href="{GITHUB_URL}" target="_blank" style="
    position:fixed; top:16px; right:16px; z-index:9999;
    background:#ffffff; border:1px solid #e2e8f0;
    border-radius:8px; padding:6px 12px;
    font-family:'Inter',sans-serif; font-size:0.75rem;
    font-weight:500; color:#475569;
    text-decoration:none; display:flex; align-items:center; gap:6px;
    box-shadow:0 1px 4px rgba(0,0,0,0.08);
    transition:border-color 0.15s,color 0.15s;"
  onmouseover="this.style.borderColor='#1e3a5f';this.style.color='#1e3a5f'"
  onmouseout="this.style.borderColor='#e2e8f0';this.style.color='#475569'">
  <svg height="14" width="14" viewBox="0 0 16 16" fill="currentColor">
    <path d="M8 0C3.58 0 0 3.58 0 8c0 3.54 2.29 6.53 5.47 7.59.4.07.55-.17.55-.38
    0-.19-.01-.82-.01-1.49-2.01.37-2.53-.49-2.69-.94-.09-.23-.48-.94-.82-1.13
    -.28-.15-.68-.52-.01-.53.63-.01 1.08.58 1.23.82.72 1.21 1.87.87 2.33.66
    .07-.52.28-.87.51-1.07-1.78-.2-3.64-.89-3.64-3.95 0-.87.31-1.59.82-2.15
    -.08-.2-.36-1.02.08-2.12 0 0 .67-.21 2.2.82.64-.18 1.32-.27 2-.27.68 0
    1.36.09 2 .27 1.53-1.04 2.2-.82 2.2-.82.44 1.1.16 1.92.08 2.12.51.56.82
    1.27.82 2.15 0 3.07-1.87 3.75-3.65 3.95.29.25.54.73.54 1.48 0 1.07-.01
    1.93-.01 2.2 0 .21.15.46.55.38A8.013 8.013 0 0016 8c0-4.42-3.58-8-8-8z"/>
  </svg>
  Source on GitHub
</a>
""", unsafe_allow_html=True)


# ─────────────────────────────────────────────
# FOOTER
# ─────────────────────────────────────────────
st.markdown("<br><br>", unsafe_allow_html=True)
st.markdown(
    f'<div style="font-family:\'JetBrains Mono\',monospace;font-size:0.68rem;'
    f'color:#cbd5e1;text-align:center;">'
    f'MPSIF Earnings Calendar · Data via Yahoo Finance · Estimates may vary · '
    f'Refreshed {datetime.now().strftime("%Y-%m-%d %H:%M")} ET'
    f'</div>',
    unsafe_allow_html=True,
)
