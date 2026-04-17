"""
send_digest.py — run by GitHub Actions every Monday at 8 AM ET.
Reads mpsif_holdings.json and mpsif_analyst_emails.json from the repo,
fetches earnings dates via Yahoo Finance, and emails the team a digest
of all positions reporting in the upcoming week.
"""

import os, json, smtplib
from datetime import date
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

import pandas as pd
import yfinance as yf


# ── Config from environment (set as GitHub Secrets) ──────────────────
SMTP_HOST = os.environ.get("SMTP_HOST", "smtp.gmail.com")
SMTP_PORT = int(os.environ.get("SMTP_PORT", "587"))
SMTP_USER = os.environ.get("SMTP_USER", "")
SMTP_PASS = os.environ.get("SMTP_PASS", "")

if not SMTP_USER or not SMTP_PASS:
    raise SystemExit("SMTP_USER and SMTP_PASS must be set as GitHub Secrets.")


# ── Load holdings and emails from JSON files in the repo ─────────────
def load_json(path, default):
    try:
        with open(path) as f:
            return json.load(f)
    except Exception:
        return default

holdings       = load_json("mpsif_holdings.json", [])
analyst_emails = load_json("mpsif_analyst_emails.json", {})
to_emails      = [e for e in analyst_emails.values() if e]

if not holdings:
    raise SystemExit("No holdings found in mpsif_holdings.json — nothing to send.")
if not to_emails:
    raise SystemExit("No analyst emails found in mpsif_analyst_emails.json.")


# ── Date helpers ──────────────────────────────────────────────────────
def get_next_week_range():
    today      = date.today()
    days_ahead = (7 - today.weekday()) % 7 or 7
    monday     = today + pd.Timedelta(days=days_ahead)
    return monday, monday + pd.Timedelta(days=6)


def days_until(d):
    return (d - date.today()).days if d else None


def fmt_eps(v):
    return f"${v:.2f}" if v is not None else None


def fmt_revenue(v):
    if v is None: return None
    if v >= 1e9:  return f"${v/1e9:.1f}B"
    if v >= 1e6:  return f"${v/1e6:.0f}M"
    return f"${v:,.0f}"


# ── Yahoo Finance fetchers ────────────────────────────────────────────
def fetch_earnings_date(ticker):
    try:
        tk          = yf.Ticker(ticker)
        today_naive = date.today()
        today_ts    = pd.Timestamp.today(tz="UTC")

        try:
            cal = tk.calendar
            if isinstance(cal, dict):
                raw = cal.get("Earnings Date")
                if raw:
                    for val in (raw if isinstance(raw, (list, tuple)) else [raw]):
                        d = pd.Timestamp(val).date()
                        if d >= today_naive - pd.Timedelta(days=5).to_pytimedelta():
                            return d
        except Exception:
            pass

        try:
            edf = tk.earnings_dates
            if edf is not None and not edf.empty:
                future = edf[edf.index >= today_ts - pd.Timedelta(days=5)]
                if not future.empty:
                    return future.index.min().date()
        except Exception:
            pass

        try:
            info = tk.info or {}
            for key in ("earningsDate", "earningsTimestamp", "earningsTimestampStart"):
                raw = info.get(key)
                if raw is None: continue
                if isinstance(raw, (list, tuple)): raw = raw[0]
                if raw:
                    from datetime import datetime
                    d = datetime.fromtimestamp(int(raw)).date()
                    if d >= today_naive:
                        return d
        except Exception:
            pass
    except Exception:
        pass
    return None


def fetch_estimates(ticker):
    out = {"eps_est": None, "rev_est": None}
    try:
        tk = yf.Ticker(ticker)
        try:
            ee = tk.earnings_estimate
            if ee is not None and not ee.empty:
                for period in ("0q", "+1q"):
                    if period in ee.index:
                        row = ee.loc[period]
                        for col in ("avg", "Avg", "average"):
                            if col in row.index and pd.notna(row[col]):
                                out["eps_est"] = float(row[col])
                                break
                        break
        except Exception:
            pass
        try:
            re_df = tk.revenue_estimate
            if re_df is not None and not re_df.empty:
                for period in ("0q", "+1q"):
                    if period in re_df.index:
                        row = re_df.loc[period]
                        for col in ("avg", "Avg", "average"):
                            if col in row.index and pd.notna(row[col]):
                                out["rev_est"] = float(row[col])
                                break
                        break
        except Exception:
            pass
        if out["eps_est"] is None or out["rev_est"] is None:
            try:
                cal = tk.calendar
                if isinstance(cal, dict):
                    if out["eps_est"] is None:
                        v = cal.get("Earnings Average") or cal.get("EPS Estimate")
                        if v: out["eps_est"] = float(v)
                    if out["rev_est"] is None:
                        v = cal.get("Revenue Average") or cal.get("Revenue Estimate")
                        if v: out["rev_est"] = float(v)
            except Exception:
                pass
    except Exception:
        pass
    return out


# ── Find holdings reporting next week ────────────────────────────────
week_start, week_end = get_next_week_range()
week_label = f"{week_start.strftime('%b %d')} – {week_end.strftime('%b %d, %Y')}"

print(f"Checking earnings for week of {week_label}…")

digest_rows = []
for h in holdings:
    ticker = h.get("ticker", "")
    if not ticker:
        continue
    ed = fetch_earnings_date(ticker)
    print(f"  {ticker}: {ed}")
    if ed and week_start <= ed <= week_end:
        ests = fetch_estimates(ticker)
        digest_rows.append({
            "ticker":        ticker,
            "company":       h.get("company", ticker),
            "earnings_date": ed,
            "eps_est":       ests["eps_est"],
            "rev_est":       ests["rev_est"],
        })

print(f"\n{len(digest_rows)} position(s) reporting {week_label}.")


# ── Build HTML email ──────────────────────────────────────────────────
n = len(digest_rows)
if not digest_rows:
    body_html = "<p style='color:#64748b;'>No MPSIF holdings report earnings this week.</p>"
else:
    rows_html = ""
    for row in sorted(digest_rows, key=lambda r: r["earnings_date"]):
        d       = row["earnings_date"]
        nd      = days_until(d)
        timing  = "TODAY" if nd == 0 else (f"in {nd}d" if nd and nd > 0 else d.strftime("%b %d"))
        eps_str = fmt_eps(row["eps_est"]) or "—"
        rev_str = fmt_revenue(row["rev_est"]) or "—"
        rows_html += f"""
        <tr style="border-bottom:1px solid #f1f5f9;">
          <td style="padding:10px 14px;">
            <span style="background:#1e3a5f;color:#fff;font-family:monospace;
              font-weight:600;font-size:13px;padding:2px 8px;border-radius:4px;">
              {row['ticker']}
            </span>
          </td>
          <td style="padding:10px 14px;font-weight:500;color:#1e293b;">{row['company']}</td>
          <td style="padding:10px 14px;font-family:monospace;font-size:13px;color:#475569;">
            {d.strftime('%A, %b %d')}
          </td>
          <td style="padding:10px 14px;color:#16a34a;font-weight:600;font-size:13px;">{timing}</td>
          <td style="padding:10px 14px;font-family:monospace;font-size:13px;color:#334155;">{eps_str}</td>
          <td style="padding:10px 14px;font-family:monospace;font-size:13px;color:#334155;">{rev_str}</td>
        </tr>"""

    body_html = f"""
    <table style="border-collapse:collapse;width:100%;background:#fff;
      border:1px solid #e2e8f0;border-radius:10px;overflow:hidden;">
      <thead>
        <tr style="background:#f8fafc;color:#94a3b8;font-size:11px;
            text-transform:uppercase;letter-spacing:.08em;">
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

html = f"""<html><body style="background:#f0f2f6;font-family:Arial,sans-serif;padding:32px;">
  <div style="max-width:660px;margin:0 auto;">
    <div style="background:linear-gradient(135deg,#1e3a5f,#2d5986);
      border-radius:12px;padding:24px 28px;margin-bottom:24px;">
      <h1 style="color:#fff;margin:0;font-size:22px;font-weight:700;">
        📅 MPSIF — Earnings This Week
      </h1>
      <p style="color:rgba(255,255,255,.65);margin:6px 0 0;font-size:13px;">
        {week_label} &nbsp;·&nbsp; {n} position{"s" if n != 1 else ""} reporting
      </p>
    </div>
    {body_html}
    <p style="color:#94a3b8;font-size:11px;margin-top:20px;text-align:center;">
      MPSIF Earnings Calendar · Data via Yahoo Finance · Estimates are next-quarter consensus
    </p>
  </div>
</body></html>"""


# ── Send ──────────────────────────────────────────────────────────────
msg = MIMEMultipart("alternative")
msg["Subject"] = f"MPSIF Earnings This Week ({week_label}) — {n} position{'s' if n != 1 else ''}"
msg["From"]    = SMTP_USER
msg["To"]      = ", ".join(to_emails)
msg.attach(MIMEText(html, "html"))

print(f"Sending to: {', '.join(to_emails)}")
with smtplib.SMTP(SMTP_HOST, SMTP_PORT) as s:
    s.starttls()
    s.login(SMTP_USER, SMTP_PASS)
    s.sendmail(SMTP_USER, to_emails, msg.as_string())

print("✓ Digest sent successfully.")
