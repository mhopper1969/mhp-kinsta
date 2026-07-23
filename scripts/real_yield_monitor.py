#!/usr/bin/env python3
"""
Long-dated real yield monitor
-----------------------------
Fetches the Bank of England's latest real (index-linked) yield curve data,
reads the long-dated real yields (20y / 25y / 30y), compares against your
trigger levels, and emails you if a trigger is breached.

Run daily (cron / GitHub Actions). It is deliberately dumb:
it ALERTS, it never decides. Confirm any alert against the BoE curve and
live dealing prices before buying anything.

SMTP creds come from environment variables (SMTP_USER, SMTP_PASS, EMAIL_TO);
in this repo those are supplied by GitHub Actions secrets.

CLI:
  --test-email   Send a one-line test email and exit.
  --dry-run      Fetch/parse and print the report; never send email.

Data source: BoE "latest yield curve data" zip, which contains the
GLC Real daily data workbook (spot curve, annual maturities).
File naming inside the zip has changed occasionally over the years,
so the script matches loosely on "real" in the filename. If the BoE
restructures the download, adjust ZIP_URL / file matching below.
"""

import io
import os
import re
import sys
import zipfile
import smtplib
import ssl
from email.message import EmailMessage
from datetime import date

import requests
import openpyxl

# ----------------------------- CONFIG ---------------------------------

# Trigger levels, in PERCENT real yield. Alert fires if ANY listed maturity
# closes at or above its threshold.
TRIGGERS = {
    20: 2.40,   # 20-year real yield >= 2.40%  -> deploy-a-tranche territory
    25: 2.40,
    30: 2.40,
}

# Optional second, louder threshold (mentioned in the email if breached)
ACCELERATE_LEVEL = 3.00   # percent

# Email settings pulled from environment (GitHub Actions secrets).
SMTP_HOST = os.environ.get("SMTP_HOST", "smtp.gmail.com")
SMTP_PORT = int(os.environ.get("SMTP_PORT", "465"))
SMTP_USER = os.environ.get("SMTP_USER", "")
SMTP_PASS = os.environ.get("SMTP_PASS", "")
EMAIL_TO  = os.environ.get("EMAIL_TO", "")

# Only email when a trigger is breached. Set to True to get a short
# "heartbeat" email every run regardless (useful for the first week).
ALWAYS_EMAIL = False

ZIP_URL = ("https://www.bankofengland.co.uk/-/media/boe/files/statistics/"
           "yield-curves/latest-yield-curve-data.zip")

# ----------------------------------------------------------------------


def fetch_real_curve():
    """Download BoE zip, return (as_of_date, {maturity_years: real_yield_pct})."""
    r = requests.get(ZIP_URL, timeout=60,
                     headers={"User-Agent": "real-yield-monitor/1.0"})
    r.raise_for_status()
    zf = zipfile.ZipFile(io.BytesIO(r.content))

    # find the real-curve workbook (name has varied: e.g. "GLC Real daily data current month.xlsx")
    name = None
    for n in zf.namelist():
        if re.search(r"real", n, re.I) and n.lower().endswith((".xlsx", ".xls")):
            name = n
            break
    if not name:
        raise RuntimeError(f"No 'real' workbook found in zip: {zf.namelist()}")

    wb = openpyxl.load_workbook(io.BytesIO(zf.read(name)), data_only=True)

    # find the spot-curve sheet
    sheet = None
    for s in wb.sheetnames:
        if "spot" in s.lower():
            sheet = wb[s]
            break
    if sheet is None:
        sheet = wb[wb.sheetnames[0]]

    # Layout (long-standing BoE format): a header row containing maturities
    # (0.5, 1, 1.5, ... years) a few rows down; dates in column A below it.
    # The maturity header is the row whose numeric cells (a) form a strictly
    # increasing sequence and (b) start near 0.5 - 1 year. Just counting
    # numerics can falsely match the first data row (yields are also numeric).
    rows = list(sheet.iter_rows(values_only=True))

    def looks_like_maturity_row(row):
        nums = [c for c in row[1:] if isinstance(c, (int, float))]
        if len(nums) < 10:
            return False
        if not (0 < nums[0] <= 1.5):
            return False
        return all(b > a for a, b in zip(nums, nums[1:]))

    header_idx, maturities = None, None
    for i, row in enumerate(rows[:15]):
        if looks_like_maturity_row(row):
            header_idx = i
            maturities = row
            break
    if header_idx is None:
        raise RuntimeError("Could not locate maturity header row in BoE workbook")

    # last row with a date in column A AND numeric yields = latest observation
    latest = None
    for row in rows[header_idx + 1:]:
        if row[0] is None:
            continue
        if not any(isinstance(c, (int, float)) for c in row[1:]):
            continue
        latest = row
    if latest is None:
        raise RuntimeError("No data rows found")

    as_of = latest[0]
    curve = {}
    for col, mat in enumerate(maturities):
        if isinstance(mat, (int, float)) and float(mat).is_integer():
            val = latest[col]
            if isinstance(val, (int, float)):
                curve[int(mat)] = float(val)
    if "--debug" in sys.argv:
        int_mats = sorted(k for k in curve.keys())
        print(f"[debug] sheet={sheet.title!r} header_row={header_idx} "
              f"max_maturity={max(int_mats) if int_mats else 'none'} "
              f"integer_maturities={int_mats}")
    return as_of, curve


def build_report(as_of, curve):
    breached, lines = [], []
    for mat, level in sorted(TRIGGERS.items()):
        y = curve.get(mat)
        if y is None:
            lines.append(f"  {mat}y: not available in curve")
            continue
        flag = ""
        if y >= level:
            breached.append((mat, y, level))
            flag = "  <<< TRIGGER"
            if y >= ACCELERATE_LEVEL:
                flag = "  <<< TRIGGER (ACCELERATE LEVEL)"
        lines.append(f"  {mat}y real yield: {y:.2f}%  (trigger {level:.2f}%){flag}")
    as_of_str = as_of.strftime("%d %b %Y") if hasattr(as_of, "strftime") else str(as_of)
    body = (f"BoE real spot curve, close of {as_of_str}\n\n"
            + "\n".join(lines)
            + "\n\nRule: breach = go and LOOK (BoE curve + live rung prices),"
              " not go and buy.\nTail candidates: 2054 / 2062 / 2068 linkers."
              "\nThis is an automated alert; verify at source before dealing.")
    return breached, body


def send_email(subject, body):
    if not (SMTP_USER and SMTP_PASS and EMAIL_TO):
        raise RuntimeError(
            "SMTP_USER / SMTP_PASS / EMAIL_TO not set in environment. "
            "In GitHub Actions, add them as repository secrets."
        )
    msg = EmailMessage()
    msg["Subject"] = subject
    msg["From"] = SMTP_USER
    msg["To"] = EMAIL_TO
    msg.set_content(body)
    ctx = ssl.create_default_context()
    with smtplib.SMTP_SSL(SMTP_HOST, SMTP_PORT, context=ctx) as s:
        s.login(SMTP_USER, SMTP_PASS)
        s.send_message(msg)


def main():
    if "--test-email" in sys.argv:
        send_email("real-yield monitor: test", "Email settings work.")
        print("Test email sent.")
        return

    dry_run = "--dry-run" in sys.argv

    try:
        as_of, curve = fetch_real_curve()
    except Exception as e:
        msg = (f"Monitor could not read the BoE data on {date.today()}:\n{e}\n"
               "Check ZIP_URL / file format at "
               "https://www.bankofengland.co.uk/statistics/yield-curves")
        if dry_run:
            print("FETCH FAILED (dry-run, no email):\n" + msg)
        else:
            send_email("real-yield monitor: FETCH FAILED", msg)
        raise

    breached, body = build_report(as_of, curve)
    if dry_run:
        print("DRY RUN - no email will be sent.\n" + body)
        if breached:
            mats = ", ".join(f"{m}y={y:.2f}%" for m, y, _ in breached)
            print(f"\n(Would have alerted: {mats})")
        return

    if breached:
        mats = ", ".join(f"{m}y={y:.2f}%" for m, y, _ in breached)
        send_email(f"REAL YIELD TRIGGER: {mats}", body)
        print("ALERT SENT\n" + body)
    elif ALWAYS_EMAIL:
        send_email("real-yield monitor: no trigger", body)
        print(body)
    else:
        print(body)


if __name__ == "__main__":
    main()
