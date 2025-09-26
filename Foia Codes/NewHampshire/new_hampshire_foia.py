#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
nh_awards_emailer.py — Email NH Purchasing (NH.Purchasing@DAS.NH.gov) once per awarded-opportunity row.

Usage examples:
  # Preview only (no emails sent)
  python nh_awards_emailer.py --input "C:/path/nh_awarded_export.xlsx"

  # Actually send first 10 rows via STARTTLS (587)
  python nh_awards_emailer.py --input "C:/path/nh_awarded_export.xlsx" --send --limit 10

  # Send using SSL/SMTPS (465) if your SMTP prefers it
  python nh_awards_emailer.py --input "C:/path/nh_awarded_export.xlsx" --send --ssl

  # Test to yourself instead of NH Purchasing
  python nh_awards_emailer.py --input "C:/path/nh_awarded_export.xlsx" --to-override "pathamaniraj97@gmail.com"

  # Resume without resending rows already SENT per a prior log
  python nh_awards_emailer.py --input "C:/path/nh_awarded_export.xlsx" --send --resume-log "C:/path/nh_awarded_export_send_log.xlsx"

Requirements:
  pip install pandas openpyxl python-dotenv python-dateutil
  # If using legacy .xls:
  pip uninstall -y xlrd && pip install xlrd==2.0.1

.env (same folder) — example for Gmail App Password:
  SMTP_HOST=smtp.gmail.com
  SMTP_PORT=587
  SMTP_USERNAME=pathamaniraj97@gmail.com
  SMTP_PASSWORD=YOUR_16_CHAR_APP_PASSWORD
  SENDER_NAME=Maniraj Patha
  SENDER_EMAIL=pathamaniraj97@gmail.com
"""

import os
import ssl
import time
import socket
import smtplib
import argparse
from datetime import datetime
from typing import Any, Dict, List, Optional, Set

import pandas as pd
from email.message import EmailMessage
from email.utils import formatdate
from dotenv import load_dotenv
from smtplib import (
    SMTP,
    SMTP_SSL,
    SMTPServerDisconnected,
    SMTPAuthenticationError,
    SMTPResponseException,
    SMTPDataError,
)

# -------------------------- REQUESTER (fixed) --------------------------
REQUESTER = {
    "name": "Maniraj Patha",
    "organization": "Southern Arkansas University",
    "address": "8181 Fannin St",
    "phone": "6824055734",
    "email": "pathamaniraj97@gmail.com",
}

# Default recipient (NH Purchasing)
DEFAULT_RECIPIENT = "NH.Purchasing@DAS.NH.gov"

# Transient SMTP codes for retry logic
TRANSIENT_SMTP_CODES = {421, 450, 451, 452, 454}

# ------------------------------- Helpers -----------------------------------
def read_table(path: str) -> pd.DataFrame:
    ext = os.path.splitext(path)[1].lower()
    if ext == ".csv":
        return pd.read_csv(path)
    if ext in {".xlsx", ".xlsm"}:
        return pd.read_excel(path, engine="openpyxl")
    if ext == ".xls":
        try:
            return pd.read_excel(path, engine="xlrd")
        except ImportError as e:
            raise RuntimeError(
                "This .xls file requires xlrd>=2.0.1.\n"
                "Install and retry:\n"
                "  pip uninstall -y xlrd && pip install xlrd==2.0.1\n"
                "Or Save As .xlsx and rerun."
            ) from e
    raise ValueError(f"Unsupported file type: {ext}. Use .csv, .xlsx, .xlsm, or .xls")

def s(val: Any) -> str:
    return "" if pd.isna(val) else str(val).strip()

# Try to guess useful columns commonly seen in your sheets
CANDIDATE_ID_COLS = ["notice_id", "solicitation_number", "bid_id", "rfp_number", "solicitation_id", "reference_no"]
CANDIDATE_TITLE_COLS = ["title", "solicitation_title", "description", "project_title"]
CANDIDATE_AWARD_DATE_COLS = ["award_date", "awarded_on", "finalize_date", "awardposteddate", "award_post_date"]
CANDIDATE_VENDOR_COLS = ["awarded_vendor", "vendor", "contractor", "supplier", "awardee"]
CANDIDATE_AMOUNT_COLS = ["award_amount", "contract_value", "amount", "value", "total_award"]
CANDIDATE_URL_COLS = ["page_url", "detail_url", "url", "source_url"]
CANDIDATE_ATTACH_COLS = ["attachments", "files", "links"]

def first_nonempty(row: pd.Series, candidates: List[str]) -> str:
    for c in candidates:
        if c in row and s(row[c]):
            return s(row[c])
    return ""

def infer_subject(row: pd.Series) -> str:
    title = first_nonempty(row, CANDIDATE_TITLE_COLS) or "Awarded Opportunity"
    nid = first_nonempty(row, CANDIDATE_ID_COLS)
    nid_seg = f" ({nid})" if nid else ""
    return f"NH Purchasing – Award Records Request – {title}{nid_seg} – {REQUESTER['organization']}"

def row_to_bullets(row: pd.Series) -> str:
    """
    Prefer common award fields first, then include the rest of non-empty columns.
    """
    preferred = []
    mapping = [
        ("Notice / Solicitation #", first_nonempty(row, CANDIDATE_ID_COLS)),
        ("Title", first_nonempty(row, CANDIDATE_TITLE_COLS)),
        ("Award Date", first_nonempty(row, CANDIDATE_AWARD_DATE_COLS)),
        ("Awarded Vendor", first_nonempty(row, CANDIDATE_VENDOR_COLS)),
        ("Award Amount", first_nonempty(row, CANDIDATE_AMOUNT_COLS)),
        ("Details URL", first_nonempty(row, CANDIDATE_URL_COLS)),
        ("Attachments", first_nonempty(row, CANDIDATE_ATTACH_COLS)),
    ]
    for label, val in mapping:
        if val:
            preferred.append(f"- {label}: {val}")

    # Add any remaining non-empty columns (avoid duplicates)
    used_cols = {c for _, cands in [
        ("id", CANDIDATE_ID_COLS),
        ("title", CANDIDATE_TITLE_COLS),
        ("award_date", CANDIDATE_AWARD_DATE_COLS),
        ("vendor", CANDIDATE_VENDOR_COLS),
        ("amount", CANDIDATE_AMOUNT_COLS),
        ("url", CANDIDATE_URL_COLS),
        ("attachments", CANDIDATE_ATTACH_COLS),
    ] for c in cands}
    extras: List[str] = []
    for col, val in row.items():
        if col.lower() in used_cols:
            continue
        txt = s(val)
        if txt:
            extras.append(f"- {col}: {txt}")

    return "\n".join(preferred + extras) if (preferred or extras) else "- (No fields present in this row)"

def build_body(row: pd.Series) -> str:
    bullets = row_to_bullets(row)
    today = datetime.now().strftime("%B %d, %Y")
    return f"""Date: {today}

To: State of New Hampshire – Purchasing (Division of Procurement and Support Services)
Email: {DEFAULT_RECIPIENT}

Subject: Award Records Request

Dear NH Purchasing Team,

I’m requesting award documentation for the opportunity summarized below. This request refers to the single row shown here:

{bullets}

Requester
---------
Name: {REQUESTER['name']}
Organization: {REQUESTER['organization']}
Address: {REQUESTER['address']}
Phone: {REQUESTER['phone']}
Email: {REQUESTER['email']}

Request
-------
Please share the available award documents: award notice, executed contract, vendor/contractor details, award amount and term (including renewals), evaluation summary, and any amendments. If these are available on a public site, a direct link is appreciated.

Delivery & Fees
---------------
Please deliver electronically to this address. If fees will exceed $50, please notify me in advance with a cost breakdown.

Thank you for your assistance.

Sincerely,
{REQUESTER['name']}
{REQUESTER['organization']}
{REQUESTER['email']}
{REQUESTER['phone']}
"""

def load_smtp(args=None) -> Dict[str, str]:
    load_dotenv()
    host = os.getenv("SMTP_HOST", "smtp.gmail.com")
    port = int(os.getenv("SMTP_PORT", "587"))
    if args and args.smtp_host:
        host = args.smtp_host
    if args and args.smtp_port:
        port = args.smtp_port
    username = os.getenv("SMTP_USERNAME")
    password = os.getenv("SMTP_PASSWORD")
    sender_name = os.getenv("SENDER_NAME", REQUESTER["name"])
    sender_email = os.getenv("SENDER_EMAIL", REQUESTER["email"])
    missing = [k for k, v in {
        "SMTP_USERNAME": username,
        "SMTP_PASSWORD": password,
        "SENDER_NAME": sender_name,
        "SENDER_EMAIL": sender_email,
    }.items() if not v]
    if missing:
        raise RuntimeError(f"Missing in .env: {', '.join(missing)}")
    return {"host": host, "port": port, "username": username, "password": password,
            "sender_name": sender_name, "sender_email": sender_email}

def _open_smtp(smtp_conf: Dict[str, str], use_ssl: bool) -> smtplib.SMTP:
    if use_ssl:
        ctx = ssl.create_default_context()
        srv = SMTP_SSL(smtp_conf["host"], smtp_conf["port"], timeout=60, context=ctx)
        srv.login(smtp_conf["username"], smtp_conf["password"])
    else:
        srv = SMTP(smtp_conf["host"], smtp_conf["port"], timeout=60)
        srv.ehlo()
        srv.starttls(context=ssl.create_default_context())
        srv.ehlo()
        srv.login(smtp_conf["username"], smtp_conf["password"])
    return srv

def send_with_retries(
    smtp_conf: Dict[str, str],
    to_addr: str,
    subject: str,
    body: str,
    max_retries: int = 5,
    prefer_ssl: bool = False,
) -> None:
    attempt = 0
    delay_seq = [5, 15, 45, 90, 180]
    last_exc: Optional[Exception] = None
    modes = [prefer_ssl, not prefer_ssl]  # try preferred, then alternate once
    mode_index = 0

    while attempt <= max_retries:
        use_ssl = modes[min(mode_index, len(modes)-1)]
        try:
            with _open_smtp(smtp_conf, use_ssl=use_ssl) as server:
                msg = EmailMessage()
                msg["From"] = f"{smtp_conf['sender_name']} <{smtp_conf['sender_email']}>"
                msg["To"] = to_addr
                msg["Date"] = formatdate(localtime=True)
                msg["Subject"] = subject
                msg.set_content(body)
                server.send_message(msg)
            return
        except (SMTPServerDisconnected, socket.timeout, socket.gaierror) as e:
            last_exc = e
            if attempt in (1, 2) and mode_index == 0:
                mode_index = 1  # flip STARTTLS<->SSL after a couple attempts
        except SMTPDataError as e:
            if e.smtp_code in TRANSIENT_SMTP_CODES:
                last_exc = e
            else:
                raise
        except SMTPAuthenticationError as e:
            # Common when Gmail App Password is missing
            raise RuntimeError(
                "SMTP authentication failed. If using Gmail, enable 2-Step Verification "
                "and use a 16-character App Password in SMTP_PASSWORD."
            ) from e
        except SMTPResponseException as e:
            if 400 <= e.smtp_code < 500 or e.smtp_code in TRANSIENT_SMTP_CODES:
                last_exc = e
            else:
                raise

        if attempt == max_retries:
            break
        time.sleep(delay_seq[min(attempt, len(delay_seq)-1)])
        attempt += 1

    raise RuntimeError(f"SMTP send failed after {max_retries+1} attempts: {last_exc}")

def load_sent_row_indices(resume_log_path: str) -> Set[int]:
    try:
        log = pd.read_excel(resume_log_path)
        sent = log.loc[log["status"].str.upper() == "SENT", "row_index"]
        return set(int(i) for i in sent.dropna().tolist())
    except Exception:
        return set()

# --------------------------------- Main ------------------------------------
def main():
    ap = argparse.ArgumentParser(description="NH awards emailer (one email per row).")
    ap.add_argument("--input", required=True, help="Path to .xlsx/.xls/.csv with awarded opportunities.")
    ap.add_argument("--send", action="store_true", help="Actually send emails (omit for dry run).")
    ap.add_argument("--to-override", help="Send to this address instead of NH.Purchasing@DAS.NH.gov.")
    ap.add_argument("--limit", type=int, default=0, help="If >0, only process this many rows.")
    ap.add_argument("--skip-blank-rows", action="store_true", help="Skip rows that are entirely blank.")
    ap.add_argument("--log-out", help="Optional path for results log (.xlsx). Default: alongside input.")
    ap.add_argument("--pause", type=float, default=1.0, help="Seconds to sleep between emails (helps throttles).")
    ap.add_argument("--resume-log", help="Existing log to skip rows already SENT.")
    ap.add_argument("--ssl", action="store_true", help="Use SMTPS (SSL) on port 465 instead of STARTTLS on 587.")
    ap.add_argument("--smtp-host", help="Override SMTP host (default from .env or smtp.gmail.com).")
    ap.add_argument("--smtp-port", type=int, help="Override SMTP port (465 for SSL, 587 for STARTTLS).")
    args = ap.parse_args()

    df = read_table(args.input)
    if args.skip_blank_rows:
        df = df.dropna(how="all")
    if args.limit and args.limit > 0:
        df = df.head(args.limit)

    skip_indices: Set[int] = set()
    if args.resume_log:
        skip_indices = load_sent_row_indices(args.resume_log)
        if skip_indices:
            print(f"[INFO] Resuming: will skip {len(skip_indices)} already-SENT rows based on {args.resume_log}")

    total = len(df)
    print(f"[INFO] Loaded {total} rows from: {args.input}")

    smtp_conf = None
    if args.send:
        smtp_conf = load_smtp(args)
        print("[INFO] SMTP loaded. SENDING mode is ON.")

    to_addr = (args.to_override or DEFAULT_RECIPIENT).strip()
    print(f"[INFO] Recipient: {to_addr}")

    results: List[Dict[str, Any]] = []
    for i, row in df.reset_index(drop=True).iterrows():
        if i in skip_indices:
            print(f"[SKIP] Row {i} already SENT per resume log.")
            results.append({
                "row_index": i, "to": to_addr, "subject": "(skipped via resume)", "status": "SKIPPED",
                "error": "", "timestamp": datetime.now().isoformat(timespec="seconds"),
            })
            continue

        subject = infer_subject(row)
        body = build_body(row)

        print("\n" + "=" * 80)
        print(f"[PREVIEW] Row {i+1}/{total}")
        print(f"TO: {to_addr}")
        print(f"SUBJECT: {subject}")
        print("-" * 80)
        print(body)
        print("=" * 80 + "\n")

        status = "DRY-RUN"
        err = ""
        if args.send:
            try:
                send_with_retries(smtp_conf, to_addr, subject, body, max_retries=5, prefer_ssl=args.ssl)
                status = "SENT"
                time.sleep(max(args.pause, 0.0))
            except Exception as e:
                status = "ERROR"
                err = str(e)
                print(f"[ERROR] Row {i}: {err}")

        results.append({
            "row_index": i,
            "to": to_addr,
            "subject": subject,
            "status": status,
            "error": err,
            "timestamp": datetime.now().isoformat(timespec="seconds"),
        })

    # Log to Excel
    log_df = pd.DataFrame(results)
    if args.log_out:
        out_path = args.log_out
    else:
        base, _ = os.path.splitext(args.input)
        out_path = f"{base}_send_log_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
    log_df.to_excel(out_path, index=False)
    print(f"[OK] Wrote log → {out_path}")

if __name__ == "__main__":
    main()
