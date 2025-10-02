#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
virginia_foia.py — Send one Virginia public-records request email per Excel/CSV row.
Robust version: SSL option, retries, resume, pacing, xls/xlsx-safe reader.

Usage examples:
  # Preview only (no sending)
  python virginia_foia.py --input "./eva_awarded_opportunities_100.xlsx"

  # Send first 10 rows via STARTTLS (587)
  python virginia_foia.py --input "./eva_awarded_opportunities_100.xlsx" --send --limit 10

  # Send via SSL (465) with pacing (recommended if .env uses 465)
  python virginia_foia.py --input "./eva_awarded_opportunities_100.xlsx" --send --ssl --smtp-port 465 --pause 2

  # Test to a specific inbox first (no mass send)
  python virginia_foia.py --input "./eva_awarded_opportunities_100.xlsx" --send --limit 1 --ssl --smtp-port 465 --to-override "someone@example.com"

  # Resume without resending rows already SENT per a prior log
  python virginia_foia.py --input "./eva_awarded_opportunities_100.xlsx" --send --resume-log "./eva_awarded_opportunities_100_send_log_YYYYMMDD_HHMM.xlsx"
  #when the default recipient is used, it is sent to
  python virginia_foia.py --input "./eva_awarded_opportunities_100.xlsx" --send --pause 2 --ssl --smtp-port 465

Requirements:
  pip install pandas openpyxl python-dotenv python-dateutil
  # If using legacy .xls files:
  pip uninstall -y xlrd && pip install xlrd==2.0.1

.env (same folder), e.g. for Gmail:
  SMTP_HOST=smtp.gmail.com
  SMTP_PORT=465
  SMTP_USERNAME=your@gmail.com
  SMTP_PASSWORD=YOUR_16_CHAR_APP_PASSWORD
  SENDER_NAME=Your Name
  SENDER_EMAIL=your@gmail.com
"""

import os
import re
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

REQUESTER = {
    "name": "Maniraj Patha",
    "organization": "Southern Arkansas University",
    "address": "8181 Fannin St",
    "city": "Houston",
    "state": "Texas",
    "zip": "77054",
    "phone": "6824055734",
    "email": "pathamaniraj97@gmail.com",
}

DEFAULT_RECIPIENT = "FOIA@governor.virginia.gov"
TRANSIENT_SMTP_CODES = {421, 450, 451, 452, 454}

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

def infer_subject(row: pd.Series) -> str:
    title = (
        s(row.get("Description"))
        or s(row.get("Title"))
        or s(row.get("solicitation_title"))
        or "Opportunity"
    )
    ident = (
        s(row.get("Solicitation ID"))
        or s(row.get("Solicitation Number"))
        or s(row.get("solicitation_number"))
        or s(row.get("notice_id"))
        or ""
    )
    suffix = f" (Solicitation {ident})" if ident else ""
    return f"Virginia Freedom of Information Act Request – {title}{suffix} – {REQUESTER['organization']}"

def row_to_bullets(row: pd.Series) -> str:
    bullets: List[str] = []
    for col, val in row.items():
        txt = s(val)
        if txt:
            bullets.append(f"- {col}: {txt}")
    return "\n".join(bullets) if bullets else "- (No fields present in this row)"

def build_body(row: pd.Series) -> str:
    bullets = row_to_bullets(row)
    today = datetime.now().strftime("%B %d, %Y")
    return f"""Date: {today}

To: State of Virginia – Public Records
Email: {DEFAULT_RECIPIENT}

Subject: Public Records Request

Dear Records Officer,

Pursuant to the Virginia Freedom of Information Act, this requests access to records related to the opportunity detailed below.
These details are taken from the attached opportunities list (this request refers to the specific row summarized here):

{bullets}

Requester Information
---------------------
Name: {REQUESTER['name']}
Organization: {REQUESTER['organization']}
Address: {REQUESTER['address']}
City/State/Zip: {REQUESTER['city']}, {REQUESTER['state']} {REQUESTER['zip']}
Phone: {REQUESTER['phone']}
E-mail: {REQUESTER['email']}

Preferred Delivery
------------------
Please provide responsive records electronically to the email address above.
If any portion of this request is unclear or overbroad, please advise so it can be narrowed.

Fees & Timing
-------------
Please inform in advance if fees exceed $50 and provide a cost breakdown.
If records are exempt in whole or in part, please cite specific statutory exemptions and release all segregable portions.

Thank you for the assistance.

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

    return {
        "host": host,
        "port": port,
        "username": username,
        "password": password,
        "sender_name": sender_name,
        "sender_email": sender_email,
    }

def _open_smtp(smtp_conf: Dict[str, str], use_ssl: bool) -> smtplib.SMTP:
    try:
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
    except SMTPAuthenticationError as e:
        msg = (e.smtp_error or b"").decode("utf-8", errors="ignore")
        if e.smtp_code == 534 or "5.7.9 Application-specific password required" in msg:
            raise RuntimeError(
                "Gmail rejected your login: an App Password is required.\n"
                "Enable 2-Step Verification and create an App Password, then put it in SMTP_PASSWORD."
            )
        raise

def normalize_email(addr: Optional[str]) -> str:
    if not addr:
        return ""
    a = addr.strip()
    a = re.sub(r"^mailto:\s*", "", a, flags=re.IGNORECASE)
    a = a.strip("[]()<>\"' \t")
    m = re.search(r"([A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,})", a, flags=re.IGNORECASE)
    return m.group(1) if m else a

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
    modes = [prefer_ssl, not prefer_ssl]  
    mode_index = 0

    to_addr = normalize_email(to_addr)

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
                mode_index = 1  
        except SMTPDataError as e:
            if e.smtp_code in TRANSIENT_SMTP_CODES:
                last_exc = e
            else:
                raise
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
        sent = log.loc[log["status"].astype(str).str.upper() == "SENT", "row_index"]
        return set(int(i) for i in sent.dropna().tolist())
    except Exception:
        return set()

def main():
    ap = argparse.ArgumentParser(description="Virginia FOIA emailer (robust).")
    ap.add_argument("--input", required=True, help="Path to .xlsx/.xls/.csv with opportunities.")
    ap.add_argument("--send", action="store_true", help="Actually send emails (omit for dry run).")
    ap.add_argument("--to-override", help="Send to this address instead of FOIA@governor.virginia.gov.")
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

    to_addr = normalize_email(args.to_override) if args.to_override else DEFAULT_RECIPIENT
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
