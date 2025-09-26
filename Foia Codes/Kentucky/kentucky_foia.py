#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
kentucky_foia.py — Send one Kentucky Open Records email per spreadsheet row.

Only the following fields are rendered in the email body (and nothing else):
  notice_id, title, publish_date, proposal_deadline, office, email, phone_number,
  industry, attachments, page_url, source, status, type

Usage examples:
  # Preview only (no sending)
  python kentucky_foia.py --input "C:/path/kentucky_awarded_rfps.xlsx"

  # Actually send first 10 rows (STARTTLS on 587 by default)
  python kentucky_foia.py --input "C:/path/kentucky_awarded_rfps.xlsx" --send --limit 10

  # Send using SMTPS/SSL on 465 (can help on some networks)
  python kentucky_foia.py --input "C:/path/kentucky_awarded_rfps.xlsx" --send --ssl --pause 2

  # Test to yourself (ignores the sheet's email column)
  python kentucky_foia.py --input "C:/path/kentucky_awarded_rfps.xlsx" --send --limit 1 --to-override "you@gmail.com"

  # Resume without resending rows already SENT per prior log
  python kentucky_foia.py --input "C:/path/kentucky_awarded_rfps.xlsx" --send --resume-log "C:/path/prev_log.xlsx"

Requirements:
  pip install pandas openpyxl python-dotenv python-dateutil
  # If using legacy .xls files:
  pip uninstall -y xlrd && pip install xlrd==2.0.1

.env (same folder as this script) should contain:
  SMTP_HOST=smtp.gmail.com
  SMTP_PORT=587
  SMTP_USERNAME=your@gmail.com
  SMTP_PASSWORD=YOUR_16_CHAR_APP_PASSWORD
  SENDER_NAME=Maniraj Patha
  SENDER_EMAIL=your@gmail.com
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

# ----------------------------- CONFIG / CONSTANTS ----------------------------

# Only these fields will be printed in the email body
ALLOWED_FIELDS = [
    "notice_id",
    "title",
    "publish_date",
    "proposal_deadline",
    "office",
    "email",
    "phone_number",
    "industry",
    "attachments",
    "page_url",
    "source",
    "status",
    "type",
]

TRANSIENT_SMTP_CODES = {421, 450, 451, 452, 454}

# ------------------------------- Helpers ------------------------------------

def read_table(path: str) -> pd.DataFrame:
    """Read CSV/XLSX/XLSM/XLS using safe engines and helpful errors."""
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
    title = s(row.get("title")) or "Opportunity"
    notice_id = s(row.get("notice_id"))
    suffix = f" (#{notice_id})" if notice_id else ""
    return f"Open Records Request: {title}{suffix}"

def row_to_allowed_bullets(row: pd.Series) -> str:
    """Render ONLY ALLOWED_FIELDS that exist & are non-empty."""
    bullets: List[str] = []
    for col in ALLOWED_FIELDS:
        if col in row.index:
            txt = s(row.get(col))
            if txt:
                bullets.append(f"- {col}: {txt}")
    return "\n".join(bullets) if bullets else "- (No allowed fields present in this row)"

def build_body(row: pd.Series, sender_name: str) -> str:
    bullets = row_to_allowed_bullets(row)
    today = datetime.now().strftime("%B %d, %Y")
    return f"""Date: {today}

Dear Records Custodian,

Please see the details below (from the awarded opportunities list):

{bullets}

Regards,
{sender_name}
"""

def load_smtp(args=None) -> Dict[str, str]:
    """Load SMTP settings from .env; allow CLI host/port overrides."""
    load_dotenv(override=True)
    host = os.getenv("SMTP_HOST", "smtp.gmail.com")
    port = int(os.getenv("SMTP_PORT", "587"))
    if args and args.smtp_host:
        host = args.smtp_host
    if args and args.smtp_port:
        port = args.smtp_port

    username = os.getenv("SMTP_USERNAME")
    password = os.getenv("SMTP_PASSWORD")
    sender_name = os.getenv("SENDER_NAME", "Your Name")
    sender_email = os.getenv("SENDER_EMAIL", username or "your@email.com")

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
    """Open SMTP connection either via STARTTLS (587) or SSL (465)."""
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

def send_with_retries(
    smtp_conf: Dict[str, str],
    to_addr: str,
    subject: str,
    body: str,
    max_retries: int = 5,
    prefer_ssl: bool = False,
) -> None:
    """Resilient sender with backoff, mode flip (STARTTLS<->SSL), and error handling."""
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

# --------------------------------- Main -------------------------------------

def main():
    ap = argparse.ArgumentParser(description="Kentucky Open Records mailer (robust).")
    ap.add_argument("--input", required=True, help="Path to .csv/.xlsx/.xlsm/.xls with opportunities.")
    ap.add_argument("--send", action="store_true", help="Actually send emails (omit for dry run).")
    ap.add_argument("--to-override", help="Send to this address instead of the row's email.")
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

    # Case-insensitive map to detect 'email' column even if capitalized
    lower_map = {c.lower(): c for c in df.columns}
    email_col = lower_map.get("email")

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
        print(f"[INFO] SMTP user: {smtp_conf['username']}  host: {smtp_conf['host']}:{smtp_conf['port']}")

    results: List[Dict[str, Any]] = []

    for i, row in df.reset_index(drop=True).iterrows():
        if i in skip_indices:
            print(f"[SKIP] Row {i} already SENT per resume log.")
            results.append({
                "row_index": i, "to": "(resume-skip)", "subject": "(skipped via resume)",
                "status": "SKIPPED", "error": "", "timestamp": datetime.now().isoformat(timespec="seconds"),
            })
            continue

        subject = infer_subject(row)
        body = build_body(row, smtp_conf["sender_name"] if smtp_conf else os.getenv("SENDER_NAME", "Your Name"))

        # Resolve recipient: --to-override wins; else sheet's 'email'
        to_addr = (args.to_override or "").strip()
        if not to_addr and email_col:
            cand = s(row.get(email_col))
            if cand and "@" in cand and "." in cand:
                to_addr = cand

        if not to_addr:
            print(f"[NO-RECIPIENT] Row {i}: no --to-override and no valid 'email' cell.")
            results.append({
                "row_index": i, "to": "(missing)", "subject": subject,
                "status": "NO-RECIPIENT", "error": "", "timestamp": datetime.now().isoformat(timespec="seconds"),
            })
            continue

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
                send_with_retries(
                    smtp_conf, to_addr, subject, body,
                    max_retries=5, prefer_ssl=args.ssl
                )
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
