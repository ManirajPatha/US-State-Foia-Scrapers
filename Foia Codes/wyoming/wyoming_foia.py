#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
wyoming_foia.py — Sends one Wyoming public records request email per row in wy_closed_bids.xlsx.

✔ Includes ALL spreadsheet fields in the email (Title, Status, End Date).
✔ Adds a new per-row "Purpose" (2 lines: winning proposals + strategies + fees).
✔ Professional wording citing the Wyoming Public Records Act (W.S. § 16-4-201 et seq.).
✔ Logs results to an Excel file and can write an augmented copy of your input with Purpose.

Usage (examples):
  # Dry-run preview (no emails sent) + write augmented sheet with Purpose
  python wyoming_foia.py --input "C:/path/wy_closed_bids.xlsx" --augment-out "C:/path/wy_closed_bids_with_purpose.xlsx"

  # Actually send (STARTTLS 587)
  python wyoming_foia.py --input "C:/path/wy_closed_bids.xlsx" --send

  # Send first 5 only
  python wyoming_foia.py --input "C:/path/wy_closed_bids.xlsx" --send --limit 5

  # Send using SSL (465)
  python wyoming_foia.py --input "C:/path/wy_closed_bids.xlsx" --send --ssl --smtp-port 465

  # Test to yourself instead of Wyoming
  python wyoming_foia.py --input "C:/path/wy_closed_bids.xlsx" --to-override "pathamaniraj97@gmail.com"

Requirements:
  pip install pandas openpyxl python-dotenv
  # If your input is legacy .xls, also:
  pip uninstall -y xlrd && pip install xlrd==2.0.1
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

# -------------------------- Requester details (fixed) --------------------------
REQUESTER = {
    "name": "Maniraj Patha",
    "organization": "Southern Arkansas University",
    "address": "8181 Fannin St, Houston, TX 77054",
    "phone": "+1 682-405-5734",
    "email": "pathamaniraj97@gmail.com",
}

# Wyoming recipient (can override via --to-override)
DEFAULT_RECIPIENT = "ai-director@wyo.gov"

# Transient SMTP codes for retry logic
TRANSIENT_SMTP_CODES = {421, 450, 451, 452, 454}

# ------------------------------- IO helpers -----------------------------------
def read_table(path: str) -> pd.DataFrame:
    ext = os.path.splitext(path)[1].lower()
    if ext == ".csv":
        return pd.read_csv(path)
    if ext in {".xlsx", ".xlsm"}:
        return pd.read_excel(path, engine="openpyxl")
    if ext == ".xls":
        try:
            return pd.read_excel(path, engine="xlrd")
        except Exception as e:
            raise RuntimeError(
                "This .xls requires xlrd==2.0.1 (or Save As .xlsx and rerun).\n"
                "Run: pip uninstall -y xlrd && pip install xlrd==2.0.1"
            ) from e
    raise ValueError(f"Unsupported file type: {ext}. Use .csv, .xlsx, .xlsm, or .xls")

def s(val: Any) -> str:
    return "" if pd.isna(val) else str(val).strip()

# ------------------------------- Column hints ---------------------------------
# Your file has: Title, Status, End Date — we still keep candidates for robustness.
CANDIDATE_TITLE_COLS = ["Title", "title", "Solicitation Title", "description", "Project Title"]
CANDIDATE_STATUS_COLS = ["Status", "status", "Bid Status", "Award Status"]
CANDIDATE_CLOSE_DATE_COLS = ["End Date", "Closing Date", "Close Date", "close_date", "end_date"]
CANDIDATE_ID_COLS = ["Notice ID", "notice_id", "Solicitation #", "solicitation_number", "Bid ID", "rfp_number"]
CANDIDATE_URL_COLS = ["Details URL", "page_url", "detail_url", "url", "source_url"]

def first_nonempty(row: pd.Series, candidates: List[str]) -> str:
    for c in candidates:
        if c in row and s(row[c]):
            return s(row[c])
    return ""

# --------------------------------- Template -----------------------------------
def infer_subject(row: pd.Series) -> str:
    title = first_nonempty(row, CANDIDATE_TITLE_COLS) or "Closed Bid"
    end_date = first_nonempty(row, CANDIDATE_CLOSE_DATE_COLS)
    when = f" – Closed {end_date}" if end_date else ""
    return f"Wyoming Public Records Request – {title}{when}"

def build_purpose(row: pd.Series) -> str:
    """
    Two-line Purpose as requested: winning proposals + strategies + fee note.
    """
    title = first_nonempty(row, CANDIDATE_TITLE_COLS) or "the referenced opportunity"
    end_date = first_nonempty(row, CANDIDATE_CLOSE_DATE_COLS)
    when = f" (closed {end_date})" if end_date else ""
    return (
        f"Requesting the winning proposal(s), evaluation documentation, and the decision rationale/strategies for {title}{when}.\n"
        f"If fees are required, please share an estimate in advance so I can approve or narrow scope."
    )

def format_record_block(row: pd.Series, purpose_text: str) -> str:
    """
    Ensures the three spreadsheet fields are always present in a clean block.
    Includes extras (ID/URL) if available.
    """
    lines: List[str] = []
    title = first_nonempty(row, CANDIDATE_TITLE_COLS)
    status = first_nonempty(row, CANDIDATE_STATUS_COLS)
    close = first_nonempty(row, CANDIDATE_CLOSE_DATE_COLS)
    rid   = first_nonempty(row, CANDIDATE_ID_COLS)
    url   = first_nonempty(row, CANDIDATE_URL_COLS)

    # Always show the spreadsheet fields explicitly
    if title: lines.append(f"- Title: {title}")
    if status: lines.append(f"- Status: {status}")
    if close: lines.append(f"- End Date: {close}")

    # Helpful extras when present
    if rid: lines.append(f"- Solicitation/Reference: {rid}")
    if url: lines.append(f"- Details URL: {url}")

    # Purpose (always)
    lines.append(f"- Purpose: {purpose_text}")

    return "\n".join(lines) if lines else "- (Row contained no values)"

def build_body(row: pd.Series, purpose_text: str) -> str:
    """
    Professional Wyoming PRA request body.
    """
    today = datetime.now().strftime("%B %d, %Y")
    subject_line = "Public Records Request – Award/Proposal Documentation"
    record_block = format_record_block(row, purpose_text)

    return f"""Date: {today}

To: State of Wyoming – Office of the Chief Information Officer
Email: {DEFAULT_RECIPIENT}

Subject: {subject_line}

Dear Records Custodian,

Pursuant to the Wyoming Public Records Act (W.S. § 16-4-201 et seq.), I respectfully request access to and copies of award-related documents for the opportunity summarized below:

{record_block}

Requested Records
-----------------
• Award notice and final executed contract (including renewals/amendments if applicable).
• Winning proposal(s) and any evaluation/scoring documentation.
• Final pricing/contract value and supplier details.
• Any addenda or post-award modifications.

Delivery & Fees
---------------
Please provide the records electronically to the email address below. If fees are required, kindly provide an estimate in advance; I am willing to pay reasonable costs to process this request.

Redactions & Exemptions
-----------------------
If any portion is exempt, please redact only the exempt portions and release the remainder, citing each specific exemption relied upon.

Requester
---------
Name: {REQUESTER['name']}
Organization: {REQUESTER['organization']}
Address: {REQUESTER['address']}
Phone: {REQUESTER['phone']}
Email: {REQUESTER['email']}

Thank you for your time and assistance.

Sincerely,
{REQUESTER['name']}
{REQUESTER['organization']}
"""

# ------------------------------- SMTP helpers ---------------------------------
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
        "host": host, "port": port, "username": username, "password": password,
        "sender_name": sender_name, "sender_email": sender_email,
    }

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
    modes = [prefer_ssl, not prefer_ssl]  # try preferred then flip once
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
                mode_index = 1
        except SMTPDataError as e:
            if e.smtp_code in TRANSIENT_SMTP_CODES:
                last_exc = e
            else:
                raise
        except SMTPAuthenticationError as e:
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

# ------------------------------- Resume helper --------------------------------
def load_sent_row_indices(resume_log_path: str) -> Set[int]:
    try:
        log = pd.read_excel(resume_log_path)
        sent = log.loc[log["status"].str.upper() == "SENT", "row_index"]
        return set(int(i) for i in sent.dropna().tolist())
    except Exception:
        return set()

# ------------------------------------ Main ------------------------------------
def main():
    ap = argparse.ArgumentParser(description="Wyoming public records emailer (one email per row).")
    ap.add_argument("--input", required=True, help="Path to wy_closed_bids.xlsx/.csv/.xlsm/.xls")
    ap.add_argument("--send", action="store_true", help="Actually send emails (omit for preview).")
    ap.add_argument("--to-override", help="Override recipient (default: ai-director@wyo.gov).")
    ap.add_argument("--limit", type=int, default=0, help="If >0, only process this many rows.")
    ap.add_argument("--skip-blank-rows", action="store_true", help="Skip rows that are entirely blank.")
    ap.add_argument("--log-out", help="Optional path for results log (.xlsx). Default is alongside input.")
    ap.add_argument("--augment-out", help="Optional path to write a copy of the input with the new Purpose column.")
    ap.add_argument("--pause", type=float, default=1.0, help="Seconds to sleep between emails (to avoid throttling).")
    ap.add_argument("--resume-log", help="Existing log to skip rows already SENT.")
    ap.add_argument("--ssl", action="store_true", help="Use SMTPS (SSL) on port 465 (instead of STARTTLS on 587).")
    ap.add_argument("--smtp-host", help="Override SMTP host (default from .env or smtp.gmail.com).")
    ap.add_argument("--smtp-port", type=int, help="Override SMTP port (465 for SSL, 587 for STARTTLS).")
    args = ap.parse_args()

    df = read_table(args.input)
    if args.skip_blank_rows:
        df = df.dropna(how="all")
    if args.limit and args.limit > 0:
        df = df.head(args.limit)

    # Build Purpose per row
    purposes: List[str] = [build_purpose(row) for _, row in df.iterrows()]
    df_with_purpose = df.copy()
    df_with_purpose["Purpose"] = purposes

    # Optionally write augmented sheet
    if args.augment_out:
        try:
            df_with_purpose.to_excel(args.augment_out, index=False)
            print(f"[INFO] Wrote augmented sheet with Purpose → {args.augment_out}")
        except Exception as e:
            print(f"[WARN] Could not write augment-out: {e}")

    # Resume support
    skip_indices: Set[int] = set()
    if args.resume_log:
        skip_indices = load_sent_row_indices(args.resume_log)
        if skip_indices:
            print(f"[INFO] Resuming: will skip {len(skip_indices)} SENT rows per {args.resume_log}")

    total = len(df_with_purpose)
    print(f"[INFO] Loaded {total} rows from: {args.input}")

    smtp_conf = None
    if args.send:
        smtp_conf = load_smtp(args)
        print("[INFO] SMTP loaded. SENDING mode is ON.")

    to_addr = (args.to_override or DEFAULT_RECIPIENT).strip()
    print(f"[INFO] Recipient: {to_addr}")

    results: List[Dict[str, Any]] = []
    for i, row in df_with_purpose.reset_index(drop=True).iterrows():
        if i in skip_indices:
            print(f"[SKIP] Row {i} already SENT per resume log.")
            results.append({
                "row_index": i, "to": to_addr, "subject": "(skipped via resume)", "status": "SKIPPED",
                "error": "", "timestamp": datetime.now().isoformat(timespec="seconds"),
            })
            continue

        subject = infer_subject(row)
        purpose_text = s(row.get("Purpose", "")) or build_purpose(row)
        body = build_body(row, purpose_text)

        print("\n" + "=" * 84)
        print(f"[PREVIEW] Row {i+1}/{total}")
        print(f"TO: {to_addr}")
        print(f"SUBJECT: {subject}")
        print("-" * 84)
        print(body)
        print("=" * 84 + "\n")

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

    # Log results
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
