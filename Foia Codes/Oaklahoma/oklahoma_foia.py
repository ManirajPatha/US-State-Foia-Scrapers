#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
oklahoma_foia_emailer.py — Send one Oklahoma Open Records request email per Excel/CSV row.

Usage (dry run):
  python oklahoma_foia_emailer.py --input "C:/path/oklahoma_results.xlsx"

Send the first 5 rows:
  python oklahoma_foia_emailer.py --input "C:/path/oklahoma_results.xlsx" --send --limit 5

Requirements:
  pip install pandas openpyxl python-dotenv python-dateutil
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

# -------------------------- REQUESTER & RECIPIENT ---------------------------
REQUESTER = {
    "name": "Maniraj Patha",
    "address": "8181 Fannin St",
    "city": "Houston",
    "state": "Texas",
    "zip": "77054",
    "phone": "+1 6824055734",
    "email": "pathamaniraj97@gmail.com",
}

DEFAULT_RECIPIENT = "openrecordsrequest@oag.ok.gov"
TRANSIENT_SMTP_CODES = {421, 450, 451, 452, 454}

# ------------------------------- Helpers -----------------------------------
def read_table(path: str) -> pd.DataFrame:
    ext = os.path.splitext(path)[1].lower()
    if ext == ".csv":
        df = pd.read_csv(path)
    elif ext in {".xlsx", ".xlsm"}:
        df = pd.read_excel(path, engine="openpyxl")
    elif ext == ".xls":
        try:
            df = pd.read_excel(path, engine="xlrd")
        except ImportError as e:
            raise RuntimeError(
                "This .xls file needs xlrd 2.0.1. Install it:\n"
                "  pip uninstall -y xlrd && pip install xlrd==2.0.1\n"
                "Or Save As .xlsx and rerun."
            ) from e
    else:
        raise ValueError(f"Unsupported file type: {ext}")

    # Drop header-like first row (your file has a duplicate header as row 0)
    def _is_header_like(row: pd.Series) -> bool:
        matches = 0
        for col in df.columns:
            val = str(row.get(col, "")).strip().lower()
            if val == str(col).strip().lower():
                matches += 1
        # If most cells mirror the column names, it's a header row
        return matches >= max(3, int(0.5 * len(df.columns)))

    if len(df) and _is_header_like(df.iloc[0]):
        df = df.iloc[1:].reset_index(drop=True)

    return df

def s(val: Any) -> str:
    return "" if pd.isna(val) else str(val).strip()

def pick_first(row: pd.Series, candidates: List[str]) -> str:
    for c in candidates:
        if c in row and s(row[c]):
            return s(row[c])
    return ""

def row_to_bullets(row: pd.Series) -> str:
    bullets: List[str] = []
    for col, val in row.items():
        txt = s(val)
        if txt:
            bullets.append(f"- {col}: {txt}")
    return "\n".join(bullets) if bullets else "- (No fields present in this row)"

def infer_subject(row: pd.Series) -> str:
    title = pick_first(row, [
        "Procurement Opportunity (text)",
        "title",
        "description",
    ]) or "Procurement Opportunity"
    req = pick_first(row, ["Requisition Number", "notice_id", "solicitation_number", "reference_number", "id", "number"])
    suffix = f" (Req {req})" if req else ""
    return f"Oklahoma Open Records Request – {title}{suffix}"

def build_information_requested(row: pd.Series) -> str:
    """
    Build a robust, per-row paragraph. Uses many columns so it's never blank.
    """
    title = pick_first(row, ["Procurement Opportunity (text)", "title", "description"])
    req = pick_first(row, ["Requisition Number", "notice_id", "solicitation_number", "reference_number", "id", "number"])
    status = pick_first(row, ["Status"])
    posted = pick_first(row, ["scraped_at"]) or pick_first(row, ["Published", "Posted", "publish_date", "date_posted"])
    close_date = pick_first(row, ["Closing Date", "close_date", "due_date", "deadline"])
    award_date = pick_first(row, ["Award Date"])
    awardees = pick_first(row, ["Awardee(s)", "awardees"])
    value = pick_first(row, ["Total Annual Contract Value", "contract_value", "value"])
    opp_url = pick_first(row, ["Procurement Opportunity (urls)", "page_url", "url", "source_url", "link"])
    amend_text = pick_first(row, ["Amendments (text)"])
    amend_urls = pick_first(row, ["Amendments (urls)"])

    lines: List[str] = []

    lines.append(
        "I respectfully request copies of all non-exempt records related to the procurement "
        "opportunity identified below, including the solicitation, all amendments, Q&A, sign-in/attendance "
        "lists, submitted bids/proposals, evaluation materials (individual and consensus), award "
        "recommendations/approvals, the executed contract and amendments, and related correspondence."
    )

    # Always include a compact “identifiers” block so staff can find it quickly
    ids: List[str] = []
    if title: ids.append(f"Title: {title}")
    if req: ids.append(f"Requisition/Ref #: {req}")
    if status: ids.append(f"Status: {status}")
    if close_date: ids.append(f"Closing/Due Date: {close_date}")
    if award_date: ids.append(f"Award Date: {award_date}")
    if value: ids.append(f"Total Contract Value: {value}")
    if awardees: ids.append(f"Awardee(s): {awardees}")
    if opp_url: ids.append(f"Opportunity URL: {opp_url}")
    if amend_text: ids.append(f"Amendments: {amend_text}")
    if amend_urls: ids.append(f"Amendment URL(s): {amend_urls}")
    if posted: ids.append(f"Collected/Scraped at: {posted}")

    if ids:
        lines.append("\nDetails for identification:\n" + "\n".join(f"- {x}" for x in ids))

    # Clarify delivery & fees every time
    lines.append(
        "\nPlease provide records electronically. If fees will exceed $50, please send a cost estimate first. "
        "If any portion is withheld, please cite each specific exemption and release all reasonably segregable material."
    )

    return "\n".join(lines)

def build_body(row: pd.Series) -> str:
    today = datetime.now().strftime("%B %d, %Y")
    info_requested = build_information_requested(row)
    bullets = row_to_bullets(row)  # full row for staff convenience, below the signature

    return f"""Date: {today}

To: Oklahoma Office of the Attorney General – Open Records
Email: {DEFAULT_RECIPIENT}

Subject: Oklahoma Open Records Request

Dear Records Officer,

Pursuant to the Oklahoma Open Records Act, 51 O.S. § 24A.1 et seq., I am submitting the following request.

Information requested
---------------------
{info_requested}

Requester information
---------------------
Full Name: {REQUESTER['name']}
Phone: {REQUESTER['phone']}
Email: {REQUESTER['email']}
Address: {REQUESTER['address']}
City: {REQUESTER['city']}
State: {REQUESTER['state']}
Zipcode: {REQUESTER['zip']}

Preferred delivery & fees
-------------------------
Please provide responsive records electronically to the email address above. If any fees will exceed $50, please let me know in advance with a cost estimate. If portions are exempt, please cite specific statutory exemptions and release all reasonably segregable portions.

If you need clarification to narrow or expedite the search, please let me know and I will respond promptly.

Thank you for your assistance.

Sincerely,
{REQUESTER['name']}

---
For your convenience, here is a summary of the spreadsheet row:
{bullets}
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
        "host": host, "port": port, "username": username, "password": password,
        "sender_name": sender_name, "sender_email": sender_email
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
    modes = [prefer_ssl, not prefer_ssl]   # try preferred, then alternate

    while attempt <= max_retries:
        use_ssl = modes[min(attempt, 1)]
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
        except SMTPDataError as e:
            if e.smtp_code not in TRANSIENT_SMTP_CODES:
                raise
            last_exc = e
        except SMTPResponseException as e:
            if not (400 <= e.smtp_code < 500 or e.smtp_code in TRANSIENT_SMTP_CODES):
                raise
            last_exc = e

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
    ap = argparse.ArgumentParser(description="Oklahoma Open Records emailer (one email per row).")
    ap.add_argument("--input", required=True, help="Path to .xlsx/.xls/.csv with opportunities.")
    ap.add_argument("--send", action="store_true", help="Actually send emails (omit for dry run).")
    ap.add_argument("--to-override", help=f"Send to this address instead of {DEFAULT_RECIPIENT}.")
    ap.add_argument("--limit", type=int, default=0, help="If >0, only process this many rows.")
    ap.add_argument("--skip-blank-rows", action="store_true", help="Skip rows that are entirely blank.")
    ap.add_argument("--log-out", help="Optional path for results log (.xlsx). Default: alongside input.")
    ap.add_argument("--pause", type=float, default=1.0, help="Seconds to sleep between emails.")
    ap.add_argument("--resume-log", help="Existing log to skip rows already SENT.")
    ap.add_argument("--ssl", action="store_true", help="Use SMTPS (SSL) on 465 instead of STARTTLS on 587.")
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
            print(f"[INFO] Resuming: will skip {len(skip_indices)} rows already SENT per {args.resume_log}")

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
