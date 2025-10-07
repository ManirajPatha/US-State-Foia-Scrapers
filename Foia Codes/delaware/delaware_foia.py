#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
delaware_foia.py — Email Delaware Department of State FOIA (dos.foia@delaware.gov)
one email per opportunity row, based on your output_delaware.xlsx.

Professionalized template + explicit Excel fields:
- Contract Number
- Contract Title
- Agency Code
- Posted Date
- Closed/Due Date

Usage (dry-run by default):
  python delaware_foia.py --input "C:/path/output_delaware.xlsx"

Send for real:
  python delaware_foia.py --input "C:/path/output_delaware.xlsx" --send

Optional:
  --limit 5
  --to-override "pathamaniraj97@gmail.com"
  --ssl (use SMTPS/465)
  --smtp-host ... --smtp-port ...
  --resume-log "C:/path/previous_send_log.xlsx"

Requires:
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

# -------------------------- REQUESTER (fixed from your prompt) --------------------------
REQUESTER = {
    "name": "Maniraj Patha",
    "address": "8181 fannin street",
    "zip": "77054",
    "phone": "+1 6824055734",
    "email": "pathamaniraj97@gmail.com",
}

# Default recipient (Delaware DOS FOIA)
DEFAULT_RECIPIENT = "dos.foia@delaware.gov"

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

# Canonical names found in your Delaware output
REQ_FIELDS = [
    "Contract Number",
    "Contract Title",
    "Agency Code",
    "Posted Date",
    "Closed/Due Date",
]

# Candidate lists for resiliency if headers differ slightly
CANDIDATE_ID_COLS = ["contract number", "notice_id", "solicitation_number", "bid_id", "rfp_number", "solicitation_id", "reference_no"]
CANDIDATE_TITLE_COLS = ["contract title", "title", "solicitation_title", "description", "project_title"]
CANDIDATE_AGENCY_COLS = ["agency code", "agency", "department"]
CANDIDATE_POSTED_DATE_COLS = ["posted date", "publish_date", "posted_on"]
CANDIDATE_CLOSED_DATE_COLS = ["closed/due date", "close_date", "closing_date", "due_date"]

def norm_cols(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    lower = {c: c.lower() for c in df.columns}
    df.rename(columns=lower, inplace=True)
    return df

def first_nonempty(row: pd.Series, candidates: List[str]) -> str:
    for c in candidates:
        if c in row and s(row[c]):
            return s(row[c])
    return ""

def ensure_purpose(df: pd.DataFrame) -> pd.DataFrame:
    """
    Create/overwrite Purpose with a concise 2-line professional description per row.
    Line 1: Ask for winning proposals and evaluation/selection strategies.
    Line 2: Fee notice up to $50 unless pre-approved.
    """
    df = df.copy()
    def make_purpose(row: pd.Series) -> str:
        title = first_nonempty(row, CANDIDATE_TITLE_COLS)
        ident = first_nonempty(row, CANDIDATE_ID_COLS)
        agency = first_nonempty(row, CANDIDATE_AGENCY_COLS)
        label_bits = [b for b in [title, ident] if b]
        label = " – ".join(label_bits) if label_bits else "this opportunity"
        agency_seg = f" (Agency: {agency})" if agency else ""
        line1 = f"Requesting copies of the winning proposal(s) and the evaluation/selection strategies for {label}{agency_seg}."
        line2 = "If fees apply, please advise in advance; I authorize costs up to $50 unless pre-approved."
        return f"{line1}\n{line2}"
    df["Purpose"] = df.apply(make_purpose, axis=1)
    return df

def infer_subject(row: pd.Series) -> str:
    title = first_nonempty(row, CANDIDATE_TITLE_COLS) or "Procurement Record"
    nid = first_nonempty(row, CANDIDATE_ID_COLS)
    nid_seg = f" ({nid})" if nid else ""
    return f"FOIA Request – Procurement Records for {title}{nid_seg}"

def opportunity_details_block(row: pd.Series) -> str:
    """
    Formats the five Excel fields in a clean, professional block.
    Falls back gracefully if any are missing.
    """
    def get(label: str, cands: List[str]) -> str:
        v = first_nonempty(row, cands)
        return v if v else "—"

    contract_no = get("Contract Number", CANDIDATE_ID_COLS)
    title       = get("Contract Title", CANDIDATE_TITLE_COLS)
    agency      = get("Agency Code", CANDIDATE_AGENCY_COLS)
    posted      = get("Posted Date", CANDIDATE_POSTED_DATE_COLS)
    closed      = get("Closed/Due Date", CANDIDATE_CLOSED_DATE_COLS)

    # Fixed labels exactly matching your Excel headers
    lines = [
        f"- Contract Number: {contract_no}",
        f"- Contract Title: {title}",
        f"- Agency Code: {agency}",
        f"- Posted Date: {posted}",
        f"- Closed/Due Date: {closed}",
    ]
    # Include Purpose (computed) for clarity:
    purpose = s(row.get("purpose", "")) or s(row.get("Purpose", ""))
    if purpose:
        lines.append(f"- Purpose: {purpose.replace('\n', ' ')}")
    return "\n".join(lines)

def build_body(row: pd.Series) -> str:
    today = datetime.now().strftime("%B %d, %Y")
    recipient_label = "Delaware Department of State — FOIA"
    to_line = f"{recipient_label}\nEmail: {DEFAULT_RECIPIENT}"

    details = opportunity_details_block(row)

    # Professional, concise, and specific scope
    records_scope = """\
1) The winning proposal(s) (including all sections, forms, and attachments).
2) Evaluation materials: scoring sheets, evaluator comments/notes, selection memoranda, and the basis of award.
3) Award documents: notice of award, contract (including all exhibits/attachments), BAFOs (if any), and pricing pages.
4) Addenda/clarifications and the final bidders list.
"""

    fee_language = (
        "If fees will exceed $50, please provide a cost estimate before proceeding. "
        "Electronic delivery via email is preferred."
    )

    requester_block = f"""\
Name: {REQUESTER['name']}
Address: {REQUESTER['address']} {REQUESTER['zip']}
Phone: {REQUESTER['phone']}
Email: {REQUESTER['email']}"""

    return f"""Date: {today}

To:
{to_line}

Subject: FOIA Request – Procurement Records

Dear FOIA Coordinator,

Pursuant to Delaware’s Freedom of Information Act, I respectfully request the following records for the opportunity identified below.

Opportunity Details
-------------------
{details}

Records Requested
-----------------
{records_scope}

Delivery & Fees
---------------
{fee_language}

Requester
---------
{requester_block}

Please let me know if any portion of this request is unclear or would benefit from narrowing. Thank you for your assistance.

Sincerely,
{REQUESTER['name']}
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
    ap = argparse.ArgumentParser(description="Delaware FOIA emailer (one email per row).")
    ap.add_argument("--input", required=True, help="Path to .xlsx/.xls/.csv with Delaware opportunities.")
    ap.add_argument("--send", action="store_true", help="Actually send emails (omit for dry run).")
    ap.add_argument("--to-override", help="Send to this address instead of dos.foia@delaware.gov.")
    ap.add_argument("--limit", type=int, default=0, help="If >0, only process this many rows.")
    ap.add_argument("--skip-blank-rows", action="store_true", help="Skip rows that are entirely blank.")
    ap.add_argument("--log-out", help="Optional path for results log (.xlsx). Default: alongside input.")
    ap.add_argument("--pause", type=float, default=1.2, help="Seconds to sleep between emails (helps throttles).")
    ap.add_argument("--resume-log", help="Existing log to skip rows already SENT.")
    ap.add_argument("--ssl", action="store_true", help="Use SMTPS (SSL) on port 465 instead of STARTTLS on 587.")
    ap.add_argument("--smtp-host", help="Override SMTP host (default from .env or smtp.gmail.com).")
    ap.add_argument("--smtp-port", type=int, help="Override SMTP port (465 for SSL, 587 for STARTTLS).")
    args = ap.parse_args()

    df = read_table(args.input)
    df = norm_cols(df)
    if args.skip_blank_rows:
        df = df.dropna(how="all")
    if args.limit and args.limit > 0:
        df = df.head(args.limit)

    # Create/refresh Purpose column
    df_with_purpose = ensure_purpose(df)

    # Write the augmented sheet next to the input
    base, ext = os.path.splitext(args.input)
    purpose_out = f"{base}_with_purpose.xlsx"
    df_with_purpose.to_excel(purpose_out, index=False)
    print(f"[OK] Wrote sheet with Purpose → {purpose_out}")

    skip_indices: Set[int] = set()
    if args.resume_log:
        skip_indices = load_sent_row_indices(args.resume_log)
        if skip_indices:
            print(f"[INFO] Resuming: will skip {len(skip_indices)} already-SENT rows based on {args.resume_log}")

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
        out_path = f"{base}_send_log_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
    log_df.to_excel(out_path, index=False)
    print(f"[OK] Wrote log → {out_path}")

if __name__ == "__main__":
    main()
