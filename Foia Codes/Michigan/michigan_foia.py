#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
michigan_foia.py — Email FOIA requests to MSPRecords@Michigan.gov
from rows in michigan_vss_results.csv (or .xlsx), adding a tailored "Purpose" column.

Key features
------------
- Reads /mnt/data/michigan_vss_results.csv (or any .csv/.xlsx you pass via --input)
- Adds a new "Purpose" column (2 professional lines per row; winning proposals + strategies; fees notice)
- Builds a professional FOIA email including:
  (1) Record Summary (key facts),
  (2) CSV Field Snapshot (ALL columns/values from the CSV row)
- Sends one email per row to MSPRecords@Michigan.gov (or --to-override for testing)
- Dry-run preview, resume, rate limiting, SSL/STARTTLS, detailed send log, annotated copy with Purpose

Install:
  pip install pandas openpyxl python-dotenv python-dateutil

.env (in same folder):
  SMTP_HOST=smtp.gmail.com
  SMTP_PORT=587
  SMTP_USERNAME=pathamaniraj97@gmail.com
  SMTP_PASSWORD=YOUR_16_CHAR_APP_PASSWORD
  SENDER_NAME=Maniraj Patha
  SENDER_EMAIL=pathamaniraj97@gmail.com

Usage examples:
  # Preview only (no send)
  python michigan_foia.py --input "/mnt/data/michigan_vss_results.csv"

  # Send first 5 rows to your own email (test)
  python michigan_foia.py --input "/mnt/data/michigan_vss_results.csv" --send --limit 5 --to-override "pathamaniraj97@gmail.com"

  # Send to MSP via SSL (465)
  python michigan_foia.py --input "/mnt/data/michigan_vss_results.csv" --send --ssl --smtp-port 465
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

# ----------------------- Requester / Recipient -----------------------
REQUESTER = {
    "name": "Maniraj Patha",
    "organization": "Southern Arkansas University",
    "address": "8181 Fannin Street, 77054",
    "phone": "6824055734",
    "email": "pathamaniraj97@gmail.com",
}
DEFAULT_RECIPIENT = "MSPRecords@Michigan.gov"

TRANSIENT_SMTP_CODES = {421, 450, 451, 452, 454}

# --------------------------- CSV column helpers ---------------------------
def s(val: Any) -> str:
    return "" if (pd.isna(val) or val is None) else str(val).strip()

# Exact columns observed in /mnt/data/michigan_vss_results.csv:
# ['Description', 'Department', 'Solicitation Number', 'Type']
# We keep generic fallbacks so this script works even if columns expand later.
CANDIDATE_ID_COLS = ["Solicitation Number", "notice_id", "solicitation_number", "bid_id", "rfp_number", "solicitation_id", "reference_no"]
CANDIDATE_TITLE_COLS = ["Description", "title", "solicitation_title", "project_title"]
CANDIDATE_DEPT_COLS = ["Department", "office", "agency", "buyer"]
CANDIDATE_TYPE_COLS = ["Type", "status", "type"]
CANDIDATE_DATE_COLS = ["award_date", "awarded_on", "finalize_date", "awardposteddate", "award_post_date", "publish_date", "proposal_deadline"]
CANDIDATE_VENDOR_COLS = ["awarded_vendor", "vendor", "contractor", "supplier", "awardee"]
CANDIDATE_AMOUNT_COLS = ["award_amount", "contract_value", "amount", "value", "total_award"]
CANDIDATE_URL_COLS = ["page_url", "detail_url", "url", "source_url"]
CANDIDATE_ATTACH_COLS = ["attachments", "files", "links"]

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
                "This .xls requires xlrd>=2.0.1 (or Save As .xlsx and retry)."
            ) from e
    raise ValueError(f"Unsupported file type: {ext}. Use .csv/.xlsx/.xlsm/.xls")

def first_nonempty(row: pd.Series, candidates: List[str]) -> str:
    for c in candidates:
        if c in row and s(row[c]):
            return s(row[c])
    return ""

def infer_subject(row: pd.Series) -> str:
    title = first_nonempty(row, CANDIDATE_TITLE_COLS) or "Awarded Opportunity"
    nid = first_nonempty(row, CANDIDATE_ID_COLS)
    nid_seg = f" ({nid})" if nid else ""
    return f"FOIA Request – Award Records – {title}{nid_seg}"

def record_summary_block(row: pd.Series) -> str:
    """
    Professional, key facts first.
    """
    lines = []
    mapping = [
        ("Title / Description", first_nonempty(row, CANDIDATE_TITLE_COLS)),
        ("Solicitation #", first_nonempty(row, CANDIDATE_ID_COLS)),
        ("Department", first_nonempty(row, CANDIDATE_DEPT_COLS)),
        ("Type / Status", first_nonempty(row, CANDIDATE_TYPE_COLS)),
        ("Relevant Date", first_nonempty(row, CANDIDATE_DATE_COLS)),
        ("Awarded Vendor", first_nonempty(row, CANDIDATE_VENDOR_COLS)),
        ("Award Amount", first_nonempty(row, CANDIDATE_AMOUNT_COLS)),
        ("Details URL", first_nonempty(row, CANDIDATE_URL_COLS)),
        ("Attachments", first_nonempty(row, CANDIDATE_ATTACH_COLS)),
    ]
    for label, val in mapping:
        if val:
            lines.append(f"- {label}: {val}")
    return "\n".join(lines) if lines else "- (No key fields present)"

def csv_snapshot_block(row: pd.Series) -> str:
    """
    Include *every* column/value from the CSV row to satisfy:
    'create a template which includes csv file fields too'.
    """
    lines = []
    for col, val in row.items():
        txt = s(val)
        if txt:
            lines.append(f"- {col}: {txt}")
        else:
            lines.append(f"- {col}: (blank)")
    return "\n".join(lines) if lines else "- (Row appears blank)"

def build_purpose(row: pd.Series) -> str:
    """
    Two professional lines tailored per row:
    - Winning proposal(s) and strategies/evaluation
    - Fees: notify in advance
    """
    title = first_nonempty(row, CANDIDATE_TITLE_COLS) or "the referenced solicitation"
    sol = first_nonempty(row, CANDIDATE_ID_COLS)
    ref = f" (Solicitation {sol})" if sol else ""
    line1 = f"Please provide the winning proposal(s) and the associated evaluation materials/selection strategies for {title}{ref}."
    line2 = "If any fees apply, kindly provide a detailed cost itemization in advance before proceeding."
    return f"{line1}\n{line2}"

def build_body(row: pd.Series, purpose_text: str, df_columns: List[str]) -> str:
    today = datetime.now().strftime("%B %d, %Y")
    subject_title = first_nonempty(row, CANDIDATE_TITLE_COLS) or "Awarded Opportunity"
    rec_summary = record_summary_block(row)
    csv_snapshot = csv_snapshot_block(row)

    return f"""Date: {today}

To: Michigan State Police — Records Section
Email: {DEFAULT_RECIPIENT}

Subject: FOIA Request — Award Records — {subject_title}

Dear Records Officer,

Pursuant to the Michigan Freedom of Information Act (MCL 15.231 et seq.), I respectfully request records related to the awarded opportunity summarized below.

Record Summary (Key Facts)
--------------------------
{rec_summary}

CSV Field Snapshot (All Columns/Values from this Row)
-----------------------------------------------------
{csv_snapshot}

Purpose of Request
------------------
{purpose_text}

Requested Records
-----------------
• Award notice and/or executed contract, including parties, term, value, and any renewals
• Evaluation committee reports, scoring sheets, recommendation memorandum, and selection rationale
• Winning vendor proposal(s) (redacted as necessary to protect exempt information)
• Amendments, addenda, best-and-final offers, and negotiation summaries
• Any publicly accessible links to these records, if available

Delivery & Fees
---------------
Please provide the records electronically to this email address. In accordance with MCL 15.234, if fees will be assessed, please send a detailed, itemized estimate before processing. If the estimate exceeds $50, kindly advise so I may confirm the scope or narrow the request if needed.

If any portion of this request is denied, please cite the specific statutory exemption(s), describe the segregable non-exempt portions that will be released, and identify the individual responsible for the denial as required by MCL 15.235 and MCL 15.244. I understand the statutory response time is five business days (with a permissible extension of up to ten business days).

Requester
---------
Name: {REQUESTER['name']}
Organization: {REQUESTER['organization']}
Address: {REQUESTER['address']}
Phone: {REQUESTER['phone']}
Email: {REQUESTER['email']}

Thank you for your assistance.

Sincerely,
{REQUESTER['name']}
{REQUESTER['organization']}
{REQUESTER['email']}
{REQUESTER['phone']}
"""

# ------------------------------ SMTP utils ------------------------------
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
    modes = [prefer_ssl, not prefer_ssl]  # try preferred then alternate
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
                "SMTP authentication failed. For Gmail, enable 2-Step Verification and use a 16-char App Password."
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

# ------------------------------ Logging / Resume ------------------------------
def load_sent_row_indices(resume_log_path: str) -> Set[int]:
    try:
        log = pd.read_excel(resume_log_path)
        sent = log.loc[log["status"].str.upper() == "SENT", "row_index"]
        return set(int(i) for i in sent.dropna().tolist())
    except Exception:
        return set()

# ----------------------------------- Main -----------------------------------
def main():
    ap = argparse.ArgumentParser(description="Michigan FOIA emailer (includes all CSV fields in the template).")
    ap.add_argument("--input", required=True, help="Path to .csv/.xlsx with opportunities (e.g., /mnt/data/michigan_vss_results.csv)")
    ap.add_argument("--send", action="store_true", help="Actually send emails (omit for dry run).")
    ap.add_argument("--to-override", help="Send to this address instead of MSPRecords@Michigan.gov (useful for testing).")
    ap.add_argument("--limit", type=int, default=0, help="If >0, process only this many rows.")
    ap.add_argument("--skip-blank-rows", action="store_true", help="Skip rows that are entirely blank.")
    ap.add_argument("--log-out", help="Optional path for results log (.xlsx). Default: alongside input.")
    ap.add_argument("--pause", type=float, default=1.0, help="Seconds to sleep between emails (throttle).")
    ap.add_argument("--resume-log", help="Existing log to skip rows already SENT.")
    ap.add_argument("--ssl", action="store_true", help="Use SMTPS (SSL) on port 465 instead of STARTTLS on 587.")
    ap.add_argument("--smtp-host", help="Override SMTP host.")
    ap.add_argument("--smtp-port", type=int, help="Override SMTP port (465 for SSL, 587 for STARTTLS).")
    args = ap.parse_args()

    df = read_table(args.input)
    if args.skip_blank_rows:
        df = df.dropna(how="all")

    # Build/overwrite Purpose column (2 lines per row)
    df["Purpose"] = [build_purpose(row) for _, row in df.iterrows()]

    # Save annotated copy
    base, _ = os.path.splitext(args.input)
    annotated_path = f"{base}_with_purpose.xlsx"
    df.to_excel(annotated_path, index=False)

    if args.limit and args.limit > 0:
        df = df.head(args.limit)

    skip_indices: Set[int] = set()
    if args.resume_log:
        skip_indices = load_sent_row_indices(args.resume_log)
        if skip_indices:
            print(f"[INFO] Resuming: will skip {len(skip_indices)} already-SENT rows based on {args.resume_log}")

    total = len(df)
    print(f"[INFO] Loaded {total} rows from: {args.input}")
    print(f"[INFO] Wrote annotated copy with 'Purpose' → {annotated_path}")

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
            results.append({"row_index": i, "to": to_addr, "subject": "(skipped via resume)",
                            "status": "SKIPPED", "error": "", "timestamp": datetime.now().isoformat(timespec="seconds")})
            continue

        subject = infer_subject(row)
        purpose_text = s(row.get("Purpose")) or build_purpose(row)
        body = build_body(row, purpose_text, list(df.columns))

        # Preview
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
    out_path = args.log_out or f"{base}_send_log_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
    log_df.to_excel(out_path, index=False)
    print(f"[OK] Wrote log → {out_path}")

if __name__ == "__main__":
    main()
