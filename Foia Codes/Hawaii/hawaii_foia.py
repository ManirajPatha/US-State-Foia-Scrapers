#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
hawaii_foia.py — Professional UIPA (Hawai‘i) records request emailer.

• Reads a spreadsheet of closed/awarded opportunities (e.g., closed-hawaii-opportunities.xlsx).
• Adds a Purpose column (2 tailored lines per row) per your requirement.
• Builds a polished, statutory UIPA request email that explicitly enumerates ALL worksheet columns
  (preserving your header names and including every non-empty value from the row).
• Sends one email per row to govoffice.uipa@hawaii.gov (or override with --to-override).
• Writes an enriched copy (with Purpose) and a send log Excel.

Author: Prepared for Maniraj Patha
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

# -------------------------- REQUESTER (fixed from user) ---------------------
REQUESTER = {
    "name": "Maniraj Patha",
    "organization": "Southern Arkansas University",
    "address_line_1": "8181 Fannin Street",
    "address_line_2": "Houston, TX 77054",
    "phone": "+1 682-405-5734",
    "email": "pathamaniraj97@gmail.com",
}

# Default recipient (Hawai‘i Governor's Office – UIPA)
DEFAULT_RECIPIENT = "govoffice.uipa@hawaii.gov"

# Transient SMTP codes for retry logic
TRANSIENT_SMTP_CODES = {421, 450, 451, 452, 454}

# ------------------------------- Helpers ------------------------------------
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
                "This .xls file requires xlrd>=2.0.1. Either install:\n"
                "  pip uninstall -y xlrd && pip install xlrd==2.0.1\n"
                "or Save As .xlsx and rerun."
            ) from e
    raise ValueError(f"Unsupported file type: {ext}. Use .csv, .xlsx, .xlsm, or .xls")

def s(val: Any) -> str:
    return "" if (pd.isna(val) or str(val).strip().lower() == "nan") else str(val).strip()

# Try to guess useful columns commonly seen in your sheets (for subject building only)
CANDIDATE_ID_COLS = ["notice_id", "solicitation_number", "bid_id", "rfp_number", "solicitation_id", "reference_no"]
CANDIDATE_TITLE_COLS = ["title", "solicitation_title", "description", "project_title"]
CANDIDATE_AGENCY_COLS = ["agency", "department", "office", "org"]

def first_nonempty(row: pd.Series, candidates: List[str]) -> str:
    for c in candidates:
        if c in row and s(row[c]):
            return s(row[c])
    return ""

def infer_subject(row: pd.Series) -> str:
    title = first_nonempty(row, CANDIDATE_TITLE_COLS) or "Closed Opportunity"
    nid = first_nonempty(row, CANDIDATE_ID_COLS)
    nid_seg = f" ({nid})" if nid else ""
    agency = first_nonempty(row, CANDIDATE_AGENCY_COLS)
    agy_seg = f" – {agency}" if agency else ""
    return f"Hawai‘i UIPA Request – Award/Contract Records – {title}{nid_seg}{agy_seg}"

def build_purpose(row: pd.Series) -> str:
    """
    2 lines: request winning proposals & evaluation strategies; mention fees if any.
    Tailored with title/solicitation/agency when available.
    """
    title = first_nonempty(row, CANDIDATE_TITLE_COLS)
    nid = first_nonempty(row, CANDIDATE_ID_COLS)
    agency = first_nonempty(row, CANDIDATE_AGENCY_COLS)

    id_part = f" (Solicitation {nid})" if nid else ""
    agy_part = f" from {agency}" if agency else ""

    line1 = (
        f"Requesting the winning proposal(s){id_part} for “{title or 'this opportunity'}”{agy_part}, "
        f"including the final executed contract and award documentation."
    )
    line2 = (
        "Please include evaluation criteria, scoring sheets, and selection rationale. "
        "If fees apply, kindly provide an estimate in advance; proceed only after approval."
    )
    return f"{line1}\n{line2}"

def row_to_all_fields_block(row: pd.Series, original_columns: List[str]) -> str:
    """
    Enumerate ALL columns from the worksheet in their original order.
    Only non-empty values are listed; empty values are omitted for clarity.
    """
    lines: List[str] = []
    for col in original_columns:
        val = s(row.get(col, ""))
        if val:
            lines.append(f"- {col}: {val}")
    return "\n".join(lines) if lines else "- (No non-empty fields in this row.)"

def build_body(row: pd.Series, purpose_text: str, original_columns: List[str]) -> str:
    today = datetime.now().strftime("%B %d, %Y")
    all_fields = row_to_all_fields_block(row, original_columns)
    return f"""Date: {today}

To: Office of the Governor, State of Hawai‘i (UIPA)
Email: {DEFAULT_RECIPIENT}

Subject: Uniform Information Practices Act (UIPA) Request – Award/Contract Records

Dear Records Coordinator,

Pursuant to the Hawai‘i Uniform Information Practices Act, HRS §92F, I respectfully request access to
and copies of public records related to the closed solicitation identified below. The row details come
directly from our working worksheet and are enumerated for precise identification:

Record Details (from our worksheet)
-----------------------------------
{all_fields}

Requester Information
---------------------
Name: {REQUESTER['name']}
Organization: {REQUESTER['organization']}
Address: {REQUESTER['address_line_1']}, {REQUESTER['address_line_2']}
Phone: {REQUESTER['phone']}
Email: {REQUESTER['email']}

Purpose of Request
------------------
{purpose_text}

Records Requested
-----------------
• Award notice and fully-executed contract (including amendments/renewals).
• Winning proposal(s) and vendor submissions releasable under UIPA.
• Evaluation criteria, evaluator score sheets/rubrics, summary sheets, and selection memorandum.
• Any determinations of responsiveness/responsibility and the final award rationale.
• If these records are publicly available online, please provide the URL(s).

Format, Segregability & Fees
----------------------------
• Please provide the records electronically via reply email when feasible.
• If portions are exempt, kindly release any reasonably segregable parts with brief justification.
• If the estimated fees are expected to exceed $50, please provide a written cost estimate and await approval before proceeding.

Acknowledgment & Point of Contact
---------------------------------
Kindly acknowledge receipt of this request and provide an estimated timeline for response. If you require
clarification, please contact me using the information above.

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
                msg["Reply-To"] = f"{REQUESTER['name']} <{REQUESTER['email']}>"
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

# --------------------------------- Main -------------------------------------
def main():
    ap = argparse.ArgumentParser(description="Hawai‘i UIPA emailer (one email per row).")
    ap.add_argument("--input", required=True, help="Path to .xlsx/.xls/.csv with closed opportunities.")
    ap.add_argument("--send", action="store_true", help="Actually send emails (omit for dry run).")
    ap.add_argument("--to-override", help="Send to this address instead of govoffice.uipa@hawaii.gov.")
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

    # Add/refresh Purpose column (2 lines) per row
    purposes: List[str] = []
    for _, row in df.fillna("").iterrows():
        purposes.append(build_purpose(row))
    df["Purpose"] = purposes

    if args.skip_blank_rows:
        df = df.dropna(how="all")
    if args.limit and args.limit > 0:
        df = df.head(args.limit)

    skip_indices: Set[int] = set()
    if args.resume_log:
        skip_indices = load_sent_row_indices(args.resume_log)
        if skip_indices:
            print(f"[INFO] Resuming: will skip {len(skip_indices)} already-SENT rows based on {args.resume_log}")

    original_columns = list(df.columns)  # preserve original order for email enumeration
    total = len(df)
    print(f"[INFO] Loaded {total} rows from: {args.input}")

    smtp_conf = None
    if args.send:
        smtp_conf = load_smtp(args)
        print("[INFO] SMTP loaded. SENDING mode is ON.")

    to_addr = (args.to_override or DEFAULT_RECIPIENT).strip()
    print(f"[INFO] Recipient: {to_addr}")

    # Save an enriched copy (with Purpose) next to the input for auditing
    base, _ = os.path.splitext(args.input)
    enriched_path = f"{base}_with_purpose_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
    try:
        df.to_excel(enriched_path, index=False)
        print(f"[OK] Enriched copy with Purpose → {enriched_path}")
    except Exception as e:
        print(f"[WARN] Could not save enriched copy: {e}")

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
        purpose_text = s(row.get("Purpose", "")) or build_purpose(row)
        body = build_body(row, purpose_text, original_columns)

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
    try:
        log_df.to_excel(out_path, index=False)
        print(f"[OK] Wrote log → {out_path}")
    except Exception as e:
        print(f"[WARN] Could not write log Excel: {e}")
        print(log_df.head())

if __name__ == "__main__":
    main()
