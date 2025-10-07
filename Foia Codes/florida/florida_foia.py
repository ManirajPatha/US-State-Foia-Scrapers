#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
florida_foia.py — Send one Florida Public Records (FOIA) request email per row.

✅ Professional template (explicit Chapter 119, F.S.)
✅ Adds/updates a "Purpose" column (2-line, tailored per row)
✅ Includes ALL CSV columns in the email body (clean summary + full list)
✅ Dry-run by default; add --send to actually send
✅ Retry logic, SSL/STARTTLS toggle, resume-safe logging

USAGE EXAMPLES (Windows):
  # Dry run (no emails sent)
  python florida_foia.py --input "C:/path/florida_bids.csv"

  # Send first 5 rows
  python florida_foia.py --input "C:/path/florida_bids.csv" --send --limit 5

  # Save augmented sheet with Purpose
  python florida_foia.py --input "C:/path/florida_bids.csv" --augment-out "C:/path/fl_bids_with_purpose.xlsx"

  # Test send to yourself instead of DOS
  python florida_foia.py --input "C:/path/florida_bids.csv" --send --to-override "pathamaniraj97@gmail.com"

  # Use SSL/465 instead of STARTTLS/587
  python florida_foia.py --input "C:/path/florida_bids.csv" --send --ssl --smtp-port 465

REQUIREMENTS:
  pip install pandas openpyxl python-dotenv python-dateutil

.env in the same folder:
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
    SMTP, SMTP_SSL, SMTPServerDisconnected, SMTPAuthenticationError,
    SMTPResponseException, SMTPDataError
)

# -------------------------- Requester details (from your message) --------------------------
REQUESTER = {
    "name": "Maniraj Patha",
    "phone": "+1 6824055734",
    "email": "pathamaniraj97@gmail.com",
    "address": "8181 fannin street",
    "zipcode": "77054",
}

# Default recipient (Florida Department of State)
DEFAULT_RECIPIENT = "PublicRecords@DOS.fl.gov"

# Transient SMTP codes for retry logic
TRANSIENT_SMTP_CODES = {421, 450, 451, 452, 454}

# ------------------------------- Helpers -----------------------------------
def read_table(path: str) -> pd.DataFrame:
    ext = os.path.splitext(path)[1].lower()
    if ext == ".csv":
        # try utf-8, fallback to latin-1 for stray encodings
        try:
            return pd.read_csv(path)
        except UnicodeDecodeError:
            return pd.read_csv(path, encoding="latin-1")
    if ext in {".xlsx", ".xlsm"}:
        return pd.read_excel(path, engine="openpyxl")
    if ext == ".xls":
        try:
            return pd.read_excel(path, engine="xlrd")
        except Exception as e:
            raise RuntimeError(
                "Reading legacy .xls requires xlrd==2.0.1. Either install it or save as .xlsx.\n"
                "  pip uninstall -y xlrd && pip install xlrd==2.0.1"
            ) from e
    raise ValueError(f"Unsupported input type: {ext}. Use .csv/.xlsx/.xlsm/.xls")

def s(val: Any) -> str:
    return "" if pd.isna(val) else str(val).strip()

# Candidate columns we’ll look for to craft a better subject/summary
CANDIDATE_ID_COLS = ["notice_id", "solicitation_number", "bid_id", "rfp_number", "solicitation_id", "reference_no", "reference_id"]
CANDIDATE_TITLE_COLS = ["title", "solicitation_title", "description", "project_title", "bid_title"]
CANDIDATE_AWARD_DATE_COLS = ["award_date", "awarded_on", "finalize_date", "award_posted_date", "award_post_date", "posted_date"]
CANDIDATE_VENDOR_COLS = ["awarded_vendor", "vendor", "contractor", "supplier", "awardee"]
CANDIDATE_AMOUNT_COLS = ["award_amount", "contract_value", "amount", "value", "total_award"]
CANDIDATE_URL_COLS = ["page_url", "detail_url", "url", "source_url"]
CANDIDATE_ATTACH_COLS = ["attachments", "files", "links"]

def first_nonempty(row: pd.Series, candidates: List[str]) -> str:
    for c in candidates:
        if c in row and s(row[c]):
            return s(row[c])
    return ""

def pad(text: str, width: int) -> str:
    # For aligned “All Provided Fields” block
    return text + " " * max(0, width - len(text))

def purpose_two_liner(row: pd.Series) -> str:
    title = first_nonempty(row, CANDIDATE_TITLE_COLS) or "the referenced opportunity"
    ident = first_nonempty(row, CANDIDATE_ID_COLS)
    id_seg = f" ({ident})" if ident else ""
    line1 = f"Requesting copies of the winning proposal(s), evaluation records, and documented strategies for “{title}”{id_seg}."
    line2 = "If fees will exceed $50, please advise in advance with an itemized estimate."
    return f"{line1}\n{line2}"

def ordered_summary_block(row: pd.Series) -> str:
    """
    Professional “Opportunity Summary”: common fields first (if present), then others.
    """
    lines = []
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
            lines.append(f"- {label}: {val}")

    # Add remaining (non-empty) columns not already covered
    used = {c.lower() for c in (
        CANDIDATE_ID_COLS + CANDIDATE_TITLE_COLS + CANDIDATE_AWARD_DATE_COLS +
        CANDIDATE_VENDOR_COLS + CANDIDATE_AMOUNT_COLS + CANDIDATE_URL_COLS +
        CANDIDATE_ATTACH_COLS
    )}
    extras = []
    for col, val in row.items():
        if col.lower() in used:
            continue
        txt = s(val)
        if txt:
            extras.append(f"- {col}: {txt}")

    return "\n".join(lines + extras) if (lines or extras) else "- (No non-empty fields found)"

def full_fields_block(row: pd.Series) -> str:
    """
    “All Provided Fields” — lists every non-empty column/value from the CSV row.
    Rendered as aligned key: value pairs for clarity (monospace-friendly).
    """
    pairs = [(str(col), s(val)) for col, val in row.items() if s(val)]
    if not pairs:
        return "(row is empty)"
    w = max(len(k) for k, _ in pairs)
    return "\n".join(f"{pad(k, w)} : {v}" for k, v in pairs)

def infer_subject(row: pd.Series) -> str:
    title = first_nonempty(row, CANDIDATE_TITLE_COLS) or "Public Records Request"
    nid = first_nonempty(row, CANDIDATE_ID_COLS)
    nid_seg = f" ({nid})" if nid else ""
    return f"Florida DOS — Public Records Request — {title}{nid_seg} — Maniraj Patha"

def build_body(row: pd.Series, purpose_text: str) -> str:
    today = datetime.now().strftime("%B %d, %Y")
    summary = ordered_summary_block(row)
    all_fields = full_fields_block(row)

    return f"""Date: {today}

To: Florida Department of State — Public Records
Email: {DEFAULT_RECIPIENT}

Re: Public Records Request (Chapter 119, Florida Statutes)

Dear Records Custodian,

Under Chapter 119, Florida Statutes, I respectfully request access to public records associated with the opportunity identified below.

OPPORTUNITY SUMMARY
-------------------
{summary}

PURPOSE OF REQUEST
------------------
{purpose_text}

DELIVERY & FEES
---------------
Please provide responsive records electronically to this email address. If fees will exceed $50, kindly notify me in advance with an itemized cost estimate. If any information is confidential or exempt, please cite the specific statutory exemption and release all reasonably segregable portions.

SCOPE REFINEMENT (IF NEEDED)
----------------------------
If narrowing is required to reduce time or cost, please advise on available options (e.g., specific date ranges, document types, or custodian offices), and I will promptly refine the scope.

REQUESTER DETAILS
-----------------
Name   : {REQUESTER['name']}
Address: {REQUESTER['address']}
ZIP    : {REQUESTER['zipcode']}
Phone  : {REQUESTER['phone']}
Email  : {REQUESTER['email']}

ALL PROVIDED FIELDS (FROM YOUR SPREADSHEET ROW)
-----------------------------------------------
{all_fields}

Thank you for your assistance. I appreciate your time and service.

Sincerely,
{REQUESTER['name']}
{REQUESTER['email']} | {REQUESTER['phone']}
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
    modes = [prefer_ssl, not prefer_ssl]  # try preferred, then alternate
    mode_index = 0

    while attempt <= max_retries:
        use_ssl = modes[min(mode_index, len(modes) - 1)]
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
                mode_index = 1  # flip STARTTLS<->SSL early
        except SMTPDataError as e:
            if e.smtp_code in TRANSIENT_SMTP_CODES:
                last_exc = e
            else:
                raise
        except SMTPAuthenticationError as e:
            raise RuntimeError(
                "SMTP authentication failed. For Gmail, use a 16-char App Password with 2FA."
            ) from e
        except SMTPResponseException as e:
            if 400 <= e.smtp_code < 500 or e.smtp_code in TRANSIENT_SMTP_CODES:
                last_exc = e
            else:
                raise

        if attempt == max_retries:
            break
        time.sleep(delay_seq[min(attempt, len(delay_seq) - 1)])
        attempt += 1

    raise RuntimeError(f"SMTP send failed after {max_retries+1} attempts: {last_exc}")

def load_sent_row_indices(resume_log_path: str) -> Set[int]:
    try:
        log = pd.read_excel(resume_log_path)
        sent = log.loc[log["status"].str.upper() == "SENT", "row_index"]
        return set(int(i) for i in sent.dropna().tolist())
    except Exception:
        return set()

# --------------------------------- CLI / Main ------------------------------------
def main():
    ap = argparse.ArgumentParser(description="Florida DOS public records emailer (one email per row).")
    ap.add_argument("--input", required=True, help="Path to .csv/.xlsx/.xlsm/.xls (e.g., florida_bids.csv).")
    ap.add_argument("--send", action="store_true", help="Actually send emails (omit for dry run).")
    ap.add_argument("--to-override", help="Override recipient (default: PublicRecords@DOS.fl.gov).")
    ap.add_argument("--limit", type=int, default=0, help="If >0, process only this many rows.")
    ap.add_argument("--skip-blank-rows", action="store_true", help="Skip rows entirely blank.")
    ap.add_argument("--log-out", help="Optional path for results log (.xlsx). Default: alongside input.")
    ap.add_argument("--pause", type=float, default=1.0, help="Seconds to sleep between sends.")
    ap.add_argument("--resume-log", help="Skip rows already SENT per prior log (.xlsx).")
    ap.add_argument("--ssl", action="store_true", help="Use SMTPS (SSL) instead of STARTTLS.")
    ap.add_argument("--smtp-host", help="Override SMTP host.")
    ap.add_argument("--smtp-port", type=int, help="Override SMTP port.")
    ap.add_argument("--augment-out", help="Write augmented table with Purpose column (.xlsx/.csv).")
    args = ap.parse_args()

    df = read_table(args.input)

    # ✅ FIXED: use underscore, not hyphen
    if args.skip_blank_rows:
        df = df.dropna(how="all")

    if args.limit and args.limit > 0:
        df = df.head(args.limit)

    # Ensure/refresh Purpose column
    purposes: List[str] = []
    for _, row in df.iterrows():
        purposes.append(purpose_two_liner(row))
    df["Purpose"] = purposes

    if args.augment_out:
        ext = os.path.splitext(args.augment_out)[1].lower()
        if ext in {".xlsx", ".xlsm"}:
            df.to_excel(args.augment_out, index=False)
        elif ext == ".csv":
            df.to_csv(args.augment_out, index=False)
        else:
            df.to_excel(f"{args.augment_out}.xlsx", index=False)
        print(f"[OK] Wrote augmented table with Purpose → {args.augment_out}")

    skip_indices: Set[int] = set()
    if args.resume_log:
        skip_indices = load_sent_row_indices(args.resume_log)
        if skip_indices:
            print(f"[INFO] Resuming: will skip {len(skip_indices)} rows already SENT (from {args.resume_log}).")

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
                "row_index": i, "to": to_addr, "subject": "(skipped via resume)",
                "status": "SKIPPED", "error": "", "timestamp": datetime.now().isoformat(timespec="seconds"),
            })
            continue

        subject = infer_subject(row)
        body = build_body(row, df.loc[i, "Purpose"])

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
            "row_index": i, "to": to_addr, "subject": subject, "status": status,
            "error": err, "timestamp": datetime.now().isoformat(timespec="seconds"),
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
