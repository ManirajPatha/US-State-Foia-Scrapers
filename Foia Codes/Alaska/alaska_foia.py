#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
alaska_foia_emailer.py — Send Alaska Public Records requests to:
  law.recordsrequest@alaska.gov

Reads alaska_vss_results.csv (or .xlsx) and sends ONE email per row.
Automatically adds a professional "Purpose" column (2 lines) per row,
and includes ALL the CSV fields in a clean, aligned section.

USAGE EXAMPLES
--------------
# Preview only (no emails sent):
python alaska_foia_emailer.py --input "C:/.../alaska_vss_results.csv"

# Actually send (first 10 rows) using STARTTLS (port 587):
python alaska_foia_emailer.py --input "C:/.../alaska_vss_results.csv" --send --limit 10

# Send via SSL (port 465):
python alaska_foia_emailer.py --input "C:/.../alaska_vss_results.csv" --send --ssl --smtp-port 465

# Test to yourself:
python alaska_foia_emailer.py --input "C:/.../alaska_vss_results.csv" --send --to-override "pathamaniraj97@gmail.com"

# Resume (skip rows already marked SENT in a prior log):
python alaska_foia_emailer.py --input "C:/.../alaska_vss_results.csv" --send --resume-log "C:/.../alaska_send_log.xlsx"


REQUIREMENTS
------------
pip install pandas openpyxl python-dotenv python-dateutil
Create a .env in the same folder:

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

# ===================== YOUR FIXED REQUESTER DETAILS ===================== #
REQUESTER = {
    "name": "Maniraj Patha",
    "organization": "Southern Arkansas University",
    "address": "8181 Fannin St, Houston, TX 77054",
    "phone": "+1 682-405-5734",
    "email": "pathamaniraj97@gmail.com",
}
DEFAULT_RECIPIENT = "law.recordsrequest@alaska.gov"

# Known transient SMTP response codes
TRANSIENT_SMTP_CODES = {421, 450, 451, 452, 454}

# ============================ CSV FIELD SET ============================= #
# These are the exact columns found in your alaska_vss_results.csv.
EXPECTED_COLS = [
    "Description",
    "Department",
    "Solicitation Number",
    "Type",
]

# ============================== HELPERS ================================= #
def read_table(path: str) -> pd.DataFrame:
    ext = os.path.splitext(path)[1].lower()
    if ext == ".csv":
        return pd.read_csv(path)
    if ext in {".xlsx", ".xlsm"}:
        return pd.read_excel(path, engine="openpyxl")
    if ext == ".xls":
        # xlrd only supports .xls with a specific version; prefer saving to .xlsx
        raise RuntimeError("Please save the .xls as .xlsx or .csv and re-run.")
    raise ValueError(f"Unsupported file type: {ext}. Use .csv/.xlsx/.xlsm")

def nz(val: Any) -> str:
    """Normalize to a clean string ('' if NaN/None/blank)."""
    if val is None:
        return ""
    text = str(val).strip()
    if text.lower() in {"nan", "none"}:
        return ""
    return text

def first_nonempty(*vals: Any) -> str:
    for v in vals:
        t = nz(v)
        if t:
            return t
    return ""

def infer_subject(row: pd.Series) -> str:
    title = nz(row.get("Description")) or "Awarded Opportunity"
    sol = nz(row.get("Solicitation Number"))
    sol_part = f" ({sol})" if sol else ""
    return f"Public Records Request – Award Materials – {title}{sol_part} – {REQUESTER['organization']}"

def build_purpose_text(row: pd.Series) -> str:
    """
    Generates the two-line Purpose text, incorporating row context.
    """
    title = nz(row.get("Description"))
    sol = nz(row.get("Solicitation Number"))
    dept = nz(row.get("Department"))

    context_bits = [b for b in [title, sol, dept] if b]
    ctx = " / ".join(context_bits) if context_bits else "the referenced award"

    line1 = "Requesting copies of the winning proposal(s), executed contract, evaluation summary, and scoring rationale."
    line2 = f"If any fees apply, please notify me in advance; electronic delivery is preferred. (Context: {ctx})"
    return f"{line1}\n{line2}"

def format_kv(label: str, value: str, labellen: int = 22) -> str:
    """
    Clean, aligned key/value line: 'Label.............: value'
    """
    label = (label or "").strip()
    value = (value or "").strip()
    dots = "." * max(1, labellen - len(label))
    return f"{label}{dots}: {value}" if value else f"{label}{dots}:"

def professional_body(row: pd.Series, to_addr: str, purpose_text: str) -> str:
    """
    Professional, Alaska-ready request letter.
    Explicitly includes ALL CSV fields you have: Description, Department, Solicitation Number, Type.
    """
    today = datetime.now().strftime("%B %d, %Y")

    # Align and print the known CSV fields
    fields_block = "\n".join([
        format_kv("Description", nz(row.get("Description"))),
        format_kv("Department", nz(row.get("Department"))),
        format_kv("Solicitation Number", nz(row.get("Solicitation Number"))),
        format_kv("Type", nz(row.get("Type"))),
    ])

    return f"""Date: {today}

To: Alaska Department of Law – Records Request
Email: {to_addr}

Subject: Public Records Request – Award Documentation

Dear Records Officer,

I respectfully request access to public records associated with the award described below. I prefer to receive records electronically via email.

Award Reference (from State records)
------------------------------------
{fields_block}

Requester Information
---------------------
{format_kv("Name", REQUESTER['name'])}
{format_kv("Organization", REQUESTER['organization'])}
{format_kv("Address", REQUESTER['address'])}
{format_kv("Phone", REQUESTER['phone'])}
{format_kv("Email", REQUESTER['email'])}

Purpose of Request
------------------
{purpose_text}

Delivery & Fees
---------------
If any fees are anticipated, please provide a cost estimate before fulfillment. If portions of the materials are already available online, a direct link will suffice. If any parts are exempt, please provide the non-exempt portions and cite the specific legal basis for any redactions.

Thank you for your time and assistance.

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
    modes = [prefer_ssl, not prefer_ssl]  # try preferred first, then alternate
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

# ================================ MAIN =================================== #
def main():
    ap = argparse.ArgumentParser(description="Alaska Public Records (FOIA) emailer – one email per row.")
    ap.add_argument("--input", required=True, help="Path to alaska_vss_results.csv or .xlsx")
    ap.add_argument("--send", action="store_true", help="Actually send emails (omit for dry run).")
    ap.add_argument("--to-override", help="Send to this address instead of law.recordsrequest@alaska.gov.")
    ap.add_argument("--limit", type=int, default=0, help="If >0, only process this many rows.")
    ap.add_argument("--skip-blank-rows", action="store_true", help="Drop rows that are entirely blank.")
    ap.add_argument("--log-out", help="Optional path for send log (.xlsx). Default: alongside input.")
    ap.add_argument("--resume-log", help="Existing log to skip rows already SENT.")
    ap.add_argument("--ssl", action="store_true", help="Use SMTPS (SSL, usually port 465) instead of STARTTLS (587).")
    ap.add_argument("--smtp-host", help="Override SMTP host (else use .env).")
    ap.add_argument("--smtp-port", type=int, help="Override SMTP port (465 for SSL, 587 for STARTTLS).")
    args = ap.parse_args()

    df = read_table(args.input)

    # Validate expected columns are present; warn if missing
    missing = [c for c in EXPECTED_COLS if c not in df.columns]
    if missing:
        raise RuntimeError(
            "Your input is missing expected columns: "
            + ", ".join(missing)
            + "\nMake sure the file has: "
            + ", ".join(EXPECTED_COLS)
        )

    if args.skip_blank_rows:
        df = df.dropna(how="all")
    if args.limit and args.limit > 0:
        df = df.head(args.limit)

    # Auto-create Purpose column (always overwrite to keep it up to date)
    df["Purpose"] = [build_purpose_text(row) for _, row in df.iterrows()]

    skip_indices: Set[int] = set()
    if args.resume_log:
        skip_indices = load_sent_row_indices(args.resume_log)
        if skip_indices:
            print(f"[INFO] Resuming: skipping {len(skip_indices)} already-SENT rows from {args.resume_log}")

    total = len(df)
    print(f"[INFO] Loaded {total} rows from: {args.input}")

    smtp_conf = None
    if args.send:
        smtp_conf = load_smtp(args)
        print("[INFO] SMTP ready. SENDING mode is ON.")

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
        purpose_text = row["Purpose"]
        body = professional_body(row, to_addr, purpose_text)

        print("\n" + "=" * 90)
        print(f"[PREVIEW] Row {i+1}/{total}")
        print(f"TO: {to_addr}")
        print(f"SUBJECT: {subject}")
        print("-" * 90)
        print(body)
        print("=" * 90 + "\n")

        status = "DRY-RUN"
        err = ""
        if args.send:
            try:
                send_with_retries(smtp_conf, to_addr, subject, body, max_retries=5, prefer_ssl=args.ssl)
                status = "SENT"
                time.sleep(1.0)
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

    # Save updated table with Purpose
    base, ext = os.path.splitext(args.input)
    updated_out = f"{base}_with_purpose_{datetime.now().strftime('%Y%m%d_%H%M')}{ext or '.csv'}"
    try:
        if ext.lower() == ".csv":
            df.to_csv(updated_out, index=False)
        else:
            df.to_excel(updated_out if ext else f"{updated_out}.xlsx", index=False)
        print(f"[OK] Saved updated table (with Purpose) → {updated_out}")
    except Exception as e:
        print(f"[WARN] Could not save updated table with Purpose: {e}")

    # Save send log
    log_df = pd.DataFrame(results)
    out_log = args.log_out or f"{base}_send_log_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
    log_df.to_excel(out_log, index=False)
    print(f"[OK] Wrote send log → {out_log}")

if __name__ == "__main__":
    main()
