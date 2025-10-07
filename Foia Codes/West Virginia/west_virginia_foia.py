#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
west_virginia_foia.py — Send WV FOIA/Public Records requests via email,
one message per row from wv_vss_results.csv (or .xlsx/.xlsm/.xls).

Key features:
- Professional WV FOIA template
- Includes ALL columns from the input file inside the email body
- Auto-creates a two-line "Purpose" per row
- Dry-run preview vs. actual send
- SSL (465) or STARTTLS (587)
- Resume from a prior send log; limit/pause controls
- Optional: attach each row as a tiny CSV file (--attach-row)

Usage examples:

  # Preview only (no emails sent)
  python west_virginia_foia.py --input "C:/path/wv_vss_results.csv"

  # Send first 5 rows via STARTTLS
  python west_virginia_foia.py --input "C:/path/wv_vss_results.csv" --send --limit 5

  # Send via SSL/465
  python west_virginia_foia.py --input "C:/path/wv_vss_results.csv" --send --ssl

  # Override recipient for testing
  python west_virginia_foia.py --input "C:/path/wv_vss_results.csv" --send \
      --to-override "pathamaniraj97@gmail.com"

  # Resume without resending rows already SENT
  python west_virginia_foia.py --input "C:/path/wv_vss_results.csv" --send \
      --resume-log "C:/path/wv_vss_results_send_log.xlsx"

  # Attach each row as CSV
  python west_virginia_foia.py --input "C:/path/wv_vss_results.csv" --send --attach-row

Dependencies:
  pip install pandas openpyxl python-dotenv python-dateutil

.env (same folder) example:
  SMTP_HOST=smtp.gmail.com
  SMTP_PORT=587
  SMTP_USERNAME=pathamaniraj97@gmail.com
  SMTP_PASSWORD=YOUR_16_CHAR_APP_PASSWORD
  SENDER_NAME=Maniraj Patha
  SENDER_EMAIL=pathamaniraj97@gmail.com
"""

import os
import ssl
import csv
import io
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

# -------------------------- REQUESTER (your details) --------------------------
REQUESTER = {
    "name": "Maniraj Patha",
    "organization": "",  # e.g., "Mani Solutions" or "Southern Arkansas University" (optional)
    "address": "8181 Fannin Street, 77054",
    "phone": "6824055734",
    "email": "pathamaniraj97@gmail.com",
}

# Default WV recipient (as provided)
DEFAULT_RECIPIENT = "Ppickensmj@wvbar.org"

# Common column candidates (used for nicer subject/purpose)
CANDIDATE_ID_COLS         = ["notice_id", "solicitation_number", "bid_id", "rfp_number", "solicitation_id", "reference_no"]
CANDIDATE_TITLE_COLS      = ["title", "solicitation_title", "description", "project_title"]
CANDIDATE_AWARD_DATE_COLS = ["award_date", "awarded_on", "finalize_date", "awardposteddate", "award_post_date"]
CANDIDATE_VENDOR_COLS     = ["awarded_vendor", "vendor", "contractor", "supplier", "awardee"]
CANDIDATE_AMOUNT_COLS     = ["award_amount", "contract_value", "amount", "value", "total_award"]
CANDIDATE_URL_COLS        = ["page_url", "detail_url", "url", "source_url"]
CANDIDATE_ATTACH_COLS     = ["attachments", "files", "links"]
CANDIDATE_DEPT_COLS       = ["department", "agency", "buyer", "office"]

TRANSIENT_SMTP_CODES = {421, 450, 451, 452, 454}

# --------------------------------- Helpers ---------------------------------

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
                "Reading legacy .xls requires 'xlrd==2.0.1'.\n"
                "Try: pip uninstall -y xlrd && pip install xlrd==2.0.1\n"
                "Or Save As .xlsx and rerun."
            ) from e
    raise ValueError(f"Unsupported file type: {ext}")

def s(val: Any) -> str:
    return "" if pd.isna(val) else str(val).strip()

def first_nonempty(row: pd.Series, candidates: List[str]) -> str:
    for c in candidates:
        if c in row and s(row[c]):
            return s(row[c])
    return ""

def make_purpose_text(row: pd.Series) -> str:
    """
    Two-line Purpose text per row: request winning proposals & strategies, and fee note.
    """
    title  = first_nonempty(row, CANDIDATE_TITLE_COLS) or "the referenced opportunity"
    sol    = first_nonempty(row, CANDIDATE_ID_COLS)
    dept   = first_nonempty(row, CANDIDATE_DEPT_COLS)
    vendor = first_nonempty(row, CANDIDATE_VENDOR_COLS)

    seg_id    = f" (Solicitation #{sol})" if sol else ""
    seg_dept  = f" from {dept}" if dept else ""
    seg_vendr = f" awarded to {vendor}" if vendor else " (winning vendor details requested)"

    line1 = f"Requesting copies of the winning proposal and evaluation/selection materials for {title}{seg_id}{seg_dept}{seg_vendr}."
    line2 = "If fees apply, please notify me in advance; electronic delivery is preferred."
    return f"{line1}\n{line2}"

def infer_subject(row: pd.Series) -> str:
    title = first_nonempty(row, CANDIDATE_TITLE_COLS) or "Awarded Opportunity"
    nid   = first_nonempty(row, CANDIDATE_ID_COLS)
    nid_seg = f" ({nid})" if nid else ""
    org  = f" – {REQUESTER['organization']}" if REQUESTER['organization'] else ""
    return f"WV FOIA/Public Records Request – {title}{nid_seg}{org}"

def render_row_data_block(row: pd.Series) -> str:
    """
    Creates a neat, aligned "key: value" block containing ALL columns
    from the input row, preserving the original column order.
    """
    items = [(str(col), s(row[col])) for col in row.index]
    # Compute label width for alignment (cap at 40 to avoid very wide labels)
    label_w = min(max((len(k) for k, _ in items), default=10), 40)
    lines = []
    for k, v in items:
        # Show empty values as '-' to make it obvious it's blank in the source
        val = v if v else "-"
        lines.append(f"{k.ljust(label_w)} : {val}")
    return "\n".join(lines)

def build_body(row: pd.Series, recipient_email: str) -> str:
    today   = datetime.now().strftime("%B %d, %Y")
    purpose = make_purpose_text(row)
    row_block = render_row_data_block(row)

    requester_org = f"\nOrganization: {REQUESTER['organization']}" if REQUESTER["organization"] else ""
    requester_block = f"""Requester
---------
Name: {REQUESTER['name']}{requester_org}
Address: {REQUESTER['address']}
Phone: {REQUESTER['phone']}
Email: {REQUESTER['email']}"""

    return f"""Date: {today}

To: West Virginia Public Records Contact
Email: {recipient_email}

Subject: WV FOIA/Public Records Request (Award Materials)

Dear Records Officer,

Pursuant to West Virginia’s Freedom of Information Act, I respectfully request public records related to the awarded/closed opportunity noted below. The following "Row Data" is a direct reflection of one record from my research dataset and is included to facilitate precise identification.

Purpose of Request
------------------
{purpose}

Requested Records
-----------------
• The winning (awarded) vendor proposal and pricing (to the extent releasable)
• Evaluation materials, scoring/ranking sheets, and recommendation/award justification
• Award or intent-to-award notices and the executed agreement/contract (including amendments/renewals)
• Any addenda that materially affected the award and the final statement of work

Delivery & Fees
---------------
Please provide records electronically via email. If there will be any fees, kindly advise with a cost breakdown and turnaround time before fulfillment.

Row Data (from wv_vss_results.csv)
----------------------------------
{row_block}

{requester_block}

Thank you for your assistance. Please let me know if any additional information is required to process this request.

Sincerely,
{REQUESTER['name']}
{REQUESTER['email']}
{REQUESTER['phone']}
"""

# ---------------------------- SMTP helpers ----------------------------

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
    sender_name  = os.getenv("SENDER_NAME", REQUESTER["name"])
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
        "host": host, "port": port,
        "username": username, "password": password,
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
    attachments: Optional[List[Dict[str, Any]]] = None,
    max_retries: int = 5,
    prefer_ssl: bool = False,
) -> None:
    attempt = 0
    delay_seq = [5, 15, 45, 90, 180]
    last_exc: Optional[Exception] = None
    modes = [prefer_ssl, not prefer_ssl]  # try preferred, then the alternate
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

                # Attachments (optional)
                if attachments:
                    for att in attachments:
                        msg.add_attachment(
                            att["data"],
                            maintype=att.get("maintype", "text"),
                            subtype=att.get("subtype", "csv"),
                            filename=att.get("filename", "row.csv"),
                        )

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
                "SMTP authentication failed. For Gmail, enable 2-Step Verification "
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

def row_as_csv_bytes(row: pd.Series) -> bytes:
    """
    Convert a single row (ALL columns) into a small CSV bytes object.
    """
    output = io.StringIO()
    writer = csv.writer(output)
    # header
    writer.writerow(list(row.index))
    # values
    writer.writerow([s(v) for v in row.values])
    return output.getvalue().encode("utf-8")

# --------------------------------- Main ------------------------------------

def main():
    ap = argparse.ArgumentParser(description="WV FOIA emailer (one email per row).")
    ap.add_argument("--input", required=True, help="Path to .csv/.xlsx/.xlsm/.xls (e.g., wv_vss_results.csv).")
    ap.add_argument("--send", action="store_true", help="Actually send emails (omit for dry run).")
    ap.add_argument("--to-override", help="Override recipient email (default WV address).")
    ap.add_argument("--limit", type=int, default=0, help="If >0, only process this many rows.")
    ap.add_argument("--skip-blank-rows", action="store_true", help="Skip rows that are entirely blank.")
    ap.add_argument("--log-out", help="Optional path for results log (.xlsx). Default: alongside input.")
    ap.add_argument("--pause", type=float, default=1.0, help="Seconds to sleep between emails.")
    ap.add_argument("--resume-log", help="Existing log to skip rows already SENT.")
    ap.add_argument("--ssl", action="store_true", help="Use SMTPS (SSL) on 465 instead of STARTTLS on 587.")
    ap.add_argument("--smtp-host", help="Override SMTP host (default from .env).")
    ap.add_argument("--smtp-port", type=int, help="Override SMTP port (465 for SSL, 587 for STARTTLS).")
    ap.add_argument("--attach-row", action="store_true", help="Attach each row as a tiny CSV file.")
    args = ap.parse_args()

    # Load data
    df = read_table(args.input)
    if args.skip_blank_rows:
        df = df.dropna(how="all")
    if args.limit and args.limit > 0:
        df = df.head(args.limit)

    # Ensure "Purpose" column exists and is populated
    df = df.copy()
    df["Purpose"] = df.apply(make_purpose_text, axis=1)

    # Persist a convenience copy with Purpose
    base, _ = os.path.splitext(args.input)
    purpose_out = f"{base}_with_purpose_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
    try:
        df.to_excel(purpose_out, index=False)
        print(f"[OK] Wrote workbook with Purpose → {purpose_out}")
    except Exception as e:
        print(f"[WARN] Could not write Purpose workbook: {e}")

    # Resume support
    skip_indices: Set[int] = set()
    if args.resume_log:
        skip_indices = load_sent_row_indices(args.resume_log)
        if skip_indices:
            print(f"[INFO] Resuming: will skip {len(skip_indices)} already-SENT rows from {args.resume_log}")

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
            print(f"[SKIP] Row {i} already SENT (resume).")
            results.append({
                "row_index": i, "to": to_addr, "subject": "(skipped via resume)",
                "status": "SKIPPED", "error": "", "timestamp": datetime.now().isoformat(timespec="seconds")
            })
            continue

        subject = infer_subject(row)
        body = build_body(row, to_addr)

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
                attachments = None
                if args.attach_row:
                    attachments = [{
                        "data": row_as_csv_bytes(row),
                        "maintype": "text",
                        "subtype": "csv",
                        "filename": f"wv_row_{i}.csv",
                    }]
                send_with_retries(
                    smtp_conf,
                    to_addr,
                    subject,
                    body,
                    attachments=attachments,
                    max_retries=5,
                    prefer_ssl=args.ssl
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

    # Write send log
    log_df = pd.DataFrame(results)
    out_path = args.log_out or f"{base}_send_log_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
    try:
        log_df.to_excel(out_path, index=False)
        print(f"[OK] Wrote send log → {out_path}")
    except Exception as e:
        print(f"[WARN] Could not write send log: {e}")

if __name__ == "__main__":
    main()
