#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
illinois_foia_emailer.py — Professional FOIA emails to Illinois IDHR (IDHR.FOIA@illinois.gov),
one per CSV/XLSX row. Adds a 'Purpose' column (2 lines) and logs send status.

Usage:
  # Preview only (no emails sent); also writes an augmented Excel with Purpose
  python illinois_foia_emailer.py --input "C:/path/il_bidbuy_results.csv"

  # Actually send (STARTTLS on 587 by default)
  python illinois_foia_emailer.py --input "C:/path/il_bidbuy_results.csv" --send

  # Useful options
  python illinois_foia_emailer.py --input "C:/path/file.csv" --send --limit 10 --pause 1.5 \
      --log-out "C:/path/il_send_log.xlsx" --to-override "yourtest@inbox.com" --cc "pathamaniraj97@gmail.com"
  python illinois_foia_emailer.py --input "C:/path/file.csv" --send --ssl  # SMTPS/465

Requires:
  pip install pandas openpyxl python-dotenv python-dateutil
  # If using legacy .xls:
  pip uninstall -y xlrd && pip install xlrd==2.0.1
"""

import os
import ssl
import time
import socket
import smtplib
import argparse
from datetime import datetime
from typing import Any, Dict, List, Optional, Set, Iterable, Tuple

import pandas as pd
from email.message import EmailMessage
from email.utils import formatdate
from dotenv import load_dotenv
from smtplib import (
    SMTP, SMTP_SSL,
    SMTPServerDisconnected, SMTPAuthenticationError,
    SMTPResponseException, SMTPDataError
)

# -------------------------- Requester (your fixed details) --------------------------
REQUESTER = {
    "name": "Maniraj Patha",
    "organization": "Southern Arkansas University",
    "address": "8181 Fannin St, Houston, TX 77054",
    "phone": "+1 6824055734",
    "email": "pathamaniraj97@gmail.com",
}

# Illinois IDHR FOIA recipient (default)
DEFAULT_RECIPIENT = "IDHR.FOIA@illinois.gov"

# Transient SMTP codes (retryable)
TRANSIENT_SMTP_CODES = {421, 450, 451, 452, 454}

# ------------------------------- Column candidates -------------------------------
CANDIDATE_ID_COLS = [
    "notice_id", "solicitation_number", "bid_id", "rfp_number",
    "solicitation_id", "reference_no", "contract_number"
]
CANDIDATE_TITLE_COLS = ["title", "solicitation_title", "description", "project_title"]
CANDIDATE_AWARD_DATE_COLS = ["award_date", "awarded_on", "finalize_date", "awardposteddate", "award_post_date"]
CANDIDATE_VENDOR_COLS = ["awarded_vendor", "vendor", "contractor", "supplier", "awardee"]
CANDIDATE_AMOUNT_COLS = ["award_amount", "contract_value", "amount", "value", "total_award"]
CANDIDATE_URL_COLS = ["page_url", "detail_url", "url", "source_url"]
CANDIDATE_ATTACH_COLS = ["attachments", "files", "links"]
CANDIDATE_AGENCY_COLS = ["agency", "department", "buyer", "issuing_office"]

# ------------------------------- Utilities -------------------------------
def s(val: Any) -> str:
    return "" if pd.isna(val) else str(val).strip()

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
                "This .xls requires xlrd==2.0.1.\n"
                "Run:\n  pip uninstall -y xlrd && pip install xlrd==2.0.1\n"
                "Or Save As .xlsx and rerun."
            ) from e
    raise ValueError(f"Unsupported file type: {ext}. Use .csv, .xlsx, .xlsm, or .xls")

def first_nonempty(row: pd.Series, candidates: Iterable[str]) -> str:
    """Case-insensitive pick of the first non-empty candidate column."""
    ci_map = {col.lower(): col for col in row.index}
    for c in candidates:
        if c.lower() in ci_map and s(row[ci_map[c.lower()]]):
            return s(row[ci_map[c.lower()]])
    return ""

def inferred_fields(row: pd.Series) -> Dict[str, str]:
    """Pull commonly useful award fields (if present)."""
    return {
        "Notice / Solicitation #": first_nonempty(row, CANDIDATE_ID_COLS),
        "Title": first_nonempty(row, CANDIDATE_TITLE_COLS),
        "Issuing Agency": first_nonempty(row, CANDIDATE_AGENCY_COLS),
        "Award Date": first_nonempty(row, CANDIDATE_AWARD_DATE_COLS),
        "Awarded Vendor": first_nonempty(row, CANDIDATE_VENDOR_COLS),
        "Award Amount": first_nonempty(row, CANDIDATE_AMOUNT_COLS),
        "Details URL": first_nonempty(row, CANDIDATE_URL_COLS),
        "Attachments": first_nonempty(row, CANDIDATE_ATTACH_COLS),
    }

def fields_block_all(row: pd.Series) -> str:
    """
    Professional 'Opportunity Details' block that:
      1) shows inferred/priority fields first (if non-empty),
      2) then lists every *other* non-empty column in the row.
    """
    # 1) Priority fields
    pri = {k: v for k, v in inferred_fields(row).items() if v}

    # Build set of "used" original column names to avoid duplication
    used_originals = set()
    for grp in [
        CANDIDATE_ID_COLS, CANDIDATE_TITLE_COLS, CANDIDATE_AGENCY_COLS,
        CANDIDATE_AWARD_DATE_COLS, CANDIDATE_VENDOR_COLS, CANDIDATE_AMOUNT_COLS,
        CANDIDATE_URL_COLS, CANDIDATE_ATTACH_COLS
    ]:
        used_originals.update(c.lower() for c in grp)

    # 2) Remaining non-empty columns (stable order by original appearance)
    extras: List[Tuple[str, str]] = []
    for col in row.index:
        if col.lower() in used_originals:
            continue
        val = s(row[col])
        if val:
            extras.append((col, val))

    # Nicely align with bullets
    lines: List[str] = []
    def add_block(items: List[Tuple[str, str]]):
        if not items:
            return
        width = max(len(k) for k, _ in items)
        for k, v in items:
            lines.append(f"• {k.ljust(width)} : {v}")

    # Priority block (preserve the curated order)
    pri_items = [(k, pri[k]) for k in [
        "Notice / Solicitation #", "Title", "Issuing Agency", "Award Date",
        "Awarded Vendor", "Award Amount", "Details URL", "Attachments"
    ] if k in pri]
    add_block(pri_items)

    # Extra fields
    add_block(extras)

    return "\n".join(lines) if lines else "• (No non-empty fields present in this row)"

def build_purpose(row: pd.Series) -> str:
    """Two concise lines tailored to the row: winning proposals + evaluation strategies + fee note."""
    title = first_nonempty(row, CANDIDATE_TITLE_COLS) or "the referenced solicitation"
    nid = first_nonempty(row, CANDIDATE_ID_COLS)
    id_seg = f" (ID: {nid})" if nid else ""
    agency = first_nonempty(row, CANDIDATE_AGENCY_COLS)
    agency_seg = f" from {agency}" if agency else ""
    line1 = f"Requesting the winning proposal(s), evaluation summaries/score sheets, and award rationale for {title}{id_seg}{agency_seg}."
    line2 = "If any fees apply, please provide an estimate and obtain approval before fulfillment."
    return f"{line1}\n{line2}"

def infer_subject(row: pd.Series) -> str:
    title = first_nonempty(row, CANDIDATE_TITLE_COLS) or "Awarded Opportunity"
    nid = first_nonempty(row, CANDIDATE_ID_COLS)
    nid_seg = f" ({nid})" if nid else ""
    return f"Illinois IDHR – FOIA Request – {title}{nid_seg} – {REQUESTER['name']}"

def build_body(row: pd.Series, purpose_text: str, to_addr: str) -> str:
    """
    Professional FOIA letter:
      - Statute reference (kept general to avoid overstating specifics)
      - Purpose block (2 lines)
      - Opportunity Details: includes ALL non-empty columns from the CSV row
      - Requester block
      - Delivery & fees / redaction language / statutory timeframe (general)
    """
    today = datetime.now().strftime("%B %d, %Y")
    details = fields_block_all(row)

    return f"""Date: {today}

To: Illinois Department of Human Rights – FOIA Office
Email: {to_addr}

Subject: Freedom of Information Act Request (Award/Contract Records)

Dear FOIA Officer,

Pursuant to the Illinois Freedom of Information Act, I respectfully request access to records associated with the opportunity described below.

Purpose of Request
------------------
{purpose_text}

Records Requested
-----------------
• Winning proposal(s)/Best-and-Final-Offer (if applicable)
• Evaluation committee score sheets, summaries, and recommendation memos
• Award justification/rationale and notice of award
• Final executed contract, including pricing schedules
• All amendments/modifications and related correspondence (if any)

Opportunity Details
-------------------
{details}

Requester Information
---------------------
Name         : {REQUESTER['name']}
Organization : {REQUESTER['organization']}
Address      : {REQUESTER['address']}
Phone        : {REQUESTER['phone']}
Email        : {REQUESTER['email']}

Format, Delivery, and Fees
--------------------------
Please provide the records electronically via email. If costs will exceed $50, kindly share a written estimate and await approval prior to fulfillment.
If any information is exempt, please redact only the exempt portions and release the remainder, citing the applicable statutory exemption(s) in your response.

Response Timeline
-----------------
I appreciate your response within the timeframe provided by Illinois FOIA.

Thank you for your assistance.

Sincerely,
{REQUESTER['name']}
{REQUESTER['organization']}
{REQUESTER['email']}
{REQUESTER['phone']}
"""

# ------------------------------- SMTP helpers -------------------------------
def load_smtp(args=None) -> Dict[str, str]:
    load_dotenv()
    host = os.getenv("SMTP_HOST", "smtp.gmail.com")
    port = int(os.getenv("SMTP_PORT", "587"))
    if getattr(args, "smtp_host", None):
        host = args.smtp_host
    if getattr(args, "smtp_port", None):
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
    cc_addrs: Optional[List[str]] = None,
    max_retries: int = 5,
    prefer_ssl: bool = False,
) -> None:
    attempt = 0
    delay_seq = [5, 15, 45, 90, 180]
    last_exc: Optional[Exception] = None
    modes = [prefer_ssl, not prefer_ssl]  # try preferred, then flip once
    mode_index = 0

    while attempt <= max_retries:
        use_ssl = modes[min(mode_index, len(modes) - 1)]
        try:
            with _open_smtp(smtp_conf, use_ssl=use_ssl) as server:
                msg = EmailMessage()
                msg["From"] = f"{smtp_conf['sender_name']} <{smtp_conf['sender_email']}>"
                msg["To"] = to_addr
                if cc_addrs:
                    msg["Cc"] = ", ".join(cc_addrs)
                msg["Date"] = formatdate(localtime=True)
                msg["Subject"] = subject
                msg.set_content(body)
                recipients = [to_addr] + (cc_addrs or [])
                server.send_message(msg, to_addrs=recipients)
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
        time.sleep(delay_seq[min(attempt, len(delay_seq) - 1)])
        attempt += 1

    raise RuntimeError(f"SMTP send failed after {max_retries + 1} attempts: {last_exc}")

def load_sent_row_indices(resume_log_path: str) -> Set[int]:
    try:
        log = pd.read_excel(resume_log_path)
        sent = log.loc[log["status"].str.upper() == "SENT", "row_index"]
        return set(int(i) for i in sent.dropna().tolist())
    except Exception:
        return set()

# --------------------------------- Main ------------------------------------
def main():
    ap = argparse.ArgumentParser(description="Illinois FOIA emailer (adds Purpose, includes ALL row fields).")
    ap.add_argument("--input", required=True, help="Path to .csv/.xlsx/.xls (e.g., il_bidbuy_results.csv).")
    ap.add_argument("--send", action="store_true", help="Actually send emails (omit for dry run).")
    ap.add_argument("--to-override", help="Send to this address instead of IDHR.FOIA@illinois.gov.")
    ap.add_argument("--cc", help="Comma-separated CC list (e.g., 'you@a.com,other@b.com').")
    ap.add_argument("--limit", type=int, default=0, help="If >0, only process this many rows.")
    ap.add_argument("--skip-blank-rows", action="store_true", help="Skip entirely blank rows.")
    ap.add_argument("--log-out", help="Path for results log (.xlsx). Default: alongside input.")
    ap.add_argument("--pause", type=float, default=1.0, help="Seconds to sleep between sends.")
    ap.add_argument("--resume-log", help="Existing log to skip rows already SENT.")
    ap.add_argument("--ssl", action="store_true", help="Use SMTPS (465) instead of STARTTLS (587).")
    ap.add_argument("--smtp-host", help="Override SMTP host.")
    ap.add_argument("--smtp-port", type=int, help="Override SMTP port.")
    args = ap.parse_args()

    df = read_table(args.input)
    if args.skip_blank_rows:
        df = df.dropna(how="all")
    if args.limit and args.limit > 0:
        df = df.head(args.limit)

    # Create/overwrite Purpose column (2 lines)
    df["Purpose"] = [build_purpose(row) for _, row in df.iterrows()]

    # Save augmented copy with Purpose
    base, _ = os.path.splitext(args.input)
    augmented_out = f"{base}_with_purpose_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
    try:
        df.to_excel(augmented_out, index=False)
        print(f"[OK] Wrote augmented file with Purpose → {augmented_out}")
    except Exception as e:
        print(f"[WARN] Could not write augmented Excel: {e}")

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
    cc_addrs = [addr.strip() for addr in (args.cc.split(",") if args.cc else []) if addr.strip()]
    if cc_addrs:
        print(f"[INFO] CC: {', '.join(cc_addrs)}")
    print(f"[INFO] Recipient: {to_addr}")

    results: List[Dict[str, Any]] = []
    for i, row in df.reset_index(drop=True).iterrows():
        if i in skip_indices:
            print(f"[SKIP] Row {i} already SENT per resume log.")
            results.append({
                "row_index": i, "to": to_addr, "subject": "(skipped via resume)",
                "status": "SKIPPED", "error": "", "timestamp": datetime.now().isoformat(timespec="seconds")
            })
            continue

        subject = infer_subject(row)
        body = build_body(row, purpose_text=df.loc[i, "Purpose"], to_addr=to_addr)

        # Preview in console
        print("\n" + "=" * 84)
        print(f"[PREVIEW] Row {i+1}/{total}")
        print(f"TO: {to_addr}")
        if cc_addrs:
            print(f"CC: {', '.join(cc_addrs)}")
        print(f"SUBJECT: {subject}")
        print("-" * 84)
        print(body)
        print("=" * 84 + "\n")

        status, err = "DRY-RUN", ""
        if args.send:
            try:
                send_with_retries(
                    smtp_conf, to_addr=to_addr, subject=subject, body=body,
                    cc_addrs=cc_addrs, max_retries=5, prefer_ssl=args.ssl
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
            "cc": ", ".join(cc_addrs) if cc_addrs else "",
            "subject": subject,
            "status": status,
            "error": err,
            "timestamp": datetime.now().isoformat(timespec="seconds"),
        })

    # Write log
    log_df = pd.DataFrame(results)
    out_path = args.log_out or f"{base}_send_log_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
    log_df.to_excel(out_path, index=False)
    print(f"[OK] Wrote log → {out_path}")

if __name__ == "__main__":
    main()
