#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Indiana Contracts Scraper (DB-free, Excel output) — strict custom date range

• Use --start-date and/or --end-date. Range is inclusive.
• --filter-on {start,end,either} controls which field the range applies to.
• No row padding: only matching rows are written. If your range has 10 rows, Excel has 10 rows.
• --max-rows is optional. If omitted (or 0), no cap is applied.

Examples:
  python indiana.py --start-date 2022-01-01 --end-date 2022-12-31 --filter-on either
  python indiana.py --start-date 2024-01-01 --filter-on start --max-rows 100 --out "indiana_2024.xlsx"
"""

import argparse
import logging
import time
from datetime import datetime
from typing import Dict, List, Optional, Union

import requests
import pandas as pd
from dateutil import parser as dateparser  # usually present with pandas

logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")

API_URL = "https://secure.in.gov/apps/idoa/contractsearch/api/contracts/search"
REFERER_URL = "https://secure.in.gov/apps/idoa/contractsearch/"
HEADERS = {
    "accept": "application/json, text/plain, */*",
    "content-type": "application/json",
    "origin": "https://secure.in.gov",
    "referer": REFERER_URL,
    "user-agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/115.0.0.0 Safari/537.36",
}

def _parse_date_time(s: Optional[str]) -> Dict[str, Optional[Union[str, datetime]]]:
    res: Dict[str, Optional[Union[str, datetime]]] = {"date": None, "time": None, "dt": None}
    if not s or s in {"-", "", "N/A"}:
        return res
    try:
        dt = dateparser.isoparse(s)
        res["date"] = dt.date().isoformat()
        res["time"] = dt.time().replace(microsecond=0).isoformat()
        res["dt"] = dt
    except Exception as e:
        logging.warning(f"Failed to parse date/time '{s}': {e}")
    return res

def _normalize_start(s: Optional[str]) -> Optional[datetime]:
    if not s:
        return None
    dt = dateparser.isoparse(s)
    return datetime(dt.year, dt.month, dt.day, 0, 0, 0)

def _normalize_end(s: Optional[str]) -> Optional[datetime]:
    if not s:
        return None
    dt = dateparser.isoparse(s)
    return datetime(dt.year, dt.month, dt.day, 23, 59, 59)

def _in_range(dt: Optional[datetime], start: Optional[datetime], end: Optional[datetime]) -> bool:
    if dt is None:
        return False
    if start and dt < start:
        return False
    if end and dt > end:
        return False
    return True

def scrape_indiana(
    api_start: datetime,
    page_size: int = 100,
    pause_sec: float = 0.75,
    max_rows: int = 0,  # 0 = unlimited
    range_start: Optional[datetime] = None,
    range_end: Optional[datetime] = None,
    filter_on: str = "start",  # 'start' | 'end' | 'either'
) -> List[Dict]:
    rows: List[Dict] = []
    kept = 0
    page_number = 1

    logging.info(
        f"Starting scrape (api_start={api_start.isoformat()}, page_size={page_size}, "
        f"max_rows={max_rows or 'unlimited'}, filter_on={filter_on}, "
        f"range_start={range_start}, range_end={range_end})"
    )

    while True:
        payload = {
            "startDate": api_start.strftime("%Y-%m-%dT%H:%M:%S.000"),
            "pageNumber": page_number,
            "pageSize": page_size,
        }
        # If the backend honors endDate, include it; harmless if ignored.
        if range_end:
            payload["endDate"] = range_end.strftime("%Y-%m-%dT%H:%M:%S.000")

        logging.info(f"Fetching page {page_number}...")
        try:
            resp = requests.post(API_URL, headers=HEADERS, json=payload, timeout=30)
            resp.raise_for_status()
        except requests.RequestException as e:
            logging.error(f"Request failed on page {page_number}: {e}")
            break

        try:
            data = resp.json()
        except ValueError:
            logging.error("Response was not valid JSON. Stopping.")
            break

        items = data.get("results", [])
        if not items:
            logging.info("No more results. Done.")
            break

        for c in items:
            s = _parse_date_time(c.get("startDate"))
            e = _parse_date_time(c.get("endDate"))

            # Strict client-side range check
            if filter_on == "start":
                keep = _in_range(s["dt"], range_start, range_end)
            elif filter_on == "end":
                keep = _in_range(e["dt"], range_start, range_end)
            else:  # either
                keep = _in_range(s["dt"], range_start, range_end) or _in_range(e["dt"], range_start, range_end)

            if not keep:
                continue

            rows.append({
                "contract_id": c.get("id"),
                "title": f"Indiana State Contract - {c.get('id')}" if c.get("id") else None,
                "vendor_name": c.get("vendorName"),
                "agency_name": c.get("agencyName"),
                "place_of_performance_zip": c.get("zipCode"),
                "start_date": s.get("date"),
                "start_time": s.get("time"),
                "end_date": e.get("date"),
                "end_time": e.get("time"),
                "pdf_url": c.get("pdfUrl"),
                "source_page_url": REFERER_URL,
            })
            kept += 1

            if max_rows and kept >= max_rows:
                logging.info(f"Reached max_rows={max_rows}. Stopping.")
                return rows

        if len(items) < page_size:
            logging.info("Reached the last page.")
            break

        page_number += 1
        time.sleep(pause_sec)

    logging.info(f"Finished. Total rows kept (after filter): {kept}")
    return rows

def _write_excel(df: pd.DataFrame, out_path: str) -> None:
    try:
        with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="contracts")
            ws = writer.sheets["contracts"]
            widths = {1:18, 2:34, 3:28, 4:28, 5:18, 6:14, 7:12, 8:14, 9:12, 10:60, 11:50}
            for col_idx, width in widths.items():
                col_letter = ws.cell(row=1, column=col_idx).column_letter
                ws.column_dimensions[col_letter].width = width
    except ModuleNotFoundError:
        # Basic write without formatting
        df.to_excel(out_path, index=False)

def main():
    p = argparse.ArgumentParser(description="Scrape Indiana contracts to Excel (strict custom date range).")
    p.add_argument("--start-date", help="Start date (2022-01-01 or ISO). If omitted, 2010-01-01 is used.")
    p.add_argument("--end-date",   help="End date (2022-12-31 or ISO). Inclusive end of day. Optional.")
    p.add_argument("--filter-on", choices=["start","end","either"], default="start",
                   help="Which field to filter on (default: start).")
    p.add_argument("--page-size", type=int, default=10, help="Contracts per page (default: 100).")
    p.add_argument("--max-rows", type=int, default=10, help="Cap rows; 0 means unlimited (default).")
    p.add_argument("--pause", type=float, default=0.75, help="Pause between pages (seconds).")
    p.add_argument("--out", help="Output Excel filename.")
    args = p.parse_args()

    # Normalize requested range
    range_start = _normalize_start(args.start_date) or datetime(2010, 1, 1)  # default lower bound
    range_end   = _normalize_end(args.end_date)

    # For the API paging, start from the requested lower bound to avoid 2010 floods
    api_start = range_start

    rows = scrape_indiana(
        api_start=api_start,
        page_size=args.page_size,
        pause_sec=args.pause,
        max_rows=args.max_rows,
        range_start=range_start,
        range_end=range_end,
        filter_on=args.filter_on,
    )

    df = pd.DataFrame(rows)
    if df.empty:
        logging.warning("No rows matched the selected date filters. Writing headers only.")

    out_path = args.out or ("indiana_contracts_" + datetime.now().strftime("%Y%m%d_%H%M") + ".xlsx")
    _write_excel(df, out_path)
    logging.info(f"Excel written to: {out_path}")

if __name__ == "__main__":
    main()