#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Scrape the OHCA Procurement opportunities table (single page) and save to Excel.

Source:
  https://oklahoma.gov/ohca/about/procurement.html

Usage:
  python ohca_oklahoma_procurement.py --out "C:/path/oklahoma_ohca_procurement.xlsx"

Requirements:
  pip install requests beautifulsoup4 pandas lxml
"""

import argparse
import sys
import time
from datetime import datetime
from typing import List, Dict, Tuple

import requests
from bs4 import BeautifulSoup
import pandas as pd

SOURCE_URL = "https://oklahoma.gov/ohca/about/procurement.html"

def fetch_html(url: str, timeout: int = 30) -> str:
    """GET the page with a friendly UA + basic retries."""
    headers = {
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
            "(KHTML, like Gecko) Chrome/120.0 Safari/537.36"
        )
    }
    last_err = None
    for _ in range(3):
        try:
            r = requests.get(url, headers=headers, timeout=timeout)
            r.raise_for_status()
            r.encoding = r.apparent_encoding or "utf-8"
            return r.text
        except Exception as e:
            last_err = e
            time.sleep(1.2)
    raise RuntimeError(f"Failed to fetch {url}: {last_err}")

def cell_text_and_links(td) -> Tuple[str, str]:
    """Return human text and joined absolute URLs from a table cell."""
    text = td.get_text(separator=" ", strip=True)

    links = []
    for a in td.find_all("a", href=True):
        href = a["href"].strip()
        if href.startswith("//"):
            href = "https:" + href
        elif href.startswith("/"):
            href = "https://oklahoma.gov" + href
        links.append(href)

    return text, " | ".join(links) if links else ""

def find_procurement_table(soup: BeautifulSoup):
    """Locate the table whose header includes 'Requisition Number'."""
    tables = soup.find_all("table")
    for tbl in tables:
        thead = tbl.find("thead")
        header_cells = []
        if thead:
            header_cells = [th.get_text(strip=True) for th in thead.find_all("th")]
        else:
            first_tr = tbl.find("tr")
            if first_tr:
                header_cells = [th.get_text(strip=True) for th in first_tr.find_all(["th", "td"])]
        headers_as_str = " | ".join(header_cells).lower()
        if "requisition" in headers_as_str and "number" in headers_as_str:
            return tbl
    return None

def parse_table(html: str) -> List[Dict]:
    soup = BeautifulSoup(html, "lxml")
    tbl = find_procurement_table(soup)
    if not tbl:
        raise RuntimeError("Could not find the procurement table (header with 'Requisition Number').")

    header_map: List[str] = []
    header_row = None

    thead = tbl.find("thead")
    if thead:
        header_row = thead.find("tr")
    if not header_row:
        header_row = tbl.find("tr")

    for th in header_row.find_all(["th", "td"]):
        header_map.append(th.get_text(separator=" ", strip=True).lower())

    expected_keys = {
        "requisition number": "Requisition Number",
        "procurement opportunity": "Procurement Opportunity",
        "amendments": "Amendments",
        "status": "Status",
        "closing date": "Closing Date",
        "award date": "Award Date",
        "total annual contract value": "Total Annual Contract Value",
        "awardee(s)": "Awardee(s)",
        "awardee(s).": "Awardee(s)",
        "awardee": "Awardee(s)",
    }

    body = tbl.find("tbody") or tbl
    rows_out: List[Dict] = []
    for tr in body.find_all("tr"):
        tds = tr.find_all("td")
        if not tds:
            continue

        record: Dict[str, str] = {
            "Requisition Number": "",
            "Procurement Opportunity (text)": "",
            "Procurement Opportunity (urls)": "",
            "Amendments (text)": "",
            "Amendments (urls)": "",
            "Status": "",
            "Closing Date": "",
            "Award Date": "",
            "Total Annual Contract Value": "",
            "Awardee(s)": "",
            "source_url": SOURCE_URL,
            "scraped_at": datetime.now().isoformat(timespec="seconds"),
        }

        for idx, td in enumerate(tds):
            if idx >= len(header_map):
                continue
            col_key_raw = header_map[idx]
            norm = expected_keys.get(col_key_raw)
            if not norm:
                for k, v in expected_keys.items():
                    if k in col_key_raw:
                        norm = v
                        break

            txt, urls = cell_text_and_links(td)

            if norm == "Procurement Opportunity":
                record["Procurement Opportunity (text)"] = txt
                record["Procurement Opportunity (urls)"] = urls
            elif norm == "Amendments":
                record["Amendments (text)"] = txt
                record["Amendments (urls)"] = urls
            elif norm:
                record[norm] = txt

        if any(v.strip() for v in record.values()):
            rows_out.append(record)

    return rows_out

def main():
    ap = argparse.ArgumentParser(description="OHCA Procurement table → Excel")
    ap.add_argument("--out", required=True, help="Output Excel path, e.g., C:/.../oklahoma_ohca_procurement.xlsx")
    args = ap.parse_args()

    html = fetch_html(SOURCE_URL)
    rows = parse_table(html)

    df = pd.DataFrame(rows)
    if df.empty:
        print("No rows found. The page may have changed or the table is empty.")
    df.to_excel(args.out, index=False)
    print(f"✅ Wrote {len(df)} row(s) → {args.out}")

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"ERROR: {e}", file=sys.stderr)
        sys.exit(1)
